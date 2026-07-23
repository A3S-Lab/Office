use a3s_use_core::UseResult;

use crate::spreadsheet_formula::structured_reference::StructuredReferenceErrorKind;
use crate::spreadsheet_formula::{
    parse_spreadsheet_formula, SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaExpression,
    SpreadsheetFormulaExpressionKind, SpreadsheetFormulaQualifier, SpreadsheetFormulaReference,
    SpreadsheetFormulaReferenceKind, MAX_SPREADSHEET_FORMULA_DEPTH,
    MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS, MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS,
};

use super::{
    accumulate_formula_text_bytes, calculation_error, invalid_array_shape, EvalValue,
    EvaluationContext, FormulaArray, FormulaCellKey, FormulaReferenceArea, ScalarValue,
    MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
};
use crate::SpreadsheetFormulaErrorLiteral;

impl EvaluationContext<'_> {
    pub(super) fn evaluate_structured_reference(
        &self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        reference: &str,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        match self.tables.resolve(
            qualifier,
            reference,
            current.sheet,
            current.column,
            current.row,
        ) {
            Ok(areas) => {
                ensure_reference_area_count(areas.len())?;
                Ok(EvalValue::Reference(
                    areas
                        .into_iter()
                        .map(|area| FormulaReferenceArea {
                            sheet: area.sheet,
                            start_column: area.start_column,
                            start_row: area.start_row,
                            end_column: area.end_column,
                            end_row: area.end_row,
                        })
                        .collect(),
                ))
            }
            Err(error) if matches!(error.kind, StructuredReferenceErrorKind::ExternalWorkbook) => {
                Err(calculation_error(
                    "use.office.spreadsheet_formula_external_reference_unsupported",
                    error.message,
                )
                .with_detail("reference", reference))
            }
            Err(error) => Err(calculation_error(
                "use.office.spreadsheet_formula_structured_reference_unsupported",
                error.message,
            )
            .with_detail("reference", reference)),
        }
    }

    pub(super) fn evaluate_reference(
        &mut self,
        reference: &SpreadsheetFormulaReference,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        let Some(sheets) = self.resolve_reference_sheets(reference.qualifier.as_ref(), current)?
        else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )));
        };
        let (start_column, start_row, end_column, end_row) = match reference.kind {
            SpreadsheetFormulaReferenceKind::Cell { column, row, .. } => (column, row, column, row),
            SpreadsheetFormulaReferenceKind::Column { column, .. } => {
                (column, 1, column, crate::spreadsheet_reference::MAX_ROWS)
            }
            SpreadsheetFormulaReferenceKind::Row { row, .. } => {
                (1, row, crate::spreadsheet_reference::MAX_COLUMNS, row)
            }
        };
        let areas = sheets
            .into_iter()
            .map(|sheet| FormulaReferenceArea {
                sheet,
                start_column,
                start_row,
                end_column,
                end_row,
            })
            .collect::<Vec<_>>();
        ensure_reference_area_count(areas.len())?;
        Ok(EvalValue::Reference(areas))
    }

    pub(super) fn evaluate_reference_operator(
        &mut self,
        operator: SpreadsheetFormulaBinaryOperator,
        left: &SpreadsheetFormulaExpression,
        right: &SpreadsheetFormulaExpression,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        if matches!(operator, SpreadsheetFormulaBinaryOperator::Range) {
            return self.evaluate_range(left, right, current);
        }
        let left = self.evaluate_expression(left, current)?;
        let right = self.evaluate_expression(right, current)?;
        let (EvalValue::Reference(mut left), EvalValue::Reference(right)) = (left, right) else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Value,
            )));
        };
        if matches!(operator, SpreadsheetFormulaBinaryOperator::Union) {
            ensure_reference_area_total(left.len(), right.len())?;
            left.extend(right);
            return Ok(EvalValue::Reference(left));
        }
        ensure_reference_comparisons(left.len(), right.len())?;
        let mut areas = Vec::new();
        for left in left {
            for right in &right {
                if let Some(area) = left.intersect(*right) {
                    push_reference_area(&mut areas, area)?;
                }
            }
        }
        Ok(EvalValue::Reference(areas))
    }

    pub(super) fn evaluate_name(
        &mut self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        name: &str,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        if qualifier.is_some_and(SpreadsheetFormulaQualifier::is_external) {
            return Err(calculation_error(
                "use.office.spreadsheet_formula_external_reference_unsupported",
                format!(
                    "Native calculation never opens external reference '{}'.",
                    qualifier.map_or_else(|| name.to_string(), qualifier_label)
                ),
            ));
        }
        let explicit_scope = if let Some(qualifier) = qualifier {
            if qualifier.is_three_dimensional() {
                return Err(calculation_error(
                    "use.office.spreadsheet_formula_named_reference_unsupported",
                    format!(
                        "Native calculation does not resolve a name through 3D qualifier '{}'.",
                        qualifier_label(qualifier)
                    ),
                ));
            }
            let Some(sheet) = self.sheet_position(&qualifier.worksheet) else {
                return Ok(EvalValue::Scalar(ScalarValue::Error(
                    SpreadsheetFormulaErrorLiteral::Reference,
                )));
            };
            Some(sheet)
        } else {
            None
        };
        let normalized = name.to_lowercase();
        let local_scope = explicit_scope.or(Some(current.sheet));
        let definition = local_scope
            .and_then(|sheet| {
                self.named_definitions
                    .get(&(Some(sheet), normalized.clone()))
            })
            .or_else(|| self.named_definitions.get(&(None, normalized.clone())))
            .cloned();
        let Some(definition) = definition else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Name,
            )));
        };
        if self.named_stack.len() >= MAX_SPREADSHEET_FORMULA_DEPTH {
            return Err(calculation_error(
                "use.office.spreadsheet_formula_named_reference_depth",
                format!(
                    "Native calculation resolves at most {MAX_SPREADSHEET_FORMULA_DEPTH} nested named references."
                ),
            )
            .with_detail("namedRange", definition.name));
        }
        let key = (definition.scope_sheet, normalized);
        if !self.named_stack.insert(key.clone()) {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Name,
            )));
        }
        let parsed = parse_spreadsheet_formula(&definition.formula).map_err(|error| {
            error
                .with_detail("namedRange", definition.name.clone())
                .with_detail(
                    "scope",
                    definition.scope_sheet.map_or_else(
                        || "workbook".to_string(),
                        |sheet| {
                            self.sheet_names
                                .get(sheet)
                                .cloned()
                                .unwrap_or_else(|| "unknown".to_string())
                        },
                    ),
                )
        })?;
        let named_current = FormulaCellKey {
            sheet: definition.scope_sheet.unwrap_or(current.sheet),
            ..current
        };
        let result = self.evaluate_expression(&parsed.root, named_current);
        self.named_stack.remove(&key);
        result
    }

    pub(super) fn spill_reference(&self, value: EvalValue) -> UseResult<EvalValue> {
        let EvalValue::Reference(areas) = value else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )));
        };
        let Some(area) = areas.first().copied().filter(|_| areas.len() == 1) else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )));
        };
        if area.start_column != area.end_column || area.start_row != area.end_row {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )));
        }
        let anchor = FormulaCellKey {
            sheet: area.sheet,
            column: area.start_column,
            row: area.start_row,
        };
        Ok(self.spills.get(&anchor).copied().map_or_else(
            || {
                EvalValue::Scalar(ScalarValue::Error(
                    SpreadsheetFormulaErrorLiteral::Reference,
                ))
            },
            |spill| EvalValue::Reference(vec![spill]),
        ))
    }

    pub(super) fn implicit_intersection(
        &self,
        value: EvalValue,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        match value {
            EvalValue::Reference(areas) => {
                let Some(area) = areas.first().copied() else {
                    return Ok(EvalValue::Scalar(ScalarValue::Error(
                        SpreadsheetFormulaErrorLiteral::Null,
                    )));
                };
                let key = if area.contains(current) {
                    current
                } else if area.start_column == area.end_column
                    && (area.start_row..=area.end_row).contains(&current.row)
                {
                    FormulaCellKey {
                        sheet: area.sheet,
                        column: area.start_column,
                        row: current.row,
                    }
                } else if area.start_row == area.end_row
                    && (area.start_column..=area.end_column).contains(&current.column)
                {
                    FormulaCellKey {
                        sheet: area.sheet,
                        column: current.column,
                        row: area.start_row,
                    }
                } else {
                    FormulaCellKey {
                        sheet: area.sheet,
                        column: area.start_column,
                        row: area.start_row,
                    }
                };
                Ok(EvalValue::Scalar(
                    self.values.get(&key).cloned().unwrap_or(ScalarValue::Blank),
                ))
            }
            EvalValue::Array(array) => Ok(EvalValue::Scalar(
                array
                    .rows
                    .first()
                    .and_then(|row| row.first())
                    .cloned()
                    .unwrap_or(ScalarValue::Blank),
            )),
            value => Ok(value),
        }
    }

    pub(super) fn materialize_areas(&self, areas: &[FormulaReferenceArea]) -> UseResult<EvalValue> {
        ensure_reference_area_count(areas.len())?;
        if areas.is_empty() {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Null,
            )));
        }
        if areas.len() > 1 {
            let mut rows = Vec::new();
            let mut text_bytes = 0_usize;
            for area in areas {
                let cells = area.cell_count().ok_or_else(reference_limit_error)?;
                if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS
                    || rows.len().saturating_add(cells) > MAX_SPREADSHEET_FORMULA_SPILL_CELLS
                {
                    return Err(reference_limit_error());
                }
                for row in area.start_row..=area.end_row {
                    for column in area.start_column..=area.end_column {
                        let value = self
                            .values
                            .get(&FormulaCellKey {
                                sheet: area.sheet,
                                column,
                                row,
                            })
                            .cloned()
                            .unwrap_or(ScalarValue::Blank);
                        accumulate_formula_text_bytes(&mut text_bytes, &value)?;
                        rows.push(vec![value]);
                    }
                }
            }
            return Ok(EvalValue::Array(
                FormulaArray::new(rows).ok_or_else(reference_limit_error)?,
            ));
        }
        let area = areas.first().copied().ok_or_else(invalid_array_shape)?;
        let cells = area.cell_count().ok_or_else(reference_limit_error)?;
        if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
            return Err(reference_limit_error().with_detail("cells", cells));
        }
        if cells == 1 {
            return Ok(EvalValue::Scalar(
                self.values
                    .get(&FormulaCellKey {
                        sheet: area.sheet,
                        column: area.start_column,
                        row: area.start_row,
                    })
                    .cloned()
                    .unwrap_or(ScalarValue::Blank),
            ));
        }
        let mut rows = Vec::new();
        let mut text_bytes = 0_usize;
        for row in area.start_row..=area.end_row {
            let mut values = Vec::new();
            for column in area.start_column..=area.end_column {
                let value = self
                    .values
                    .get(&FormulaCellKey {
                        sheet: area.sheet,
                        column,
                        row,
                    })
                    .cloned()
                    .unwrap_or(ScalarValue::Blank);
                accumulate_formula_text_bytes(&mut text_bytes, &value)?;
                values.push(value);
            }
            rows.push(values);
        }
        Ok(EvalValue::Array(FormulaArray { rows }))
    }

    pub(super) fn populated_reference_values(
        &self,
        areas: &[FormulaReferenceArea],
    ) -> UseResult<Vec<ScalarValue>> {
        ensure_reference_area_count(areas.len())?;
        let mut values = Vec::new();
        for area in areas {
            let mut area_values = Vec::new();
            for (key, value) in self.values.iter().filter(|(key, _)| area.contains(**key)) {
                let count = values
                    .len()
                    .checked_add(area_values.len())
                    .and_then(|count| count.checked_add(1))
                    .ok_or_else(reference_limit_error)?;
                if count > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
                    return Err(reference_limit_error().with_detail("cells", count));
                }
                area_values.push((*key, value.clone()));
            }
            area_values.sort_by_key(|(key, _)| (key.row, key.column));
            values.extend(area_values.into_iter().map(|(_, value)| value));
        }
        Ok(values)
    }

    fn evaluate_range(
        &mut self,
        left: &SpreadsheetFormulaExpression,
        right: &SpreadsheetFormulaExpression,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        let (Some(left), Some(right)) = (endpoint_reference(left), endpoint_reference(right))
        else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )));
        };
        let qualifier = match (&left.qualifier, &right.qualifier) {
            (Some(left), Some(right)) if left != right => {
                return Ok(EvalValue::Scalar(ScalarValue::Error(
                    SpreadsheetFormulaErrorLiteral::Reference,
                )));
            }
            (Some(qualifier), _) | (_, Some(qualifier)) => Some(qualifier),
            (None, None) => None,
        };
        let Some(sheets) = self.resolve_reference_sheets(qualifier, current)? else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )));
        };
        let coordinates = match (left.kind, right.kind) {
            (
                SpreadsheetFormulaReferenceKind::Cell {
                    column: left_column,
                    row: left_row,
                    ..
                },
                SpreadsheetFormulaReferenceKind::Cell {
                    column: right_column,
                    row: right_row,
                    ..
                },
            ) => (
                left_column.min(right_column),
                left_row.min(right_row),
                left_column.max(right_column),
                left_row.max(right_row),
            ),
            (
                SpreadsheetFormulaReferenceKind::Column {
                    column: left_column,
                    ..
                },
                SpreadsheetFormulaReferenceKind::Column {
                    column: right_column,
                    ..
                },
            ) => (
                left_column.min(right_column),
                1,
                left_column.max(right_column),
                crate::spreadsheet_reference::MAX_ROWS,
            ),
            (
                SpreadsheetFormulaReferenceKind::Row { row: left_row, .. },
                SpreadsheetFormulaReferenceKind::Row { row: right_row, .. },
            ) => (
                1,
                left_row.min(right_row),
                crate::spreadsheet_reference::MAX_COLUMNS,
                left_row.max(right_row),
            ),
            _ => {
                return Ok(EvalValue::Scalar(ScalarValue::Error(
                    SpreadsheetFormulaErrorLiteral::Reference,
                )));
            }
        };
        let areas = sheets
            .into_iter()
            .map(|sheet| FormulaReferenceArea {
                sheet,
                start_column: coordinates.0,
                start_row: coordinates.1,
                end_column: coordinates.2,
                end_row: coordinates.3,
            })
            .collect::<Vec<_>>();
        ensure_reference_area_count(areas.len())?;
        Ok(EvalValue::Reference(areas))
    }

    fn resolve_reference_sheets(
        &self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        current: FormulaCellKey,
    ) -> UseResult<Option<Vec<usize>>> {
        let Some(qualifier) = qualifier else {
            return Ok(Some(vec![current.sheet]));
        };
        if qualifier.is_external() {
            return Err(calculation_error(
                "use.office.spreadsheet_formula_external_reference_unsupported",
                format!(
                    "Native calculation never opens external reference '{}'.",
                    qualifier_label(qualifier)
                ),
            ));
        }
        let Some(start) = self.sheet_position(&qualifier.worksheet) else {
            return Ok(None);
        };
        let Some(end_name) = qualifier.worksheet_end.as_deref() else {
            return Ok(Some(vec![start]));
        };
        let Some(end) = self.sheet_position(end_name) else {
            return Ok(None);
        };
        let low = start.min(end);
        let high = start.max(end);
        let areas = high
            .checked_sub(low)
            .and_then(|distance| distance.checked_add(1))
            .ok_or_else(reference_area_limit_error)?;
        ensure_reference_area_count(areas)?;
        Ok(Some((low..=high).collect()))
    }

    fn sheet_position(&self, name: &str) -> Option<usize> {
        self.sheet_names
            .iter()
            .position(|sheet| sheet.eq_ignore_ascii_case(name))
    }
}

fn endpoint_reference(
    expression: &SpreadsheetFormulaExpression,
) -> Option<&SpreadsheetFormulaReference> {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Reference(reference) => Some(reference),
        SpreadsheetFormulaExpressionKind::Parenthesized(inner) => endpoint_reference(inner),
        _ => None,
    }
}

fn qualifier_label(qualifier: &SpreadsheetFormulaQualifier) -> String {
    let workbook = qualifier.workbook.as_deref().unwrap_or_default();
    let worksheet = qualifier.worksheet_end.as_ref().map_or_else(
        || qualifier.worksheet.clone(),
        |end| format!("{}:{end}", qualifier.worksheet),
    );
    format!("{workbook}{worksheet}")
}

fn reference_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_reference_limit",
        format!(
            "Native calculation materializes at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} referenced cells per value."
        ),
    )
}

fn push_reference_area(
    areas: &mut Vec<FormulaReferenceArea>,
    area: FormulaReferenceArea,
) -> UseResult<()> {
    let total = areas
        .len()
        .checked_add(1)
        .ok_or_else(reference_area_limit_error)?;
    ensure_reference_area_count(total)?;
    areas.push(area);
    Ok(())
}

fn ensure_reference_area_total(left: usize, right: usize) -> UseResult<()> {
    let total = left
        .checked_add(right)
        .ok_or_else(reference_area_limit_error)?;
    ensure_reference_area_count(total)
}

fn ensure_reference_area_count(areas: usize) -> UseResult<()> {
    if areas > MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS {
        return Err(reference_area_limit_error().with_detail("areas", areas));
    }
    Ok(())
}

fn ensure_reference_comparisons(left: usize, right: usize) -> UseResult<()> {
    let visits = left
        .checked_mul(right)
        .ok_or_else(reference_visit_limit_error)?;
    if visits > MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS {
        return Err(reference_visit_limit_error().with_detail("visits", visits));
    }
    Ok(())
}

fn reference_area_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_reference_area_limit",
        format!(
            "Native calculation retains at most {MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS} reference areas per value."
        ),
    )
}

fn reference_visit_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_reference_visit_limit",
        format!(
            "Native calculation reference operators visit at most {MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS} area pairs."
        ),
    )
}
