use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use crate::spreadsheet_reference::{MAX_COLUMNS, MAX_ROWS};

use super::{
    graph_error, SpreadsheetFormulaUnresolvedReference, SpreadsheetFormulaUnresolvedReferenceKind,
    MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS,
};
use crate::spreadsheet_formula::structured_reference::{
    FormulaTableCatalog, StructuredReferenceErrorKind,
};
use crate::spreadsheet_formula::{
    parse_spreadsheet_formula, SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaExpression,
    SpreadsheetFormulaExpressionKind, SpreadsheetFormulaPostfixOperator,
    SpreadsheetFormulaQualifier, SpreadsheetFormulaReference, SpreadsheetFormulaReferenceKind,
    SpreadsheetFormulaUnaryOperator, MAX_SPREADSHEET_FORMULA_DEPTH,
    MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS,
};

#[derive(Debug, Clone)]
pub(super) struct FormulaNamedDefinition {
    pub(super) name: String,
    pub(super) formula: String,
    pub(super) scope_sheet: Option<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct FormulaReferenceArea {
    pub(super) sheet_index: usize,
    pub(super) start_column: u32,
    pub(super) start_row: u32,
    pub(super) end_column: u32,
    pub(super) end_row: u32,
}

impl FormulaReferenceArea {
    fn intersect(self, other: Self) -> Option<Self> {
        if self.sheet_index != other.sheet_index {
            return None;
        }
        let start_column = self.start_column.max(other.start_column);
        let start_row = self.start_row.max(other.start_row);
        let end_column = self.end_column.min(other.end_column);
        let end_row = self.end_row.min(other.end_row);
        (start_column <= end_column && start_row <= end_row).then_some(Self {
            sheet_index: self.sheet_index,
            start_column,
            start_row,
            end_column,
            end_row,
        })
    }
}

pub(super) struct FormulaReferenceCollection {
    pub(super) areas: Vec<FormulaReferenceArea>,
    pub(super) unresolved_references: Vec<SpreadsheetFormulaUnresolvedReference>,
}

pub(super) fn collect_references(
    expression: &SpreadsheetFormulaExpression,
    current_sheet: usize,
    current_column: u32,
    current_row: u32,
    sheet_names: &[String],
    named_definitions: &BTreeMap<(Option<usize>, String), FormulaNamedDefinition>,
    tables: &FormulaTableCatalog,
) -> UseResult<FormulaReferenceCollection> {
    let mut collector = ReferenceCollector {
        sheet_names,
        named_definitions,
        tables,
        current_column,
        current_row,
        areas: Vec::new(),
        unresolved: Vec::new(),
        visited_names: BTreeSet::new(),
    };
    collector.collect_expression(expression, current_sheet)?;
    collector.areas.sort();
    collector.areas.dedup();
    Ok(FormulaReferenceCollection {
        areas: collector.areas,
        unresolved_references: collector.unresolved,
    })
}

struct ReferenceCollector<'a> {
    sheet_names: &'a [String],
    named_definitions: &'a BTreeMap<(Option<usize>, String), FormulaNamedDefinition>,
    tables: &'a FormulaTableCatalog,
    current_column: u32,
    current_row: u32,
    areas: Vec<FormulaReferenceArea>,
    unresolved: Vec<SpreadsheetFormulaUnresolvedReference>,
    visited_names: BTreeSet<(Option<usize>, String)>,
}

impl ReferenceCollector<'_> {
    fn collect_expression(
        &mut self,
        expression: &SpreadsheetFormulaExpression,
        current_sheet: usize,
    ) -> UseResult<()> {
        if let Some(areas) = self.try_reference_areas(expression, current_sheet)? {
            self.extend_areas(areas)?;
            return Ok(());
        }
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Name { qualifier, name } => {
                self.collect_named_reference(qualifier.as_ref(), name, current_sheet)
            }
            SpreadsheetFormulaExpressionKind::StructuredReference {
                qualifier,
                reference,
            } => {
                match self.tables.resolve(
                    qualifier.as_ref(),
                    reference,
                    current_sheet,
                    self.current_column,
                    self.current_row,
                ) {
                    Ok(areas) => self.extend_areas(
                        areas
                            .into_iter()
                            .map(|area| FormulaReferenceArea {
                                sheet_index: area.sheet,
                                start_column: area.start_column,
                                start_row: area.start_row,
                                end_column: area.end_column,
                                end_row: area.end_row,
                            })
                            .collect(),
                    )?,
                    Err(error) => {
                        let kind =
                            if matches!(error.kind, StructuredReferenceErrorKind::ExternalWorkbook)
                            {
                                SpreadsheetFormulaUnresolvedReferenceKind::ExternalWorkbook
                            } else {
                                SpreadsheetFormulaUnresolvedReferenceKind::StructuredReference
                            };
                        self.push_unresolved(kind, reference);
                    }
                }
                Ok(())
            }
            SpreadsheetFormulaExpressionKind::Unary { operand, .. }
            | SpreadsheetFormulaExpressionKind::Postfix { operand, .. }
            | SpreadsheetFormulaExpressionKind::Parenthesized(operand) => {
                self.collect_expression(operand, current_sheet)
            }
            SpreadsheetFormulaExpressionKind::Binary {
                operator,
                left,
                right,
            } => {
                if matches!(operator, SpreadsheetFormulaBinaryOperator::Range) {
                    self.push_unresolved(
                        SpreadsheetFormulaUnresolvedReferenceKind::UnsupportedReference,
                        "range with non-A1 endpoints",
                    );
                }
                self.collect_expression(left, current_sheet)?;
                self.collect_expression(right, current_sheet)
            }
            SpreadsheetFormulaExpressionKind::FunctionCall {
                name, arguments, ..
            } => {
                if matches!(
                    name.to_ascii_uppercase().as_str(),
                    "INDIRECT" | "OFFSET" | "INDEX"
                ) {
                    self.push_unresolved(
                        SpreadsheetFormulaUnresolvedReferenceKind::DynamicReference,
                        name,
                    );
                }
                for argument in arguments.iter().flatten() {
                    self.collect_expression(argument, current_sheet)?;
                }
                Ok(())
            }
            SpreadsheetFormulaExpressionKind::Array { rows } => {
                for value in rows.iter().flatten() {
                    self.collect_expression(value, current_sheet)?;
                }
                Ok(())
            }
            SpreadsheetFormulaExpressionKind::Literal(_)
            | SpreadsheetFormulaExpressionKind::Reference(_) => Ok(()),
        }
    }

    fn try_reference_areas(
        &mut self,
        expression: &SpreadsheetFormulaExpression,
        current_sheet: usize,
    ) -> UseResult<Option<Vec<FormulaReferenceArea>>> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Reference(reference) => {
                Ok(Some(self.single_reference_areas(reference, current_sheet)?))
            }
            SpreadsheetFormulaExpressionKind::Binary {
                operator:
                    SpreadsheetFormulaBinaryOperator::Range
                    | SpreadsheetFormulaBinaryOperator::Union
                    | SpreadsheetFormulaBinaryOperator::Intersection,
                left,
                right,
            } => {
                let SpreadsheetFormulaExpressionKind::Binary { operator, .. } = &expression.kind
                else {
                    return Ok(None);
                };
                if matches!(operator, SpreadsheetFormulaBinaryOperator::Range) {
                    return self.range_areas(left, right, current_sheet);
                }
                let Some(mut left) = self.try_reference_areas(left, current_sheet)? else {
                    return Ok(None);
                };
                let Some(right) = self.try_reference_areas(right, current_sheet)? else {
                    return Ok(None);
                };
                if matches!(operator, SpreadsheetFormulaBinaryOperator::Union) {
                    ensure_reference_area_total(left.len(), right.len())?;
                    left.extend(right);
                    Ok(Some(left))
                } else {
                    ensure_reference_comparisons(left.len(), right.len())?;
                    let mut areas = Vec::new();
                    for left in left {
                        for right in &right {
                            if let Some(area) = left.intersect(*right) {
                                push_reference_area(&mut areas, area)?;
                            }
                        }
                    }
                    Ok(Some(areas))
                }
            }
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.try_reference_areas(inner, current_sheet)
            }
            SpreadsheetFormulaExpressionKind::Unary {
                operator: SpreadsheetFormulaUnaryOperator::ImplicitIntersection,
                operand,
            }
            | SpreadsheetFormulaExpressionKind::Postfix {
                operator: SpreadsheetFormulaPostfixOperator::Spill,
                operand,
            } => self.try_reference_areas(operand, current_sheet),
            _ => Ok(None),
        }
    }

    fn single_reference_areas(
        &mut self,
        reference: &SpreadsheetFormulaReference,
        current_sheet: usize,
    ) -> UseResult<Vec<FormulaReferenceArea>> {
        let (start_column, start_row, end_column, end_row) = match reference.kind {
            SpreadsheetFormulaReferenceKind::Cell { column, row, .. } => (column, row, column, row),
            SpreadsheetFormulaReferenceKind::Column { column, .. } => (column, 1, column, MAX_ROWS),
            SpreadsheetFormulaReferenceKind::Row { row, .. } => (1, row, MAX_COLUMNS, row),
        };
        Ok(self
            .resolve_sheets(reference.qualifier.as_ref(), current_sheet)?
            .into_iter()
            .map(|sheet_index| FormulaReferenceArea {
                sheet_index,
                start_column,
                start_row,
                end_column,
                end_row,
            })
            .collect())
    }

    fn range_areas(
        &mut self,
        left: &SpreadsheetFormulaExpression,
        right: &SpreadsheetFormulaExpression,
        current_sheet: usize,
    ) -> UseResult<Option<Vec<FormulaReferenceArea>>> {
        let Some(left) = endpoint_reference(left) else {
            return Ok(None);
        };
        let Some(right) = endpoint_reference(right) else {
            return Ok(None);
        };
        let qualifier = match (&left.qualifier, &right.qualifier) {
            (Some(left), Some(right)) if left != right => {
                self.push_unresolved(
                    SpreadsheetFormulaUnresolvedReferenceKind::UnsupportedReference,
                    "range with different worksheet qualifiers",
                );
                return Ok(Some(Vec::new()));
            }
            (Some(qualifier), _) | (_, Some(qualifier)) => Some(qualifier),
            (None, None) => None,
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
                MAX_ROWS,
            ),
            (
                SpreadsheetFormulaReferenceKind::Row { row: left_row, .. },
                SpreadsheetFormulaReferenceKind::Row { row: right_row, .. },
            ) => (
                1,
                left_row.min(right_row),
                MAX_COLUMNS,
                left_row.max(right_row),
            ),
            _ => return Ok(None),
        };
        Ok(Some(
            self.resolve_sheets(qualifier, current_sheet)?
                .into_iter()
                .map(|sheet_index| FormulaReferenceArea {
                    sheet_index,
                    start_column: coordinates.0,
                    start_row: coordinates.1,
                    end_column: coordinates.2,
                    end_row: coordinates.3,
                })
                .collect(),
        ))
    }

    fn resolve_sheets(
        &mut self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        current_sheet: usize,
    ) -> UseResult<Vec<usize>> {
        let Some(qualifier) = qualifier else {
            return Ok(vec![current_sheet]);
        };
        if qualifier.is_external() {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::ExternalWorkbook,
                qualifier_label(qualifier),
            );
            return Ok(Vec::new());
        }
        let Some(start) = self.sheet_position(&qualifier.worksheet) else {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::MissingWorksheet,
                &qualifier.worksheet,
            );
            return Ok(Vec::new());
        };
        let Some(end_name) = qualifier.worksheet_end.as_deref() else {
            return Ok(vec![start]);
        };
        let Some(end) = self.sheet_position(end_name) else {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::MissingWorksheet,
                end_name,
            );
            return Ok(Vec::new());
        };
        let low = start.min(end);
        let high = start.max(end);
        let areas = high
            .checked_sub(low)
            .and_then(|distance| distance.checked_add(1))
            .ok_or_else(reference_area_limit)?;
        if areas > MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS {
            return Err(reference_area_limit().with_detail("areas", areas));
        }
        Ok((low..=high).collect())
    }

    fn collect_named_reference(
        &mut self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        name: &str,
        current_sheet: usize,
    ) -> UseResult<()> {
        if qualifier.is_some_and(SpreadsheetFormulaQualifier::is_external) {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::ExternalWorkbook,
                qualifier.map_or(name.to_string(), qualifier_label),
            );
            return Ok(());
        }
        let explicit_scope = qualifier.and_then(|qualifier| {
            if qualifier.is_three_dimensional() {
                None
            } else {
                self.sheet_position(&qualifier.worksheet)
            }
        });
        if qualifier.is_some() && explicit_scope.is_none() {
            if let Some(qualifier) = qualifier {
                let kind = if self.sheet_position(&qualifier.worksheet).is_none() {
                    SpreadsheetFormulaUnresolvedReferenceKind::MissingWorksheet
                } else {
                    SpreadsheetFormulaUnresolvedReferenceKind::UnsupportedReference
                };
                self.push_unresolved(kind, qualifier_label(qualifier));
            }
            return Ok(());
        }
        let normalized = name.to_lowercase();
        let local_scope = explicit_scope.or_else(|| qualifier.is_none().then_some(current_sheet));
        let definition = local_scope
            .and_then(|scope| {
                self.named_definitions
                    .get(&(Some(scope), normalized.clone()))
            })
            .or_else(|| self.named_definitions.get(&(None, normalized.clone())))
            .cloned();
        let Some(definition) = definition else {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::UndefinedName,
                name,
            );
            return Ok(());
        };
        if self.visited_names.len() >= MAX_SPREADSHEET_FORMULA_DEPTH {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::NamedRangeDepth,
                &definition.name,
            );
            return Ok(());
        }
        let key = (definition.scope_sheet, normalized);
        if !self.visited_names.insert(key.clone()) {
            self.push_unresolved(
                SpreadsheetFormulaUnresolvedReferenceKind::NamedRangeCycle,
                &definition.name,
            );
            return Ok(());
        }
        let parsed = parse_spreadsheet_formula(&definition.formula).map_err(|error| {
            error
                .with_detail("namedRange", definition.name.clone())
                .with_detail(
                    "scope",
                    definition.scope_sheet.map_or_else(
                        || "workbook".to_string(),
                        |scope| {
                            self.sheet_names
                                .get(scope)
                                .cloned()
                                .unwrap_or_else(|| "unknown".to_string())
                        },
                    ),
                )
        })?;
        let definition_sheet = definition.scope_sheet.unwrap_or(current_sheet);
        let result = self.collect_expression(&parsed.root, definition_sheet);
        self.visited_names.remove(&key);
        result
    }

    fn sheet_position(&self, name: &str) -> Option<usize> {
        self.sheet_names
            .iter()
            .position(|sheet| sheet.eq_ignore_ascii_case(name))
    }

    fn push_unresolved(
        &mut self,
        kind: SpreadsheetFormulaUnresolvedReferenceKind,
        reference: impl AsRef<str>,
    ) {
        self.unresolved.push(SpreadsheetFormulaUnresolvedReference {
            kind,
            reference: reference.as_ref().to_string(),
        });
    }

    fn extend_areas(&mut self, areas: Vec<FormulaReferenceArea>) -> UseResult<()> {
        ensure_reference_area_total(self.areas.len(), areas.len())?;
        self.areas.extend(areas);
        Ok(())
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
    let worksheets = qualifier.worksheet_end.as_ref().map_or_else(
        || qualifier.worksheet.clone(),
        |end| format!("{}:{end}", qualifier.worksheet),
    );
    format!("{workbook}{worksheets}")
}

fn push_reference_area(
    areas: &mut Vec<FormulaReferenceArea>,
    area: FormulaReferenceArea,
) -> UseResult<()> {
    let total = areas
        .len()
        .checked_add(1)
        .ok_or_else(reference_area_limit)?;
    if total > MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS {
        return Err(reference_area_limit().with_detail("areas", total));
    }
    areas.push(area);
    Ok(())
}

fn ensure_reference_area_total(left: usize, right: usize) -> UseResult<()> {
    let total = left.checked_add(right).ok_or_else(reference_area_limit)?;
    if total > MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS {
        return Err(reference_area_limit().with_detail("areas", total));
    }
    Ok(())
}

fn ensure_reference_comparisons(left: usize, right: usize) -> UseResult<()> {
    let visits = left.checked_mul(right).ok_or_else(reference_visit_limit)?;
    if visits > MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS {
        return Err(reference_visit_limit().with_detail("visits", visits));
    }
    Ok(())
}

fn reference_area_limit() -> a3s_use_core::UseError {
    graph_error(
        "use.office.spreadsheet_formula_reference_area_limit",
        format!(
            "Spreadsheet formulas retain at most {MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS} static reference areas."
        ),
    )
}

fn reference_visit_limit() -> a3s_use_core::UseError {
    graph_error(
        "use.office.spreadsheet_formula_reference_visit_limit",
        format!(
            "Spreadsheet formula reference operators visit at most {MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS} area pairs."
        ),
    )
}
