use a3s_office_formula_parser::{
    SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaExpression,
    SpreadsheetFormulaExpressionKind, SpreadsheetFormulaLiteral, SpreadsheetFormulaPostfixOperator,
    SpreadsheetFormulaReference, SpreadsheetFormulaReferenceKind, SpreadsheetFormulaUnaryOperator,
};

use super::value::{apply_binary, finite_number, scalar_number, unsupported, EvaluatedValue};
use super::{
    CellKey, EvaluationFailure, SpreadsheetEvaluator, SpreadsheetValue, MAX_SPREADSHEET_CELLS,
};

impl SpreadsheetEvaluator<'_> {
    pub(super) fn evaluate_expression(
        &mut self,
        expression: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<SpreadsheetValue, EvaluationFailure> {
        let value = self.evaluate_value(expression, current)?;
        Ok(value.into_scalar())
    }

    pub(super) fn evaluate_value(
        &mut self,
        expression: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Literal(literal) => {
                Ok(EvaluatedValue::Scalar(literal_value(literal)?))
            }
            SpreadsheetFormulaExpressionKind::Reference(reference) => {
                let key = self.reference_key(reference, current.sheet)?;
                Ok(EvaluatedValue::Scalar(self.evaluate_dependency(&key)?))
            }
            SpreadsheetFormulaExpressionKind::Name { name, .. } => Err(unsupported(format!(
                "Named formula or range '{name}' is outside the first browser calculation slice."
            ))),
            SpreadsheetFormulaExpressionKind::StructuredReference { reference, .. } => {
                Err(unsupported(format!(
                    "Structured reference '{reference}' is not yet calculated in the browser kernel."
                )))
            }
            SpreadsheetFormulaExpressionKind::Unary { operator, operand } => {
                let value = self
                    .evaluate_value(operand, current)?
                    .into_scalar();
                let number = match scalar_number(value) {
                    Ok(number) => number,
                    Err(error) => {
                        return Ok(EvaluatedValue::Scalar(SpreadsheetValue::error(error)));
                    }
                };
                let number = match operator {
                    SpreadsheetFormulaUnaryOperator::Positive
                    | SpreadsheetFormulaUnaryOperator::ImplicitIntersection => number,
                    SpreadsheetFormulaUnaryOperator::Negative => -number,
                };
                Ok(EvaluatedValue::Scalar(finite_number(number)))
            }
            SpreadsheetFormulaExpressionKind::Postfix { operator, operand } => match operator {
                SpreadsheetFormulaPostfixOperator::Percent => {
                    let value = self
                        .evaluate_value(operand, current)?
                        .into_scalar();
                    let number = match scalar_number(value) {
                        Ok(number) => number,
                        Err(error) => {
                            return Ok(EvaluatedValue::Scalar(SpreadsheetValue::error(error)));
                        }
                    };
                    Ok(EvaluatedValue::Scalar(finite_number(number / 100.0)))
                }
                SpreadsheetFormulaPostfixOperator::Spill => Err(unsupported(
                    "Dynamic-array spill references are not yet calculated in the browser kernel.",
                )),
            },
            SpreadsheetFormulaExpressionKind::Binary {
                operator:
                    operator @ (SpreadsheetFormulaBinaryOperator::Range
                    | SpreadsheetFormulaBinaryOperator::Intersection
                    | SpreadsheetFormulaBinaryOperator::Union),
                left,
                right,
            } => self.evaluate_reference_operator(*operator, left, right, current),
            SpreadsheetFormulaExpressionKind::Binary {
                operator,
                left,
                right,
            } => {
                let left = self.evaluate_value(left, current)?.into_scalar();
                let right = self.evaluate_value(right, current)?.into_scalar();
                Ok(EvaluatedValue::Scalar(apply_binary(*operator, left, right)))
            }
            SpreadsheetFormulaExpressionKind::FunctionCall {
                qualifier,
                name,
                arguments,
            } => {
                if qualifier.is_some() {
                    return Err(unsupported(format!(
                        "Qualified function '{name}' is not supported."
                    )));
                }
                self.evaluate_function(name, arguments, current)
            }
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.evaluate_value(inner, current)
            }
            SpreadsheetFormulaExpressionKind::Array { .. } => Err(unsupported(
                "Array constants are not yet calculated in the browser kernel.",
            )),
        }
    }

    fn evaluate_reference_operator(
        &mut self,
        operator: SpreadsheetFormulaBinaryOperator,
        left: &SpreadsheetFormulaExpression,
        right: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        if operator != SpreadsheetFormulaBinaryOperator::Range {
            return Err(unsupported(
                "Reference unions and intersections are not yet calculated in the browser kernel.",
            ));
        }
        let left = reference_expression(left)
            .ok_or_else(|| unsupported("A range endpoint must be a concrete A1 cell reference."))?;
        let left_key = self.reference_key(left, current.sheet)?;
        let right = reference_expression(right)
            .ok_or_else(|| unsupported("A range endpoint must be a concrete A1 cell reference."))?;
        let right_key = self.reference_key(right, left_key.sheet)?;
        if left_key.sheet != right_key.sheet {
            return Err(unsupported(
                "A browser calculation range cannot span multiple worksheets.",
            ));
        }
        let start_row = left_key.row.min(right_key.row);
        let end_row = left_key.row.max(right_key.row);
        let start_column = left_key.column.min(right_key.column);
        let end_column = left_key.column.max(right_key.column);
        let rows = usize::try_from(end_row - start_row + 1)
            .map_err(|_| unsupported("Spreadsheet range height is outside supported limits."))?;
        let columns = usize::try_from(end_column - start_column + 1)
            .map_err(|_| unsupported("Spreadsheet range width is outside supported limits."))?;
        let cells = rows
            .checked_mul(columns)
            .ok_or_else(|| unsupported("Spreadsheet range size is outside supported limits."))?;
        if cells > MAX_SPREADSHEET_CELLS {
            return Err(unsupported(format!(
                "One browser formula range may materialize at most {MAX_SPREADSHEET_CELLS} cells."
            )));
        }
        self.reserve_materialized_range_cells(cells)?;
        let mut values = Vec::with_capacity(cells);
        for row in start_row..=end_row {
            for column in start_column..=end_column {
                values.push(self.evaluate_dependency(&CellKey {
                    sheet: left_key.sheet,
                    row,
                    column,
                })?);
            }
        }
        Ok(EvaluatedValue::Range(values))
    }

    fn reserve_materialized_range_cells(&mut self, cells: usize) -> Result<(), EvaluationFailure> {
        let Some(current) = self.materialized_range_cells.last_mut() else {
            return Err(unsupported(
                "A Spreadsheet range was evaluated outside a formula calculation.",
            ));
        };
        let Some(total) = current.checked_add(cells) else {
            return Err(unsupported(
                "Spreadsheet range materialization exceeds supported limits.",
            ));
        };
        if total > MAX_SPREADSHEET_CELLS {
            return Err(unsupported(format!(
                "One browser formula may materialize at most {MAX_SPREADSHEET_CELLS} cells across all ranges."
            )));
        }
        *current = total;
        Ok(())
    }

    fn reference_key(
        &self,
        reference: &SpreadsheetFormulaReference,
        default_sheet: usize,
    ) -> Result<CellKey, EvaluationFailure> {
        let sheet = match &reference.qualifier {
            None => default_sheet,
            Some(qualifier) if qualifier.is_external() => {
                return Err(unsupported(
                    "External-workbook references are retained but not refreshed by the browser kernel.",
                ));
            }
            Some(qualifier) if qualifier.is_three_dimensional() => {
                return Err(unsupported(
                    "Three-dimensional worksheet references are not yet calculated in the browser kernel.",
                ));
            }
            Some(qualifier) => *self
                .session
                .sheet_names
                .get(&qualifier.worksheet.to_lowercase())
                .ok_or_else(|| {
                    unsupported(format!(
                        "Worksheet '{}' does not exist in this calculation request.",
                        qualifier.worksheet
                    ))
                })?,
        };
        let SpreadsheetFormulaReferenceKind::Cell { column, row, .. } = reference.kind else {
            return Err(unsupported(
                "Whole-row and whole-column references are not yet materialized by the browser kernel.",
            ));
        };
        Ok(CellKey {
            sheet,
            row: row.saturating_sub(1),
            column: column.saturating_sub(1),
        })
    }
}

fn reference_expression(
    expression: &SpreadsheetFormulaExpression,
) -> Option<&SpreadsheetFormulaReference> {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Reference(reference) => Some(reference),
        SpreadsheetFormulaExpressionKind::Parenthesized(inner) => reference_expression(inner),
        _ => None,
    }
}

fn literal_value(
    literal: &SpreadsheetFormulaLiteral,
) -> Result<SpreadsheetValue, EvaluationFailure> {
    Ok(match literal {
        SpreadsheetFormulaLiteral::Number(value) => {
            let number = value.parse::<f64>().map_err(|_| {
                unsupported(format!("Formula number literal '{value}' is not finite."))
            })?;
            finite_number(number)
        }
        SpreadsheetFormulaLiteral::Text(value) => SpreadsheetValue::Text {
            value: value.clone(),
        },
        SpreadsheetFormulaLiteral::Boolean(value) => SpreadsheetValue::Boolean { value: *value },
        SpreadsheetFormulaLiteral::Error(error) => SpreadsheetValue::error(*error),
    })
}
