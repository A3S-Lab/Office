mod context;
mod function;
mod operators;
mod reference;
mod spill;

use context::{
    build_context, public_cell_key, EvalValue, EvaluationContext, FormulaArray, FormulaCellKey,
    FormulaReferenceArea, ScalarValue,
};
use operators::{
    accumulate_formula_text_bytes, broadcast_dimension, calculation_error,
    calculation_text_limit_error, checked_array_cells, ensure_formula_text_limit,
    finite_or_number_error, format_number, into_array, invalid_array_shape, public_scalar,
    scalar_binary, spill_limit_error,
};

use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use crate::semantic::NativeOfficeDocument;
use crate::DocumentKind;

use super::{
    SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaCalculatedCell,
    SpreadsheetFormulaCalculation, SpreadsheetFormulaCell, SpreadsheetFormulaErrorLiteral,
    SpreadsheetFormulaExpression, SpreadsheetFormulaExpressionKind,
    SpreadsheetFormulaFunctionRegistry, SpreadsheetFormulaLiteral,
    SpreadsheetFormulaPostfixOperator, SpreadsheetFormulaUnaryOperator,
};

/// Maximum cells in one dynamic-array result and cumulative spill children in
/// one native calculation pass.
pub const MAX_SPREADSHEET_FORMULA_SPILL_CELLS: usize = 100_000;

/// Maximum UTF-8 bytes produced by one native formula text result.
pub const MAX_SPREADSHEET_FORMULA_TEXT_BYTES: usize = 1024 * 1024;

/// Maximum cumulative UTF-8 text-result bytes in one native calculation pass.
pub const MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES: usize = 8 * 1024 * 1024;

impl NativeOfficeDocument {
    /// Calculates supported formulas in memory without changing package bytes.
    pub fn calculate_spreadsheet_formulas(&self) -> UseResult<SpreadsheetFormulaCalculation> {
        self.calculate_spreadsheet_formulas_with_registry(
            &SpreadsheetFormulaFunctionRegistry::default(),
        )
    }

    /// Calculates supported formulas using an explicit typed function registry.
    pub fn calculate_spreadsheet_formulas_with_registry(
        &self,
        registry: &SpreadsheetFormulaFunctionRegistry,
    ) -> UseResult<SpreadsheetFormulaCalculation> {
        calculate(self, registry)
    }
}

fn calculate(
    document: &NativeOfficeDocument,
    registry: &SpreadsheetFormulaFunctionRegistry,
) -> UseResult<SpreadsheetFormulaCalculation> {
    if document.kind() != DocumentKind::Spreadsheet {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_calculation_type_unsupported",
            "Native formula calculation is available only for Spreadsheet documents.",
        ));
    }
    let graph = document.formula_dependency_graph()?;
    if !graph.cycles.is_empty() {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_cycle",
            "Spreadsheet formula dependency graph contains a circular reference.",
        )
        .with_detail(
            "cycles",
            serde_json::to_value(
                graph
                    .cycles
                    .iter()
                    .map(|cycle| {
                        cycle
                            .iter()
                            .map(SpreadsheetFormulaCell::path)
                            .collect::<Vec<_>>()
                    })
                    .collect::<Vec<_>>(),
            )
            .unwrap_or(serde_json::Value::Null),
        ));
    }
    let (mut context, records) = build_context(document, registry, &graph.nodes)?;
    let record_indexes = records
        .iter()
        .enumerate()
        .map(|(index, record)| (record.key, index))
        .collect::<BTreeMap<_, _>>();
    let mut cells = Vec::with_capacity(records.len());
    let mut spill_cell_count = 0_usize;
    let mut text_result_bytes = 0_usize;
    for cell in &graph.calculation_order {
        let key = public_cell_key(cell, &context.sheet_names)?;
        let record = record_indexes
            .get(&key)
            .and_then(|index| records.get(*index))
            .ok_or_else(|| {
                calculation_error(
                    "use.office.spreadsheet_formula_calculation_invalid",
                    format!(
                        "Calculation order references missing formula cell '{}'.",
                        cell.path()
                    ),
                )
            })?;
        let evaluated = context
            .evaluate_expression(&record.formula.root, key)
            .map_err(|error| error.with_detail("cell", cell.path()))?;
        let (value, spill_range, spill_cells, text_bytes) =
            context.finalize_formula_result(key, evaluated, spill_cell_count, text_result_bytes)?;
        spill_cell_count = spill_cell_count
            .checked_add(spill_cells)
            .ok_or_else(spill_limit_error)?;
        text_result_bytes = text_result_bytes
            .checked_add(text_bytes)
            .ok_or_else(calculation_text_limit_error)?;
        cells.push(SpreadsheetFormulaCalculatedCell {
            cell: cell.clone(),
            value,
            spill_range,
        });
    }
    Ok(SpreadsheetFormulaCalculation {
        formula_count: records.len(),
        spill_cell_count,
        calculation_order: graph.calculation_order,
        cells,
    })
}

impl EvaluationContext<'_> {
    pub(super) fn evaluate_expression(
        &mut self,
        expression: &SpreadsheetFormulaExpression,
        current: FormulaCellKey,
    ) -> UseResult<EvalValue> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Literal(literal) => {
                Ok(EvalValue::Scalar(match literal {
                    SpreadsheetFormulaLiteral::Number(value) => value
                        .parse::<f64>()
                        .ok()
                        .filter(|value| value.is_finite())
                        .map_or(
                            ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Number),
                            ScalarValue::Number,
                        ),
                    SpreadsheetFormulaLiteral::Text(value) => ScalarValue::Text(value.clone()),
                    SpreadsheetFormulaLiteral::Boolean(value) => ScalarValue::Boolean(*value),
                    SpreadsheetFormulaLiteral::Error(error) => ScalarValue::Error(*error),
                }))
            }
            SpreadsheetFormulaExpressionKind::Reference(reference) => {
                self.evaluate_reference(reference, current)
            }
            SpreadsheetFormulaExpressionKind::Name { qualifier, name } => {
                self.evaluate_name(qualifier.as_ref(), name, current)
            }
            SpreadsheetFormulaExpressionKind::StructuredReference {
                qualifier,
                reference,
            } => self.evaluate_structured_reference(qualifier.as_ref(), reference, current),
            SpreadsheetFormulaExpressionKind::Unary { operator, operand } => match operator {
                SpreadsheetFormulaUnaryOperator::ImplicitIntersection => {
                    let value = self.evaluate_expression(operand, current)?;
                    self.implicit_intersection(value, current)
                }
                SpreadsheetFormulaUnaryOperator::Positive => {
                    let value = self.evaluate_expression(operand, current)?;
                    self.map_numeric(value, finite_or_number_error)
                }
                SpreadsheetFormulaUnaryOperator::Negative => {
                    let value = self.evaluate_expression(operand, current)?;
                    self.map_numeric(value, |number| finite_or_number_error(-number))
                }
            },
            SpreadsheetFormulaExpressionKind::Postfix { operator, operand } => match operator {
                SpreadsheetFormulaPostfixOperator::Percent => {
                    let value = self.evaluate_expression(operand, current)?;
                    self.map_numeric(value, |number| finite_or_number_error(number / 100.0))
                }
                SpreadsheetFormulaPostfixOperator::Spill => {
                    let value = self.evaluate_expression(operand, current)?;
                    self.spill_reference(value)
                }
            },
            SpreadsheetFormulaExpressionKind::Binary {
                operator,
                left,
                right,
            } if matches!(
                operator,
                SpreadsheetFormulaBinaryOperator::Range
                    | SpreadsheetFormulaBinaryOperator::Intersection
                    | SpreadsheetFormulaBinaryOperator::Union
            ) =>
            {
                self.evaluate_reference_operator(*operator, left, right, current)
            }
            SpreadsheetFormulaExpressionKind::Binary {
                operator,
                left,
                right,
            } => {
                let left = self.evaluate_expression(left, current)?;
                let right = self.evaluate_expression(right, current)?;
                self.apply_binary(*operator, left, right)
            }
            SpreadsheetFormulaExpressionKind::FunctionCall {
                qualifier,
                name,
                arguments,
            } => {
                if qualifier.is_some() {
                    return Err(calculation_error(
                        "use.office.spreadsheet_formula_function_unsupported",
                        format!("Native calculation does not execute qualified function '{name}'."),
                    )
                    .with_detail("function", name.clone()));
                }
                function::evaluate_function(self, name, arguments, current)
            }
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.evaluate_expression(inner, current)
            }
            SpreadsheetFormulaExpressionKind::Array { rows } => {
                let mut values = Vec::with_capacity(rows.len());
                for row in rows {
                    let mut output = Vec::with_capacity(row.len());
                    for value in row {
                        let evaluated = self.evaluate_expression(value, current)?;
                        output.push(self.require_scalar(evaluated)?);
                    }
                    values.push(output);
                }
                Ok(EvalValue::Array(FormulaArray::new(values).ok_or_else(
                    || {
                        calculation_error(
                            "use.office.spreadsheet_formula_array_invalid",
                            "Formula array constant is not rectangular.",
                        )
                    },
                )?))
            }
        }
    }

    fn apply_binary(
        &self,
        operator: SpreadsheetFormulaBinaryOperator,
        left: EvalValue,
        right: EvalValue,
    ) -> UseResult<EvalValue> {
        let left = self.materialize(left)?;
        let right = self.materialize(right)?;
        self.broadcast_binary(left, right, |left, right| {
            scalar_binary(operator, left, right)
        })
    }

    pub(super) fn map_numeric(
        &self,
        value: EvalValue,
        operation: impl Fn(f64) -> ScalarValue + Copy,
    ) -> UseResult<EvalValue> {
        let value = self.materialize(value)?;
        Ok(match value {
            EvalValue::Scalar(value) => EvalValue::Scalar(map_numeric_scalar(value, operation)),
            EvalValue::Array(array) => EvalValue::Array(FormulaArray {
                rows: array
                    .rows
                    .into_iter()
                    .map(|row| {
                        row.into_iter()
                            .map(|value| map_numeric_scalar(value, operation))
                            .collect()
                    })
                    .collect(),
            }),
            EvalValue::Reference(_) => {
                return Err(calculation_error(
                    "use.office.spreadsheet_formula_calculation_invalid",
                    "Reference materialization did not produce a value.",
                ));
            }
        })
    }

    pub(super) fn materialize(&self, value: EvalValue) -> UseResult<EvalValue> {
        match value {
            EvalValue::Reference(areas) => self.materialize_areas(&areas),
            value => Ok(value),
        }
    }

    pub(super) fn require_scalar(&self, value: EvalValue) -> UseResult<ScalarValue> {
        match self.materialize(value)? {
            EvalValue::Scalar(value) => Ok(value),
            EvalValue::Array(array) if array.height() == 1 && array.width() == 1 => Ok(array
                .rows
                .into_iter()
                .next()
                .and_then(|row| row.into_iter().next())
                .unwrap_or(ScalarValue::Blank)),
            EvalValue::Array(_) => Ok(ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value)),
            EvalValue::Reference(_) => Ok(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Reference,
            )),
        }
    }

    pub(super) fn broadcast_binary(
        &self,
        left: EvalValue,
        right: EvalValue,
        operation: impl Fn(ScalarValue, ScalarValue) -> UseResult<ScalarValue> + Copy,
    ) -> UseResult<EvalValue> {
        let left = into_array(left)?;
        let right = into_array(right)?;
        let height = broadcast_dimension(left.height(), right.height());
        let width = broadcast_dimension(left.width(), right.width());
        let (Some(height), Some(width)) = (height, width) else {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::Value,
            )));
        };
        checked_array_cells(height, width)?;
        let mut rows = Vec::with_capacity(height);
        let mut text_bytes = 0_usize;
        for row in 0..height {
            let mut values = Vec::with_capacity(width);
            for column in 0..width {
                let left_value = left
                    .broadcast_value(row, column)
                    .cloned()
                    .ok_or_else(invalid_array_shape)?;
                let right_value = right
                    .broadcast_value(row, column)
                    .cloned()
                    .ok_or_else(invalid_array_shape)?;
                let value = operation(left_value, right_value)?;
                accumulate_formula_text_bytes(&mut text_bytes, &value)?;
                values.push(value);
            }
            rows.push(values);
        }
        let array = FormulaArray { rows };
        if array.height() == 1 && array.width() == 1 {
            Ok(EvalValue::Scalar(
                array
                    .rows
                    .first()
                    .and_then(|row| row.first())
                    .cloned()
                    .ok_or_else(invalid_array_shape)?,
            ))
        } else {
            Ok(EvalValue::Array(array))
        }
    }
}

fn map_numeric_scalar(value: ScalarValue, operation: impl Fn(f64) -> ScalarValue) -> ScalarValue {
    match scalar_number(value) {
        Ok(number) => operation(number),
        Err(error) => ScalarValue::Error(error),
    }
}

fn scalar_number(value: ScalarValue) -> Result<f64, SpreadsheetFormulaErrorLiteral> {
    match value {
        ScalarValue::Blank => Ok(0.0),
        ScalarValue::Number(value) => Ok(value),
        ScalarValue::Boolean(value) => Ok(if value { 1.0 } else { 0.0 }),
        ScalarValue::Text(value) => value
            .parse::<f64>()
            .ok()
            .filter(|value| value.is_finite())
            .ok_or(SpreadsheetFormulaErrorLiteral::Value),
        ScalarValue::Error(error) => Err(error),
    }
}

fn scalar_boolean(value: ScalarValue) -> Result<bool, SpreadsheetFormulaErrorLiteral> {
    match value {
        ScalarValue::Blank => Ok(false),
        ScalarValue::Number(value) => Ok(value != 0.0),
        ScalarValue::Boolean(value) => Ok(value),
        ScalarValue::Text(value) if value.eq_ignore_ascii_case("TRUE") => Ok(true),
        ScalarValue::Text(value) if value.eq_ignore_ascii_case("FALSE") => Ok(false),
        ScalarValue::Text(_) => Err(SpreadsheetFormulaErrorLiteral::Value),
        ScalarValue::Error(error) => Err(error),
    }
}

fn scalar_text(value: ScalarValue) -> Result<String, SpreadsheetFormulaErrorLiteral> {
    match value {
        ScalarValue::Blank => Ok(String::new()),
        ScalarValue::Number(value) => Ok(format_number(value)),
        ScalarValue::Text(value) => Ok(value),
        ScalarValue::Boolean(true) => Ok("TRUE".to_string()),
        ScalarValue::Boolean(false) => Ok("FALSE".to_string()),
        ScalarValue::Error(error) => Err(error),
    }
}
