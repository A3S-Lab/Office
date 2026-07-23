use a3s_use_core::UseResult;

use crate::SpreadsheetFormulaErrorLiteral;

use super::super::{
    calculation_error, scalar_boolean, scalar_number, scalar_text, EvalValue, EvaluationContext,
    FormulaArray, ScalarValue, MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
};

#[derive(Debug, Clone, Copy)]
pub(super) enum Aggregate {
    Sum,
    Average,
    Minimum,
    Maximum,
}

pub(super) fn aggregate(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
    operation: Aggregate,
) -> UseResult<EvalValue> {
    let mut count = 0_usize;
    let mut sum = 0.0_f64;
    let mut minimum = None;
    let mut maximum = None;
    let mut materialized_cells = 0_usize;
    for argument in arguments {
        match argument {
            EvalValue::Reference(areas) => {
                let values = context.populated_reference_values(areas)?;
                add_function_cells(&mut materialized_cells, values.len())?;
                for value in values {
                    match value {
                        ScalarValue::Number(value) => {
                            record_number(&mut count, &mut sum, &mut minimum, &mut maximum, value)
                        }
                        ScalarValue::Error(error) => {
                            return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                        }
                        _ => {}
                    }
                }
            }
            EvalValue::Array(array) => {
                add_function_cells(
                    &mut materialized_cells,
                    array.height().saturating_mul(array.width()),
                )?;
                for value in array.rows.iter().flatten().cloned() {
                    match scalar_number(value) {
                        Ok(value) => {
                            record_number(&mut count, &mut sum, &mut minimum, &mut maximum, value)
                        }
                        Err(error) => {
                            return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                        }
                    }
                }
            }
            EvalValue::Scalar(value) => {
                add_function_cells(&mut materialized_cells, 1)?;
                match scalar_number(value.clone()) {
                    Ok(value) => {
                        record_number(&mut count, &mut sum, &mut minimum, &mut maximum, value)
                    }
                    Err(error) => return Ok(EvalValue::Scalar(ScalarValue::Error(error))),
                }
            }
        }
    }
    let value = match operation {
        Aggregate::Sum => sum,
        Aggregate::Average if count == 0 => {
            return Ok(EvalValue::Scalar(ScalarValue::Error(
                SpreadsheetFormulaErrorLiteral::DivisionByZero,
            )));
        }
        Aggregate::Average => sum / count as f64,
        Aggregate::Minimum => minimum.unwrap_or(0.0),
        Aggregate::Maximum => maximum.unwrap_or(0.0),
    };
    Ok(EvalValue::Scalar(super::super::finite_or_number_error(
        value,
    )))
}

pub(super) fn count(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let mut count = 0_usize;
    let mut materialized_cells = 0_usize;
    for argument in arguments {
        match argument {
            EvalValue::Reference(areas) => {
                let values = context.populated_reference_values(areas)?;
                add_function_cells(&mut materialized_cells, values.len())?;
                count = count.saturating_add(
                    values
                        .iter()
                        .filter(|value| matches!(value, ScalarValue::Number(_)))
                        .count(),
                );
            }
            EvalValue::Array(array) => {
                add_function_cells(
                    &mut materialized_cells,
                    array.height().saturating_mul(array.width()),
                )?;
                count = count.saturating_add(
                    array
                        .rows
                        .iter()
                        .flatten()
                        .filter(|value| matches!(value, ScalarValue::Number(_)))
                        .count(),
                );
            }
            EvalValue::Scalar(value) => {
                add_function_cells(&mut materialized_cells, 1)?;
                match value {
                    ScalarValue::Number(_) | ScalarValue::Boolean(_) => {
                        count = count.saturating_add(1);
                    }
                    ScalarValue::Text(value)
                        if value
                            .parse::<f64>()
                            .ok()
                            .is_some_and(|value| value.is_finite()) =>
                    {
                        count = count.saturating_add(1);
                    }
                    ScalarValue::Error(error) => {
                        return Ok(EvalValue::Scalar(ScalarValue::Error(*error)));
                    }
                    ScalarValue::Blank | ScalarValue::Text(_) => {}
                }
            }
        }
    }
    Ok(EvalValue::Scalar(ScalarValue::Number(count as f64)))
}

pub(super) fn count_a(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let mut count = 0_usize;
    let mut materialized_cells = 0_usize;
    for argument in arguments {
        match argument {
            EvalValue::Reference(areas) => {
                let values = context.populated_reference_values(areas)?;
                add_function_cells(&mut materialized_cells, values.len())?;
                count = count.saturating_add(
                    values
                        .iter()
                        .filter(|value| !matches!(value, ScalarValue::Blank))
                        .count(),
                );
            }
            EvalValue::Array(array) => {
                add_function_cells(
                    &mut materialized_cells,
                    array.height().saturating_mul(array.width()),
                )?;
                count = count.saturating_add(
                    array
                        .rows
                        .iter()
                        .flatten()
                        .filter(|value| !matches!(value, ScalarValue::Blank))
                        .count(),
                );
            }
            EvalValue::Scalar(ScalarValue::Blank) => {
                add_function_cells(&mut materialized_cells, 1)?;
            }
            EvalValue::Scalar(_) => {
                add_function_cells(&mut materialized_cells, 1)?;
                count = count.saturating_add(1);
            }
        }
    }
    Ok(EvalValue::Scalar(ScalarValue::Number(count as f64)))
}

pub(super) fn logical_aggregate(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
    and: bool,
) -> UseResult<EvalValue> {
    let mut observed = false;
    let mut result = and;
    let mut materialized_cells = 0_usize;
    for argument in arguments {
        match argument {
            EvalValue::Reference(areas) => {
                let values = context.populated_reference_values(areas)?;
                add_function_cells(&mut materialized_cells, values.len())?;
                for value in values {
                    match value {
                        ScalarValue::Blank | ScalarValue::Text(_) => {}
                        ScalarValue::Error(error) => {
                            return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                        }
                        value => {
                            observed = true;
                            match scalar_boolean(value) {
                                Ok(value) => update_logical_result(&mut result, and, value),
                                Err(error) => {
                                    return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                                }
                            }
                        }
                    }
                }
            }
            EvalValue::Array(array) => {
                add_function_cells(
                    &mut materialized_cells,
                    array.height().saturating_mul(array.width()),
                )?;
                for value in array.rows.iter().flatten().cloned() {
                    match value {
                        ScalarValue::Blank | ScalarValue::Text(_) => {}
                        ScalarValue::Error(error) => {
                            return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                        }
                        value => {
                            observed = true;
                            match scalar_boolean(value) {
                                Ok(value) => update_logical_result(&mut result, and, value),
                                Err(error) => {
                                    return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                                }
                            }
                        }
                    }
                }
            }
            EvalValue::Scalar(value) => {
                add_function_cells(&mut materialized_cells, 1)?;
                match scalar_boolean(value.clone()) {
                    Ok(value) => {
                        observed = true;
                        update_logical_result(&mut result, and, value);
                    }
                    Err(error) => {
                        return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                    }
                }
            }
        }
    }
    if !observed {
        return Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::Value,
        )));
    }
    Ok(EvalValue::Scalar(ScalarValue::Boolean(result)))
}

pub(super) fn concatenate(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let mut output = String::new();
    let mut materialized_cells = 0_usize;
    for argument in arguments {
        let values = argument_scalars(context, argument)?;
        add_function_cells(&mut materialized_cells, values.len())?;
        for value in values {
            match scalar_text(value) {
                Ok(value) => append_formula_text(&mut output, &value)?,
                Err(error) => return Ok(EvalValue::Scalar(ScalarValue::Error(error))),
            }
        }
    }
    Ok(EvalValue::Scalar(ScalarValue::Text(output)))
}

pub(super) fn logical_not(
    context: &EvaluationContext<'_>,
    value: EvalValue,
) -> UseResult<EvalValue> {
    let value = context.materialize(value)?;
    Ok(match value {
        EvalValue::Scalar(value) => EvalValue::Scalar(not_scalar(value)),
        EvalValue::Array(array) => EvalValue::Array(FormulaArray {
            rows: array
                .rows
                .into_iter()
                .map(|row| row.into_iter().map(not_scalar).collect())
                .collect(),
        }),
        EvalValue::Reference(_) => {
            return Err(calculation_error(
                "use.office.spreadsheet_formula_calculation_invalid",
                "Reference materialization did not produce a logical value.",
            ));
        }
    })
}

pub(super) fn add_function_cells(total: &mut usize, added: usize) -> UseResult<()> {
    *total = total
        .checked_add(added)
        .ok_or_else(function_argument_limit)?;
    if *total > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
        return Err(function_argument_limit().with_detail("cells", *total));
    }
    Ok(())
}

fn record_number(
    count: &mut usize,
    sum: &mut f64,
    minimum: &mut Option<f64>,
    maximum: &mut Option<f64>,
    value: f64,
) {
    *count = count.saturating_add(1);
    *sum += value;
    *minimum = Some(minimum.map_or(value, |current| current.min(value)));
    *maximum = Some(maximum.map_or(value, |current| current.max(value)));
}

fn function_argument_limit() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_function_array_limit",
        format!(
            "Native formula functions materialize at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} argument cells per call."
        ),
    )
}

fn update_logical_result(result: &mut bool, and: bool, value: bool) {
    if and {
        *result &= value;
    } else {
        *result |= value;
    }
}

fn not_scalar(value: ScalarValue) -> ScalarValue {
    match scalar_boolean(value) {
        Ok(value) => ScalarValue::Boolean(!value),
        Err(error) => ScalarValue::Error(error),
    }
}

fn argument_scalars(
    context: &EvaluationContext<'_>,
    argument: &EvalValue,
) -> UseResult<Vec<ScalarValue>> {
    match argument {
        EvalValue::Scalar(value) => Ok(vec![value.clone()]),
        EvalValue::Array(array) => Ok(array.rows.iter().flatten().cloned().collect()),
        EvalValue::Reference(areas) => {
            let value = context.materialize(EvalValue::Reference(areas.clone()))?;
            match value {
                EvalValue::Scalar(value) => Ok(vec![value]),
                EvalValue::Array(array) => Ok(array.rows.into_iter().flatten().collect()),
                EvalValue::Reference(_) => Ok(Vec::new()),
            }
        }
    }
}

fn append_formula_text(output: &mut String, value: &str) -> UseResult<()> {
    let bytes = output
        .len()
        .checked_add(value.len())
        .ok_or_else(formula_text_limit)?;
    if bytes > super::super::MAX_SPREADSHEET_FORMULA_TEXT_BYTES {
        return Err(formula_text_limit().with_detail("bytes", bytes));
    }
    output.push_str(value);
    Ok(())
}

fn formula_text_limit() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_text_limit",
        format!(
            "Native formula text results support at most {} UTF-8 bytes.",
            super::super::MAX_SPREADSHEET_FORMULA_TEXT_BYTES
        ),
    )
}
