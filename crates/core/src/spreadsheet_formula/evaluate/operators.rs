use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;
use crate::{
    SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaValue,
};

use super::{
    scalar_number, scalar_text, EvalValue, FormulaArray, ScalarValue,
    MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES, MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
    MAX_SPREADSHEET_FORMULA_TEXT_BYTES,
};

pub(super) fn scalar_binary(
    operator: SpreadsheetFormulaBinaryOperator,
    left: ScalarValue,
    right: ScalarValue,
) -> UseResult<ScalarValue> {
    if let ScalarValue::Error(error) = &left {
        return Ok(ScalarValue::Error(*error));
    }
    if let ScalarValue::Error(error) = &right {
        return Ok(ScalarValue::Error(*error));
    }
    Ok(match operator {
        SpreadsheetFormulaBinaryOperator::Add
        | SpreadsheetFormulaBinaryOperator::Subtract
        | SpreadsheetFormulaBinaryOperator::Multiply
        | SpreadsheetFormulaBinaryOperator::Divide
        | SpreadsheetFormulaBinaryOperator::Power => {
            let (Ok(left), Ok(right)) = (scalar_number(left), scalar_number(right)) else {
                return Ok(ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value));
            };
            let value = match operator {
                SpreadsheetFormulaBinaryOperator::Add => left + right,
                SpreadsheetFormulaBinaryOperator::Subtract => left - right,
                SpreadsheetFormulaBinaryOperator::Multiply => left * right,
                SpreadsheetFormulaBinaryOperator::Divide if right == 0.0 => {
                    return Ok(ScalarValue::Error(
                        SpreadsheetFormulaErrorLiteral::DivisionByZero,
                    ));
                }
                SpreadsheetFormulaBinaryOperator::Divide => left / right,
                SpreadsheetFormulaBinaryOperator::Power => left.powf(right),
                _ => return Ok(ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value)),
            };
            finite_or_number_error(value)
        }
        SpreadsheetFormulaBinaryOperator::Concatenate => {
            match (scalar_text(left), scalar_text(right)) {
                (Ok(left), Ok(right)) => {
                    let bytes = left
                        .len()
                        .checked_add(right.len())
                        .ok_or_else(formula_text_limit_error)?;
                    if bytes > MAX_SPREADSHEET_FORMULA_TEXT_BYTES {
                        return Err(formula_text_limit_error().with_detail("bytes", bytes));
                    }
                    let mut value = String::with_capacity(bytes);
                    value.push_str(&left);
                    value.push_str(&right);
                    ScalarValue::Text(value)
                }
                (Err(error), _) | (_, Err(error)) => ScalarValue::Error(error),
            }
        }
        SpreadsheetFormulaBinaryOperator::Equal
        | SpreadsheetFormulaBinaryOperator::NotEqual
        | SpreadsheetFormulaBinaryOperator::LessThan
        | SpreadsheetFormulaBinaryOperator::LessThanOrEqual
        | SpreadsheetFormulaBinaryOperator::GreaterThan
        | SpreadsheetFormulaBinaryOperator::GreaterThanOrEqual => {
            let ordering = compare_scalars(&left, &right);
            let value = match operator {
                SpreadsheetFormulaBinaryOperator::Equal => ordering == std::cmp::Ordering::Equal,
                SpreadsheetFormulaBinaryOperator::NotEqual => ordering != std::cmp::Ordering::Equal,
                SpreadsheetFormulaBinaryOperator::LessThan => ordering.is_lt(),
                SpreadsheetFormulaBinaryOperator::LessThanOrEqual => !ordering.is_gt(),
                SpreadsheetFormulaBinaryOperator::GreaterThan => ordering.is_gt(),
                SpreadsheetFormulaBinaryOperator::GreaterThanOrEqual => !ordering.is_lt(),
                _ => false,
            };
            ScalarValue::Boolean(value)
        }
        _ => ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value),
    })
}

fn compare_scalars(left: &ScalarValue, right: &ScalarValue) -> std::cmp::Ordering {
    match (left, right) {
        (ScalarValue::Blank, ScalarValue::Blank) => std::cmp::Ordering::Equal,
        (ScalarValue::Number(left), ScalarValue::Number(right)) => {
            left.partial_cmp(right).unwrap_or(std::cmp::Ordering::Equal)
        }
        (ScalarValue::Boolean(left), ScalarValue::Boolean(right)) => left.cmp(right),
        (ScalarValue::Text(left), ScalarValue::Text(right)) => {
            left.to_lowercase().cmp(&right.to_lowercase())
        }
        (ScalarValue::Blank, ScalarValue::Number(value))
        | (ScalarValue::Number(value), ScalarValue::Blank)
            if *value == 0.0 =>
        {
            std::cmp::Ordering::Equal
        }
        _ => scalar_rank(left).cmp(&scalar_rank(right)),
    }
}

fn scalar_rank(value: &ScalarValue) -> u8 {
    match value {
        ScalarValue::Blank => 0,
        ScalarValue::Number(_) => 1,
        ScalarValue::Text(_) => 2,
        ScalarValue::Boolean(_) => 3,
        ScalarValue::Error(_) => 4,
    }
}

pub(super) fn finite_or_number_error(value: f64) -> ScalarValue {
    if value.is_finite() {
        ScalarValue::Number(if value == 0.0 { 0.0 } else { value })
    } else {
        ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Number)
    }
}

pub(super) fn into_array(value: EvalValue) -> UseResult<FormulaArray> {
    match value {
        EvalValue::Scalar(value) => Ok(FormulaArray::scalar(value)),
        EvalValue::Array(array) => Ok(array),
        EvalValue::Reference(_) => Err(calculation_error(
            "use.office.spreadsheet_formula_calculation_invalid",
            "Reference must be materialized before array broadcasting.",
        )),
    }
}

pub(super) fn broadcast_dimension(left: usize, right: usize) -> Option<usize> {
    if left == right {
        Some(left)
    } else if left == 1 {
        Some(right)
    } else if right == 1 {
        Some(left)
    } else {
        None
    }
}

pub(super) fn checked_array_cells(height: usize, width: usize) -> UseResult<usize> {
    let cells = height.checked_mul(width).ok_or_else(spill_limit_error)?;
    if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
        return Err(spill_limit_error().with_detail("cells", cells));
    }
    Ok(cells)
}

pub(super) fn public_scalar(value: &ScalarValue) -> SpreadsheetFormulaValue {
    match value {
        ScalarValue::Blank => SpreadsheetFormulaValue::Blank,
        ScalarValue::Number(value) => SpreadsheetFormulaValue::Number {
            value: format_number(*value),
        },
        ScalarValue::Text(value) => SpreadsheetFormulaValue::Text {
            value: value.clone(),
        },
        ScalarValue::Boolean(value) => SpreadsheetFormulaValue::Boolean { value: *value },
        ScalarValue::Error(error) => SpreadsheetFormulaValue::Error { error: *error },
    }
}

pub(super) fn ensure_formula_text_limit(value: &ScalarValue) -> UseResult<()> {
    let ScalarValue::Text(value) = value else {
        return Ok(());
    };
    if value.len() > MAX_SPREADSHEET_FORMULA_TEXT_BYTES {
        return Err(formula_text_limit_error().with_detail("bytes", value.len()));
    }
    Ok(())
}

pub(super) fn accumulate_formula_text_bytes(
    bytes: &mut usize,
    value: &ScalarValue,
) -> UseResult<()> {
    ensure_formula_text_limit(value)?;
    let ScalarValue::Text(value) = value else {
        return Ok(());
    };
    *bytes = bytes
        .checked_add(value.len())
        .ok_or_else(calculation_text_limit_error)?;
    if *bytes > MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES {
        return Err(calculation_text_limit_error().with_detail("bytes", *bytes));
    }
    Ok(())
}

pub(super) fn format_number(value: f64) -> String {
    if value == 0.0 {
        "0".to_string()
    } else {
        value.to_string()
    }
}

pub(super) fn spill_limit_error() -> UseError {
    calculation_error(
        "use.office.spreadsheet_formula_spill_limit",
        format!(
            "Native formula spills support at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} cells within worksheet limits."
        ),
    )
}

fn formula_text_limit_error() -> UseError {
    calculation_error(
        "use.office.spreadsheet_formula_text_limit",
        format!(
            "Native formula text results support at most {MAX_SPREADSHEET_FORMULA_TEXT_BYTES} UTF-8 bytes."
        ),
    )
}

pub(super) fn calculation_text_limit_error() -> UseError {
    calculation_error(
        "use.office.spreadsheet_formula_text_limit",
        format!(
            "Native formula calculation produces at most {MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES} cumulative UTF-8 text-result bytes."
        ),
    )
}

pub(super) fn invalid_array_shape() -> UseError {
    calculation_error(
        "use.office.spreadsheet_formula_calculation_invalid",
        "Native formula calculation produced an invalid array shape.",
    )
}

pub(super) fn calculation_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}
