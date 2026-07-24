use std::cmp::Ordering;

use a3s_office_formula_parser::{SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaErrorLiteral};

use super::{EvaluationFailure, SpreadsheetValue, MAX_SPREADSHEET_TEXT_BYTES};

#[derive(Debug, Clone)]
pub(super) enum EvaluatedValue {
    Scalar(SpreadsheetValue),
    Range(Vec<SpreadsheetValue>),
}

impl EvaluatedValue {
    pub(super) fn into_scalar(self) -> SpreadsheetValue {
        match self {
            Self::Scalar(value) => value,
            Self::Range(values) => values.into_iter().next().unwrap_or(SpreadsheetValue::Blank),
        }
    }

    pub(super) fn into_values(self) -> Vec<SpreadsheetValue> {
        match self {
            Self::Scalar(value) => vec![value],
            Self::Range(values) => values,
        }
    }
}

pub(super) fn apply_binary(
    operator: SpreadsheetFormulaBinaryOperator,
    left: SpreadsheetValue,
    right: SpreadsheetValue,
) -> SpreadsheetValue {
    if let SpreadsheetValue::Error { .. } = left {
        return left;
    }
    if let SpreadsheetValue::Error { .. } = right {
        return right;
    }
    match operator {
        SpreadsheetFormulaBinaryOperator::Add
        | SpreadsheetFormulaBinaryOperator::Subtract
        | SpreadsheetFormulaBinaryOperator::Multiply
        | SpreadsheetFormulaBinaryOperator::Divide
        | SpreadsheetFormulaBinaryOperator::Power => {
            let (Ok(left), Ok(right)) = (scalar_number(left), scalar_number(right)) else {
                return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Value);
            };
            let result = match operator {
                SpreadsheetFormulaBinaryOperator::Add => left + right,
                SpreadsheetFormulaBinaryOperator::Subtract => left - right,
                SpreadsheetFormulaBinaryOperator::Multiply => left * right,
                SpreadsheetFormulaBinaryOperator::Divide if right == 0.0 => {
                    return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero);
                }
                SpreadsheetFormulaBinaryOperator::Divide => left / right,
                SpreadsheetFormulaBinaryOperator::Power => left.powf(right),
                _ => unreachable!("arithmetic operator was matched above"),
            };
            finite_number(result)
        }
        SpreadsheetFormulaBinaryOperator::Concatenate => {
            let (Ok(left), Ok(right)) = (scalar_text(left), scalar_text(right)) else {
                return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Value);
            };
            let Some(length) = left.len().checked_add(right.len()) else {
                return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Value);
            };
            if length > MAX_SPREADSHEET_TEXT_BYTES {
                return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Value);
            }
            SpreadsheetValue::Text {
                value: format!("{left}{right}"),
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
                SpreadsheetFormulaBinaryOperator::Equal => ordering == Ordering::Equal,
                SpreadsheetFormulaBinaryOperator::NotEqual => ordering != Ordering::Equal,
                SpreadsheetFormulaBinaryOperator::LessThan => ordering.is_lt(),
                SpreadsheetFormulaBinaryOperator::LessThanOrEqual => !ordering.is_gt(),
                SpreadsheetFormulaBinaryOperator::GreaterThan => ordering.is_gt(),
                SpreadsheetFormulaBinaryOperator::GreaterThanOrEqual => !ordering.is_lt(),
                _ => unreachable!("comparison operator was matched above"),
            };
            SpreadsheetValue::Boolean { value }
        }
        _ => SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Value),
    }
}

pub(super) fn scalar_number(
    value: SpreadsheetValue,
) -> Result<f64, SpreadsheetFormulaErrorLiteral> {
    match value {
        SpreadsheetValue::Blank => Ok(0.0),
        SpreadsheetValue::Number { value } => Ok(value),
        SpreadsheetValue::Boolean { value } => Ok(if value { 1.0 } else { 0.0 }),
        SpreadsheetValue::Text { value } => value
            .parse::<f64>()
            .ok()
            .filter(|value| value.is_finite())
            .ok_or(SpreadsheetFormulaErrorLiteral::Value),
        SpreadsheetValue::Error { value } => Err(SpreadsheetFormulaErrorLiteral::parse(&value)
            .unwrap_or(SpreadsheetFormulaErrorLiteral::Value)),
    }
}

pub(super) fn scalar_boolean(
    value: SpreadsheetValue,
) -> Result<bool, SpreadsheetFormulaErrorLiteral> {
    match value {
        SpreadsheetValue::Blank => Ok(false),
        SpreadsheetValue::Number { value } => Ok(value != 0.0),
        SpreadsheetValue::Boolean { value } => Ok(value),
        SpreadsheetValue::Text { value } if value.eq_ignore_ascii_case("TRUE") => Ok(true),
        SpreadsheetValue::Text { value } if value.eq_ignore_ascii_case("FALSE") => Ok(false),
        SpreadsheetValue::Text { .. } => Err(SpreadsheetFormulaErrorLiteral::Value),
        SpreadsheetValue::Error { value } => Err(SpreadsheetFormulaErrorLiteral::parse(&value)
            .unwrap_or(SpreadsheetFormulaErrorLiteral::Value)),
    }
}

pub(super) fn scalar_text(
    value: SpreadsheetValue,
) -> Result<String, SpreadsheetFormulaErrorLiteral> {
    match value {
        SpreadsheetValue::Blank => Ok(String::new()),
        SpreadsheetValue::Number { value } => Ok(format_number(value)),
        SpreadsheetValue::Text { value } => Ok(value),
        SpreadsheetValue::Boolean { value: true } => Ok("TRUE".into()),
        SpreadsheetValue::Boolean { value: false } => Ok("FALSE".into()),
        SpreadsheetValue::Error { value } => Err(SpreadsheetFormulaErrorLiteral::parse(&value)
            .unwrap_or(SpreadsheetFormulaErrorLiteral::Value)),
    }
}

pub(super) fn finite_number(value: f64) -> SpreadsheetValue {
    if value.is_finite() {
        SpreadsheetValue::Number {
            value: if value == 0.0 { 0.0 } else { value },
        }
    } else {
        SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Number)
    }
}

pub(super) fn value_error() -> SpreadsheetValue {
    SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Value)
}

pub(super) fn unsupported(message: impl Into<String>) -> EvaluationFailure {
    EvaluationFailure::unsupported(message)
}

fn compare_scalars(left: &SpreadsheetValue, right: &SpreadsheetValue) -> Ordering {
    match (left, right) {
        (SpreadsheetValue::Blank, SpreadsheetValue::Blank) => Ordering::Equal,
        (SpreadsheetValue::Number { value: left }, SpreadsheetValue::Number { value: right }) => {
            left.partial_cmp(right).unwrap_or(Ordering::Equal)
        }
        (SpreadsheetValue::Boolean { value: left }, SpreadsheetValue::Boolean { value: right }) => {
            left.cmp(right)
        }
        (SpreadsheetValue::Text { value: left }, SpreadsheetValue::Text { value: right }) => {
            left.to_lowercase().cmp(&right.to_lowercase())
        }
        (SpreadsheetValue::Blank, SpreadsheetValue::Number { value })
        | (SpreadsheetValue::Number { value }, SpreadsheetValue::Blank)
            if *value == 0.0 =>
        {
            Ordering::Equal
        }
        _ => scalar_rank(left).cmp(&scalar_rank(right)),
    }
}

fn scalar_rank(value: &SpreadsheetValue) -> u8 {
    match value {
        SpreadsheetValue::Blank => 0,
        SpreadsheetValue::Number { .. } => 1,
        SpreadsheetValue::Text { .. } => 2,
        SpreadsheetValue::Boolean { .. } => 3,
        SpreadsheetValue::Error { .. } => 4,
    }
}

fn format_number(value: f64) -> String {
    if value == 0.0 {
        "0".into()
    } else {
        value.to_string()
    }
}
