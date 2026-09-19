use a3s_office_formula_parser::{
    SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaExpression,
    SpreadsheetFormulaExpressionKind,
};

use super::value::unsupported;
use super::{EvaluationFailure, SpreadsheetValue};

#[derive(Debug, Clone, Copy)]
pub(super) struct SumIfRectangle {
    pub(super) sheet: usize,
    pub(super) start_row: u32,
    pub(super) start_column: u32,
    pub(super) end_row: u32,
    pub(super) end_column: u32,
}

pub(super) fn conditional_argument<'a>(
    name: &str,
    arguments: &'a [Option<SpreadsheetFormulaExpression>],
    index: usize,
) -> Result<&'a SpreadsheetFormulaExpression, EvaluationFailure> {
    arguments
        .get(index)
        .and_then(Option::as_ref)
        .ok_or_else(|| unsupported(format!("{name} requires a criteria range and a criterion.")))
}

pub(super) fn ensure_conditional_criterion(
    name: &str,
    expression: &SpreadsheetFormulaExpression,
) -> Result<(), EvaluationFailure> {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
            ensure_conditional_criterion(name, inner)
        }
        SpreadsheetFormulaExpressionKind::Binary {
            operator: SpreadsheetFormulaBinaryOperator::Range,
            ..
        }
        | SpreadsheetFormulaExpressionKind::StructuredReference { .. }
        | SpreadsheetFormulaExpressionKind::Name { .. }
        | SpreadsheetFormulaExpressionKind::Array { .. } => Err(unsupported(format!(
            "{name} criteria must be one value, not a range."
        ))),
        _ => Ok(()),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SumIfCompare {
    Equal,
    NotEqual,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
}

pub(super) enum SumIfCriterion {
    Number(SumIfCompare, f64),
    Text(SumIfCompare, String),
    Boolean(bool),
    Blank,
}

impl SumIfCriterion {
    pub(super) fn matches(&self, value: &SpreadsheetValue) -> bool {
        if matches!(value, SpreadsheetValue::Error { .. }) {
            return false;
        }
        match self {
            Self::Blank => matches!(value, SpreadsheetValue::Blank),
            Self::Boolean(expected) => {
                matches!(value, SpreadsheetValue::Boolean { value } if value == expected)
            }
            Self::Number(compare, expected) => sum_if_number_matches(value, *compare, *expected),
            Self::Text(compare, expected) => sum_if_text_matches(value, *compare, expected),
        }
    }
}

pub(super) fn parse_conditional_criterion(
    name: &str,
    value: SpreadsheetValue,
) -> Result<SumIfCriterion, EvaluationFailure> {
    match value {
        SpreadsheetValue::Error { value } => Err(unsupported(format!(
            "{name} criteria cannot be the error {value}."
        ))),
        SpreadsheetValue::Blank => Ok(SumIfCriterion::Blank),
        SpreadsheetValue::Boolean { value } => Ok(SumIfCriterion::Boolean(value)),
        SpreadsheetValue::Number { value } => {
            Ok(SumIfCriterion::Number(SumIfCompare::Equal, value))
        }
        SpreadsheetValue::Text { value } => parse_conditional_text(name, &value),
    }
}

fn parse_conditional_text(name: &str, value: &str) -> Result<SumIfCriterion, EvaluationFailure> {
    let (compare, operand) = split_sum_if_compare(value);
    if sum_if_contains_wildcard(operand) {
        return Err(unsupported(format!(
            "{name} wildcard criteria are not calculated. Pass an exact value or a numeric comparison."
        )));
    }
    let operand = unescape_sum_if_literal(operand);
    if operand.is_empty() && compare == SumIfCompare::Equal {
        return Ok(SumIfCriterion::Blank);
    }
    if let Some(number) = operand
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())
    {
        return Ok(SumIfCriterion::Number(compare, number));
    }
    Ok(SumIfCriterion::Text(compare, operand))
}

fn split_sum_if_compare(value: &str) -> (SumIfCompare, &str) {
    if let Some(rest) = value.strip_prefix(">=") {
        (SumIfCompare::GreaterOrEqual, rest)
    } else if let Some(rest) = value.strip_prefix("<=") {
        (SumIfCompare::LessOrEqual, rest)
    } else if let Some(rest) = value.strip_prefix("<>") {
        (SumIfCompare::NotEqual, rest)
    } else if let Some(rest) = value.strip_prefix('>') {
        (SumIfCompare::Greater, rest)
    } else if let Some(rest) = value.strip_prefix('<') {
        (SumIfCompare::Less, rest)
    } else if let Some(rest) = value.strip_prefix('=') {
        (SumIfCompare::Equal, rest)
    } else {
        (SumIfCompare::Equal, value)
    }
}

fn sum_if_contains_wildcard(value: &str) -> bool {
    let mut characters = value.chars();
    while let Some(character) = characters.next() {
        if character == '~' {
            characters.next();
            continue;
        }
        if character == '*' || character == '?' {
            return true;
        }
    }
    false
}

fn unescape_sum_if_literal(value: &str) -> String {
    let mut characters = value.chars();
    let mut literal = String::new();
    while let Some(character) = characters.next() {
        if character == '~' {
            if let Some(escaped) = characters.next() {
                literal.push(escaped);
            }
            continue;
        }
        literal.push(character);
    }
    literal
}

fn sum_if_number_matches(value: &SpreadsheetValue, compare: SumIfCompare, expected: f64) -> bool {
    match value {
        SpreadsheetValue::Number { value } => sum_if_compare_numbers(*value, expected, compare),
        SpreadsheetValue::Blank
        | SpreadsheetValue::Text { .. }
        | SpreadsheetValue::Boolean { .. } => compare == SumIfCompare::NotEqual,
        SpreadsheetValue::Error { .. } => false,
    }
}

fn sum_if_text_matches(value: &SpreadsheetValue, compare: SumIfCompare, expected: &str) -> bool {
    let actual = match value {
        SpreadsheetValue::Text { value } => value.as_str(),
        SpreadsheetValue::Blank => "",
        SpreadsheetValue::Number { .. } | SpreadsheetValue::Boolean { .. } => {
            return compare == SumIfCompare::NotEqual;
        }
        SpreadsheetValue::Error { .. } => return false,
    };
    let ordering = actual
        .to_ascii_lowercase()
        .cmp(&expected.to_ascii_lowercase());
    match compare {
        SumIfCompare::Equal => ordering.is_eq(),
        SumIfCompare::NotEqual => ordering.is_ne(),
        SumIfCompare::Greater => ordering.is_gt(),
        SumIfCompare::GreaterOrEqual => ordering.is_ge(),
        SumIfCompare::Less => ordering.is_lt(),
        SumIfCompare::LessOrEqual => ordering.is_le(),
    }
}

fn sum_if_compare_numbers(actual: f64, expected: f64, compare: SumIfCompare) -> bool {
    match compare {
        SumIfCompare::Equal => actual == expected,
        SumIfCompare::NotEqual => actual != expected,
        SumIfCompare::Greater => actual > expected,
        SumIfCompare::GreaterOrEqual => actual >= expected,
        SumIfCompare::Less => actual < expected,
        SumIfCompare::LessOrEqual => actual <= expected,
    }
}
