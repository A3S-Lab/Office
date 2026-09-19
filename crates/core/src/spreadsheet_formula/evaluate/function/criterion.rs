use a3s_use_core::UseResult;

use super::super::{
    calculation_error, EvalValue, EvaluationContext, FormulaCellKey, FormulaReferenceArea,
    ScalarValue,
};

#[derive(Clone, Copy)]
pub(super) struct CriteriaCall {
    pub name: &'static str,
    pub criteria_code: &'static str,
    pub range_code: &'static str,
}

pub(super) const SUM_IF: CriteriaCall = CriteriaCall {
    name: "SUMIF",
    criteria_code: "use.office.spreadsheet_formula_sumif_criteria_unsupported",
    range_code: "use.office.spreadsheet_formula_sumif_range_unsupported",
};

pub(super) const COUNT_IF: CriteriaCall = CriteriaCall {
    name: "COUNTIF",
    criteria_code: "use.office.spreadsheet_formula_countif_criteria_unsupported",
    range_code: "use.office.spreadsheet_formula_countif_range_unsupported",
};

pub(super) const AVERAGE_IF: CriteriaCall = CriteriaCall {
    name: "AVERAGEIF",
    criteria_code: "use.office.spreadsheet_formula_averageif_criteria_unsupported",
    range_code: "use.office.spreadsheet_formula_averageif_range_unsupported",
};

pub(super) fn one_rectangle(
    value: &EvalValue,
    role: &str,
    call: CriteriaCall,
) -> UseResult<FormulaReferenceArea> {
    let EvalValue::Reference(areas) = value else {
        return Err(calculation_error(
            call.range_code,
            format!("{role} must be one worksheet rectangle."),
        ));
    };
    if areas.len() != 1 {
        return Err(calculation_error(
            call.range_code,
            format!("{role} must be one worksheet rectangle."),
        ));
    }
    Ok(areas[0])
}

pub(super) fn criterion_scalar(
    context: &EvaluationContext<'_>,
    value: &EvalValue,
    call: CriteriaCall,
) -> UseResult<ScalarValue> {
    match value {
        EvalValue::Scalar(value) => Ok(value.clone()),
        EvalValue::Reference(areas) if areas.len() == 1 && areas[0].cell_count() == Some(1) => {
            Ok(cell_value(
                context,
                FormulaCellKey {
                    sheet: areas[0].sheet,
                    row: areas[0].start_row,
                    column: areas[0].start_column,
                },
            ))
        }
        _ => Err(calculation_error(
            call.criteria_code,
            format!("{} criteria must be one value, not a range.", call.name),
        )),
    }
}

pub(super) fn cell_value(context: &EvaluationContext<'_>, key: FormulaCellKey) -> ScalarValue {
    context
        .values
        .get(&key)
        .cloned()
        .unwrap_or(ScalarValue::Blank)
}

pub(super) fn parse_criterion(call: CriteriaCall, value: ScalarValue) -> UseResult<Criterion> {
    match value {
        ScalarValue::Error(error) => Err(calculation_error(
            call.criteria_code,
            format!("{} criteria cannot be the error {error:?}.", call.name),
        )),
        ScalarValue::Blank => Ok(Criterion::Blank),
        ScalarValue::Boolean(value) => Ok(Criterion::Boolean(value)),
        ScalarValue::Number(value) => Ok(Criterion::Number(Compare::Equal, value)),
        ScalarValue::Text(value) => parse_text_criterion(call, &value),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Compare {
    Equal,
    NotEqual,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
}

pub(super) enum Criterion {
    Number(Compare, f64),
    Text(Compare, String),
    Boolean(bool),
    Blank,
}

impl Criterion {
    pub(super) fn matches(&self, value: &ScalarValue) -> bool {
        if matches!(value, ScalarValue::Error(_)) {
            return false;
        }
        match self {
            Self::Blank => matches!(value, ScalarValue::Blank),
            Self::Boolean(expected) => {
                matches!(value, ScalarValue::Boolean(actual) if actual == expected)
            }
            Self::Number(compare, expected) => number_matches(value, *compare, *expected),
            Self::Text(compare, expected) => text_matches(value, *compare, expected),
        }
    }
}

fn parse_text_criterion(call: CriteriaCall, value: &str) -> UseResult<Criterion> {
    let (compare, operand) = split_compare(value);
    if contains_wildcard(operand) {
        return Err(calculation_error(
            call.criteria_code,
            format!(
                "{} wildcard criteria are not calculated. Pass an exact value or a numeric comparison.",
                call.name
            ),
        ));
    }
    let operand = unescape_literal(operand);
    if operand.is_empty() && compare == Compare::Equal {
        return Ok(Criterion::Blank);
    }
    if let Some(number) = parse_finite_number(&operand) {
        return Ok(Criterion::Number(compare, number));
    }
    Ok(Criterion::Text(compare, operand))
}

fn split_compare(value: &str) -> (Compare, &str) {
    if let Some(rest) = value.strip_prefix(">=") {
        (Compare::GreaterOrEqual, rest)
    } else if let Some(rest) = value.strip_prefix("<=") {
        (Compare::LessOrEqual, rest)
    } else if let Some(rest) = value.strip_prefix("<>") {
        (Compare::NotEqual, rest)
    } else if let Some(rest) = value.strip_prefix('>') {
        (Compare::Greater, rest)
    } else if let Some(rest) = value.strip_prefix('<') {
        (Compare::Less, rest)
    } else if let Some(rest) = value.strip_prefix('=') {
        (Compare::Equal, rest)
    } else {
        (Compare::Equal, value)
    }
}

fn contains_wildcard(value: &str) -> bool {
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

fn unescape_literal(value: &str) -> String {
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

fn parse_finite_number(value: &str) -> Option<f64> {
    let number = value.parse::<f64>().ok()?;
    number.is_finite().then_some(number)
}

fn number_matches(value: &ScalarValue, compare: Compare, expected: f64) -> bool {
    match value {
        ScalarValue::Number(actual) => compare_numbers(*actual, expected, compare),
        ScalarValue::Blank | ScalarValue::Text(_) | ScalarValue::Boolean(_) => {
            compare == Compare::NotEqual
        }
        ScalarValue::Error(_) => false,
    }
}

fn text_matches(value: &ScalarValue, compare: Compare, expected: &str) -> bool {
    let actual = match value {
        ScalarValue::Text(value) => value.as_str(),
        ScalarValue::Blank => "",
        ScalarValue::Number(_) | ScalarValue::Boolean(_) => {
            return compare == Compare::NotEqual;
        }
        ScalarValue::Error(_) => return false,
    };
    let ordering = actual
        .to_ascii_lowercase()
        .cmp(&expected.to_ascii_lowercase());
    match compare {
        Compare::Equal => ordering.is_eq(),
        Compare::NotEqual => ordering.is_ne(),
        Compare::Greater => ordering.is_gt(),
        Compare::GreaterOrEqual => ordering.is_ge(),
        Compare::Less => ordering.is_lt(),
        Compare::LessOrEqual => ordering.is_le(),
    }
}

fn compare_numbers(actual: f64, expected: f64, compare: Compare) -> bool {
    match compare {
        Compare::Equal => actual == expected,
        Compare::NotEqual => actual != expected,
        Compare::Greater => actual > expected,
        Compare::GreaterOrEqual => actual >= expected,
        Compare::Less => actual < expected,
        Compare::LessOrEqual => actual <= expected,
    }
}
