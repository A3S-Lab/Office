use a3s_use_core::UseResult;

use crate::spreadsheet_formula::evaluate::MAX_SPREADSHEET_FORMULA_SPILL_CELLS;
use crate::spreadsheet_reference::{MAX_COLUMNS, MAX_ROWS};

use super::super::{
    calculation_error, finite_or_number_error, EvalValue, EvaluationContext, FormulaCellKey,
    FormulaReferenceArea, ScalarValue,
};

/// `SUMIF(range, criteria, [sum_range])`.
///
/// The summed block has the criteria range's shape and starts at the sum
/// range's top-left cell. Omitting `sum_range` sums the criteria range.
/// Wildcard criteria fail closed; they are not treated as literal text.
pub(super) fn sum_if(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let criteria_area = one_rectangle(argument(arguments, 0)?, "SUMIF criteria range")?;
    let criterion = parse_criterion(criterion_scalar(context, argument(arguments, 1)?)?)?;
    let sum_anchor = match arguments.get(2) {
        Some(value) => one_rectangle(value, "SUMIF sum range")?,
        None => criteria_area,
    };
    let height = criteria_area.end_row - criteria_area.start_row;
    let width = criteria_area.end_column - criteria_area.start_column;
    let cells = criteria_area.cell_count().ok_or_else(sum_if_limit_error)?;
    if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
        return Err(sum_if_limit_error().with_detail("cells", cells));
    }
    let _sum_end_row = sum_anchor
        .start_row
        .checked_add(height)
        .filter(|row| *row <= MAX_ROWS)
        .ok_or_else(sum_if_limit_error)?;
    let _sum_end_column = sum_anchor
        .start_column
        .checked_add(width)
        .filter(|column| *column <= MAX_COLUMNS)
        .ok_or_else(sum_if_limit_error)?;

    let mut sum = 0.0_f64;
    for row_offset in 0..=height {
        for column_offset in 0..=width {
            let criteria_key = FormulaCellKey {
                sheet: criteria_area.sheet,
                row: criteria_area.start_row + row_offset,
                column: criteria_area.start_column + column_offset,
            };
            let criteria_value = cell_value(context, criteria_key);
            if !criterion.matches(&criteria_value) {
                continue;
            }
            let sum_key = FormulaCellKey {
                sheet: sum_anchor.sheet,
                row: sum_anchor.start_row + row_offset,
                column: sum_anchor.start_column + column_offset,
            };
            match cell_value(context, sum_key) {
                ScalarValue::Number(value) => sum += value,
                ScalarValue::Error(error) => {
                    return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                }
                ScalarValue::Blank | ScalarValue::Text(_) | ScalarValue::Boolean(_) => {}
            }
        }
    }
    Ok(EvalValue::Scalar(finite_or_number_error(sum)))
}

fn argument(arguments: &[EvalValue], index: usize) -> UseResult<&EvalValue> {
    arguments.get(index).ok_or_else(|| {
        calculation_error(
            "use.office.spreadsheet_formula_function_arity",
            "SUMIF requires a criteria range and a criterion.",
        )
    })
}

fn one_rectangle(value: &EvalValue, role: &str) -> UseResult<FormulaReferenceArea> {
    let EvalValue::Reference(areas) = value else {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_sumif_range_unsupported",
            format!("{role} must be one worksheet rectangle."),
        ));
    };
    if areas.len() != 1 {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_sumif_range_unsupported",
            format!("{role} must be one worksheet rectangle."),
        ));
    }
    Ok(areas[0])
}

fn criterion_scalar(context: &EvaluationContext<'_>, value: &EvalValue) -> UseResult<ScalarValue> {
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
            "use.office.spreadsheet_formula_sumif_criteria_unsupported",
            "SUMIF criteria must be one value, not a range.",
        )),
    }
}

fn cell_value(context: &EvaluationContext<'_>, key: FormulaCellKey) -> ScalarValue {
    context
        .values
        .get(&key)
        .cloned()
        .unwrap_or(ScalarValue::Blank)
}

fn sum_if_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_spill_limit",
        format!(
            "SUMIF supports at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} criteria cells within worksheet limits."
        ),
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Compare {
    Equal,
    NotEqual,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
}

enum Criterion {
    Number(Compare, f64),
    Text(Compare, String),
    Boolean(bool),
    Blank,
}

impl Criterion {
    fn matches(&self, value: &ScalarValue) -> bool {
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

fn parse_criterion(value: ScalarValue) -> UseResult<Criterion> {
    match value {
        ScalarValue::Error(error) => Err(calculation_error(
            "use.office.spreadsheet_formula_sumif_criteria_unsupported",
            format!("SUMIF criteria cannot be the error {error:?}."),
        )),
        ScalarValue::Blank => Ok(Criterion::Blank),
        ScalarValue::Boolean(value) => Ok(Criterion::Boolean(value)),
        ScalarValue::Number(value) => Ok(Criterion::Number(Compare::Equal, value)),
        ScalarValue::Text(value) => parse_text_criterion(&value),
    }
}

fn parse_text_criterion(value: &str) -> UseResult<Criterion> {
    let (compare, operand) = split_compare(value);
    if contains_wildcard(operand) {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_sumif_criteria_unsupported",
            "SUMIF wildcard criteria are not calculated. Pass an exact value or a numeric comparison.",
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
