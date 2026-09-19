use a3s_office_formula_parser::{
    SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaExpression,
    SpreadsheetFormulaExpressionKind,
};

use super::value::{
    finite_number, scalar_boolean, scalar_number, scalar_text, unsupported, value_error,
    EvaluatedValue,
};
use super::{
    CellKey, EvaluationFailure, SpreadsheetEvaluator, SpreadsheetValue, MAX_SPREADSHEET_CELLS,
    MAX_SPREADSHEET_COLUMNS, MAX_SPREADSHEET_ROWS, MAX_SPREADSHEET_TEXT_BYTES,
};

impl SpreadsheetEvaluator<'_> {
    pub(super) fn evaluate_function(
        &mut self,
        name: &str,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let normalized = normalize_function_name(name);
        if normalized == "IF" {
            validate_arity(&normalized, arguments.len(), 2, Some(3))?;
            return self.evaluate_if(arguments, current);
        }
        if normalized == "IFERROR" {
            validate_arity(&normalized, arguments.len(), 2, Some(2))?;
            return self.evaluate_if_error(arguments, current);
        }
        if normalized == "SUMIF" {
            validate_arity(&normalized, arguments.len(), 2, Some(3))?;
            return self.evaluate_sum_if(arguments, current);
        }
        let (minimum, maximum) = function_arity(&normalized).ok_or_else(|| {
            unsupported(format!("Formula function '{normalized}' is not supported."))
        })?;
        validate_arity(&normalized, arguments.len(), minimum, maximum)?;
        let values = self.evaluate_arguments(arguments, current)?;
        let result = match normalized.as_str() {
            "SUM" => aggregate(&values, Aggregate::Sum),
            "SUBTOTAL" => subtotal(&values),
            "AVERAGE" => aggregate(&values, Aggregate::Average),
            "MIN" => aggregate(&values, Aggregate::Minimum),
            "MAX" => aggregate(&values, Aggregate::Maximum),
            "COUNT" => count(&values),
            "COUNTA" => count_a(&values),
            "ABS" => unary_numeric(&values, f64::abs),
            "SQRT" => unary_numeric_value(&values, |value| {
                if value < 0.0 {
                    SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Number)
                } else {
                    finite_number(value.sqrt())
                }
            }),
            "POWER" => numeric_binary(&values, |left, right| finite_number(left.powf(right))),
            "MOD" => numeric_binary(&values, |left, right| {
                if right == 0.0 {
                    SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero)
                } else {
                    finite_number(left - right * (left / right).floor())
                }
            }),
            "ROUND" => numeric_binary(&values, round_number),
            "AND" => logical_aggregate(&values, true),
            "OR" => logical_aggregate(&values, false),
            "NOT" => logical_not(&values),
            "CONCAT" | "CONCATENATE" => concatenate(&values),
            "ROW" => row_or_column(current, true),
            "COLUMN" => row_or_column(current, false),
            "PI" => finite_number(std::f64::consts::PI),
            "NA" => SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::NotAvailable),
            "TRUE" => SpreadsheetValue::Boolean { value: true },
            "FALSE" => SpreadsheetValue::Boolean { value: false },
            _ => {
                return Err(unsupported(format!(
                    "Formula function '{normalized}' is not supported."
                )));
            }
        };
        Ok(EvaluatedValue::Scalar(result))
    }

    fn evaluate_arguments(
        &mut self,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<Vec<EvaluatedValue>, EvaluationFailure> {
        arguments
            .iter()
            .map(|argument| {
                argument.as_ref().map_or_else(
                    || Ok(EvaluatedValue::Scalar(SpreadsheetValue::Blank)),
                    |argument| self.evaluate_value(argument, current),
                )
            })
            .collect()
    }

    fn evaluate_if(
        &mut self,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let condition =
            self.evaluate_optional(arguments.first(), current, SpreadsheetValue::Blank)?;
        match scalar_boolean(condition.into_scalar()) {
            Ok(true) => self.evaluate_optional(arguments.get(1), current, SpreadsheetValue::Blank),
            Ok(false) => self.evaluate_optional(
                arguments.get(2),
                current,
                SpreadsheetValue::Boolean { value: false },
            ),
            Err(error) => Ok(EvaluatedValue::Scalar(SpreadsheetValue::error(error))),
        }
    }

    fn evaluate_if_error(
        &mut self,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let value = self.evaluate_optional(arguments.first(), current, SpreadsheetValue::Blank)?;
        if matches!(
            value,
            EvaluatedValue::Scalar(SpreadsheetValue::Error { .. })
        ) {
            self.evaluate_optional(arguments.get(1), current, SpreadsheetValue::Blank)
        } else {
            Ok(value)
        }
    }

    fn evaluate_sum_if(
        &mut self,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let criteria_range = sum_if_argument(arguments, 0)?;
        let criterion = sum_if_argument(arguments, 1)?;
        let criteria_area = self.sum_if_rectangle(criteria_range, current)?;
        ensure_sum_if_criterion(criterion)?;
        let criterion = parse_sum_if_criterion(self.evaluate_expression(criterion, current)?)?;
        let sum_anchor = match arguments.get(2).and_then(Option::as_ref) {
            Some(sum_range) => self.sum_if_rectangle(sum_range, current)?,
            None => criteria_area,
        };
        let row_span = criteria_area
            .end_row
            .checked_sub(criteria_area.start_row)
            .ok_or_else(|| unsupported("SUMIF criteria range is not a rectangle."))?;
        let column_span = criteria_area
            .end_column
            .checked_sub(criteria_area.start_column)
            .ok_or_else(|| unsupported("SUMIF criteria range is not a rectangle."))?;
        let rows = usize::try_from(row_span + 1)
            .map_err(|_| unsupported("SUMIF criteria range is outside supported limits."))?;
        let columns = usize::try_from(column_span + 1)
            .map_err(|_| unsupported("SUMIF criteria range is outside supported limits."))?;
        let cells = rows
            .checked_mul(columns)
            .ok_or_else(|| unsupported("SUMIF criteria range is outside supported limits."))?;
        if cells > MAX_SPREADSHEET_CELLS {
            return Err(unsupported(format!(
                "SUMIF supports at most {MAX_SPREADSHEET_CELLS} criteria cells."
            )));
        }
        let sum_end_row = sum_anchor
            .start_row
            .checked_add(row_span)
            .ok_or_else(|| unsupported("SUMIF sum range extends past the worksheet row limit."))?;
        let sum_end_column = sum_anchor
            .start_column
            .checked_add(column_span)
            .ok_or_else(|| {
                unsupported("SUMIF sum range extends past the worksheet column limit.")
            })?;
        if sum_end_row >= MAX_SPREADSHEET_ROWS || sum_end_column >= MAX_SPREADSHEET_COLUMNS {
            return Err(unsupported(
                "SUMIF sum range extends past the worksheet limits.",
            ));
        }

        let mut sum = 0.0_f64;
        for row_offset in 0..=row_span {
            for column_offset in 0..=column_span {
                let criteria_value = self.evaluate_dependency(&CellKey {
                    sheet: criteria_area.sheet,
                    row: criteria_area.start_row + row_offset,
                    column: criteria_area.start_column + column_offset,
                })?;
                if !criterion.matches(&criteria_value) {
                    continue;
                }
                let summed = self.evaluate_dependency(&CellKey {
                    sheet: sum_anchor.sheet,
                    row: sum_anchor.start_row + row_offset,
                    column: sum_anchor.start_column + column_offset,
                })?;
                match summed {
                    SpreadsheetValue::Number { value } => sum += value,
                    SpreadsheetValue::Error { .. } => {
                        return Ok(EvaluatedValue::Scalar(summed));
                    }
                    SpreadsheetValue::Blank
                    | SpreadsheetValue::Text { .. }
                    | SpreadsheetValue::Boolean { .. } => {}
                }
            }
        }
        Ok(EvaluatedValue::Scalar(finite_number(sum)))
    }

    fn sum_if_rectangle(
        &self,
        expression: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<SumIfRectangle, EvaluationFailure> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.sum_if_rectangle(inner, current)
            }
            SpreadsheetFormulaExpressionKind::Reference(_) => {
                let key = self.sum_if_endpoint(expression, current)?;
                Ok(SumIfRectangle {
                    sheet: key.sheet,
                    start_row: key.row,
                    start_column: key.column,
                    end_row: key.row,
                    end_column: key.column,
                })
            }
            SpreadsheetFormulaExpressionKind::Binary {
                operator: SpreadsheetFormulaBinaryOperator::Range,
                left,
                right,
            } => {
                let left = self.sum_if_endpoint(left, current)?;
                let right = self.sum_if_endpoint(right, current)?;
                if left.sheet != right.sheet {
                    return Err(unsupported(
                        "A SUMIF range cannot span multiple worksheets.",
                    ));
                }
                Ok(SumIfRectangle {
                    sheet: left.sheet,
                    start_row: left.row.min(right.row),
                    start_column: left.column.min(right.column),
                    end_row: left.row.max(right.row),
                    end_column: left.column.max(right.column),
                })
            }
            _ => Err(unsupported("SUMIF ranges must be one worksheet rectangle.")),
        }
    }

    fn sum_if_endpoint(
        &self,
        expression: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<CellKey, EvaluationFailure> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.sum_if_endpoint(inner, current)
            }
            SpreadsheetFormulaExpressionKind::Reference(reference) => {
                self.reference_key(reference, current.sheet)
            }
            _ => Err(unsupported("SUMIF ranges must be one worksheet rectangle.")),
        }
    }

    fn evaluate_optional(
        &mut self,
        argument: Option<&Option<SpreadsheetFormulaExpression>>,
        current: &CellKey,
        absent: SpreadsheetValue,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        match argument {
            Some(Some(argument)) => self.evaluate_value(argument, current),
            Some(None) => Ok(EvaluatedValue::Scalar(SpreadsheetValue::Blank)),
            None => Ok(EvaluatedValue::Scalar(absent)),
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum Aggregate {
    Sum,
    Average,
    Minimum,
    Maximum,
}

fn aggregate(values: &[EvaluatedValue], operation: Aggregate) -> SpreadsheetValue {
    let mut count = 0_usize;
    let mut sum = 0.0_f64;
    let mut minimum: Option<f64> = None;
    let mut maximum: Option<f64> = None;
    for argument in values {
        match argument {
            EvaluatedValue::Range(values) => {
                for value in values {
                    match value {
                        SpreadsheetValue::Number { value } => {
                            record_number(&mut count, &mut sum, &mut minimum, &mut maximum, *value);
                        }
                        SpreadsheetValue::Error { .. } => return value.clone(),
                        _ => {}
                    }
                }
            }
            EvaluatedValue::Scalar(value) => match scalar_number(value.clone()) {
                Ok(value) => record_number(&mut count, &mut sum, &mut minimum, &mut maximum, value),
                Err(error) => return SpreadsheetValue::error(error),
            },
        }
    }
    match operation {
        Aggregate::Sum => finite_number(sum),
        Aggregate::Average if count == 0 => {
            SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero)
        }
        Aggregate::Average => finite_number(sum / count as f64),
        Aggregate::Minimum => finite_number(minimum.unwrap_or(0.0)),
        Aggregate::Maximum => finite_number(maximum.unwrap_or(0.0)),
    }
}

/// Evaluate the bounded native `SUBTOTAL(function_num, ref)` family. Hidden
/// row metadata is not currently part of the kernel request, so the resolved
/// reference values are aggregated as-is while preserving Excel's function
/// number mapping.
fn subtotal(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let Some(code) = values
        .first()
        .cloned()
        .map(EvaluatedValue::into_scalar)
        .and_then(|value| scalar_number(value).ok())
        .filter(|value| value.is_finite())
        .map(|value| value.trunc() as i32)
    else {
        return value_error();
    };
    let arguments = &values[1..];
    let flattened = arguments
        .iter()
        .cloned()
        .flat_map(EvaluatedValue::into_values)
        .collect::<Vec<_>>();
    match code {
        1 | 101 => subtotal_numeric(&flattened, SubtotalOperation::Average),
        2 | 102 => subtotal_count(&flattened, false),
        3 | 103 => subtotal_count(&flattened, true),
        4 | 104 => subtotal_numeric(&flattened, SubtotalOperation::Maximum),
        5 | 105 => subtotal_numeric(&flattened, SubtotalOperation::Minimum),
        6 | 106 => subtotal_numeric(&flattened, SubtotalOperation::Product),
        7 | 107 => subtotal_numeric(&flattened, SubtotalOperation::StdDev),
        8 | 108 => subtotal_numeric(&flattened, SubtotalOperation::StdDevP),
        9 | 109 => subtotal_numeric(&flattened, SubtotalOperation::Sum),
        10 | 110 => subtotal_numeric(&flattened, SubtotalOperation::Var),
        11 | 111 => subtotal_numeric(&flattened, SubtotalOperation::VarP),
        _ => value_error(),
    }
}

#[derive(Debug, Clone, Copy)]
enum SubtotalOperation {
    Average,
    Maximum,
    Minimum,
    Product,
    StdDev,
    StdDevP,
    Sum,
    Var,
    VarP,
}

fn subtotal_numeric(values: &[SpreadsheetValue], operation: SubtotalOperation) -> SpreadsheetValue {
    let mut numbers = Vec::new();
    for value in values {
        match value {
            SpreadsheetValue::Number { value } => numbers.push(*value),
            SpreadsheetValue::Error { value } => {
                return SpreadsheetValue::Error {
                    value: value.clone(),
                };
            }
            SpreadsheetValue::Blank
            | SpreadsheetValue::Text { .. }
            | SpreadsheetValue::Boolean { .. } => {}
        }
    }
    let count = numbers.len();
    let sum = numbers.iter().sum::<f64>();
    let result = match operation {
        SubtotalOperation::Sum => sum,
        SubtotalOperation::Average => {
            if count == 0 {
                return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero);
            }
            sum / count as f64
        }
        SubtotalOperation::Maximum => numbers.iter().copied().reduce(f64::max).unwrap_or(0.0),
        SubtotalOperation::Minimum => numbers.iter().copied().reduce(f64::min).unwrap_or(0.0),
        SubtotalOperation::Product => {
            if count == 0 {
                0.0
            } else {
                numbers.iter().product()
            }
        }
        SubtotalOperation::StdDev | SubtotalOperation::Var if count < 2 => {
            return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero);
        }
        SubtotalOperation::StdDevP | SubtotalOperation::VarP if count == 0 => {
            return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero);
        }
        SubtotalOperation::StdDev | SubtotalOperation::StdDevP => {
            let mean = sum / count as f64;
            let divisor = if matches!(operation, SubtotalOperation::StdDev) {
                (count - 1) as f64
            } else {
                count as f64
            };
            (numbers
                .iter()
                .map(|number| (number - mean).powi(2))
                .sum::<f64>()
                / divisor)
                .sqrt()
        }
        SubtotalOperation::Var | SubtotalOperation::VarP => {
            let mean = sum / count as f64;
            let divisor = if matches!(operation, SubtotalOperation::Var) {
                (count - 1) as f64
            } else {
                count as f64
            };
            numbers
                .iter()
                .map(|number| (number - mean).powi(2))
                .sum::<f64>()
                / divisor
        }
    };
    finite_number(result)
}

fn subtotal_count(values: &[SpreadsheetValue], include_non_numeric: bool) -> SpreadsheetValue {
    let count = values
        .iter()
        .filter(|value| {
            if include_non_numeric {
                !matches!(value, SpreadsheetValue::Blank)
            } else {
                matches!(value, SpreadsheetValue::Number { .. })
            }
        })
        .count();
    finite_number(count as f64)
}

fn count(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let mut count = 0_usize;
    for argument in values {
        match argument {
            EvaluatedValue::Range(values) => {
                count += values
                    .iter()
                    .filter(|value| matches!(value, SpreadsheetValue::Number { .. }))
                    .count();
            }
            EvaluatedValue::Scalar(SpreadsheetValue::Error { value }) => {
                return SpreadsheetValue::Error {
                    value: value.clone(),
                };
            }
            EvaluatedValue::Scalar(value) => {
                if scalar_number(value.clone()).is_ok() {
                    count += 1;
                }
            }
        }
    }
    finite_number(count as f64)
}

fn count_a(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let count = values
        .iter()
        .map(|argument| match argument {
            EvaluatedValue::Scalar(SpreadsheetValue::Blank) => 0,
            EvaluatedValue::Scalar(_) => 1,
            EvaluatedValue::Range(values) => values
                .iter()
                .filter(|value| !matches!(value, SpreadsheetValue::Blank))
                .count(),
        })
        .sum::<usize>();
    finite_number(count as f64)
}

fn unary_numeric(values: &[EvaluatedValue], operation: fn(f64) -> f64) -> SpreadsheetValue {
    unary_numeric_value(values, |value| finite_number(operation(value)))
}

fn unary_numeric_value(
    values: &[EvaluatedValue],
    operation: impl Fn(f64) -> SpreadsheetValue,
) -> SpreadsheetValue {
    let Some(value) = values.first().cloned() else {
        return value_error();
    };
    match scalar_number(value.into_scalar()) {
        Ok(value) => operation(value),
        Err(error) => SpreadsheetValue::error(error),
    }
}

fn numeric_binary(
    values: &[EvaluatedValue],
    operation: impl Fn(f64, f64) -> SpreadsheetValue,
) -> SpreadsheetValue {
    let Some(left) = values.first().cloned() else {
        return value_error();
    };
    let Some(right) = values.get(1).cloned() else {
        return value_error();
    };
    match (
        scalar_number(left.into_scalar()),
        scalar_number(right.into_scalar()),
    ) {
        (Ok(left), Ok(right)) => operation(left, right),
        (Err(error), _) | (_, Err(error)) => SpreadsheetValue::error(error),
    }
}

fn round_number(number: f64, digits: f64) -> SpreadsheetValue {
    if !(-308.0..=308.0).contains(&digits) {
        return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Number);
    }
    let factor = 10_f64.powi(digits.trunc() as i32);
    if !factor.is_finite() || factor == 0.0 {
        return SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::Number);
    }
    finite_number((number * factor).round() / factor)
}

fn logical_aggregate(values: &[EvaluatedValue], and: bool) -> SpreadsheetValue {
    let mut observed = false;
    let mut result = and;
    for value in values.iter().cloned().flat_map(EvaluatedValue::into_values) {
        if matches!(
            value,
            SpreadsheetValue::Blank | SpreadsheetValue::Text { .. }
        ) {
            continue;
        }
        match scalar_boolean(value) {
            Ok(value) => {
                observed = true;
                if and {
                    result &= value;
                } else {
                    result |= value;
                }
            }
            Err(error) => return SpreadsheetValue::error(error),
        }
    }
    if observed {
        SpreadsheetValue::Boolean { value: result }
    } else {
        value_error()
    }
}

fn logical_not(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let Some(value) = values.first().cloned() else {
        return value_error();
    };
    match scalar_boolean(value.into_scalar()) {
        Ok(value) => SpreadsheetValue::Boolean { value: !value },
        Err(error) => SpreadsheetValue::error(error),
    }
}

fn concatenate(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let mut output = String::new();
    for value in values.iter().cloned().flat_map(EvaluatedValue::into_values) {
        let Ok(value) = scalar_text(value) else {
            return value_error();
        };
        let Some(length) = output.len().checked_add(value.len()) else {
            return value_error();
        };
        if length > MAX_SPREADSHEET_TEXT_BYTES {
            return value_error();
        }
        output.push_str(&value);
    }
    SpreadsheetValue::Text { value: output }
}

fn row_or_column(current: &CellKey, row: bool) -> SpreadsheetValue {
    let value = if row {
        current.row + 1
    } else {
        current.column + 1
    };
    finite_number(f64::from(value))
}

fn record_number(
    count: &mut usize,
    sum: &mut f64,
    minimum: &mut Option<f64>,
    maximum: &mut Option<f64>,
    value: f64,
) {
    *count += 1;
    *sum += value;
    *minimum = Some(minimum.map_or(value, |current| current.min(value)));
    *maximum = Some(maximum.map_or(value, |current| current.max(value)));
}

fn function_arity(name: &str) -> Option<(usize, Option<usize>)> {
    Some(match name {
        "SUM" | "AVERAGE" | "MIN" | "MAX" | "COUNT" | "COUNTA" | "AND" | "OR" | "CONCAT"
        | "CONCATENATE" => (1, Some(255)),
        "SUBTOTAL" => (2, Some(255)),
        "SUMIF" => (2, Some(3)),
        "ABS" | "SQRT" | "NOT" => (1, Some(1)),
        "POWER" | "MOD" | "ROUND" => (2, Some(2)),
        "ROW" | "COLUMN" => (0, Some(0)),
        "FALSE" | "PI" | "NA" | "TRUE" => (0, Some(0)),
        _ => return None,
    })
}

fn validate_arity(
    name: &str,
    received: usize,
    minimum: usize,
    maximum: Option<usize>,
) -> Result<(), EvaluationFailure> {
    if received < minimum || maximum.is_some_and(|maximum| received > maximum) {
        return Err(unsupported(format!(
            "Formula function '{name}' received {received} arguments outside its supported arity."
        )));
    }
    Ok(())
}

fn normalize_function_name(name: &str) -> String {
    let mut normalized = name.to_ascii_uppercase();
    loop {
        let stripped = ["_XLFN.", "_XLWS."]
            .into_iter()
            .find_map(|prefix| normalized.strip_prefix(prefix).map(ToOwned::to_owned));
        let Some(stripped) = stripped else {
            return normalized;
        };
        normalized = stripped;
    }
}

#[derive(Debug, Clone, Copy)]
struct SumIfRectangle {
    sheet: usize,
    start_row: u32,
    start_column: u32,
    end_row: u32,
    end_column: u32,
}

fn sum_if_argument(
    arguments: &[Option<SpreadsheetFormulaExpression>],
    index: usize,
) -> Result<&SpreadsheetFormulaExpression, EvaluationFailure> {
    arguments
        .get(index)
        .and_then(Option::as_ref)
        .ok_or_else(|| unsupported("SUMIF requires a criteria range and a criterion."))
}

fn ensure_sum_if_criterion(
    expression: &SpreadsheetFormulaExpression,
) -> Result<(), EvaluationFailure> {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Parenthesized(inner) => ensure_sum_if_criterion(inner),
        SpreadsheetFormulaExpressionKind::Binary {
            operator: SpreadsheetFormulaBinaryOperator::Range,
            ..
        }
        | SpreadsheetFormulaExpressionKind::StructuredReference { .. }
        | SpreadsheetFormulaExpressionKind::Name { .. }
        | SpreadsheetFormulaExpressionKind::Array { .. } => Err(unsupported(
            "SUMIF criteria must be one value, not a range.",
        )),
        _ => Ok(()),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SumIfCompare {
    Equal,
    NotEqual,
    Greater,
    GreaterOrEqual,
    Less,
    LessOrEqual,
}

enum SumIfCriterion {
    Number(SumIfCompare, f64),
    Text(SumIfCompare, String),
    Boolean(bool),
    Blank,
}

impl SumIfCriterion {
    fn matches(&self, value: &SpreadsheetValue) -> bool {
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

fn parse_sum_if_criterion(value: SpreadsheetValue) -> Result<SumIfCriterion, EvaluationFailure> {
    match value {
        SpreadsheetValue::Error { value } => Err(unsupported(format!(
            "SUMIF criteria cannot be the error {value}."
        ))),
        SpreadsheetValue::Blank => Ok(SumIfCriterion::Blank),
        SpreadsheetValue::Boolean { value } => Ok(SumIfCriterion::Boolean(value)),
        SpreadsheetValue::Number { value } => {
            Ok(SumIfCriterion::Number(SumIfCompare::Equal, value))
        }
        SpreadsheetValue::Text { value } => parse_sum_if_text(&value),
    }
}

fn parse_sum_if_text(value: &str) -> Result<SumIfCriterion, EvaluationFailure> {
    let (compare, operand) = split_sum_if_compare(value);
    if sum_if_contains_wildcard(operand) {
        return Err(unsupported(
            "SUMIF wildcard criteria are not calculated. Pass an exact value or a numeric comparison.",
        ));
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
