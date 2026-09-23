use a3s_office_formula_parser::{
    SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaExpression,
    SpreadsheetFormulaExpressionKind, SpreadsheetFormulaReference, SpreadsheetFormulaReferenceKind,
};

use super::conditional::{
    conditional_argument, ensure_conditional_criterion, parse_conditional_criterion, SumIfRectangle,
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
        if normalized == "COUNTIF" {
            validate_arity(&normalized, arguments.len(), 2, Some(2))?;
            return self.evaluate_count_if(arguments, current);
        }
        if normalized == "AVERAGEIF" {
            validate_arity(&normalized, arguments.len(), 2, Some(3))?;
            return self.evaluate_average_if(arguments, current);
        }
        if normalized == "ROW" || normalized == "COLUMN" {
            validate_arity(&normalized, arguments.len(), 0, Some(1))?;
            return self.evaluate_row_or_column(normalized == "ROW", arguments, current);
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
            "LEFT" => text_left(&values),
            "RIGHT" => text_right(&values),
            "LEN" => text_len(&values),
            "MID" => text_mid(&values),
            "ROW" => row_or_column(current, true),
            "COLUMN" => row_or_column(current, false),
            "PI" => finite_number(std::f64::consts::PI),
            "NA" => SpreadsheetValue::error(SpreadsheetFormulaErrorLiteral::NotAvailable),
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
        let criteria_range = conditional_argument("SUMIF", arguments, 0)?;
        let criterion = conditional_argument("SUMIF", arguments, 1)?;
        let criteria_area = self.conditional_rectangle("SUMIF", criteria_range, current)?;
        ensure_conditional_criterion("SUMIF", criterion)?;
        let criterion =
            parse_conditional_criterion("SUMIF", self.evaluate_expression(criterion, current)?)?;
        let sum_anchor = match arguments.get(2).and_then(Option::as_ref) {
            Some(sum_range) => self.conditional_rectangle("SUMIF", sum_range, current)?,
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

    fn evaluate_count_if(
        &mut self,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let criteria_range = conditional_argument("COUNTIF", arguments, 0)?;
        let criterion = conditional_argument("COUNTIF", arguments, 1)?;
        let criteria_area = self.conditional_rectangle("COUNTIF", criteria_range, current)?;
        ensure_conditional_criterion("COUNTIF", criterion)?;
        let criterion =
            parse_conditional_criterion("COUNTIF", self.evaluate_expression(criterion, current)?)?;
        let row_span = criteria_area
            .end_row
            .checked_sub(criteria_area.start_row)
            .ok_or_else(|| unsupported("COUNTIF criteria range is not a rectangle."))?;
        let column_span = criteria_area
            .end_column
            .checked_sub(criteria_area.start_column)
            .ok_or_else(|| unsupported("COUNTIF criteria range is not a rectangle."))?;
        let rows = usize::try_from(row_span + 1)
            .map_err(|_| unsupported("COUNTIF criteria range is outside supported limits."))?;
        let columns = usize::try_from(column_span + 1)
            .map_err(|_| unsupported("COUNTIF criteria range is outside supported limits."))?;
        let cells = rows
            .checked_mul(columns)
            .ok_or_else(|| unsupported("COUNTIF criteria range is outside supported limits."))?;
        if cells > MAX_SPREADSHEET_CELLS {
            return Err(unsupported(format!(
                "COUNTIF supports at most {MAX_SPREADSHEET_CELLS} criteria cells."
            )));
        }

        let mut count = 0.0_f64;
        for row_offset in 0..=row_span {
            for column_offset in 0..=column_span {
                let value = self.evaluate_dependency(&CellKey {
                    sheet: criteria_area.sheet,
                    row: criteria_area.start_row + row_offset,
                    column: criteria_area.start_column + column_offset,
                })?;
                if criterion.matches(&value) {
                    count += 1.0;
                }
            }
        }
        Ok(EvaluatedValue::Scalar(finite_number(count)))
    }

    fn evaluate_average_if(
        &mut self,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let criteria_range = conditional_argument("AVERAGEIF", arguments, 0)?;
        let criterion = conditional_argument("AVERAGEIF", arguments, 1)?;
        let criteria_area = self.conditional_rectangle("AVERAGEIF", criteria_range, current)?;
        ensure_conditional_criterion("AVERAGEIF", criterion)?;
        let criterion = parse_conditional_criterion(
            "AVERAGEIF",
            self.evaluate_expression(criterion, current)?,
        )?;
        let average_anchor = match arguments.get(2).and_then(Option::as_ref) {
            Some(average_range) => {
                self.conditional_rectangle("AVERAGEIF", average_range, current)?
            }
            None => criteria_area,
        };
        let row_span = criteria_area
            .end_row
            .checked_sub(criteria_area.start_row)
            .ok_or_else(|| unsupported("AVERAGEIF criteria range is not a rectangle."))?;
        let column_span = criteria_area
            .end_column
            .checked_sub(criteria_area.start_column)
            .ok_or_else(|| unsupported("AVERAGEIF criteria range is not a rectangle."))?;
        let rows = usize::try_from(row_span + 1)
            .map_err(|_| unsupported("AVERAGEIF criteria range is outside supported limits."))?;
        let columns = usize::try_from(column_span + 1)
            .map_err(|_| unsupported("AVERAGEIF criteria range is outside supported limits."))?;
        let cells = rows
            .checked_mul(columns)
            .ok_or_else(|| unsupported("AVERAGEIF criteria range is outside supported limits."))?;
        if cells > MAX_SPREADSHEET_CELLS {
            return Err(unsupported(format!(
                "AVERAGEIF supports at most {MAX_SPREADSHEET_CELLS} criteria cells."
            )));
        }
        let average_end_row = average_anchor
            .start_row
            .checked_add(row_span)
            .ok_or_else(|| {
                unsupported("AVERAGEIF average range extends past the worksheet row limit.")
            })?;
        let average_end_column = average_anchor
            .start_column
            .checked_add(column_span)
            .ok_or_else(|| {
                unsupported("AVERAGEIF average range extends past the worksheet column limit.")
            })?;
        if average_end_row >= MAX_SPREADSHEET_ROWS || average_end_column >= MAX_SPREADSHEET_COLUMNS
        {
            return Err(unsupported(
                "AVERAGEIF average range extends past the worksheet limits.",
            ));
        }

        let mut sum = 0.0_f64;
        let mut count = 0.0_f64;
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
                let averaged = self.evaluate_dependency(&CellKey {
                    sheet: average_anchor.sheet,
                    row: average_anchor.start_row + row_offset,
                    column: average_anchor.start_column + column_offset,
                })?;
                match averaged {
                    SpreadsheetValue::Number { value } => {
                        sum += value;
                        count += 1.0;
                    }
                    SpreadsheetValue::Error { .. } => {
                        return Ok(EvaluatedValue::Scalar(averaged));
                    }
                    SpreadsheetValue::Blank
                    | SpreadsheetValue::Text { .. }
                    | SpreadsheetValue::Boolean { .. } => {}
                }
            }
        }
        if count == 0.0 {
            return Ok(EvaluatedValue::Scalar(SpreadsheetValue::error(
                SpreadsheetFormulaErrorLiteral::DivisionByZero,
            )));
        }
        Ok(EvaluatedValue::Scalar(finite_number(sum / count)))
    }

    fn conditional_rectangle(
        &self,
        name: &str,
        expression: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<SumIfRectangle, EvaluationFailure> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.conditional_rectangle(name, inner, current)
            }
            SpreadsheetFormulaExpressionKind::Reference(_) => {
                let key = self.conditional_endpoint(name, expression, current)?;
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
                let left = self.conditional_endpoint(name, left, current)?;
                let right = self.conditional_endpoint(name, right, current)?;
                if left.sheet != right.sheet {
                    return Err(unsupported(format!(
                        "A {name} range cannot span multiple worksheets."
                    )));
                }
                Ok(SumIfRectangle {
                    sheet: left.sheet,
                    start_row: left.row.min(right.row),
                    start_column: left.column.min(right.column),
                    end_row: left.row.max(right.row),
                    end_column: left.column.max(right.column),
                })
            }
            _ => Err(unsupported(format!(
                "{name} ranges must be one worksheet rectangle."
            ))),
        }
    }

    fn conditional_endpoint(
        &self,
        name: &str,
        expression: &SpreadsheetFormulaExpression,
        current: &CellKey,
    ) -> Result<CellKey, EvaluationFailure> {
        match &expression.kind {
            SpreadsheetFormulaExpressionKind::Parenthesized(inner) => {
                self.conditional_endpoint(name, inner, current)
            }
            SpreadsheetFormulaExpressionKind::Reference(reference) => {
                self.reference_key(reference, current.sheet)
            }
            _ => Err(unsupported(format!(
                "{name} ranges must be one worksheet rectangle."
            ))),
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

    fn evaluate_row_or_column(
        &self,
        row: bool,
        arguments: &[Option<SpreadsheetFormulaExpression>],
        current: &CellKey,
    ) -> Result<EvaluatedValue, EvaluationFailure> {
        let name = if row { "ROW" } else { "COLUMN" };
        let Some(argument) = arguments.first().and_then(Option::as_ref) else {
            return Ok(EvaluatedValue::Scalar(row_or_column(current, row)));
        };
        let Some(reference) = single_cell_reference(argument) else {
            return Err(unsupported(format!(
                "Formula function '{name}' in the browser kernel accepts one single cell, not a range."
            )));
        };
        let key = self.reference_key(reference, current.sheet)?;
        let value = if row { key.row + 1 } else { key.column + 1 };
        Ok(EvaluatedValue::Scalar(finite_number(f64::from(value))))
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

fn single_cell_reference(
    expression: &SpreadsheetFormulaExpression,
) -> Option<&SpreadsheetFormulaReference> {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Parenthesized(inner) => single_cell_reference(inner),
        SpreadsheetFormulaExpressionKind::Reference(reference)
            if matches!(reference.kind, SpreadsheetFormulaReferenceKind::Cell { .. }) =>
        {
            Some(reference)
        }
        _ => None,
    }
}

fn text_left(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let Ok(text) = first_scalar_text(values) else {
        return value_error();
    };
    let count = optional_text_count(values, 1).unwrap_or(1.0);
    if !count.is_finite() || count < 0.0 {
        return value_error();
    }
    let take = count.floor() as usize;
    SpreadsheetValue::Text {
        value: text.chars().take(take).collect(),
    }
}

fn text_right(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let Ok(text) = first_scalar_text(values) else {
        return value_error();
    };
    let count = optional_text_count(values, 1).unwrap_or(1.0);
    if !count.is_finite() || count < 0.0 {
        return value_error();
    }
    let take = count.floor() as usize;
    let total = text.chars().count();
    let skip = total.saturating_sub(take);
    SpreadsheetValue::Text {
        value: text.chars().skip(skip).collect(),
    }
}

fn text_len(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let Ok(text) = first_scalar_text(values) else {
        return value_error();
    };
    finite_number(text.chars().count() as f64)
}

fn text_mid(values: &[EvaluatedValue]) -> SpreadsheetValue {
    let Ok(text) = first_scalar_text(values) else {
        return value_error();
    };
    let Some(start) = optional_text_count(values, 1) else {
        return value_error();
    };
    let Some(count) = optional_text_count(values, 2) else {
        return value_error();
    };
    if !start.is_finite() || start < 1.0 || !count.is_finite() || count < 0.0 {
        return value_error();
    }
    let start_index = start.floor() as usize - 1;
    let take = count.floor() as usize;
    SpreadsheetValue::Text {
        value: text.chars().skip(start_index).take(take).collect(),
    }
}

fn first_scalar_text(values: &[EvaluatedValue]) -> Result<String, ()> {
    let Some(value) = values.first().cloned().and_then(|value| {
        value.into_values().into_iter().next()
    }) else {
        return Err(());
    };
    scalar_text(value).map_err(|_| ())
}

fn optional_text_count(values: &[EvaluatedValue], index: usize) -> Option<f64> {
    let value = values.get(index)?.clone().into_values().into_iter().next()?;
    scalar_number(value).ok()
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
        "COUNTIF" => (2, Some(2)),
        "AVERAGEIF" => (2, Some(3)),
        "ABS" | "SQRT" | "NOT" | "LEN" => (1, Some(1)),
        "LEFT" | "RIGHT" => (1, Some(2)),
        "MID" => (3, Some(3)),
        "ROW" | "COLUMN" => (0, Some(0)),
        "POWER" | "MOD" | "ROUND" => (2, Some(2)),
        "PI" | "NA" => (0, Some(0)),
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
