use a3s_use_core::UseResult;

use crate::spreadsheet_formula::evaluate::MAX_SPREADSHEET_FORMULA_SPILL_CELLS;
use crate::spreadsheet_reference::{MAX_COLUMNS, MAX_ROWS};

use super::super::{
    calculation_error, finite_or_number_error, EvalValue, EvaluationContext, FormulaCellKey,
    ScalarValue,
};
use super::criterion::{cell_value, criterion_scalar, one_rectangle, parse_criterion, SUM_IF};

/// `SUMIF(range, criteria, [sum_range])`.
///
/// The summed block has the criteria range's shape and starts at the sum
/// range's top-left cell. Omitting `sum_range` sums the criteria range.
/// Wildcard criteria fail closed; they are not treated as literal text.
pub(super) fn sum_if(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let criteria_area = one_rectangle(argument(arguments, 0)?, "SUMIF criteria range", SUM_IF)?;
    let criterion = parse_criterion(
        SUM_IF,
        criterion_scalar(context, argument(arguments, 1)?, SUM_IF)?,
    )?;
    let sum_anchor = match arguments.get(2) {
        Some(value) => one_rectangle(value, "SUMIF sum range", SUM_IF)?,
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
        .with_suggestion(
            "Match the closed-registry argument count for this function. Recalculate in-process; do not pad, omit, or evaluate arguments outside the native engine.",
        )
    })
}

fn sum_if_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_spill_limit",
        format!(
            "SUMIF supports at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} criteria cells within worksheet limits."
        ),
    )
}
