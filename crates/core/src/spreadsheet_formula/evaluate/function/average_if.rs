use a3s_use_core::UseResult;

use crate::spreadsheet_formula::evaluate::MAX_SPREADSHEET_FORMULA_SPILL_CELLS;
use crate::spreadsheet_reference::{MAX_COLUMNS, MAX_ROWS};
use crate::SpreadsheetFormulaErrorLiteral;

use super::super::{
    calculation_error, finite_or_number_error, EvalValue, EvaluationContext, FormulaCellKey,
    ScalarValue,
};
use super::criterion::{cell_value, criterion_scalar, one_rectangle, parse_criterion, AVERAGE_IF};

/// `AVERAGEIF(range, criteria, [average_range])`.
///
/// The averaged block has the criteria range's shape and starts at the
/// average range's top-left cell. Omitting `average_range` averages numeric
/// cells in the criteria range. Text and blanks are ignored. No numeric match
/// is `#DIV/0!`. Wildcard criteria fail closed.
pub(super) fn average_if(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let criteria_area = one_rectangle(
        argument(arguments, 0)?,
        "AVERAGEIF criteria range",
        AVERAGE_IF,
    )?;
    let criterion = parse_criterion(
        AVERAGE_IF,
        criterion_scalar(context, argument(arguments, 1)?, AVERAGE_IF)?,
    )?;
    let average_anchor = match arguments.get(2) {
        Some(value) => one_rectangle(value, "AVERAGEIF average range", AVERAGE_IF)?,
        None => criteria_area,
    };
    let height = criteria_area.end_row - criteria_area.start_row;
    let width = criteria_area.end_column - criteria_area.start_column;
    let cells = criteria_area
        .cell_count()
        .ok_or_else(average_if_limit_error)?;
    if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
        return Err(average_if_limit_error().with_detail("cells", cells));
    }
    let _average_end_row = average_anchor
        .start_row
        .checked_add(height)
        .filter(|row| *row <= MAX_ROWS)
        .ok_or_else(average_if_limit_error)?;
    let _average_end_column = average_anchor
        .start_column
        .checked_add(width)
        .filter(|column| *column <= MAX_COLUMNS)
        .ok_or_else(average_if_limit_error)?;

    let mut sum = 0.0_f64;
    let mut count = 0.0_f64;
    for row_offset in 0..=height {
        for column_offset in 0..=width {
            let criteria_key = FormulaCellKey {
                sheet: criteria_area.sheet,
                row: criteria_area.start_row + row_offset,
                column: criteria_area.start_column + column_offset,
            };
            if !criterion.matches(&cell_value(context, criteria_key)) {
                continue;
            }
            let average_key = FormulaCellKey {
                sheet: average_anchor.sheet,
                row: average_anchor.start_row + row_offset,
                column: average_anchor.start_column + column_offset,
            };
            match cell_value(context, average_key) {
                ScalarValue::Number(value) => {
                    sum += value;
                    count += 1.0;
                }
                ScalarValue::Error(error) => {
                    return Ok(EvalValue::Scalar(ScalarValue::Error(error)));
                }
                ScalarValue::Blank | ScalarValue::Text(_) | ScalarValue::Boolean(_) => {}
            }
        }
    }
    if count == 0.0 {
        return Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::DivisionByZero,
        )));
    }
    Ok(EvalValue::Scalar(finite_or_number_error(sum / count)))
}

fn argument(arguments: &[EvalValue], index: usize) -> UseResult<&EvalValue> {
    arguments.get(index).ok_or_else(|| {
        calculation_error(
            "use.office.spreadsheet_formula_function_arity",
            "AVERAGEIF requires a criteria range and a criterion.",
        )
        .with_suggestion(
            "Match the closed-registry argument count for this function. Recalculate in-process; do not pad, omit, or evaluate arguments outside the native engine.",
        )
    })
}

fn average_if_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_spill_limit",
        format!(
            "AVERAGEIF supports at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} criteria cells within worksheet limits."
        ),
    )
}
