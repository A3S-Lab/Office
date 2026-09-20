use a3s_use_core::UseResult;

use crate::spreadsheet_formula::evaluate::MAX_SPREADSHEET_FORMULA_SPILL_CELLS;

use super::super::{
    calculation_error, finite_or_number_error, EvalValue, EvaluationContext, FormulaCellKey,
};
use super::criterion::{cell_value, criterion_scalar, one_rectangle, parse_criterion, COUNT_IF};

/// `COUNTIF(range, criteria)` counts cells in one rectangle.
///
/// Comparison and fail-closed wildcard rules match `SUMIF`. There is no
/// second range to realign.
pub(super) fn count_if(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let criteria_area = one_rectangle(argument(arguments, 0)?, "COUNTIF criteria range", COUNT_IF)?;
    let criterion = parse_criterion(
        COUNT_IF,
        criterion_scalar(context, argument(arguments, 1)?, COUNT_IF)?,
    )?;
    let height = criteria_area.end_row - criteria_area.start_row;
    let width = criteria_area.end_column - criteria_area.start_column;
    let cells = criteria_area
        .cell_count()
        .ok_or_else(count_if_limit_error)?;
    if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
        return Err(count_if_limit_error().with_detail("cells", cells));
    }

    let mut count = 0.0_f64;
    for row_offset in 0..=height {
        for column_offset in 0..=width {
            let key = FormulaCellKey {
                sheet: criteria_area.sheet,
                row: criteria_area.start_row + row_offset,
                column: criteria_area.start_column + column_offset,
            };
            if criterion.matches(&cell_value(context, key)) {
                count += 1.0;
            }
        }
    }
    Ok(EvalValue::Scalar(finite_or_number_error(count)))
}

fn argument(arguments: &[EvalValue], index: usize) -> UseResult<&EvalValue> {
    arguments.get(index).ok_or_else(|| {
        calculation_error(
            "use.office.spreadsheet_formula_function_arity",
            "COUNTIF requires a criteria range and a criterion.",
        )
        .with_suggestion(
            "Match the closed-registry argument count for this function. Recalculate in-process; do not pad, omit, or evaluate arguments outside the native engine.",
        )
    })
}

fn count_if_limit_error() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_spill_limit",
        format!(
            "COUNTIF supports at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} criteria cells within worksheet limits."
        ),
    )
}
