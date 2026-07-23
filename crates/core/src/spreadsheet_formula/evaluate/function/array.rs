use a3s_use_core::UseResult;

use crate::SpreadsheetFormulaErrorLiteral;

use super::super::{
    calculation_error, finite_or_number_error, into_array, invalid_array_shape, scalar_number,
    EvalValue, EvaluationContext, FormulaArray, FormulaCellKey, ScalarValue,
    MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
};
use super::{argument, missing_argument};

pub(super) fn row_or_column(
    _context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
    current: FormulaCellKey,
    row: bool,
) -> UseResult<EvalValue> {
    if arguments.is_empty() {
        return Ok(EvalValue::Scalar(ScalarValue::Number(f64::from(if row {
            current.row
        } else {
            current.column
        }))));
    }
    let EvalValue::Reference(areas) = arguments.first().ok_or_else(missing_argument)? else {
        return Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::Value,
        )));
    };
    let mut output = Vec::new();
    for area in areas {
        let cells = area.cell_count().ok_or_else(function_array_limit)?;
        if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS
            || output.len().saturating_add(cells) > MAX_SPREADSHEET_FORMULA_SPILL_CELLS
        {
            return Err(function_array_limit());
        }
        for source_row in area.start_row..=area.end_row {
            let mut values = Vec::new();
            for source_column in area.start_column..=area.end_column {
                values.push(ScalarValue::Number(f64::from(if row {
                    source_row
                } else {
                    source_column
                })));
            }
            output.push(values);
        }
    }
    collapse_array(FormulaArray::new(output).ok_or_else(function_array_limit)?)
}

pub(super) fn sequence(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let rows = positive_dimension(context.require_scalar(argument(arguments, 0)?)?);
    let columns = if let Some(value) = arguments.get(1) {
        positive_dimension(context.require_scalar(value.clone())?)
    } else {
        Ok(1)
    };
    let start = if let Some(value) = arguments.get(2) {
        scalar_number(context.require_scalar(value.clone())?)
    } else {
        Ok(1.0)
    };
    let step = if let Some(value) = arguments.get(3) {
        scalar_number(context.require_scalar(value.clone())?)
    } else {
        Ok(1.0)
    };
    let (Ok(rows), Ok(columns), Ok(start), Ok(step)) = (rows, columns, start, step) else {
        return Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::Value,
        )));
    };
    let cells = rows.checked_mul(columns).ok_or_else(function_array_limit)?;
    if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
        return Err(function_array_limit().with_detail("cells", cells));
    }
    let mut output = Vec::with_capacity(rows);
    for row in 0..rows {
        let mut values = Vec::with_capacity(columns);
        for column in 0..columns {
            let offset = row
                .checked_mul(columns)
                .and_then(|value| value.checked_add(column))
                .ok_or_else(function_array_limit)?;
            values.push(finite_or_number_error(start + step * offset as f64));
        }
        output.push(values);
    }
    Ok(EvalValue::Array(FormulaArray { rows: output }))
}

pub(super) fn transpose(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
) -> UseResult<EvalValue> {
    let array = into_array(context.materialize(argument(arguments, 0)?)?)?;
    let mut output = vec![vec![ScalarValue::Blank; array.height()]; array.width()];
    for (row, values) in array.rows.into_iter().enumerate() {
        for (column, value) in values.into_iter().enumerate() {
            let target = output
                .get_mut(column)
                .and_then(|values| values.get_mut(row))
                .ok_or_else(invalid_array_shape)?;
            *target = value;
        }
    }
    collapse_array(FormulaArray { rows: output })
}

pub(super) fn collapse_array(array: FormulaArray) -> UseResult<EvalValue> {
    if array.height() == 1 && array.width() == 1 {
        Ok(EvalValue::Scalar(
            array
                .rows
                .into_iter()
                .next()
                .and_then(|row| row.into_iter().next())
                .unwrap_or(ScalarValue::Blank),
        ))
    } else {
        Ok(EvalValue::Array(array))
    }
}

fn positive_dimension(value: ScalarValue) -> Result<usize, SpreadsheetFormulaErrorLiteral> {
    let value = scalar_number(value)?;
    if !value.is_finite() || value < 1.0 || value.fract() != 0.0 {
        return Err(SpreadsheetFormulaErrorLiteral::Value);
    }
    usize::try_from(value as u64).map_err(|_| SpreadsheetFormulaErrorLiteral::Number)
}

fn function_array_limit() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_function_array_limit",
        format!(
            "Native formula functions return at most {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} array cells."
        ),
    )
}
