use a3s_use_core::UseResult;

use crate::spreadsheet_formula::SpreadsheetFormulaExpression;
use crate::SpreadsheetFormulaErrorLiteral;

use super::super::{
    accumulate_formula_text_bytes, broadcast_dimension, checked_array_cells, into_array,
    invalid_array_shape, scalar_boolean, EvalValue, EvaluationContext, FormulaArray,
    FormulaCellKey, ScalarValue,
};
use super::array::collapse_array;

pub(super) fn evaluate_if(
    context: &mut EvaluationContext<'_>,
    arguments: &[Option<SpreadsheetFormulaExpression>],
    current: FormulaCellKey,
) -> UseResult<EvalValue> {
    let condition =
        evaluate_optional_argument(context, arguments.first(), current, ScalarValue::Blank)?;
    let condition = context.materialize(condition)?;
    if let EvalValue::Scalar(condition) = condition {
        return match scalar_boolean(condition) {
            Ok(true) => {
                evaluate_optional_argument(context, arguments.get(1), current, ScalarValue::Blank)
            }
            Ok(false) => evaluate_optional_argument(
                context,
                arguments.get(2),
                current,
                ScalarValue::Boolean(false),
            ),
            Err(error) => Ok(EvalValue::Scalar(ScalarValue::Error(error))),
        };
    }
    let true_value =
        evaluate_optional_argument(context, arguments.get(1), current, ScalarValue::Blank)?;
    let false_value = evaluate_optional_argument(
        context,
        arguments.get(2),
        current,
        ScalarValue::Boolean(false),
    )?;
    select_array(context, condition, true_value, false_value)
}

pub(super) fn evaluate_if_error(
    context: &mut EvaluationContext<'_>,
    arguments: &[Option<SpreadsheetFormulaExpression>],
    current: FormulaCellKey,
) -> UseResult<EvalValue> {
    let value =
        evaluate_optional_argument(context, arguments.first(), current, ScalarValue::Blank)?;
    let value = context.materialize(value)?;
    if let EvalValue::Scalar(ScalarValue::Error(_)) = value {
        return evaluate_optional_argument(context, arguments.get(1), current, ScalarValue::Blank);
    }
    let EvalValue::Array(array) = &value else {
        return Ok(value);
    };
    if !array
        .rows
        .iter()
        .flatten()
        .any(|value| matches!(value, ScalarValue::Error(_)))
    {
        return Ok(value);
    }
    let fallback =
        evaluate_optional_argument(context, arguments.get(1), current, ScalarValue::Blank)?;
    let fallback = into_array(context.materialize(fallback)?)?;
    let height = broadcast_dimension(array.height(), fallback.height());
    let width = broadcast_dimension(array.width(), fallback.width());
    let (Some(height), Some(width)) = (height, width) else {
        return Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::Value,
        )));
    };
    checked_array_cells(height, width)?;
    let mut rows = Vec::with_capacity(height);
    let mut text_bytes = 0_usize;
    for row in 0..height {
        let mut values = Vec::with_capacity(width);
        for column in 0..width {
            let original = array
                .broadcast_value(row, column)
                .cloned()
                .ok_or_else(invalid_array_shape)?;
            let value = if matches!(original, ScalarValue::Error(_)) {
                fallback
                    .broadcast_value(row, column)
                    .cloned()
                    .ok_or_else(invalid_array_shape)?
            } else {
                original
            };
            accumulate_formula_text_bytes(&mut text_bytes, &value)?;
            values.push(value);
        }
        rows.push(values);
    }
    collapse_array(FormulaArray { rows })
}

fn evaluate_optional_argument(
    context: &mut EvaluationContext<'_>,
    argument: Option<&Option<SpreadsheetFormulaExpression>>,
    current: FormulaCellKey,
    absent: ScalarValue,
) -> UseResult<EvalValue> {
    match argument {
        Some(Some(argument)) => context.evaluate_expression(argument, current),
        Some(None) => Ok(EvalValue::Scalar(ScalarValue::Blank)),
        None => Ok(EvalValue::Scalar(absent)),
    }
}

fn select_array(
    context: &EvaluationContext<'_>,
    condition: EvalValue,
    true_value: EvalValue,
    false_value: EvalValue,
) -> UseResult<EvalValue> {
    let condition = into_array(condition)?;
    let true_value = into_array(context.materialize(true_value)?)?;
    let false_value = into_array(context.materialize(false_value)?)?;
    let height = [
        condition.height(),
        true_value.height(),
        false_value.height(),
    ]
    .into_iter()
    .reduce(|left, right| broadcast_dimension(left, right).unwrap_or(usize::MAX))
    .filter(|value| *value != usize::MAX);
    let width = [condition.width(), true_value.width(), false_value.width()]
        .into_iter()
        .reduce(|left, right| broadcast_dimension(left, right).unwrap_or(usize::MAX))
        .filter(|value| *value != usize::MAX);
    let (Some(height), Some(width)) = (height, width) else {
        return Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::Value,
        )));
    };
    checked_array_cells(height, width)?;
    let mut rows = Vec::with_capacity(height);
    let mut text_bytes = 0_usize;
    for row in 0..height {
        let mut values = Vec::with_capacity(width);
        for column in 0..width {
            let condition = condition
                .broadcast_value(row, column)
                .cloned()
                .ok_or_else(invalid_array_shape)?;
            let value = match scalar_boolean(condition) {
                Ok(true) => true_value
                    .broadcast_value(row, column)
                    .cloned()
                    .ok_or_else(invalid_array_shape)?,
                Ok(false) => false_value
                    .broadcast_value(row, column)
                    .cloned()
                    .ok_or_else(invalid_array_shape)?,
                Err(error) => ScalarValue::Error(error),
            };
            accumulate_formula_text_bytes(&mut text_bytes, &value)?;
            values.push(value);
        }
        rows.push(values);
    }
    collapse_array(FormulaArray { rows })
}
