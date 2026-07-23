mod aggregate;
mod array;
mod lazy;

use a3s_use_core::UseResult;

use crate::spreadsheet_formula::registry::BuiltinFunction;
use crate::spreadsheet_formula::SpreadsheetFormulaExpression;
use crate::SpreadsheetFormulaErrorLiteral;

use super::{
    calculation_error, checked_array_cells, finite_or_number_error, scalar_number, EvalValue,
    EvaluationContext, FormulaCellKey, ScalarValue,
};
use aggregate::{
    add_function_cells, aggregate, concatenate, count, count_a, logical_aggregate, logical_not,
    Aggregate,
};
use array::{row_or_column, sequence, transpose};
use lazy::{evaluate_if, evaluate_if_error};

pub(super) fn evaluate_function(
    context: &mut EvaluationContext<'_>,
    name: &str,
    arguments: &[Option<SpreadsheetFormulaExpression>],
    current: FormulaCellKey,
) -> UseResult<EvalValue> {
    let Some(definition) = context.registry.get(name) else {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_function_unsupported",
            format!("Native calculation does not implement function '{name}'."),
        )
        .with_detail("function", name));
    };
    if arguments.len() < definition.minimum_arguments
        || definition
            .maximum_arguments
            .is_some_and(|maximum| arguments.len() > maximum)
    {
        return Err(calculation_error(
            "use.office.spreadsheet_formula_function_arity",
            format!(
                "Function '{}' accepts {}{} arguments, but received {}.",
                definition.name,
                definition.minimum_arguments,
                definition.maximum_arguments.map_or_else(
                    || " or more".to_string(),
                    |maximum| {
                        if maximum == definition.minimum_arguments {
                            String::new()
                        } else {
                            format!("-{maximum}")
                        }
                    }
                ),
                arguments.len()
            ),
        )
        .with_detail("function", definition.name.clone())
        .with_detail("arguments", arguments.len()));
    }
    let function = context.registry.function(name).ok_or_else(|| {
        calculation_error(
            "use.office.spreadsheet_formula_function_unsupported",
            format!("Native function registry has no implementation for '{name}'."),
        )
    })?;
    if matches!(function, BuiltinFunction::If) {
        return evaluate_if(context, arguments, current);
    }
    if matches!(function, BuiltinFunction::IfError) {
        return evaluate_if_error(context, arguments, current);
    }
    let values = evaluate_arguments(context, arguments, current)?;
    match function {
        BuiltinFunction::Sum => aggregate(context, &values, Aggregate::Sum),
        BuiltinFunction::Average => aggregate(context, &values, Aggregate::Average),
        BuiltinFunction::Minimum => aggregate(context, &values, Aggregate::Minimum),
        BuiltinFunction::Maximum => aggregate(context, &values, Aggregate::Maximum),
        BuiltinFunction::Count => count(context, &values),
        BuiltinFunction::CountA => count_a(context, &values),
        BuiltinFunction::Absolute => context.map_numeric(argument(&values, 0)?, |number| {
            finite_or_number_error(number.abs())
        }),
        BuiltinFunction::SquareRoot => context.map_numeric(argument(&values, 0)?, |number| {
            if number < 0.0 {
                ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Number)
            } else {
                finite_or_number_error(number.sqrt())
            }
        }),
        BuiltinFunction::Power => numeric_binary(context, &values, |left, right| {
            finite_or_number_error(left.powf(right))
        }),
        BuiltinFunction::Modulo => numeric_binary(context, &values, |left, right| {
            if right == 0.0 {
                ScalarValue::Error(SpreadsheetFormulaErrorLiteral::DivisionByZero)
            } else {
                finite_or_number_error(left - right * (left / right).floor())
            }
        }),
        BuiltinFunction::Round => numeric_binary(context, &values, |number, digits| {
            if !(-308.0..=308.0).contains(&digits) {
                return ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Number);
            }
            let digits = digits.trunc() as i32;
            let factor = 10_f64.powi(digits);
            if !factor.is_finite() || factor == 0.0 {
                return ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Number);
            }
            finite_or_number_error((number * factor).round() / factor)
        }),
        BuiltinFunction::And => logical_aggregate(context, &values, true),
        BuiltinFunction::Or => logical_aggregate(context, &values, false),
        BuiltinFunction::Not => logical_not(context, argument(&values, 0)?),
        BuiltinFunction::Concatenate => concatenate(context, &values),
        BuiltinFunction::Row => row_or_column(context, &values, current, true),
        BuiltinFunction::Column => row_or_column(context, &values, current, false),
        BuiltinFunction::Sequence => sequence(context, &values),
        BuiltinFunction::Transpose => transpose(context, &values),
        BuiltinFunction::Pi => Ok(EvalValue::Scalar(ScalarValue::Number(std::f64::consts::PI))),
        BuiltinFunction::NotAvailable => Ok(EvalValue::Scalar(ScalarValue::Error(
            SpreadsheetFormulaErrorLiteral::NotAvailable,
        ))),
        BuiltinFunction::If | BuiltinFunction::IfError => Err(calculation_error(
            "use.office.spreadsheet_formula_calculation_invalid",
            "Lazy function reached eager dispatch.",
        )),
    }
}

fn evaluate_arguments(
    context: &mut EvaluationContext<'_>,
    arguments: &[Option<SpreadsheetFormulaExpression>],
    current: FormulaCellKey,
) -> UseResult<Vec<EvalValue>> {
    let mut values = Vec::with_capacity(arguments.len());
    let mut materialized_cells = 0_usize;
    for argument in arguments {
        let value = argument.as_ref().map_or_else(
            || Ok(EvalValue::Scalar(ScalarValue::Blank)),
            |argument| context.evaluate_expression(argument, current),
        )?;
        if let EvalValue::Array(array) = &value {
            let cells = checked_array_cells(array.height(), array.width())?;
            add_function_cells(&mut materialized_cells, cells)?;
        }
        values.push(value);
    }
    Ok(values)
}

fn numeric_binary(
    context: &EvaluationContext<'_>,
    arguments: &[EvalValue],
    operation: impl Fn(f64, f64) -> ScalarValue + Copy,
) -> UseResult<EvalValue> {
    let left = context.materialize(argument(arguments, 0)?)?;
    let right = context.materialize(argument(arguments, 1)?)?;
    context.broadcast_binary(left, right, |left, right| {
        let (Ok(left), Ok(right)) = (scalar_number(left), scalar_number(right)) else {
            return Ok(ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value));
        };
        Ok(operation(left, right))
    })
}

fn argument(arguments: &[EvalValue], index: usize) -> UseResult<EvalValue> {
    arguments.get(index).cloned().ok_or_else(missing_argument)
}

fn missing_argument() -> a3s_use_core::UseError {
    calculation_error(
        "use.office.spreadsheet_formula_calculation_invalid",
        "Native formula function dispatch is missing a validated argument.",
    )
}
