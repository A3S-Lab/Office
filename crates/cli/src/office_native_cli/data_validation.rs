use a3s_office::{
    DocumentNode, NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
};
use a3s_use_core::UseResult;

use super::{usage_error, ParsedArguments};

pub(super) fn canonical_path(path: &str) -> Option<String> {
    let (sheet, segment) = path.rsplit_once('/')?;
    let (name, index) = segment.split_once('[')?;
    if !matches!(
        name.to_ascii_lowercase().as_str(),
        "datavalidation" | "validation"
    ) {
        return None;
    }
    let index = index
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)?;
    Some(format!("{sheet}/dataValidation[{index}]"))
}

pub(super) fn build_new(parsed: &ParsedArguments) -> UseResult<NativeSpreadsheetDataValidation> {
    let validation_type = parsed
        .validation_type
        .as_deref()
        .ok_or_else(|| usage_error("native data-validation add requires --validation-type"))
        .and_then(parse_type)?;
    let ranges = parsed_ranges(parsed)?;
    if ranges.is_empty() {
        return Err(usage_error(
            "native data-validation add requires at least one --range",
        ));
    }
    let formula1 = parsed
        .validation_formula1
        .clone()
        .ok_or_else(|| usage_error("native data-validation add requires --formula1"))?;
    Ok(NativeSpreadsheetDataValidation {
        validation_type,
        ranges,
        operator: parsed
            .validation_operator
            .as_deref()
            .map(parse_operator)
            .transpose()?,
        formula1,
        formula2: optional_value(parsed.validation_formula2.as_deref()),
        allow_blank: parse_optional_bool(
            "--allow-blank",
            parsed.validation_allow_blank.as_deref(),
            true,
        )?,
        show_input: parse_optional_bool(
            "--show-input",
            parsed.validation_show_input.as_deref(),
            true,
        )?,
        show_error: parse_optional_bool(
            "--show-error",
            parsed.validation_show_error.as_deref(),
            true,
        )?,
        prompt_title: optional_value(parsed.validation_prompt_title.as_deref()),
        prompt: optional_value(parsed.validation_prompt.as_deref()),
        error_title: optional_value(parsed.validation_error_title.as_deref()),
        error: optional_value(parsed.validation_error_message.as_deref()),
        error_style: parsed
            .validation_error_style
            .as_deref()
            .map(parse_error_style)
            .transpose()?
            .unwrap_or_default(),
        in_cell_dropdown: parse_optional_bool(
            "--in-cell-dropdown",
            parsed.validation_in_cell_dropdown.as_deref(),
            true,
        )?,
    })
}

pub(super) fn merge_existing(
    node: &DocumentNode,
    parsed: &ParsedArguments,
) -> UseResult<NativeSpreadsheetDataValidation> {
    let mut validation = from_node(node)?;
    let previous_type = validation.validation_type;
    if let Some(validation_type) = parsed.validation_type.as_deref() {
        validation.validation_type = parse_type(validation_type)?;
    }
    let type_changed = validation.validation_type != previous_type;
    let ranges = parsed_ranges(parsed)?;
    if !ranges.is_empty() {
        validation.ranges = ranges;
    }
    if let Some(operator) = parsed.validation_operator.as_deref() {
        validation.operator = if operator.eq_ignore_ascii_case("none") || operator.is_empty() {
            None
        } else {
            Some(parse_operator(operator)?)
        };
    } else if type_changed
        && matches!(
            validation.validation_type,
            NativeSpreadsheetDataValidationType::List | NativeSpreadsheetDataValidationType::Custom
        )
    {
        validation.operator = None;
    }
    if let Some(formula1) = &parsed.validation_formula1 {
        validation.formula1 = formula1.clone();
    }
    if let Some(formula2) = parsed.validation_formula2.as_deref() {
        validation.formula2 = optional_value(Some(formula2));
    } else if (type_changed
        && matches!(
            validation.validation_type,
            NativeSpreadsheetDataValidationType::List | NativeSpreadsheetDataValidationType::Custom
        ))
        || parsed
            .validation_operator
            .as_deref()
            .is_some_and(|operator| !matches_between(operator))
    {
        validation.formula2 = None;
    }
    update_bool(
        "--allow-blank",
        parsed.validation_allow_blank.as_deref(),
        &mut validation.allow_blank,
    )?;
    update_bool(
        "--show-input",
        parsed.validation_show_input.as_deref(),
        &mut validation.show_input,
    )?;
    update_bool(
        "--show-error",
        parsed.validation_show_error.as_deref(),
        &mut validation.show_error,
    )?;
    update_bool(
        "--in-cell-dropdown",
        parsed.validation_in_cell_dropdown.as_deref(),
        &mut validation.in_cell_dropdown,
    )?;
    if type_changed
        && validation.validation_type != NativeSpreadsheetDataValidationType::List
        && parsed.validation_in_cell_dropdown.is_none()
    {
        validation.in_cell_dropdown = true;
    }
    update_optional(
        &mut validation.prompt_title,
        parsed.validation_prompt_title.as_deref(),
    );
    update_optional(&mut validation.prompt, parsed.validation_prompt.as_deref());
    update_optional(
        &mut validation.error_title,
        parsed.validation_error_title.as_deref(),
    );
    update_optional(
        &mut validation.error,
        parsed.validation_error_message.as_deref(),
    );
    if let Some(style) = parsed.validation_error_style.as_deref() {
        validation.error_style = parse_error_style(style)?;
    }
    Ok(validation)
}

fn from_node(node: &DocumentNode) -> UseResult<NativeSpreadsheetDataValidation> {
    let validation_type = node
        .format
        .get("type")
        .ok_or_else(|| usage_error("data-validation node has no type"))
        .and_then(|value| parse_type(value))?;
    let ranges = node
        .format
        .get("ref")
        .ok_or_else(|| usage_error("data-validation node has no ref"))?
        .split_ascii_whitespace()
        .map(str::to_string)
        .collect();
    Ok(NativeSpreadsheetDataValidation {
        validation_type,
        ranges,
        operator: node
            .format
            .get("operator")
            .map(|value| parse_operator(value))
            .transpose()?,
        formula1: node
            .format
            .get("formula1")
            .cloned()
            .ok_or_else(|| usage_error("data-validation node has no formula1"))?,
        formula2: node.format.get("formula2").cloned(),
        allow_blank: node_bool(node, "allowBlank")?,
        show_input: node_bool(node, "showInput")?,
        show_error: node_bool(node, "showError")?,
        prompt_title: node.format.get("promptTitle").cloned(),
        prompt: node.format.get("prompt").cloned(),
        error_title: node.format.get("errorTitle").cloned(),
        error: node.format.get("error").cloned(),
        error_style: node
            .format
            .get("errorStyle")
            .map(|value| parse_error_style(value))
            .transpose()?
            .unwrap_or_default(),
        in_cell_dropdown: node_bool(node, "inCellDropdown")?,
    })
}

fn parsed_ranges(parsed: &ParsedArguments) -> UseResult<Vec<String>> {
    let ranges = parsed
        .validation_ranges
        .iter()
        .flat_map(|value| value.split_ascii_whitespace())
        .map(str::to_string)
        .collect::<Vec<_>>();
    if parsed
        .validation_ranges
        .iter()
        .any(|value| value.split_ascii_whitespace().next().is_none())
    {
        return Err(usage_error("--range cannot be empty"));
    }
    Ok(ranges)
}

fn parse_type(value: &str) -> UseResult<NativeSpreadsheetDataValidationType> {
    match value.to_ascii_lowercase().replace(['-', '_'], "").as_str() {
        "list" => Ok(NativeSpreadsheetDataValidationType::List),
        "whole" | "integer" => Ok(NativeSpreadsheetDataValidationType::Whole),
        "decimal" | "number" => Ok(NativeSpreadsheetDataValidationType::Decimal),
        "date" => Ok(NativeSpreadsheetDataValidationType::Date),
        "time" => Ok(NativeSpreadsheetDataValidationType::Time),
        "textlength" | "length" => Ok(NativeSpreadsheetDataValidationType::TextLength),
        "custom" | "formula" => Ok(NativeSpreadsheetDataValidationType::Custom),
        _ => Err(usage_error(format!(
            "--validation-type requires list, whole, decimal, date, time, text-length, or custom, received '{value}'"
        ))),
    }
}

fn parse_operator(value: &str) -> UseResult<NativeSpreadsheetDataValidationOperator> {
    match value.to_ascii_lowercase().replace(['-', '_'], "").as_str() {
        "between" => Ok(NativeSpreadsheetDataValidationOperator::Between),
        "notbetween" => Ok(NativeSpreadsheetDataValidationOperator::NotBetween),
        "equal" | "eq" => Ok(NativeSpreadsheetDataValidationOperator::Equal),
        "notequal" | "ne" => Ok(NativeSpreadsheetDataValidationOperator::NotEqual),
        "greaterthan" | "gt" => Ok(NativeSpreadsheetDataValidationOperator::GreaterThan),
        "greaterthanorequal" | "gte" => {
            Ok(NativeSpreadsheetDataValidationOperator::GreaterThanOrEqual)
        }
        "lessthan" | "lt" => Ok(NativeSpreadsheetDataValidationOperator::LessThan),
        "lessthanorequal" | "lte" => {
            Ok(NativeSpreadsheetDataValidationOperator::LessThanOrEqual)
        }
        _ => Err(usage_error(format!(
            "--operator requires between, not-between, equal, not-equal, greater-than, greater-than-or-equal, less-than, or less-than-or-equal, received '{value}'"
        ))),
    }
}

fn parse_error_style(value: &str) -> UseResult<NativeSpreadsheetDataValidationErrorStyle> {
    match value.to_ascii_lowercase().as_str() {
        "stop" => Ok(NativeSpreadsheetDataValidationErrorStyle::Stop),
        "warning" | "warn" => Ok(NativeSpreadsheetDataValidationErrorStyle::Warning),
        "information" | "info" => Ok(NativeSpreadsheetDataValidationErrorStyle::Information),
        _ => Err(usage_error(format!(
            "--error-style requires stop, warning, or information, received '{value}'"
        ))),
    }
}

fn parse_optional_bool(option: &str, value: Option<&str>, default: bool) -> UseResult<bool> {
    value.map_or(Ok(default), |value| parse_bool(option, value))
}

fn update_bool(option: &str, value: Option<&str>, target: &mut bool) -> UseResult<()> {
    if let Some(value) = value {
        *target = parse_bool(option, value)?;
    }
    Ok(())
}

fn parse_bool(option: &str, value: &str) -> UseResult<bool> {
    match value.to_ascii_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(usage_error(format!(
            "{option} requires true or false, received '{value}'"
        ))),
    }
}

fn node_bool(node: &DocumentNode, key: &str) -> UseResult<bool> {
    let value = node
        .format
        .get(key)
        .ok_or_else(|| usage_error(format!("data-validation node has no {key} value")))?;
    parse_bool(key, value)
}

fn optional_value(value: Option<&str>) -> Option<String> {
    value
        .filter(|value| !value.is_empty() && !value.eq_ignore_ascii_case("none"))
        .map(str::to_string)
}

fn update_optional(target: &mut Option<String>, value: Option<&str>) {
    if let Some(value) = value {
        *target = optional_value(Some(value));
    }
}

fn matches_between(value: &str) -> bool {
    matches!(
        value.to_ascii_lowercase().replace(['-', '_'], "").as_str(),
        "between" | "notbetween"
    )
}
