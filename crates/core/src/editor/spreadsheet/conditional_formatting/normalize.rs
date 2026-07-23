use a3s_use_core::UseResult;

use super::super::editor_error;
use crate::editor::{
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetConditionalFormatThresholdKind,
};
use crate::spreadsheet_reference::{first_intersecting_ranges, CellRange};

pub(super) const MAX_CONDITIONAL_FORMATS: usize = 65_534;
pub(super) const MAX_CONDITIONAL_FORMAT_RANGES: usize = 1_024;
const MAX_FORMULA_CHARS: usize = 8_192;
const MAX_TEXT_CHARS: usize = 255;

pub(super) fn normalize(
    value: &NativeSpreadsheetConditionalFormat,
) -> UseResult<NativeSpreadsheetConditionalFormat> {
    let mut value = value.clone();
    if value.ranges.is_empty() || value.ranges.len() > MAX_CONDITIONAL_FORMAT_RANGES {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_range_invalid",
            format!(
                "A conditional format requires 1-{MAX_CONDITIONAL_FORMAT_RANGES} rectangular A1 ranges."
            ),
        )
        .with_detail("ranges", value.ranges.len()));
    }
    let ranges = value
        .ranges
        .iter()
        .map(|range| CellRange::parse(range))
        .collect::<UseResult<Vec<_>>>()?;
    if let Some((left, right)) = first_intersecting_ranges(&ranges) {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_range_overlap",
            format!(
                "Conditional-format ranges '{}' and '{}' overlap inside one rule.",
                ranges[left].a1(),
                ranges[right].a1()
            ),
        ));
    }
    value.ranges = ranges.iter().map(|range| range.a1()).collect();

    match &mut value.rule {
        NativeSpreadsheetConditionalFormatRule::CellIs {
            operator,
            formula1,
            formula2,
            ..
        } => {
            validate_formula(formula1, "formula1")?;
            let needs_second = matches!(
                operator,
                NativeSpreadsheetConditionalFormatOperator::Between
                    | NativeSpreadsheetConditionalFormatOperator::NotBetween
            );
            if needs_second && formula2.is_none() {
                return Err(editor_error(
                    "use.office.spreadsheet_conditional_format_formula2_required",
                    "Between and not-between conditional formats require formula2.",
                ));
            }
            if !needs_second && formula2.is_some() {
                return Err(editor_error(
                    "use.office.spreadsheet_conditional_format_formula2_unsupported",
                    "Only between and not-between conditional formats accept formula2.",
                ));
            }
            if let Some(formula2) = formula2 {
                validate_formula(formula2, "formula2")?;
            }
        }
        NativeSpreadsheetConditionalFormatRule::Formula { formula, .. } => {
            validate_formula(formula, "formula")?;
        }
        NativeSpreadsheetConditionalFormatRule::ContainsText { text, .. }
        | NativeSpreadsheetConditionalFormatRule::NotContainsText { text, .. }
        | NativeSpreadsheetConditionalFormatRule::BeginsWith { text, .. }
        | NativeSpreadsheetConditionalFormatRule::EndsWith { text, .. } => {
            validate_text(text)?;
        }
        NativeSpreadsheetConditionalFormatRule::Top { rank, percent, .. } => {
            if *rank == 0 || *rank > 1_000 || (*percent && *rank > 100) {
                return Err(editor_error(
                    "use.office.spreadsheet_conditional_format_rank_invalid",
                    "Top/bottom rank must be 1-1000, or 1-100 for a percentage rule.",
                )
                .with_detail("rank", *rank)
                .with_detail("percent", *percent));
            }
        }
        NativeSpreadsheetConditionalFormatRule::AboveAverage {
            standard_deviations,
            ..
        } => {
            if standard_deviations.is_some_and(|value| value > 3) {
                return Err(editor_error(
                    "use.office.spreadsheet_conditional_format_std_dev_invalid",
                    "Average conditional formats accept 0-3 standard deviations.",
                ));
            }
        }
        NativeSpreadsheetConditionalFormatRule::DataBar {
            min,
            max,
            min_length,
            max_length,
            ..
        } => {
            validate_threshold(min, "min")?;
            validate_threshold(max, "max")?;
            validate_lengths(*min_length, *max_length)?;
        }
        NativeSpreadsheetConditionalFormatRule::ColorScale {
            min,
            mid,
            mid_color,
            max,
            ..
        } => {
            validate_threshold(min, "min")?;
            validate_threshold(max, "max")?;
            if mid.is_some() != mid_color.is_some() {
                return Err(editor_error(
                    "use.office.spreadsheet_conditional_format_midpoint_invalid",
                    "A three-color scale requires both mid and midColor; a two-color scale accepts neither.",
                ));
            }
            if let Some(mid) = mid {
                validate_threshold(mid, "mid")?;
            }
        }
        NativeSpreadsheetConditionalFormatRule::IconSet {
            icon_set,
            thresholds,
            ..
        } => {
            if thresholds.is_empty() {
                let count = icon_set.icon_count();
                *thresholds = (0..count)
                    .map(|index| NativeSpreadsheetConditionalFormatThreshold {
                        kind: NativeSpreadsheetConditionalFormatThresholdKind::Percent,
                        value: Some((index * 100 / count).to_string()),
                    })
                    .collect();
            }
            if thresholds.len() != icon_set.icon_count() {
                return Err(editor_error(
                    "use.office.spreadsheet_conditional_format_icon_threshold_invalid",
                    format!(
                        "Icon set '{}' requires exactly {} thresholds.",
                        icon_set.token(),
                        icon_set.icon_count()
                    ),
                )
                .with_detail("thresholds", thresholds.len()));
            }
            for threshold in thresholds {
                validate_threshold(threshold, "threshold")?;
            }
        }
        NativeSpreadsheetConditionalFormatRule::DuplicateValues { .. }
        | NativeSpreadsheetConditionalFormatRule::UniqueValues { .. }
        | NativeSpreadsheetConditionalFormatRule::ContainsBlanks { .. }
        | NativeSpreadsheetConditionalFormatRule::NotContainsBlanks { .. }
        | NativeSpreadsheetConditionalFormatRule::ContainsErrors { .. }
        | NativeSpreadsheetConditionalFormatRule::NotContainsErrors { .. }
        | NativeSpreadsheetConditionalFormatRule::TimePeriod { .. } => {}
    }
    Ok(value)
}

fn validate_formula(value: &str, field: &str) -> UseResult<()> {
    if value.is_empty()
        || value.chars().count() > MAX_FORMULA_CHARS
        || value.trim() != value
        || value.starts_with('=')
    {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_formula_invalid",
            format!(
                "Conditional-format {field} must contain 1-{MAX_FORMULA_CHARS} characters without surrounding whitespace or a leading '='."
            ),
        ));
    }
    validate_xml_text(value, field)
}

fn validate_text(value: &str) -> UseResult<()> {
    if value.is_empty() || value.chars().count() > MAX_TEXT_CHARS {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_text_invalid",
            format!("Conditional-format text must contain 1-{MAX_TEXT_CHARS} characters."),
        ));
    }
    validate_xml_text(value, "text")
}

fn validate_threshold(
    threshold: &NativeSpreadsheetConditionalFormatThreshold,
    field: &str,
) -> UseResult<()> {
    match threshold.kind {
        NativeSpreadsheetConditionalFormatThresholdKind::Min
        | NativeSpreadsheetConditionalFormatThresholdKind::Max => {
            if threshold.value.is_some() {
                return Err(threshold_error(
                    field,
                    "min/max thresholds do not accept a value",
                ));
            }
        }
        NativeSpreadsheetConditionalFormatThresholdKind::Number => {
            let value = required_threshold_value(threshold, field)?;
            parse_finite(value, field)?;
        }
        NativeSpreadsheetConditionalFormatThresholdKind::Percent
        | NativeSpreadsheetConditionalFormatThresholdKind::Percentile => {
            let value = required_threshold_value(threshold, field)?;
            let number = parse_finite(value, field)?;
            if !(0.0..=100.0).contains(&number) {
                return Err(threshold_error(
                    field,
                    "percent values must be between 0 and 100",
                ));
            }
        }
        NativeSpreadsheetConditionalFormatThresholdKind::Formula => {
            validate_formula(required_threshold_value(threshold, field)?, field)?;
        }
    }
    Ok(())
}

fn required_threshold_value<'a>(
    threshold: &'a NativeSpreadsheetConditionalFormatThreshold,
    field: &str,
) -> UseResult<&'a str> {
    threshold
        .value
        .as_deref()
        .filter(|value| !value.is_empty())
        .ok_or_else(|| threshold_error(field, "this threshold kind requires a value"))
}

fn parse_finite(value: &str, field: &str) -> UseResult<f64> {
    let parsed = value
        .parse::<f64>()
        .map_err(|_| threshold_error(field, "expected a finite decimal number"))?;
    if !parsed.is_finite() {
        return Err(threshold_error(field, "expected a finite decimal number"));
    }
    Ok(parsed)
}

fn validate_lengths(min: Option<u8>, max: Option<u8>) -> UseResult<()> {
    if min.is_some_and(|value| value > 100)
        || max.is_some_and(|value| value > 100)
        || matches!((min, max), (Some(min), Some(max)) if min > max)
    {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_bar_length_invalid",
            "Data-bar lengths must be between 0 and 100 with minLength <= maxLength.",
        ));
    }
    Ok(())
}

fn validate_xml_text(value: &str, field: &str) -> UseResult<()> {
    if let Some(character) = value.chars().find(|character| {
        !matches!(*character, '\u{9}' | '\u{a}' | '\u{d}')
            && (*character < '\u{20}' || matches!(*character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(editor_error(
            "use.office.spreadsheet_conditional_format_text_invalid",
            format!(
                "Conditional-format {field} contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        ));
    }
    Ok(())
}

fn threshold_error(field: &str, reason: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_conditional_format_threshold_invalid",
        format!("Conditional-format {field} threshold is invalid: {reason}."),
    )
}
