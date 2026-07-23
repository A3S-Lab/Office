use a3s_office::{
    DocumentNode, NativeOfficeRgbColor, NativeSpreadsheetConditionalFormat,
    NativeSpreadsheetConditionalFormatIconSet, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetConditionalFormatThresholdKind, NativeSpreadsheetConditionalFormatTimePeriod,
    NativeSpreadsheetDifferentialFormat,
};
use a3s_use_core::UseResult;

use super::format::{parse_format_boolean, parse_rgb_color};
use super::{usage_error, ParsedArguments};

pub(super) fn canonical_path(path: &str) -> Option<String> {
    let (sheet, segment) = path.rsplit_once('/')?;
    let (name, index) = segment.split_once('[')?;
    if !matches!(
        normalize_token(name).as_str(),
        "cf" | "conditionalformat" | "conditionalformatting"
    ) {
        return None;
    }
    let index = index
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|index| *index > 0)?;
    Some(format!("{sheet}/cf[{index}]"))
}

pub(super) fn build_new(parsed: &ParsedArguments) -> UseResult<NativeSpreadsheetConditionalFormat> {
    let rule_type = parsed
        .conditional_format_type
        .as_deref()
        .ok_or_else(|| usage_error("native conditional-format add requires --rule-type"))?;
    let ranges = parsed_ranges(parsed)?;
    if ranges.is_empty() {
        return Err(usage_error(
            "native conditional-format add requires at least one --range",
        ));
    }
    Ok(NativeSpreadsheetConditionalFormat {
        ranges,
        stop_if_true: optional_bool(
            "--stop-if-true",
            parsed.conditional_stop_if_true.as_deref(),
            false,
        )?,
        rule: build_rule(parsed, rule_type, None)?,
    })
}

pub(super) fn merge_existing(
    node: &DocumentNode,
    parsed: &ParsedArguments,
) -> UseResult<NativeSpreadsheetConditionalFormat> {
    let mut current =
        NativeSpreadsheetConditionalFormat::from_semantic_node(node).map_err(|error| {
            usage_error(format!("conditional-format node is not editable: {error}"))
        })?;
    let ranges = parsed_ranges(parsed)?;
    if !ranges.is_empty() {
        current.ranges = ranges;
    }
    if let Some(value) = parsed.conditional_stop_if_true.as_deref() {
        current.stop_if_true = parse_bool("--stop-if-true", value)?;
    }
    let current_type = rule_type(&current.rule);
    let requested_type = parsed
        .conditional_format_type
        .as_deref()
        .unwrap_or(current_type);
    let previous =
        (normalize_token(requested_type) == normalize_token(current_type)).then_some(&current.rule);
    current.rule = build_rule(parsed, requested_type, previous)?;
    Ok(current)
}

fn build_rule(
    parsed: &ParsedArguments,
    requested_type: &str,
    previous: Option<&NativeSpreadsheetConditionalFormatRule>,
) -> UseResult<NativeSpreadsheetConditionalFormatRule> {
    let kind = normalize_token(requested_type);
    match kind.as_str() {
        "cellis" | "comparison" => {
            reject_unused(parsed, &["operator", "formula", "formula1", "formula2"])?;
            let old = match previous {
                Some(NativeSpreadsheetConditionalFormatRule::CellIs {
                    operator,
                    formula1,
                    formula2,
                    format,
                }) => Some((*operator, formula1, formula2, format)),
                _ => None,
            };
            let operator = parsed
                .validation_operator
                .as_deref()
                .map(parse_operator)
                .transpose()?
                .or_else(|| old.map(|value| value.0))
                .ok_or_else(|| usage_error("cell-is conditional format requires --operator"))?;
            let formula1 = one_formula_input(parsed)?
                .map(str::to_string)
                .or_else(|| old.map(|value| value.1.clone()))
                .ok_or_else(|| usage_error("cell-is conditional format requires --formula1"))?;
            let formula2 = match parsed.validation_formula2.as_deref() {
                Some(value) if is_none(value) => None,
                Some(value) => Some(value.to_string()),
                None => old.and_then(|value| value.2.clone()),
            };
            Ok(NativeSpreadsheetConditionalFormatRule::CellIs {
                operator,
                formula1,
                formula2,
                format: differential_format(parsed, old.map(|value| value.3))?,
            })
        }
        "formula" | "expression" => {
            reject_unused(parsed, &["formula", "formula1"])?;
            let old = match previous {
                Some(NativeSpreadsheetConditionalFormatRule::Formula { formula, format }) => {
                    Some((formula, format))
                }
                _ => None,
            };
            let formula = one_formula_input(parsed)?
                .map(str::to_string)
                .or_else(|| old.map(|value| value.0.clone()))
                .ok_or_else(|| usage_error("formula conditional format requires --formula"))?;
            Ok(NativeSpreadsheetConditionalFormatRule::Formula {
                formula,
                format: differential_format(parsed, old.map(|value| value.1))?,
            })
        }
        "containstext" | "notcontainstext" | "beginswith" | "endswith" => {
            reject_unused(parsed, &["text"])?;
            let (old_text, old_format) = match previous {
                Some(NativeSpreadsheetConditionalFormatRule::ContainsText { text, format })
                | Some(NativeSpreadsheetConditionalFormatRule::NotContainsText {
                    text,
                    format,
                })
                | Some(NativeSpreadsheetConditionalFormatRule::BeginsWith { text, format })
                | Some(NativeSpreadsheetConditionalFormatRule::EndsWith { text, format }) => {
                    (Some(text), Some(format))
                }
                _ => (None, None),
            };
            let text = parsed
                .text
                .clone()
                .or_else(|| old_text.cloned())
                .ok_or_else(|| usage_error("text conditional format requires --text"))?;
            let format = differential_format(parsed, old_format)?;
            Ok(match kind.as_str() {
                "containstext" => NativeSpreadsheetConditionalFormatRule::ContainsText {
                    text,
                    format,
                },
                "notcontainstext" => NativeSpreadsheetConditionalFormatRule::NotContainsText {
                    text,
                    format,
                },
                "beginswith" => {
                    NativeSpreadsheetConditionalFormatRule::BeginsWith { text, format }
                }
                "endswith" => NativeSpreadsheetConditionalFormatRule::EndsWith { text, format },
                _ => {
                    return Err(usage_error(format!(
                        "unsupported conditional-format rule type '{requested_type}'"
                    )))
                }
            })
        }
        "top" | "topn" | "top10" | "toppercent" | "bottom" | "bottomn"
        | "bottompercent" => {
            reject_unused(parsed, &["rank", "percent", "bottom"])?;
            let old = match previous {
                Some(NativeSpreadsheetConditionalFormatRule::Top {
                    rank,
                    percent,
                    bottom,
                    format,
                }) => Some((*rank, *percent, *bottom, format)),
                _ => None,
            };
            let routed_percent = matches!(kind.as_str(), "toppercent" | "bottompercent");
            let routed_bottom = matches!(kind.as_str(), "bottom" | "bottomn" | "bottompercent");
            Ok(NativeSpreadsheetConditionalFormatRule::Top {
                rank: parsed
                    .conditional_rank
                    .or_else(|| old.map(|value| value.0))
                    .unwrap_or(10),
                percent: parsed
                    .conditional_percent
                    .as_deref()
                    .map(|value| parse_bool("--percent", value))
                    .transpose()?
                    .unwrap_or_else(|| old.map_or(routed_percent, |value| value.1)),
                bottom: parsed
                    .conditional_bottom
                    .as_deref()
                    .map(|value| parse_bool("--bottom", value))
                    .transpose()?
                    .unwrap_or_else(|| old.map_or(routed_bottom, |value| value.2)),
                format: differential_format(parsed, old.map(|value| value.3))?,
            })
        }
        "aboveaverage" | "belowaverage" => {
            reject_unused(parsed, &["above", "equal", "standardDeviations"])?;
            let old = match previous {
                Some(NativeSpreadsheetConditionalFormatRule::AboveAverage {
                    above,
                    equal,
                    standard_deviations,
                    format,
                }) => Some((*above, *equal, *standard_deviations, format)),
                _ => None,
            };
            Ok(NativeSpreadsheetConditionalFormatRule::AboveAverage {
                above: parsed
                    .conditional_above
                    .as_deref()
                    .map(|value| parse_bool("--above", value))
                    .transpose()?
                    .unwrap_or_else(|| old.map_or(kind != "belowaverage", |value| value.0)),
                equal: parsed
                    .conditional_equal
                    .as_deref()
                    .map(|value| parse_bool("--equal-average", value))
                    .transpose()?
                    .unwrap_or_else(|| old.is_some_and(|value| value.1)),
                standard_deviations: parsed
                    .conditional_standard_deviations
                    .or_else(|| old.and_then(|value| value.2)),
                format: differential_format(parsed, old.map(|value| value.3))?,
            })
        }
        "duplicatevalues" | "uniquevalues" | "containsblanks" | "notcontainsblanks"
        | "containserrors" | "notcontainserrors" => {
            reject_unused(parsed, &[])?;
            let old_format = previous.and_then(classic_format);
            let format = differential_format(parsed, old_format)?;
            Ok(match kind.as_str() {
                "duplicatevalues" => {
                    NativeSpreadsheetConditionalFormatRule::DuplicateValues { format }
                }
                "uniquevalues" => NativeSpreadsheetConditionalFormatRule::UniqueValues { format },
                "containsblanks" => {
                    NativeSpreadsheetConditionalFormatRule::ContainsBlanks { format }
                }
                "notcontainsblanks" => {
                    NativeSpreadsheetConditionalFormatRule::NotContainsBlanks { format }
                }
                "containserrors" => {
                    NativeSpreadsheetConditionalFormatRule::ContainsErrors { format }
                }
                "notcontainserrors" => {
                    NativeSpreadsheetConditionalFormatRule::NotContainsErrors { format }
                }
                _ => {
                    return Err(usage_error(format!(
                        "unsupported conditional-format rule type '{requested_type}'"
                    )))
                }
            })
        }
        "dateoccurring" | "timeperiod" | "today" | "yesterday" | "tomorrow"
        | "last7days" | "thisweek" | "lastweek" | "nextweek" | "thismonth"
        | "lastmonth" | "nextmonth" => {
            reject_unused(parsed, &["period"])?;
            let old = match previous {
                Some(NativeSpreadsheetConditionalFormatRule::TimePeriod { period, format }) => {
                    Some((*period, format))
                }
                _ => None,
            };
            let implicit_period = if matches!(kind.as_str(), "dateoccurring" | "timeperiod") {
                None
            } else {
                Some(kind.as_str())
            };
            let period = parsed
                .conditional_period
                .as_deref()
                .or(implicit_period)
                .map(parse_period)
                .transpose()?
                .or_else(|| old.map(|value| value.0))
                .unwrap_or(NativeSpreadsheetConditionalFormatTimePeriod::Today);
            Ok(NativeSpreadsheetConditionalFormatRule::TimePeriod {
                period,
                format: differential_format(parsed, old.map(|value| value.1))?,
            })
        }
        "databar" => build_data_bar(parsed, previous),
        "colorscale" => build_color_scale(parsed, previous),
        "iconset" => build_icon_set(parsed, previous),
        _ => Err(usage_error(format!(
            "--rule-type requires cell-is, formula, contains-text, not-contains-text, begins-with, ends-with, top, bottom, above-average, below-average, duplicate-values, unique-values, contains-blanks, not-contains-blanks, contains-errors, not-contains-errors, time-period, data-bar, color-scale, or icon-set; received '{requested_type}'"
        ))),
    }
}

fn build_data_bar(
    parsed: &ParsedArguments,
    previous: Option<&NativeSpreadsheetConditionalFormatRule>,
) -> UseResult<NativeSpreadsheetConditionalFormatRule> {
    reject_unused(
        parsed,
        &["color", "min", "max", "showValue", "minLength", "maxLength"],
    )?;
    reject_differential_options(parsed, "data-bar")?;
    let old = match previous {
        Some(NativeSpreadsheetConditionalFormatRule::DataBar {
            color,
            min,
            max,
            show_value,
            min_length,
            max_length,
        }) => Some((*color, min, max, *show_value, *min_length, *max_length)),
        _ => None,
    };
    Ok(NativeSpreadsheetConditionalFormatRule::DataBar {
        color: optional_color("--color", parsed.conditional_color.as_deref())?
            .or_else(|| old.map(|value| value.0))
            .unwrap_or(NativeOfficeRgbColor::new(99, 142, 198)),
        min: parsed
            .conditional_min
            .as_deref()
            .map(|value| parse_threshold(value, ThresholdDefault::Min))
            .transpose()?
            .or_else(|| old.map(|value| value.1.clone()))
            .unwrap_or_else(NativeSpreadsheetConditionalFormatThreshold::min),
        max: parsed
            .conditional_max
            .as_deref()
            .map(|value| parse_threshold(value, ThresholdDefault::Max))
            .transpose()?
            .or_else(|| old.map(|value| value.2.clone()))
            .unwrap_or_else(NativeSpreadsheetConditionalFormatThreshold::max),
        show_value: parsed
            .conditional_show_value
            .as_deref()
            .map(|value| parse_bool("--show-value", value))
            .transpose()?
            .unwrap_or_else(|| old.is_none_or(|value| value.3)),
        min_length: optional_u8("--min-length", parsed.conditional_min_length)?
            .or_else(|| old.and_then(|value| value.4)),
        max_length: optional_u8("--max-length", parsed.conditional_max_length)?
            .or_else(|| old.and_then(|value| value.5)),
    })
}

fn build_color_scale(
    parsed: &ParsedArguments,
    previous: Option<&NativeSpreadsheetConditionalFormatRule>,
) -> UseResult<NativeSpreadsheetConditionalFormatRule> {
    reject_unused(
        parsed,
        &["min", "max", "minColor", "midColor", "maxColor", "midpoint"],
    )?;
    reject_differential_options(parsed, "color-scale")?;
    let old = match previous {
        Some(NativeSpreadsheetConditionalFormatRule::ColorScale {
            min,
            min_color,
            mid,
            mid_color,
            max,
            max_color,
        }) => Some((min, *min_color, mid, *mid_color, max, *max_color)),
        _ => None,
    };
    let explicit_mid_clear = parsed.conditional_midpoint.as_deref().is_some_and(is_none)
        || parsed.conditional_mid_color.as_deref().is_some_and(is_none);
    let wants_mid = !explicit_mid_clear
        && (parsed.conditional_midpoint.is_some()
            || parsed.conditional_mid_color.is_some()
            || old.is_some_and(|value| value.2.is_some()));
    let mid = if wants_mid {
        Some(
            parsed
                .conditional_midpoint
                .as_deref()
                .filter(|value| !is_none(value))
                .map(|value| parse_threshold(value, ThresholdDefault::Percentile))
                .transpose()?
                .or_else(|| old.and_then(|value| value.2.clone()))
                .unwrap_or_else(|| NativeSpreadsheetConditionalFormatThreshold::percentile("50")),
        )
    } else {
        None
    };
    let mid_color = if wants_mid {
        optional_color("--mid-color", parsed.conditional_mid_color.as_deref())?
            .or_else(|| old.and_then(|value| value.3))
            .or(Some(NativeOfficeRgbColor::new(255, 235, 132)))
    } else {
        None
    };
    Ok(NativeSpreadsheetConditionalFormatRule::ColorScale {
        min: parsed
            .conditional_min
            .as_deref()
            .map(|value| parse_threshold(value, ThresholdDefault::Min))
            .transpose()?
            .or_else(|| old.map(|value| value.0.clone()))
            .unwrap_or_else(NativeSpreadsheetConditionalFormatThreshold::min),
        min_color: optional_color("--min-color", parsed.conditional_min_color.as_deref())?
            .or_else(|| old.map(|value| value.1))
            .unwrap_or(NativeOfficeRgbColor::new(248, 105, 107)),
        mid,
        mid_color,
        max: parsed
            .conditional_max
            .as_deref()
            .map(|value| parse_threshold(value, ThresholdDefault::Max))
            .transpose()?
            .or_else(|| old.map(|value| value.4.clone()))
            .unwrap_or_else(NativeSpreadsheetConditionalFormatThreshold::max),
        max_color: optional_color("--max-color", parsed.conditional_max_color.as_deref())?
            .or_else(|| old.map(|value| value.5))
            .unwrap_or(NativeOfficeRgbColor::new(99, 190, 123)),
    })
}

fn build_icon_set(
    parsed: &ParsedArguments,
    previous: Option<&NativeSpreadsheetConditionalFormatRule>,
) -> UseResult<NativeSpreadsheetConditionalFormatRule> {
    reject_unused(parsed, &["iconSet", "thresholds", "reverse", "showValue"])?;
    reject_differential_options(parsed, "icon-set")?;
    let old = match previous {
        Some(NativeSpreadsheetConditionalFormatRule::IconSet {
            icon_set,
            thresholds,
            reverse,
            show_value,
        }) => Some((*icon_set, thresholds, *reverse, *show_value)),
        _ => None,
    };
    Ok(NativeSpreadsheetConditionalFormatRule::IconSet {
        icon_set: parsed
            .conditional_icon_set
            .as_deref()
            .map(parse_icon_set)
            .transpose()?
            .or_else(|| old.map(|value| value.0))
            .unwrap_or_default(),
        thresholds: if parsed.conditional_thresholds.is_empty() {
            old.map_or_else(Vec::new, |value| value.1.clone())
        } else {
            parsed
                .conditional_thresholds
                .iter()
                .map(|value| parse_threshold(value, ThresholdDefault::Percent))
                .collect::<UseResult<Vec<_>>>()?
        },
        reverse: parsed
            .conditional_reverse
            .as_deref()
            .map(|value| parse_bool("--reverse", value))
            .transpose()?
            .unwrap_or_else(|| old.is_some_and(|value| value.2)),
        show_value: parsed
            .conditional_show_value
            .as_deref()
            .map(|value| parse_bool("--show-value", value))
            .transpose()?
            .unwrap_or_else(|| old.is_none_or(|value| value.3)),
    })
}

fn differential_format(
    parsed: &ParsedArguments,
    previous: Option<&NativeSpreadsheetDifferentialFormat>,
) -> UseResult<NativeSpreadsheetDifferentialFormat> {
    let mut format = previous.cloned().unwrap_or_default();
    if let Some(value) = parsed.fill.as_deref() {
        format.fill = optional_color("--fill", Some(value))?;
    }
    if let Some(value) = parsed.text_color.as_deref() {
        format.font_color = optional_color("--text-color", Some(value))?;
    }
    if let Some(value) = parsed.bold.as_deref() {
        format.bold = if is_none(value) {
            None
        } else {
            Some(parse_format_boolean("--bold", value)?)
        };
    }
    Ok(format)
}

fn classic_format(
    rule: &NativeSpreadsheetConditionalFormatRule,
) -> Option<&NativeSpreadsheetDifferentialFormat> {
    match rule {
        NativeSpreadsheetConditionalFormatRule::DuplicateValues { format }
        | NativeSpreadsheetConditionalFormatRule::UniqueValues { format }
        | NativeSpreadsheetConditionalFormatRule::ContainsBlanks { format }
        | NativeSpreadsheetConditionalFormatRule::NotContainsBlanks { format }
        | NativeSpreadsheetConditionalFormatRule::ContainsErrors { format }
        | NativeSpreadsheetConditionalFormatRule::NotContainsErrors { format } => Some(format),
        _ => None,
    }
}

fn reject_differential_options(parsed: &ParsedArguments, kind: &str) -> UseResult<()> {
    if parsed.fill.is_some() || parsed.text_color.is_some() || parsed.bold.is_some() {
        return Err(usage_error(format!(
            "{kind} conditional formats do not accept --fill, --text-color, or --bold"
        )));
    }
    Ok(())
}

fn reject_unused(parsed: &ParsedArguments, allowed: &[&str]) -> UseResult<()> {
    let options = [
        (
            parsed.validation_operator.is_some(),
            "operator",
            "--operator",
        ),
        (parsed.formula.is_some(), "formula", "--formula"),
        (
            parsed.validation_formula1.is_some(),
            "formula1",
            "--formula1",
        ),
        (
            parsed.validation_formula2.is_some(),
            "formula2",
            "--formula2",
        ),
        (parsed.text.is_some(), "text", "--text"),
        (parsed.conditional_rank.is_some(), "rank", "--rank"),
        (parsed.conditional_percent.is_some(), "percent", "--percent"),
        (parsed.conditional_bottom.is_some(), "bottom", "--bottom"),
        (parsed.conditional_above.is_some(), "above", "--above"),
        (
            parsed.conditional_equal.is_some(),
            "equal",
            "--equal-average",
        ),
        (
            parsed.conditional_standard_deviations.is_some(),
            "standardDeviations",
            "--std-dev",
        ),
        (parsed.conditional_period.is_some(), "period", "--period"),
        (parsed.conditional_color.is_some(), "color", "--color"),
        (parsed.conditional_min.is_some(), "min", "--min"),
        (parsed.conditional_max.is_some(), "max", "--max"),
        (
            parsed.conditional_show_value.is_some(),
            "showValue",
            "--show-value",
        ),
        (
            parsed.conditional_min_length.is_some(),
            "minLength",
            "--min-length",
        ),
        (
            parsed.conditional_max_length.is_some(),
            "maxLength",
            "--max-length",
        ),
        (
            parsed.conditional_min_color.is_some(),
            "minColor",
            "--min-color",
        ),
        (
            parsed.conditional_mid_color.is_some(),
            "midColor",
            "--mid-color",
        ),
        (
            parsed.conditional_max_color.is_some(),
            "maxColor",
            "--max-color",
        ),
        (
            parsed.conditional_midpoint.is_some(),
            "midpoint",
            "--midpoint",
        ),
        (
            parsed.conditional_icon_set.is_some(),
            "iconSet",
            "--icon-set",
        ),
        (parsed.conditional_reverse.is_some(), "reverse", "--reverse"),
        (
            !parsed.conditional_thresholds.is_empty(),
            "thresholds",
            "--threshold",
        ),
    ];
    if let Some((_, _, option)) = options
        .iter()
        .find(|(present, key, _)| *present && !allowed.contains(key))
    {
        return Err(usage_error(format!(
            "conditional-format rule type does not accept {option}"
        )));
    }
    Ok(())
}

fn parsed_ranges(parsed: &ParsedArguments) -> UseResult<Vec<String>> {
    if parsed
        .validation_ranges
        .iter()
        .any(|value| value.split_ascii_whitespace().next().is_none())
    {
        return Err(usage_error("--range cannot be empty"));
    }
    Ok(parsed
        .validation_ranges
        .iter()
        .flat_map(|value| value.split_ascii_whitespace())
        .map(str::to_string)
        .collect())
}

fn one_formula_input(parsed: &ParsedArguments) -> UseResult<Option<&str>> {
    if parsed.formula.is_some() && parsed.validation_formula1.is_some() {
        return Err(usage_error(
            "conditional format accepts at most one of --formula and --formula1",
        ));
    }
    Ok(parsed
        .formula
        .as_deref()
        .or(parsed.validation_formula1.as_deref()))
}

fn parse_operator(value: &str) -> UseResult<NativeSpreadsheetConditionalFormatOperator> {
    match normalize_token(value).as_str() {
        "between" => Ok(NativeSpreadsheetConditionalFormatOperator::Between),
        "notbetween" => Ok(NativeSpreadsheetConditionalFormatOperator::NotBetween),
        "equal" | "eq" => Ok(NativeSpreadsheetConditionalFormatOperator::Equal),
        "notequal" | "ne" => Ok(NativeSpreadsheetConditionalFormatOperator::NotEqual),
        "greaterthan" | "gt" => Ok(NativeSpreadsheetConditionalFormatOperator::GreaterThan),
        "greaterthanorequal" | "gte" => {
            Ok(NativeSpreadsheetConditionalFormatOperator::GreaterThanOrEqual)
        }
        "lessthan" | "lt" => Ok(NativeSpreadsheetConditionalFormatOperator::LessThan),
        "lessthanorequal" | "lte" => {
            Ok(NativeSpreadsheetConditionalFormatOperator::LessThanOrEqual)
        }
        _ => Err(usage_error(format!(
            "--operator requires between, not-between, equal, not-equal, greater-than, greater-than-or-equal, less-than, or less-than-or-equal; received '{value}'"
        ))),
    }
}

fn parse_period(value: &str) -> UseResult<NativeSpreadsheetConditionalFormatTimePeriod> {
    match normalize_token(value).as_str() {
        "today" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Today),
        "yesterday" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Yesterday),
        "tomorrow" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Tomorrow),
        "last7days" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Last7Days),
        "thisweek" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::ThisWeek),
        "lastweek" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::LastWeek),
        "nextweek" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::NextWeek),
        "thismonth" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::ThisMonth),
        "lastmonth" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::LastMonth),
        "nextmonth" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::NextMonth),
        _ => Err(usage_error(format!(
            "--period requires today, yesterday, tomorrow, last-7-days, this-week, last-week, next-week, this-month, last-month, or next-month; received '{value}'"
        ))),
    }
}

fn parse_icon_set(value: &str) -> UseResult<NativeSpreadsheetConditionalFormatIconSet> {
    match normalize_token(value).as_str() {
        "3arrows" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeArrows),
        "3arrowsgray" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeArrowsGray),
        "3flags" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeFlags),
        "3trafficlights1" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeTrafficLights1),
        "3trafficlights2" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeTrafficLights2),
        "3signs" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeSigns),
        "3symbols" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeSymbols),
        "3symbols2" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeSymbols2),
        "4arrows" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourArrows),
        "4arrowsgray" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourArrowsGray),
        "4redtoblack" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourRedToBlack),
        "4rating" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourRating),
        "4trafficlights" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourTrafficLights),
        "5arrows" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveArrows),
        "5arrowsgray" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveArrowsGray),
        "5rating" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveRating),
        "5quarters" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveQuarters),
        _ => Err(usage_error(format!(
            "--icon-set received unsupported standard icon set '{value}'"
        ))),
    }
}

#[derive(Debug, Clone, Copy)]
enum ThresholdDefault {
    Min,
    Max,
    Percent,
    Percentile,
}

fn parse_threshold(
    value: &str,
    default: ThresholdDefault,
) -> UseResult<NativeSpreadsheetConditionalFormatThreshold> {
    let value = value.trim();
    if value.eq_ignore_ascii_case("auto") {
        return Ok(match default {
            ThresholdDefault::Max => NativeSpreadsheetConditionalFormatThreshold::max(),
            ThresholdDefault::Min | ThresholdDefault::Percent | ThresholdDefault::Percentile => {
                NativeSpreadsheetConditionalFormatThreshold::min()
            }
        });
    }
    if value.eq_ignore_ascii_case("min") {
        return Ok(NativeSpreadsheetConditionalFormatThreshold::min());
    }
    if value.eq_ignore_ascii_case("max") {
        return Ok(NativeSpreadsheetConditionalFormatThreshold::max());
    }
    let (kind, scalar) = value
        .split_once(':')
        .map_or((None, value), |(kind, scalar)| {
            (Some(normalize_token(kind)), scalar)
        });
    let kind = match kind.as_deref() {
        Some("number" | "num") => NativeSpreadsheetConditionalFormatThresholdKind::Number,
        Some("percent") => NativeSpreadsheetConditionalFormatThresholdKind::Percent,
        Some("percentile") => NativeSpreadsheetConditionalFormatThresholdKind::Percentile,
        Some("formula") => NativeSpreadsheetConditionalFormatThresholdKind::Formula,
        Some(kind) => {
            return Err(usage_error(format!(
                "unknown conditional-format threshold kind '{kind}'"
            )))
        }
        None if value.ends_with('%') => NativeSpreadsheetConditionalFormatThresholdKind::Percent,
        None => match default {
            ThresholdDefault::Percent => NativeSpreadsheetConditionalFormatThresholdKind::Percent,
            ThresholdDefault::Percentile => {
                NativeSpreadsheetConditionalFormatThresholdKind::Percentile
            }
            ThresholdDefault::Min | ThresholdDefault::Max => {
                NativeSpreadsheetConditionalFormatThresholdKind::Number
            }
        },
    };
    let scalar = scalar.strip_suffix('%').unwrap_or(scalar).trim();
    if scalar.is_empty() {
        return Err(usage_error(
            "conditional-format threshold value cannot be empty",
        ));
    }
    Ok(NativeSpreadsheetConditionalFormatThreshold {
        kind,
        value: Some(scalar.into()),
    })
}

fn optional_color(option: &str, value: Option<&str>) -> UseResult<Option<NativeOfficeRgbColor>> {
    value
        .filter(|value| !is_none(value))
        .map(|value| parse_rgb_color(option, value))
        .transpose()
}

fn optional_bool(option: &str, value: Option<&str>, default: bool) -> UseResult<bool> {
    value.map_or(Ok(default), |value| parse_bool(option, value))
}

fn parse_bool(option: &str, value: &str) -> UseResult<bool> {
    parse_format_boolean(option, value)
}

fn optional_u8(option: &str, value: Option<u32>) -> UseResult<Option<u8>> {
    value
        .map(|value| {
            u8::try_from(value).map_err(|_| usage_error(format!("{option} must be from 0 to 100")))
        })
        .transpose()
}

fn rule_type(rule: &NativeSpreadsheetConditionalFormatRule) -> &'static str {
    match rule {
        NativeSpreadsheetConditionalFormatRule::CellIs { .. } => "cell-is",
        NativeSpreadsheetConditionalFormatRule::Formula { .. } => "formula",
        NativeSpreadsheetConditionalFormatRule::ContainsText { .. } => "contains-text",
        NativeSpreadsheetConditionalFormatRule::NotContainsText { .. } => "not-contains-text",
        NativeSpreadsheetConditionalFormatRule::BeginsWith { .. } => "begins-with",
        NativeSpreadsheetConditionalFormatRule::EndsWith { .. } => "ends-with",
        NativeSpreadsheetConditionalFormatRule::Top { .. } => "top",
        NativeSpreadsheetConditionalFormatRule::AboveAverage { .. } => "above-average",
        NativeSpreadsheetConditionalFormatRule::DuplicateValues { .. } => "duplicate-values",
        NativeSpreadsheetConditionalFormatRule::UniqueValues { .. } => "unique-values",
        NativeSpreadsheetConditionalFormatRule::ContainsBlanks { .. } => "contains-blanks",
        NativeSpreadsheetConditionalFormatRule::NotContainsBlanks { .. } => "not-contains-blanks",
        NativeSpreadsheetConditionalFormatRule::ContainsErrors { .. } => "contains-errors",
        NativeSpreadsheetConditionalFormatRule::NotContainsErrors { .. } => "not-contains-errors",
        NativeSpreadsheetConditionalFormatRule::TimePeriod { .. } => "time-period",
        NativeSpreadsheetConditionalFormatRule::DataBar { .. } => "data-bar",
        NativeSpreadsheetConditionalFormatRule::ColorScale { .. } => "color-scale",
        NativeSpreadsheetConditionalFormatRule::IconSet { .. } => "icon-set",
    }
}

fn normalize_token(value: &str) -> String {
    value
        .chars()
        .filter(|character| character.is_ascii_alphanumeric())
        .flat_map(char::to_lowercase)
        .collect()
}

fn is_none(value: &str) -> bool {
    value.is_empty() || value.eq_ignore_ascii_case("none")
}
