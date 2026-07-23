use std::collections::BTreeSet;

use a3s_use_core::UseResult;

use crate::discovery::office_error;
use crate::{
    DocumentNode, NativeOfficeRgbColor, NativeSpreadsheetConditionalFormat,
    NativeSpreadsheetConditionalFormatIconSet, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetConditionalFormatThresholdKind, NativeSpreadsheetConditionalFormatTimePeriod,
    NativeSpreadsheetDifferentialFormat, OfficeNodeType,
};

impl NativeSpreadsheetConditionalFormat {
    /// Converts a supported native semantic node back to its complete typed
    /// conditional-format value. Read-only priority and dxf indexes are not
    /// part of the returned mutation value.
    pub fn from_semantic_node(node: &DocumentNode) -> UseResult<Self> {
        if node.node_type != OfficeNodeType::ConditionalFormatting
            || node.tag != "conditionalFormatting"
            || node.style.is_some()
            || !node.children.is_empty()
        {
            return Err(invalid_node(
                node,
                "is not a leaf Spreadsheet conditional-formatting node",
            ));
        }
        if node.format.get("nativeMutable").map(String::as_str) != Some("true") {
            return Err(invalid_node(
                node,
                "contains a rule or differential format outside the native typed subset",
            ));
        }
        let ranges = required(node, "ref")?
            .split_ascii_whitespace()
            .map(str::to_string)
            .collect::<Vec<_>>();
        let stop_if_true = boolean(node, "stopIfTrue", false)?;
        let format = differential_format(node)?;
        let rule_type = required(node, "type")?;
        let rule = match rule_type {
            "cellIs" => NativeSpreadsheetConditionalFormatRule::CellIs {
                operator: parse_operator(required(node, "operator")?, node)?,
                formula1: required(node, "formula1")?.into(),
                formula2: node.format.get("formula2").cloned(),
                format,
            },
            "expression" => NativeSpreadsheetConditionalFormatRule::Formula {
                formula: required(node, "formula")?.into(),
                format,
            },
            "containsText" => NativeSpreadsheetConditionalFormatRule::ContainsText {
                text: required(node, "text")?.into(),
                format,
            },
            "notContainsText" => NativeSpreadsheetConditionalFormatRule::NotContainsText {
                text: required(node, "text")?.into(),
                format,
            },
            "beginsWith" => NativeSpreadsheetConditionalFormatRule::BeginsWith {
                text: required(node, "text")?.into(),
                format,
            },
            "endsWith" => NativeSpreadsheetConditionalFormatRule::EndsWith {
                text: required(node, "text")?.into(),
                format,
            },
            "top10" => NativeSpreadsheetConditionalFormatRule::Top {
                rank: number(node, "rank")?,
                percent: boolean(node, "percent", false)?,
                bottom: boolean(node, "bottom", false)?,
                format,
            },
            "aboveAverage" => NativeSpreadsheetConditionalFormatRule::AboveAverage {
                above: boolean(node, "above", true)?,
                equal: boolean(node, "equal", false)?,
                standard_deviations: optional_number(node, "standardDeviations")?,
                format,
            },
            "duplicateValues" => NativeSpreadsheetConditionalFormatRule::DuplicateValues { format },
            "uniqueValues" => NativeSpreadsheetConditionalFormatRule::UniqueValues { format },
            "containsBlanks" => NativeSpreadsheetConditionalFormatRule::ContainsBlanks { format },
            "notContainsBlanks" => {
                NativeSpreadsheetConditionalFormatRule::NotContainsBlanks { format }
            }
            "containsErrors" => NativeSpreadsheetConditionalFormatRule::ContainsErrors { format },
            "notContainsErrors" => {
                NativeSpreadsheetConditionalFormatRule::NotContainsErrors { format }
            }
            "timePeriod" => NativeSpreadsheetConditionalFormatRule::TimePeriod {
                period: parse_period(required(node, "period")?, node)?,
                format,
            },
            "dataBar" => NativeSpreadsheetConditionalFormatRule::DataBar {
                color: color(node, "color")?,
                min: threshold(required(node, "min")?, node)?,
                max: threshold(required(node, "max")?, node)?,
                show_value: boolean(node, "showValue", true)?,
                min_length: optional_number(node, "minLength")?
                    .map(|value| {
                        u8::try_from(value)
                            .map_err(|_| invalid_node(node, "has minLength outside 0-255"))
                    })
                    .transpose()?,
                max_length: optional_number(node, "maxLength")?
                    .map(|value| {
                        u8::try_from(value)
                            .map_err(|_| invalid_node(node, "has maxLength outside 0-255"))
                    })
                    .transpose()?,
            },
            "colorScale" => NativeSpreadsheetConditionalFormatRule::ColorScale {
                min: threshold(required(node, "min")?, node)?,
                min_color: color(node, "minColor")?,
                mid: node
                    .format
                    .get("mid")
                    .map(|value| threshold(value, node))
                    .transpose()?,
                mid_color: node
                    .format
                    .get("midColor")
                    .map(|_| color(node, "midColor"))
                    .transpose()?,
                max: threshold(required(node, "max")?, node)?,
                max_color: color(node, "maxColor")?,
            },
            "iconSet" => NativeSpreadsheetConditionalFormatRule::IconSet {
                icon_set: parse_icon_set(required(node, "iconSet")?, node)?,
                thresholds: required(node, "thresholds")?
                    .split(';')
                    .map(|value| threshold(value, node))
                    .collect::<UseResult<Vec<_>>>()?,
                reverse: boolean(node, "reverse", false)?,
                show_value: boolean(node, "showValue", true)?,
            },
            value => {
                return Err(invalid_node(
                    node,
                    format!("has unsupported rule type '{value}'"),
                ))
            }
        };
        validate_known_properties(node, rule_type)?;
        Ok(Self {
            ranges,
            stop_if_true,
            rule,
        })
    }
}

fn differential_format(node: &DocumentNode) -> UseResult<NativeSpreadsheetDifferentialFormat> {
    Ok(NativeSpreadsheetDifferentialFormat {
        fill: node
            .format
            .get("fill")
            .map(|_| color(node, "fill"))
            .transpose()?,
        font_color: node
            .format
            .get("fontColor")
            .map(|_| color(node, "fontColor"))
            .transpose()?,
        bold: node
            .format
            .get("fontBold")
            .map(|_| boolean(node, "fontBold", false))
            .transpose()?,
    })
}

fn validate_known_properties(node: &DocumentNode, rule_type: &str) -> UseResult<()> {
    let mut allowed = BTreeSet::from([
        "ref",
        "type",
        "priority",
        "stopIfTrue",
        "nativeMutable",
        "dxfId",
        "fill",
        "fontColor",
        "fontBold",
    ]);
    let extra: &[&str] = match rule_type {
        "cellIs" => &["operator", "formula1", "formula2"],
        "expression" => &["formula"],
        "containsText" | "notContainsText" | "beginsWith" | "endsWith" => &["text"],
        "top10" => &["rank", "percent", "bottom"],
        "aboveAverage" => &["above", "equal", "standardDeviations"],
        "timePeriod" => &["period"],
        "dataBar" => &["color", "min", "max", "showValue", "minLength", "maxLength"],
        "colorScale" => &["min", "minColor", "mid", "midColor", "max", "maxColor"],
        "iconSet" => &["iconSet", "thresholds", "reverse", "showValue"],
        _ => &[],
    };
    allowed.extend(extra.iter().copied());
    if let Some(key) = node
        .format
        .keys()
        .find(|key| !allowed.contains(key.as_str()))
    {
        return Err(invalid_node(
            node,
            format!("contains unsupported semantic property '{key}'"),
        ));
    }
    Ok(())
}

fn threshold(
    value: &str,
    node: &DocumentNode,
) -> UseResult<NativeSpreadsheetConditionalFormatThreshold> {
    let (kind, value) = value
        .split_once(':')
        .map_or((value, None), |(kind, value)| {
            (kind, Some(value.to_string()))
        });
    let kind = match kind {
        "min" => NativeSpreadsheetConditionalFormatThresholdKind::Min,
        "max" => NativeSpreadsheetConditionalFormatThresholdKind::Max,
        "number" => NativeSpreadsheetConditionalFormatThresholdKind::Number,
        "percent" => NativeSpreadsheetConditionalFormatThresholdKind::Percent,
        "percentile" => NativeSpreadsheetConditionalFormatThresholdKind::Percentile,
        "formula" => NativeSpreadsheetConditionalFormatThresholdKind::Formula,
        _ => {
            return Err(invalid_node(
                node,
                format!("contains unknown threshold kind '{kind}'"),
            ))
        }
    };
    Ok(NativeSpreadsheetConditionalFormatThreshold { kind, value })
}

fn parse_operator(
    value: &str,
    node: &DocumentNode,
) -> UseResult<NativeSpreadsheetConditionalFormatOperator> {
    match value {
        "between" => Ok(NativeSpreadsheetConditionalFormatOperator::Between),
        "notBetween" => Ok(NativeSpreadsheetConditionalFormatOperator::NotBetween),
        "equal" => Ok(NativeSpreadsheetConditionalFormatOperator::Equal),
        "notEqual" => Ok(NativeSpreadsheetConditionalFormatOperator::NotEqual),
        "greaterThan" => Ok(NativeSpreadsheetConditionalFormatOperator::GreaterThan),
        "greaterThanOrEqual" => Ok(NativeSpreadsheetConditionalFormatOperator::GreaterThanOrEqual),
        "lessThan" => Ok(NativeSpreadsheetConditionalFormatOperator::LessThan),
        "lessThanOrEqual" => Ok(NativeSpreadsheetConditionalFormatOperator::LessThanOrEqual),
        _ => Err(invalid_node(
            node,
            format!("contains unsupported operator '{value}'"),
        )),
    }
}

fn parse_period(
    value: &str,
    node: &DocumentNode,
) -> UseResult<NativeSpreadsheetConditionalFormatTimePeriod> {
    match value {
        "today" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Today),
        "yesterday" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Yesterday),
        "tomorrow" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Tomorrow),
        "last7Days" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::Last7Days),
        "thisWeek" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::ThisWeek),
        "lastWeek" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::LastWeek),
        "nextWeek" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::NextWeek),
        "thisMonth" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::ThisMonth),
        "lastMonth" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::LastMonth),
        "nextMonth" => Ok(NativeSpreadsheetConditionalFormatTimePeriod::NextMonth),
        _ => Err(invalid_node(
            node,
            format!("contains unsupported time period '{value}'"),
        )),
    }
}

fn parse_icon_set(
    value: &str,
    node: &DocumentNode,
) -> UseResult<NativeSpreadsheetConditionalFormatIconSet> {
    match value {
        "3Arrows" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeArrows),
        "3ArrowsGray" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeArrowsGray),
        "3Flags" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeFlags),
        "3TrafficLights1" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeTrafficLights1),
        "3TrafficLights2" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeTrafficLights2),
        "3Signs" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeSigns),
        "3Symbols" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeSymbols),
        "3Symbols2" => Ok(NativeSpreadsheetConditionalFormatIconSet::ThreeSymbols2),
        "4Arrows" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourArrows),
        "4ArrowsGray" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourArrowsGray),
        "4RedToBlack" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourRedToBlack),
        "4Rating" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourRating),
        "4TrafficLights" => Ok(NativeSpreadsheetConditionalFormatIconSet::FourTrafficLights),
        "5Arrows" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveArrows),
        "5ArrowsGray" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveArrowsGray),
        "5Rating" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveRating),
        "5Quarters" => Ok(NativeSpreadsheetConditionalFormatIconSet::FiveQuarters),
        _ => Err(invalid_node(
            node,
            format!("contains unsupported icon set '{value}'"),
        )),
    }
}

fn color(node: &DocumentNode, key: &str) -> UseResult<NativeOfficeRgbColor> {
    let value = required(node, key)?;
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(invalid_node(
            node,
            format!("contains invalid RGB property '{key}={value}'"),
        ));
    }
    Ok(NativeOfficeRgbColor::new(
        u8::from_str_radix(&value[0..2], 16).map_err(|_| invalid_node(node, "has invalid red"))?,
        u8::from_str_radix(&value[2..4], 16)
            .map_err(|_| invalid_node(node, "has invalid green"))?,
        u8::from_str_radix(&value[4..6], 16).map_err(|_| invalid_node(node, "has invalid blue"))?,
    ))
}

fn boolean(node: &DocumentNode, key: &str, default: bool) -> UseResult<bool> {
    match node.format.get(key).map(String::as_str) {
        None => Ok(default),
        Some("true") => Ok(true),
        Some("false") => Ok(false),
        Some(value) => Err(invalid_node(
            node,
            format!("contains non-boolean property '{key}={value}'"),
        )),
    }
}

fn number(node: &DocumentNode, key: &str) -> UseResult<u32> {
    required(node, key)?
        .parse::<u32>()
        .map_err(|_| invalid_node(node, format!("contains non-integer property '{key}'")))
}

fn optional_number(node: &DocumentNode, key: &str) -> UseResult<Option<u32>> {
    node.format.get(key).map(|_| number(node, key)).transpose()
}

fn required<'a>(node: &'a DocumentNode, key: &str) -> UseResult<&'a str> {
    node.format
        .get(key)
        .map(String::as_str)
        .ok_or_else(|| invalid_node(node, format!("has no required semantic property '{key}'")))
}

fn invalid_node(node: &DocumentNode, reason: impl Into<String>) -> a3s_use_core::UseError {
    office_error(
        "use.office.spreadsheet_conditional_format_semantic_invalid",
        format!(
            "Spreadsheet conditional-format node '{}' {}.",
            node.path,
            reason.into()
        ),
    )
    .with_detail("path", node.path.clone())
}
