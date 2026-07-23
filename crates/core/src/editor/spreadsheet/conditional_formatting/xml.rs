use a3s_use_core::UseResult;

use super::super::{editor_error, qualified};
use crate::editor::{
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetConditionalFormatThresholdKind, NativeSpreadsheetConditionalFormatTimePeriod,
};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{escape_attribute, escape_text, IndexedXmlElement};
use crate::LosslessXmlPart;

const KNOWN_RULE_ATTRIBUTES: &[&str] = &[
    "type",
    "dxfId",
    "priority",
    "stopIfTrue",
    "aboveAverage",
    "percent",
    "bottom",
    "operator",
    "text",
    "timePeriod",
    "rank",
    "stdDev",
    "equalAverage",
];

pub(super) fn rule_fragment(
    namespace_prefix: Option<&str>,
    existing: Option<&IndexedXmlElement>,
    value: &NativeSpreadsheetConditionalFormat,
    priority: u32,
    dxf_id: Option<usize>,
) -> UseResult<String> {
    let mut attributes = existing
        .map(|element| element.qualified_attributes.clone())
        .unwrap_or_default();
    for attribute in KNOWN_RULE_ATTRIBUTES {
        attributes.remove(*attribute);
    }
    attributes.insert("priority".into(), priority.to_string());
    if value.stop_if_true {
        attributes.insert("stopIfTrue".into(), "1".into());
    }
    if let Some(dxf_id) = dxf_id {
        attributes.insert("dxfId".into(), dxf_id.to_string());
    }

    let mut children = String::new();
    match &value.rule {
        NativeSpreadsheetConditionalFormatRule::CellIs {
            operator,
            formula1,
            formula2,
            ..
        } => {
            attributes.insert("type".into(), "cellIs".into());
            attributes.insert("operator".into(), operator_token(*operator).into());
            children.push_str(&formula_fragment(namespace_prefix, formula1));
            if let Some(formula2) = formula2 {
                children.push_str(&formula_fragment(namespace_prefix, formula2));
            }
        }
        NativeSpreadsheetConditionalFormatRule::Formula { formula, .. } => {
            attributes.insert("type".into(), "expression".into());
            children.push_str(&formula_fragment(namespace_prefix, formula));
        }
        NativeSpreadsheetConditionalFormatRule::ContainsText { text, .. } => {
            append_text_rule(
                &mut attributes,
                &mut children,
                namespace_prefix,
                value,
                text,
                "containsText",
                "containsText",
            )?;
        }
        NativeSpreadsheetConditionalFormatRule::NotContainsText { text, .. } => {
            append_text_rule(
                &mut attributes,
                &mut children,
                namespace_prefix,
                value,
                text,
                "notContainsText",
                "notContains",
            )?;
        }
        NativeSpreadsheetConditionalFormatRule::BeginsWith { text, .. } => {
            append_text_rule(
                &mut attributes,
                &mut children,
                namespace_prefix,
                value,
                text,
                "beginsWith",
                "beginsWith",
            )?;
        }
        NativeSpreadsheetConditionalFormatRule::EndsWith { text, .. } => {
            append_text_rule(
                &mut attributes,
                &mut children,
                namespace_prefix,
                value,
                text,
                "endsWith",
                "endsWith",
            )?;
        }
        NativeSpreadsheetConditionalFormatRule::Top {
            rank,
            percent,
            bottom,
            ..
        } => {
            attributes.insert("type".into(), "top10".into());
            attributes.insert("rank".into(), rank.to_string());
            if *percent {
                attributes.insert("percent".into(), "1".into());
            }
            if *bottom {
                attributes.insert("bottom".into(), "1".into());
            }
        }
        NativeSpreadsheetConditionalFormatRule::AboveAverage {
            above,
            equal,
            standard_deviations,
            ..
        } => {
            attributes.insert("type".into(), "aboveAverage".into());
            attributes.insert("aboveAverage".into(), bool_token(*above).into());
            if *equal {
                attributes.insert("equalAverage".into(), "1".into());
            }
            if let Some(value) = standard_deviations {
                attributes.insert("stdDev".into(), value.to_string());
            }
        }
        NativeSpreadsheetConditionalFormatRule::DuplicateValues { .. } => {
            attributes.insert("type".into(), "duplicateValues".into());
        }
        NativeSpreadsheetConditionalFormatRule::UniqueValues { .. } => {
            attributes.insert("type".into(), "uniqueValues".into());
        }
        NativeSpreadsheetConditionalFormatRule::ContainsBlanks { .. } => append_predicate_rule(
            &mut attributes,
            &mut children,
            namespace_prefix,
            value,
            "containsBlanks",
        )?,
        NativeSpreadsheetConditionalFormatRule::NotContainsBlanks { .. } => append_predicate_rule(
            &mut attributes,
            &mut children,
            namespace_prefix,
            value,
            "notContainsBlanks",
        )?,
        NativeSpreadsheetConditionalFormatRule::ContainsErrors { .. } => append_predicate_rule(
            &mut attributes,
            &mut children,
            namespace_prefix,
            value,
            "containsErrors",
        )?,
        NativeSpreadsheetConditionalFormatRule::NotContainsErrors { .. } => append_predicate_rule(
            &mut attributes,
            &mut children,
            namespace_prefix,
            value,
            "notContainsErrors",
        )?,
        NativeSpreadsheetConditionalFormatRule::TimePeriod { period, .. } => {
            attributes.insert("type".into(), "timePeriod".into());
            attributes.insert("timePeriod".into(), time_period_token(*period).into());
        }
        NativeSpreadsheetConditionalFormatRule::DataBar {
            color,
            min,
            max,
            show_value,
            min_length,
            max_length,
        } => {
            attributes.insert("type".into(), "dataBar".into());
            let tag = qualified(namespace_prefix, "dataBar");
            let mut nested_attributes = String::new();
            if !*show_value {
                nested_attributes.push_str(" showValue=\"0\"");
            }
            if let Some(value) = min_length {
                nested_attributes.push_str(&format!(" minLength=\"{value}\""));
            }
            if let Some(value) = max_length {
                nested_attributes.push_str(&format!(" maxLength=\"{value}\""));
            }
            children.push_str(&format!(
                "<{tag}{nested_attributes}>{}{}{}</{tag}>",
                threshold_fragment(namespace_prefix, min),
                threshold_fragment(namespace_prefix, max),
                color_fragment(namespace_prefix, *color),
            ));
        }
        NativeSpreadsheetConditionalFormatRule::ColorScale {
            min,
            min_color,
            mid,
            mid_color,
            max,
            max_color,
        } => {
            attributes.insert("type".into(), "colorScale".into());
            let tag = qualified(namespace_prefix, "colorScale");
            let mut nested = threshold_fragment(namespace_prefix, min);
            if let Some(mid) = mid {
                nested.push_str(&threshold_fragment(namespace_prefix, mid));
            }
            nested.push_str(&threshold_fragment(namespace_prefix, max));
            nested.push_str(&color_fragment(namespace_prefix, *min_color));
            if let Some(mid_color) = mid_color {
                nested.push_str(&color_fragment(namespace_prefix, *mid_color));
            }
            nested.push_str(&color_fragment(namespace_prefix, *max_color));
            children.push_str(&format!("<{tag}>{nested}</{tag}>"));
        }
        NativeSpreadsheetConditionalFormatRule::IconSet {
            icon_set,
            thresholds,
            reverse,
            show_value,
        } => {
            attributes.insert("type".into(), "iconSet".into());
            let tag = qualified(namespace_prefix, "iconSet");
            let mut nested_attributes = format!(" iconSet=\"{}\"", icon_set.token());
            if *reverse {
                nested_attributes.push_str(" reverse=\"1\"");
            }
            if !*show_value {
                nested_attributes.push_str(" showValue=\"0\"");
            }
            let nested = thresholds
                .iter()
                .map(|threshold| threshold_fragment(namespace_prefix, threshold))
                .collect::<String>();
            children.push_str(&format!("<{tag}{nested_attributes}>{nested}</{tag}>"));
        }
    }

    let rule_name = existing.map_or_else(
        || qualified(namespace_prefix, "cfRule"),
        |element| element.qualified_name.clone(),
    );
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", escape_attribute(&value)))
        .collect::<String>();
    Ok(format!("<{rule_name}{attributes}>{children}</{rule_name}>"))
}

pub(super) fn reject_unknown_rule_children(
    part_name: &str,
    part: &LosslessXmlPart,
    rule: &IndexedXmlElement,
) -> UseResult<()> {
    let bytes = part.parse_bytes();
    let mut cursor = rule.content_range.start;
    for child in &rule.children {
        let gap = bytes
            .get(cursor..child.full_range.start)
            .ok_or_else(|| invalid(part_name, "has invalid rule child ranges"))?;
        if !gap.iter().all(u8::is_ascii_whitespace)
            || child.namespace != rule.namespace
            || !matches!(
                child.local_name.as_str(),
                "formula" | "dataBar" | "colorScale" | "iconSet"
            )
        {
            return Err(unknown_content(part_name, &child.qualified_name));
        }
        cursor = child.full_range.end;
    }
    let trailing = bytes
        .get(cursor..rule.content_range.end)
        .ok_or_else(|| invalid(part_name, "has invalid trailing rule content"))?;
    if !trailing.iter().all(u8::is_ascii_whitespace) {
        return Err(unknown_content(part_name, "non-element child content"));
    }
    Ok(())
}

fn formula_fragment(namespace_prefix: Option<&str>, value: &str) -> String {
    let tag = qualified(namespace_prefix, "formula");
    format!("<{tag}>{}</{tag}>", escape_text(value))
}

fn append_text_rule(
    attributes: &mut std::collections::BTreeMap<String, String>,
    children: &mut String,
    namespace_prefix: Option<&str>,
    value: &NativeSpreadsheetConditionalFormat,
    text: &str,
    rule_type: &str,
    operator: &str,
) -> UseResult<()> {
    attributes.insert("type".into(), rule_type.into());
    attributes.insert("operator".into(), operator.into());
    attributes.insert("text".into(), text.into());
    children.push_str(&formula_fragment(
        namespace_prefix,
        &text_rule_formula(&value.ranges[0], text, rule_type)?,
    ));
    Ok(())
}

fn append_predicate_rule(
    attributes: &mut std::collections::BTreeMap<String, String>,
    children: &mut String,
    namespace_prefix: Option<&str>,
    value: &NativeSpreadsheetConditionalFormat,
    rule_type: &str,
) -> UseResult<()> {
    attributes.insert("type".into(), rule_type.into());
    children.push_str(&formula_fragment(
        namespace_prefix,
        &predicate_rule_formula(&value.ranges[0], rule_type)?,
    ));
    Ok(())
}

fn threshold_fragment(
    namespace_prefix: Option<&str>,
    threshold: &NativeSpreadsheetConditionalFormatThreshold,
) -> String {
    let tag = qualified(namespace_prefix, "cfvo");
    let kind = match threshold.kind {
        NativeSpreadsheetConditionalFormatThresholdKind::Min => "min",
        NativeSpreadsheetConditionalFormatThresholdKind::Max => "max",
        NativeSpreadsheetConditionalFormatThresholdKind::Number => "num",
        NativeSpreadsheetConditionalFormatThresholdKind::Percent => "percent",
        NativeSpreadsheetConditionalFormatThresholdKind::Percentile => "percentile",
        NativeSpreadsheetConditionalFormatThresholdKind::Formula => "formula",
    };
    threshold.value.as_ref().map_or_else(
        || format!("<{tag} type=\"{kind}\"/>"),
        |value| {
            format!(
                "<{tag} type=\"{kind}\" val=\"{}\"/>",
                escape_attribute(value)
            )
        },
    )
}

fn color_fragment(
    namespace_prefix: Option<&str>,
    color: crate::editor::NativeOfficeRgbColor,
) -> String {
    let tag = qualified(namespace_prefix, "color");
    format!("<{tag} rgb=\"FF{}\"/>", color.hex())
}

fn text_rule_formula(range: &str, text: &str, rule_type: &str) -> UseResult<String> {
    let anchor = CellRange::parse(range)?.start.a1();
    let text = text.replace('"', "\"\"");
    Ok(match rule_type {
        "containsText" => format!("NOT(ISERROR(SEARCH(\"{text}\",{anchor})))"),
        "notContainsText" => format!("ISERROR(SEARCH(\"{text}\",{anchor}))"),
        "beginsWith" => format!("LEFT({anchor},{})=\"{text}\"", text.chars().count()),
        "endsWith" => format!("RIGHT({anchor},{})=\"{text}\"", text.chars().count()),
        _ => {
            return Err(editor_error(
                "use.office.spreadsheet_conditional_format_type_invalid",
                format!("Unsupported text conditional-format type '{rule_type}'."),
            ))
        }
    })
}

fn predicate_rule_formula(range: &str, rule_type: &str) -> UseResult<String> {
    let anchor = CellRange::parse(range)?.start.a1();
    Ok(match rule_type {
        "containsBlanks" => format!("LEN(TRIM({anchor}))=0"),
        "notContainsBlanks" => format!("LEN(TRIM({anchor}))>0"),
        "containsErrors" => format!("ISERROR({anchor})"),
        "notContainsErrors" => format!("NOT(ISERROR({anchor}))"),
        _ => {
            return Err(editor_error(
                "use.office.spreadsheet_conditional_format_type_invalid",
                format!("Unsupported predicate conditional-format type '{rule_type}'."),
            ))
        }
    })
}

const fn operator_token(value: NativeSpreadsheetConditionalFormatOperator) -> &'static str {
    match value {
        NativeSpreadsheetConditionalFormatOperator::Between => "between",
        NativeSpreadsheetConditionalFormatOperator::NotBetween => "notBetween",
        NativeSpreadsheetConditionalFormatOperator::Equal => "equal",
        NativeSpreadsheetConditionalFormatOperator::NotEqual => "notEqual",
        NativeSpreadsheetConditionalFormatOperator::GreaterThan => "greaterThan",
        NativeSpreadsheetConditionalFormatOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        NativeSpreadsheetConditionalFormatOperator::LessThan => "lessThan",
        NativeSpreadsheetConditionalFormatOperator::LessThanOrEqual => "lessThanOrEqual",
    }
}

const fn time_period_token(value: NativeSpreadsheetConditionalFormatTimePeriod) -> &'static str {
    match value {
        NativeSpreadsheetConditionalFormatTimePeriod::Today => "today",
        NativeSpreadsheetConditionalFormatTimePeriod::Yesterday => "yesterday",
        NativeSpreadsheetConditionalFormatTimePeriod::Tomorrow => "tomorrow",
        NativeSpreadsheetConditionalFormatTimePeriod::Last7Days => "last7Days",
        NativeSpreadsheetConditionalFormatTimePeriod::ThisWeek => "thisWeek",
        NativeSpreadsheetConditionalFormatTimePeriod::LastWeek => "lastWeek",
        NativeSpreadsheetConditionalFormatTimePeriod::NextWeek => "nextWeek",
        NativeSpreadsheetConditionalFormatTimePeriod::ThisMonth => "thisMonth",
        NativeSpreadsheetConditionalFormatTimePeriod::LastMonth => "lastMonth",
        NativeSpreadsheetConditionalFormatTimePeriod::NextMonth => "nextMonth",
    }
}

const fn bool_token(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn unknown_content(part_name: &str, child: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_conditional_format_unknown_content",
        "The conditional format cannot be replaced without discarding unknown child content.",
    )
    .with_detail("part", part_name)
    .with_detail("child", child)
}

fn invalid(part_name: &str, reason: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_conditional_format_invalid",
        format!("Spreadsheet worksheet conditional formatting {reason}."),
    )
    .with_detail("part", part_name)
}
