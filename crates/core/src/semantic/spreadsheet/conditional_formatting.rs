use a3s_use_core::UseResult;

use super::{direct_text, semantic_error, DocumentNode, OfficeNodeType, XmlElement};
use crate::spreadsheet_reference::CellRange;

use super::style::DifferentialFormat;

const MAX_CONDITIONAL_FORMATS: usize = 65_534;
const MAX_CONDITIONAL_FORMAT_RANGES: usize = 1_024;

pub(super) fn read(
    worksheet: &XmlElement,
    part_name: &str,
    sheet_path: &str,
    differential_formats: &[DifferentialFormat],
) -> UseResult<Vec<DocumentNode>> {
    let mut nodes = Vec::new();
    for container in worksheet.child_elements().filter(|element| {
        element.local_name == "conditionalFormatting" && element.namespace == worksheet.namespace
    }) {
        let (reference, ranges_supported) =
            normalized_ranges(unqualified_attribute(container, "sqref").unwrap_or_default());
        for rule in container.child_elements().filter(|element| {
            element.local_name == "cfRule" && element.namespace == worksheet.namespace
        }) {
            if nodes.len() >= MAX_CONDITIONAL_FORMATS {
                return Err(semantic_error(
                    "use.office.spreadsheet_conditional_format_limit",
                    format!(
                        "Worksheet part '{part_name}' contains more than {MAX_CONDITIONAL_FORMATS} conditional formats."
                    ),
                ));
            }
            let index = nodes.len() + 1;
            let rule_type = unqualified_attribute(rule, "type").unwrap_or("unknown");
            let mut node = DocumentNode::new(
                format!("{sheet_path}/cf[{index}]"),
                "conditionalFormatting",
                OfficeNodeType::ConditionalFormatting,
            );
            node.text = reference.clone();
            node.preview = Some(format!("{rule_type} ({reference})"));
            node.format.insert("ref".into(), reference.clone());
            node.format.insert("type".into(), rule_type.into());
            let mut supported = ranges_supported;

            if let Some(priority) = unqualified_attribute(rule, "priority") {
                if priority
                    .parse::<u32>()
                    .ok()
                    .filter(|value| *value > 0)
                    .is_some()
                {
                    node.format.insert("priority".into(), priority.into());
                } else {
                    supported = false;
                }
            } else {
                supported = false;
            }
            match bool_attribute(rule, "stopIfTrue", false) {
                Some(value) => {
                    node.format.insert("stopIfTrue".into(), value.to_string());
                }
                None => supported = false,
            }
            if let Some(dxf_id) = unqualified_attribute(rule, "dxfId") {
                node.format.insert("dxfId".into(), dxf_id.into());
                match dxf_id
                    .parse::<usize>()
                    .ok()
                    .and_then(|index| differential_formats.get(index))
                {
                    Some(format) => {
                        node.format.extend(format.values.clone());
                        supported &= format.supported;
                    }
                    None => supported = false,
                }
            }

            supported &= read_rule(rule, rule_type, &mut node);
            node.format
                .insert("nativeMutable".into(), supported.to_string());
            nodes.push(node);
        }
    }
    Ok(nodes)
}

fn read_rule(rule: &XmlElement, rule_type: &str, node: &mut DocumentNode) -> bool {
    match rule_type {
        "cellIs" => read_cell_is(rule, node),
        "expression" => one_formula(rule, node, "formula"),
        "containsText" | "notContainsText" | "beginsWith" | "endsWith" => {
            let Some(text) = unqualified_attribute(rule, "text") else {
                return false;
            };
            node.format.insert("text".into(), text.into());
            true
        }
        "top10" => read_top(rule, node),
        "aboveAverage" => read_average(rule, node),
        "duplicateValues" | "uniqueValues" | "containsBlanks" | "notContainsBlanks"
        | "containsErrors" | "notContainsErrors" => true,
        "timePeriod" => {
            let Some(period) = unqualified_attribute(rule, "timePeriod") else {
                return false;
            };
            if !matches!(
                period,
                "today"
                    | "yesterday"
                    | "tomorrow"
                    | "last7Days"
                    | "thisWeek"
                    | "lastWeek"
                    | "nextWeek"
                    | "thisMonth"
                    | "lastMonth"
                    | "nextMonth"
            ) {
                return false;
            }
            node.format.insert("period".into(), period.into());
            true
        }
        "dataBar" => read_data_bar(rule, node),
        "colorScale" => read_color_scale(rule, node),
        "iconSet" => read_icon_set(rule, node),
        _ => false,
    }
}

fn read_cell_is(rule: &XmlElement, node: &mut DocumentNode) -> bool {
    let Some(operator) = unqualified_attribute(rule, "operator") else {
        return false;
    };
    if !matches!(
        operator,
        "between"
            | "notBetween"
            | "equal"
            | "notEqual"
            | "greaterThan"
            | "greaterThanOrEqual"
            | "lessThan"
            | "lessThanOrEqual"
    ) {
        return false;
    }
    node.format.insert("operator".into(), operator.into());
    let formulas = formulas(rule);
    let expected = if matches!(operator, "between" | "notBetween") {
        2
    } else {
        1
    };
    if formulas.len() != expected {
        return false;
    }
    node.format.insert("formula1".into(), formulas[0].clone());
    if let Some(formula2) = formulas.get(1) {
        node.format.insert("formula2".into(), formula2.clone());
    }
    true
}

fn one_formula(rule: &XmlElement, node: &mut DocumentNode, key: &str) -> bool {
    let formulas = formulas(rule);
    if formulas.len() != 1 {
        return false;
    }
    node.format.insert(key.into(), formulas[0].clone());
    true
}

fn read_top(rule: &XmlElement, node: &mut DocumentNode) -> bool {
    let Some(rank) = unqualified_attribute(rule, "rank") else {
        return false;
    };
    if rank
        .parse::<u32>()
        .ok()
        .filter(|value| *value > 0)
        .is_none()
    {
        return false;
    }
    node.format.insert("rank".into(), rank.into());
    for (attribute, key) in [("percent", "percent"), ("bottom", "bottom")] {
        let Some(value) = bool_attribute(rule, attribute, false) else {
            return false;
        };
        node.format.insert(key.into(), value.to_string());
    }
    true
}

fn read_average(rule: &XmlElement, node: &mut DocumentNode) -> bool {
    let Some(above) = bool_attribute(rule, "aboveAverage", true) else {
        return false;
    };
    let Some(equal) = bool_attribute(rule, "equalAverage", false) else {
        return false;
    };
    node.format.insert("above".into(), above.to_string());
    node.format.insert("equal".into(), equal.to_string());
    if let Some(value) = unqualified_attribute(rule, "stdDev") {
        if value
            .parse::<u32>()
            .ok()
            .filter(|value| *value <= 3)
            .is_none()
        {
            return false;
        }
        node.format
            .insert("standardDeviations".into(), value.into());
    }
    true
}

fn read_data_bar(rule: &XmlElement, node: &mut DocumentNode) -> bool {
    let children = rule.children_named("dataBar").collect::<Vec<_>>();
    let Some(bar) = children.first().copied().filter(|_| children.len() == 1) else {
        return false;
    };
    let thresholds = threshold_tokens(bar);
    let colors = colors(bar);
    if thresholds.len() != 2 || colors.len() != 1 {
        return false;
    }
    node.format.insert("min".into(), thresholds[0].clone());
    node.format.insert("max".into(), thresholds[1].clone());
    node.format.insert("color".into(), colors[0].clone());
    let Some(show_value) = bool_attribute(bar, "showValue", true) else {
        return false;
    };
    node.format
        .insert("showValue".into(), show_value.to_string());
    for (attribute, key) in [("minLength", "minLength"), ("maxLength", "maxLength")] {
        if let Some(value) = unqualified_attribute(bar, attribute) {
            if value
                .parse::<u8>()
                .ok()
                .filter(|value| *value <= 100)
                .is_none()
            {
                return false;
            }
            node.format.insert(key.into(), value.into());
        }
    }
    true
}

fn read_color_scale(rule: &XmlElement, node: &mut DocumentNode) -> bool {
    let children = rule.children_named("colorScale").collect::<Vec<_>>();
    let Some(scale) = children.first().copied().filter(|_| children.len() == 1) else {
        return false;
    };
    let thresholds = threshold_tokens(scale);
    let colors = colors(scale);
    if thresholds.len() != colors.len() || !matches!(thresholds.len(), 2 | 3) {
        return false;
    }
    node.format.insert("min".into(), thresholds[0].clone());
    node.format.insert("minColor".into(), colors[0].clone());
    if thresholds.len() == 3 {
        node.format.insert("mid".into(), thresholds[1].clone());
        node.format.insert("midColor".into(), colors[1].clone());
    }
    node.format
        .insert("max".into(), thresholds[thresholds.len() - 1].clone());
    node.format
        .insert("maxColor".into(), colors[colors.len() - 1].clone());
    true
}

fn read_icon_set(rule: &XmlElement, node: &mut DocumentNode) -> bool {
    let children = rule.children_named("iconSet").collect::<Vec<_>>();
    let Some(icon) = children.first().copied().filter(|_| children.len() == 1) else {
        return false;
    };
    let name = unqualified_attribute(icon, "iconSet").unwrap_or("3TrafficLights1");
    let count = match name.as_bytes().first() {
        Some(b'3') => 3,
        Some(b'4') => 4,
        Some(b'5') => 5,
        _ => return false,
    };
    if !matches!(
        name,
        "3Arrows"
            | "3ArrowsGray"
            | "3Flags"
            | "3TrafficLights1"
            | "3TrafficLights2"
            | "3Signs"
            | "3Symbols"
            | "3Symbols2"
            | "4Arrows"
            | "4ArrowsGray"
            | "4RedToBlack"
            | "4Rating"
            | "4TrafficLights"
            | "5Arrows"
            | "5ArrowsGray"
            | "5Rating"
            | "5Quarters"
    ) {
        return false;
    }
    let thresholds = threshold_tokens(icon);
    if thresholds.len() != count {
        return false;
    }
    let Some(reverse) = bool_attribute(icon, "reverse", false) else {
        return false;
    };
    let Some(show_value) = bool_attribute(icon, "showValue", true) else {
        return false;
    };
    node.format.insert("iconSet".into(), name.into());
    node.format
        .insert("thresholds".into(), thresholds.join(";"));
    node.format.insert("reverse".into(), reverse.to_string());
    node.format
        .insert("showValue".into(), show_value.to_string());
    true
}

fn formulas(rule: &XmlElement) -> Vec<String> {
    rule.child_elements()
        .filter(|child| child.local_name == "formula" && child.namespace == rule.namespace)
        .map(direct_text)
        .collect()
}

fn threshold_tokens(parent: &XmlElement) -> Vec<String> {
    parent
        .children_named("cfvo")
        .filter_map(|threshold| {
            let kind = match unqualified_attribute(threshold, "type")? {
                "min" => "min",
                "max" => "max",
                "num" => "number",
                "percent" => "percent",
                "percentile" => "percentile",
                "formula" => "formula",
                _ => return None,
            };
            Some(
                unqualified_attribute(threshold, "val")
                    .map_or_else(|| kind.to_string(), |value| format!("{kind}:{value}")),
            )
        })
        .collect()
}

fn colors(parent: &XmlElement) -> Vec<String> {
    parent
        .children_named("color")
        .filter_map(|color| unqualified_attribute(color, "rgb"))
        .filter_map(normalize_rgb)
        .collect()
}

fn normalize_rgb(value: &str) -> Option<String> {
    if !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    match value.len() {
        6 => Some(value.to_ascii_uppercase()),
        8 => Some(value[2..].to_ascii_uppercase()),
        _ => None,
    }
}

fn normalized_ranges(value: &str) -> (String, bool) {
    let tokens = value.split_ascii_whitespace().collect::<Vec<_>>();
    if tokens.is_empty() || tokens.len() > MAX_CONDITIONAL_FORMAT_RANGES {
        return (value.to_string(), false);
    }
    let mut supported = true;
    let normalized = tokens
        .into_iter()
        .map(|token| match CellRange::parse(token) {
            Ok(range) => range.a1(),
            Err(_) => {
                supported = false;
                token.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ");
    (normalized, supported)
}

fn bool_attribute(element: &XmlElement, name: &str, default: bool) -> Option<bool> {
    match unqualified_attribute(element, name) {
        None => Some(default),
        Some("1" | "true") => Some(true),
        Some("0" | "false") => Some(false),
        Some(_) => None,
    }
}

fn unqualified_attribute<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| attribute.namespace.is_none() && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}
