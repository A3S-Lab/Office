use std::collections::BTreeSet;

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};
use crate::spreadsheet_filter::{criteria_type, decode_custom_pattern};
use crate::spreadsheet_reference::CellRange;
use crate::xml_tree::XmlElement;
use crate::{NativeSpreadsheetDynamicFilter, NativeSpreadsheetFilterCriteria};

pub(super) fn read(
    worksheet: &XmlElement,
    part_name: &str,
    sheet_path: &str,
) -> UseResult<Option<DocumentNode>> {
    let filters = worksheet
        .children_named("autoFilter")
        .filter(|filter| filter.namespace == worksheet.namespace)
        .collect::<Vec<_>>();
    if filters.len() > 1 {
        return Err(filter_error(
            part_name,
            "contains multiple worksheet AutoFilter elements",
        ));
    }
    filters
        .first()
        .map(|filter| read_element(filter, part_name, &format!("{sheet_path}/autofilter")))
        .transpose()
}

pub(super) fn read_element(
    filter: &XmlElement,
    part_name: &str,
    path: &str,
) -> UseResult<DocumentNode> {
    let reference = filter.attribute("ref").ok_or_else(|| {
        filter_error(
            part_name,
            "contains an AutoFilter without a range reference",
        )
    })?;
    let range = CellRange::parse(reference).map_err(|error| {
        filter_error(
            part_name,
            format!("contains invalid AutoFilter range '{reference}': {error}"),
        )
    })?;
    let mut mutable = filter
        .attributes
        .iter()
        .all(|attribute| attribute.namespace.is_none() && attribute.local_name == "ref");
    let width = range.end.column - range.start.column + 1;
    let mut observed_columns = BTreeSet::new();
    let mut node = DocumentNode::new(path, "autofilter", OfficeNodeType::AutoFilter);
    node.text = range.a1();
    node.format.insert("ref".into(), range.a1());

    for child in filter.child_elements() {
        if child.namespace != filter.namespace || child.local_name != "filterColumn" {
            mutable = false;
            continue;
        }
        let (column, column_mutable) = read_column(child, part_name, &node.path)?;
        if column
            .format
            .get("column")
            .and_then(|value| value.parse::<u32>().ok())
            .is_none_or(|column| column >= width || !observed_columns.insert(column))
        {
            mutable = false;
        }
        mutable &= column_mutable;
        node.children.push(column);
    }
    node.format
        .insert("filterColumnCount".into(), node.children.len().to_string());
    node.format
        .insert("nativeMutable".into(), mutable.to_string());
    Ok(node)
}

fn read_column(
    column: &XmlElement,
    part_name: &str,
    filter_path: &str,
) -> UseResult<(DocumentNode, bool)> {
    let id = column
        .attribute("colId")
        .ok_or_else(|| filter_error(part_name, "contains a filterColumn without colId"))?
        .parse::<u32>()
        .map_err(|error| {
            filter_error(
                part_name,
                format!("contains invalid filterColumn colId: {error}"),
            )
        })?;
    let mut mutable = column
        .attributes
        .iter()
        .all(|attribute| attribute.namespace.is_none() && attribute.local_name == "colId");
    let criteria_children = column
        .child_elements()
        .filter(|child| child.namespace == column.namespace)
        .collect::<Vec<_>>();
    let criteria = if criteria_children.len() == 1 {
        parse_criteria(criteria_children[0], part_name)?
    } else {
        None
    };
    let mut node = DocumentNode::new(
        format!("{filter_path}/filterColumn[{}]", id + 1),
        "filtercolumn",
        OfficeNodeType::FilterColumn,
    );
    node.format.insert("column".into(), id.to_string());
    if let Some(criteria) = criteria {
        if criteria.validate().is_err() {
            mutable = false;
        }
        apply_criteria(&mut node, &criteria);
    } else {
        mutable = false;
        node.format
            .insert("criteriaType".into(), "unsupported".into());
        node.text = "unsupported".into();
    }
    Ok((node, mutable))
}

fn parse_criteria(
    element: &XmlElement,
    part_name: &str,
) -> UseResult<Option<NativeSpreadsheetFilterCriteria>> {
    Ok(match element.local_name.as_str() {
        "filters" => parse_values(element, part_name)?,
        "customFilters" => parse_custom_filters(element, part_name)?,
        "top10" => parse_top(element, part_name)?,
        "dynamicFilter" => parse_dynamic(element),
        _ => None,
    })
}

fn parse_values(
    element: &XmlElement,
    part_name: &str,
) -> UseResult<Option<NativeSpreadsheetFilterCriteria>> {
    if element
        .attributes
        .iter()
        .any(|attribute| attribute.namespace.is_some() || attribute.local_name != "blank")
    {
        return Ok(None);
    }
    let include_blanks = boolean_attribute(element, "blank", false, part_name)?;
    let mut values = Vec::new();
    for child in element.child_elements() {
        if child.namespace != element.namespace
            || child.local_name != "filter"
            || child
                .attributes
                .iter()
                .any(|attribute| attribute.namespace.is_some() || attribute.local_name != "val")
            || child.child_elements().next().is_some()
        {
            return Ok(None);
        }
        let Some(value) = child.attribute("val") else {
            return Ok(None);
        };
        values.push(value.to_string());
    }
    if values.is_empty() && include_blanks {
        Ok(Some(NativeSpreadsheetFilterCriteria::Blanks))
    } else if values.is_empty() {
        Ok(None)
    } else {
        Ok(Some(NativeSpreadsheetFilterCriteria::Values {
            values,
            include_blanks,
        }))
    }
}

fn parse_custom_filters(
    element: &XmlElement,
    part_name: &str,
) -> UseResult<Option<NativeSpreadsheetFilterCriteria>> {
    if element
        .attributes
        .iter()
        .any(|attribute| attribute.namespace.is_some() || attribute.local_name != "and")
    {
        return Ok(None);
    }
    let and = boolean_attribute(element, "and", false, part_name)?;
    let filters = element
        .children_named("customFilter")
        .filter(|filter| filter.namespace == element.namespace)
        .collect::<Vec<_>>();
    if filters.len() != element.child_elements().count() || !(1..=2).contains(&filters.len()) {
        return Ok(None);
    }
    let parsed = filters
        .iter()
        .map(|filter| parse_custom_filter(filter))
        .collect::<Option<Vec<_>>>();
    let Some(parsed) = parsed else {
        return Ok(None);
    };
    if parsed.len() == 1 {
        let (operator, value) = &parsed[0];
        return Ok(match operator.as_str() {
            "equal" => decode_custom_pattern(value, false),
            "notEqual" if value.is_empty() => Some(NativeSpreadsheetFilterCriteria::NonBlanks),
            "notEqual" => decode_custom_pattern(value, true),
            "greaterThan" => Some(NativeSpreadsheetFilterCriteria::GreaterThan {
                value: value.clone(),
            }),
            "greaterThanOrEqual" => Some(NativeSpreadsheetFilterCriteria::GreaterThanOrEqual {
                value: value.clone(),
            }),
            "lessThan" => Some(NativeSpreadsheetFilterCriteria::LessThan {
                value: value.clone(),
            }),
            "lessThanOrEqual" => Some(NativeSpreadsheetFilterCriteria::LessThanOrEqual {
                value: value.clone(),
            }),
            _ => None,
        });
    }
    let first = (&parsed[0].0, &parsed[0].1);
    let second = (&parsed[1].0, &parsed[1].1);
    if and && first.0 == "greaterThanOrEqual" && second.0 == "lessThanOrEqual" {
        Ok(Some(NativeSpreadsheetFilterCriteria::Between {
            lower: first.1.clone(),
            upper: second.1.clone(),
        }))
    } else if !and && first.0 == "lessThan" && second.0 == "greaterThan" {
        Ok(Some(NativeSpreadsheetFilterCriteria::NotBetween {
            lower: first.1.clone(),
            upper: second.1.clone(),
        }))
    } else {
        Ok(None)
    }
}

fn parse_custom_filter(element: &XmlElement) -> Option<(String, String)> {
    if element.attributes.iter().any(|attribute| {
        attribute.namespace.is_some()
            || !matches!(attribute.local_name.as_str(), "operator" | "val")
    }) || element.child_elements().next().is_some()
    {
        return None;
    }
    Some((
        element.attribute("operator").unwrap_or("equal").to_string(),
        element.attribute("val")?.to_string(),
    ))
}

fn parse_top(
    element: &XmlElement,
    part_name: &str,
) -> UseResult<Option<NativeSpreadsheetFilterCriteria>> {
    if element.attributes.iter().any(|attribute| {
        attribute.namespace.is_some()
            || !matches!(attribute.local_name.as_str(), "top" | "percent" | "val")
    }) || element.child_elements().next().is_some()
    {
        return Ok(None);
    }
    let top = boolean_attribute(element, "top", true, part_name)?;
    let percent = boolean_attribute(element, "percent", false, part_name)?;
    let Some(value) = element
        .attribute("val")
        .and_then(|value| value.parse::<u16>().ok())
    else {
        return Ok(None);
    };
    Ok(Some(match (top, percent) {
        (true, false) => NativeSpreadsheetFilterCriteria::Top { count: value },
        (true, true) => NativeSpreadsheetFilterCriteria::TopPercent {
            percent: u8::try_from(value).unwrap_or(u8::MAX),
        },
        (false, false) => NativeSpreadsheetFilterCriteria::Bottom { count: value },
        (false, true) => NativeSpreadsheetFilterCriteria::BottomPercent {
            percent: u8::try_from(value).unwrap_or(u8::MAX),
        },
    }))
}

fn parse_dynamic(element: &XmlElement) -> Option<NativeSpreadsheetFilterCriteria> {
    if element
        .attributes
        .iter()
        .any(|attribute| attribute.namespace.is_some() || attribute.local_name != "type")
        || element.child_elements().next().is_some()
    {
        return None;
    }
    Some(NativeSpreadsheetFilterCriteria::Dynamic {
        kind: NativeSpreadsheetDynamicFilter::from_ooxml_name(element.attribute("type")?)?,
    })
}

fn apply_criteria(node: &mut DocumentNode, criteria: &NativeSpreadsheetFilterCriteria) {
    let kind = criteria_type(criteria);
    node.text = kind.to_string();
    node.format.insert("criteriaType".into(), kind.into());
    match criteria {
        NativeSpreadsheetFilterCriteria::Values {
            values,
            include_blanks,
        } => {
            node.format
                .insert("includeBlanks".into(), include_blanks.to_string());
            for (index, value) in values.iter().enumerate() {
                let mut child = DocumentNode::new(
                    format!("{}/filterValue[{}]", node.path, index + 1),
                    "filtervalue",
                    OfficeNodeType::FilterValue,
                );
                child.text = value.clone();
                node.children.push(child);
            }
        }
        NativeSpreadsheetFilterCriteria::Equals { value }
        | NativeSpreadsheetFilterCriteria::NotEquals { value }
        | NativeSpreadsheetFilterCriteria::Contains { value }
        | NativeSpreadsheetFilterCriteria::DoesNotContain { value }
        | NativeSpreadsheetFilterCriteria::BeginsWith { value }
        | NativeSpreadsheetFilterCriteria::EndsWith { value }
        | NativeSpreadsheetFilterCriteria::GreaterThan { value }
        | NativeSpreadsheetFilterCriteria::GreaterThanOrEqual { value }
        | NativeSpreadsheetFilterCriteria::LessThan { value }
        | NativeSpreadsheetFilterCriteria::LessThanOrEqual { value } => {
            node.format.insert("value".into(), value.clone());
        }
        NativeSpreadsheetFilterCriteria::Between { lower, upper }
        | NativeSpreadsheetFilterCriteria::NotBetween { lower, upper } => {
            node.format.insert("lower".into(), lower.clone());
            node.format.insert("upper".into(), upper.clone());
        }
        NativeSpreadsheetFilterCriteria::Top { count }
        | NativeSpreadsheetFilterCriteria::Bottom { count } => {
            node.format.insert("count".into(), count.to_string());
        }
        NativeSpreadsheetFilterCriteria::TopPercent { percent }
        | NativeSpreadsheetFilterCriteria::BottomPercent { percent } => {
            node.format.insert("percent".into(), percent.to_string());
        }
        NativeSpreadsheetFilterCriteria::Dynamic { kind } => {
            node.format
                .insert("dynamicKind".into(), kind.ooxml_name().into());
        }
        NativeSpreadsheetFilterCriteria::Blanks | NativeSpreadsheetFilterCriteria::NonBlanks => {}
    }
}

fn boolean_attribute(
    element: &XmlElement,
    name: &str,
    default: bool,
    part_name: &str,
) -> UseResult<bool> {
    match element.attribute(name) {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(filter_error(
            part_name,
            format!("contains invalid boolean {name} '{value}'"),
        )),
    }
}

fn filter_error(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(
        "use.office.spreadsheet_filter_invalid",
        format!(
            "Spreadsheet AutoFilter in part '{part_name}' {}.",
            reason.into()
        ),
    )
    .with_detail("part", part_name)
}
