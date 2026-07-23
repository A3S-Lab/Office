use a3s_use_core::UseResult;

use crate::spreadsheet_filter::escape_wildcard_literal;
use crate::xml_edit::{escape_attribute, IndexedXmlElement};
use crate::{
    LosslessXmlPart, NativeSpreadsheetAutoFilter, NativeSpreadsheetFilterColumn,
    NativeSpreadsheetFilterCriteria,
};

pub(super) fn fragment(
    prefix: Option<&str>,
    filter: &NativeSpreadsheetAutoFilter,
) -> UseResult<String> {
    filter.validate()?;
    let name = qualified(prefix, "autoFilter");
    let mut columns = filter.columns.iter().collect::<Vec<_>>();
    columns.sort_by_key(|column| column.column);
    if columns.is_empty() {
        return Ok(format!(
            "<{name} ref=\"{}\"/>",
            escape_attribute(&filter.range)
        ));
    }
    let children = columns
        .into_iter()
        .map(|column| filter_column_fragment(prefix, column))
        .collect::<String>();
    Ok(format!(
        "<{name} ref=\"{}\">{children}</{name}>",
        escape_attribute(&filter.range)
    ))
}

pub(super) fn validate_mutable(
    part: &LosslessXmlPart,
    filter: &IndexedXmlElement,
) -> UseResult<()> {
    if filter.qualified_attributes.keys().any(|name| name != "ref")
        || !whitespace_outside_children(part, filter)
    {
        return Err(unknown_content(part.name()));
    }
    for column in &filter.children {
        if column.local_name != "filterColumn"
            || column.namespace != filter.namespace
            || column
                .qualified_attributes
                .keys()
                .any(|name| name != "colId")
            || column.children.len() != 1
            || !whitespace_outside_children(part, column)
            || !criterion_is_owned(part, &column.children[0], &filter.namespace)
        {
            return Err(unknown_content(part.name()));
        }
    }
    Ok(())
}

fn filter_column_fragment(prefix: Option<&str>, filter: &NativeSpreadsheetFilterColumn) -> String {
    let name = qualified(prefix, "filterColumn");
    let criteria = criteria_fragment(prefix, &filter.criteria);
    format!("<{name} colId=\"{}\">{criteria}</{name}>", filter.column)
}

fn criteria_fragment(prefix: Option<&str>, criteria: &NativeSpreadsheetFilterCriteria) -> String {
    match criteria {
        NativeSpreadsheetFilterCriteria::Values {
            values,
            include_blanks,
        } => {
            let filters = qualified(prefix, "filters");
            let filter = qualified(prefix, "filter");
            let blank = if *include_blanks { " blank=\"1\"" } else { "" };
            let children = values
                .iter()
                .map(|value| format!("<{filter} val=\"{}\"/>", escape_attribute(value)))
                .collect::<String>();
            format!("<{filters}{blank}>{children}</{filters}>")
        }
        NativeSpreadsheetFilterCriteria::Equals { value } => {
            custom_single(prefix, "equal", &escape_wildcard_literal(value))
        }
        NativeSpreadsheetFilterCriteria::NotEquals { value } => {
            custom_single(prefix, "notEqual", &escape_wildcard_literal(value))
        }
        NativeSpreadsheetFilterCriteria::Contains { value } => custom_single(
            prefix,
            "equal",
            &format!("*{}*", escape_wildcard_literal(value)),
        ),
        NativeSpreadsheetFilterCriteria::DoesNotContain { value } => custom_single(
            prefix,
            "notEqual",
            &format!("*{}*", escape_wildcard_literal(value)),
        ),
        NativeSpreadsheetFilterCriteria::BeginsWith { value } => custom_single(
            prefix,
            "equal",
            &format!("{}*", escape_wildcard_literal(value)),
        ),
        NativeSpreadsheetFilterCriteria::EndsWith { value } => custom_single(
            prefix,
            "equal",
            &format!("*{}", escape_wildcard_literal(value)),
        ),
        NativeSpreadsheetFilterCriteria::GreaterThan { value } => {
            custom_single(prefix, "greaterThan", value)
        }
        NativeSpreadsheetFilterCriteria::GreaterThanOrEqual { value } => {
            custom_single(prefix, "greaterThanOrEqual", value)
        }
        NativeSpreadsheetFilterCriteria::LessThan { value } => {
            custom_single(prefix, "lessThan", value)
        }
        NativeSpreadsheetFilterCriteria::LessThanOrEqual { value } => {
            custom_single(prefix, "lessThanOrEqual", value)
        }
        NativeSpreadsheetFilterCriteria::Between { lower, upper } => custom_pair(
            prefix,
            true,
            ("greaterThanOrEqual", lower),
            ("lessThanOrEqual", upper),
        ),
        NativeSpreadsheetFilterCriteria::NotBetween { lower, upper } => {
            custom_pair(prefix, false, ("lessThan", lower), ("greaterThan", upper))
        }
        NativeSpreadsheetFilterCriteria::Blanks => {
            let name = qualified(prefix, "filters");
            format!("<{name} blank=\"1\"/>")
        }
        NativeSpreadsheetFilterCriteria::NonBlanks => custom_single(prefix, "notEqual", ""),
        NativeSpreadsheetFilterCriteria::Top { count } => top_fragment(prefix, true, false, *count),
        NativeSpreadsheetFilterCriteria::TopPercent { percent } => {
            top_fragment(prefix, true, true, u16::from(*percent))
        }
        NativeSpreadsheetFilterCriteria::Bottom { count } => {
            top_fragment(prefix, false, false, *count)
        }
        NativeSpreadsheetFilterCriteria::BottomPercent { percent } => {
            top_fragment(prefix, false, true, u16::from(*percent))
        }
        NativeSpreadsheetFilterCriteria::Dynamic { kind } => {
            let name = qualified(prefix, "dynamicFilter");
            format!("<{name} type=\"{}\"/>", escape_attribute(kind.ooxml_name()))
        }
    }
}

fn custom_single(prefix: Option<&str>, operator: &str, value: &str) -> String {
    let collection = qualified(prefix, "customFilters");
    let filter = qualified(prefix, "customFilter");
    format!(
        "<{collection}><{filter} operator=\"{operator}\" val=\"{}\"/></{collection}>",
        escape_attribute(value)
    )
}

fn custom_pair(
    prefix: Option<&str>,
    and: bool,
    first: (&str, &str),
    second: (&str, &str),
) -> String {
    let collection = qualified(prefix, "customFilters");
    let filter = qualified(prefix, "customFilter");
    format!(
        "<{collection} and=\"{}\"><{filter} operator=\"{}\" val=\"{}\"/><{filter} operator=\"{}\" val=\"{}\"/></{collection}>",
        bool_token(and),
        first.0,
        escape_attribute(first.1),
        second.0,
        escape_attribute(second.1)
    )
}

fn top_fragment(prefix: Option<&str>, top: bool, percent: bool, value: u16) -> String {
    let name = qualified(prefix, "top10");
    format!(
        "<{name} percent=\"{}\" top=\"{}\" val=\"{value}\"/>",
        bool_token(percent),
        bool_token(top)
    )
}

fn criterion_is_owned(
    part: &LosslessXmlPart,
    criterion: &IndexedXmlElement,
    namespace: &Option<String>,
) -> bool {
    if &criterion.namespace != namespace || !whitespace_outside_children(part, criterion) {
        return false;
    }
    match criterion.local_name.as_str() {
        "filters" => {
            criterion
                .qualified_attributes
                .keys()
                .all(|name| name == "blank")
                && criterion.children.iter().all(|child| {
                    child.local_name == "filter"
                        && child.namespace == *namespace
                        && child.qualified_attributes.keys().all(|name| name == "val")
                        && child.children.is_empty()
                        && whitespace_outside_children(part, child)
                })
        }
        "customFilters" => {
            criterion
                .qualified_attributes
                .keys()
                .all(|name| name == "and")
                && (1..=2).contains(&criterion.children.len())
                && criterion.children.iter().all(|child| {
                    child.local_name == "customFilter"
                        && child.namespace == *namespace
                        && child
                            .qualified_attributes
                            .keys()
                            .all(|name| matches!(name.as_str(), "operator" | "val"))
                        && child.children.is_empty()
                        && whitespace_outside_children(part, child)
                })
        }
        "top10" => {
            criterion
                .qualified_attributes
                .keys()
                .all(|name| matches!(name.as_str(), "percent" | "top" | "val"))
                && criterion.children.is_empty()
        }
        "dynamicFilter" => {
            criterion
                .qualified_attributes
                .keys()
                .all(|name| name == "type")
                && criterion.children.is_empty()
        }
        _ => false,
    }
}

fn whitespace_outside_children(part: &LosslessXmlPart, element: &IndexedXmlElement) -> bool {
    let bytes = part.parse_bytes();
    let mut cursor = element.content_range.start;
    for child in &element.children {
        if bytes
            .get(cursor..child.full_range.start)
            .is_none_or(|slice| !slice.iter().all(u8::is_ascii_whitespace))
        {
            return false;
        }
        cursor = child.full_range.end;
    }
    bytes
        .get(cursor..element.content_range.end)
        .is_some_and(|slice| slice.iter().all(u8::is_ascii_whitespace))
}

fn qualified(prefix: Option<&str>, local_name: &str) -> String {
    prefix.map_or_else(
        || local_name.to_string(),
        |prefix| format!("{prefix}:{local_name}"),
    )
}

const fn bool_token(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

fn unknown_content(part_name: &str) -> a3s_use_core::UseError {
    super::editor_error(
        "use.office.spreadsheet_filter_unknown_content",
        "The Spreadsheet AutoFilter contains sort state, date groups, color/icon filters, extensions, or other content outside the typed filter contract.",
    )
    .with_detail("part", part_name)
}
