use std::collections::BTreeSet;

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};
use crate::spreadsheet_reference::{column_name, CellRange};
use crate::xml_tree::XmlElement;
use crate::NativeSpreadsheetSortDirection;

const MAX_SORT_KEYS: usize = 64;

pub(super) fn read(
    worksheet: &XmlElement,
    part_name: &str,
    sheet_path: &str,
) -> UseResult<Option<DocumentNode>> {
    let states = worksheet
        .children_named("sortState")
        .filter(|state| state.namespace == worksheet.namespace)
        .collect::<Vec<_>>();
    if states.len() > 1 {
        return Err(sort_error(
            part_name,
            "contains multiple worksheet sortState elements",
        ));
    }
    states
        .first()
        .map(|state| read_state(state, part_name, sheet_path))
        .transpose()
}

fn read_state(state: &XmlElement, part_name: &str, sheet_path: &str) -> UseResult<DocumentNode> {
    let reference = state
        .attribute("ref")
        .ok_or_else(|| sort_error(part_name, "contains a worksheet sortState without ref"))?;
    let range = CellRange::parse(reference).map_err(|error| {
        sort_error(
            part_name,
            format!("contains invalid worksheet sortState ref '{reference}': {error}"),
        )
    })?;
    let case_sensitive = boolean_attribute(state, "caseSensitive", false, part_name)?;
    let column_sort = boolean_attribute(state, "columnSort", false, part_name)?;
    let mut mutable = !column_sort
        && state.attributes.iter().all(|attribute| {
            attribute.namespace.is_none()
                && matches!(
                    attribute.local_name.as_str(),
                    "ref" | "caseSensitive" | "columnSort"
                )
        });
    let conditions = state.child_elements().collect::<Vec<_>>();
    if conditions.is_empty() || conditions.len() > MAX_SORT_KEYS {
        mutable = false;
    }

    let mut node = DocumentNode::new(
        format!("{sheet_path}/sort"),
        "sort",
        OfficeNodeType::SortState,
    );
    node.format.insert("ref".into(), range.a1());
    node.format
        .insert("caseSensitive".into(), case_sensitive.to_string());
    let mut observed_columns = BTreeSet::new();
    let mut observed_rows = None;
    let mut labels = Vec::new();
    for (offset, condition) in conditions.into_iter().enumerate() {
        if condition.namespace != state.namespace || condition.local_name != "sortCondition" {
            mutable = false;
            continue;
        }
        let parsed = read_condition(condition, part_name, &node.path, offset + 1)?;
        mutable &= parsed.mutable;
        if !observed_columns.insert(parsed.column) {
            mutable = false;
        }
        if !(range.start.column..=range.end.column).contains(&parsed.column)
            || parsed.start_row < range.start.row
            || parsed.end_row > range.end.row
        {
            mutable = false;
        }
        match observed_rows {
            None => observed_rows = Some((parsed.start_row, parsed.end_row)),
            Some(rows) if rows != (parsed.start_row, parsed.end_row) => mutable = false,
            Some(_) => {}
        }
        labels.push(format!(
            "{} {}",
            column_name(parsed.column),
            parsed.direction.semantic_name()
        ));
        node.children.push(parsed.node);
    }
    let header = match observed_rows {
        Some((start, end)) if start == range.start.row && end == range.end.row => false,
        Some((start, end))
            if start == range.start.row.saturating_add(1) && end == range.end.row =>
        {
            true
        }
        _ => {
            mutable = false;
            false
        }
    };
    node.text = labels.join(", ");
    node.format.insert("header".into(), header.to_string());
    node.format
        .insert("keyCount".into(), node.children.len().to_string());
    node.format
        .insert("nativeMutable".into(), mutable.to_string());
    Ok(node)
}

struct ParsedCondition {
    node: DocumentNode,
    column: u32,
    start_row: u32,
    end_row: u32,
    direction: NativeSpreadsheetSortDirection,
    mutable: bool,
}

fn read_condition(
    condition: &XmlElement,
    part_name: &str,
    parent_path: &str,
    index: usize,
) -> UseResult<ParsedCondition> {
    let reference = condition
        .attribute("ref")
        .ok_or_else(|| sort_error(part_name, "contains a sortCondition without ref"))?;
    let range = CellRange::parse(reference).map_err(|error| {
        sort_error(
            part_name,
            format!("contains invalid sortCondition ref '{reference}': {error}"),
        )
    })?;
    let descending = boolean_attribute(condition, "descending", false, part_name)?;
    let direction = if descending {
        NativeSpreadsheetSortDirection::Descending
    } else {
        NativeSpreadsheetSortDirection::Ascending
    };
    let mutable = range.start.column == range.end.column
        && condition.child_elements().next().is_none()
        && condition.attributes.iter().all(|attribute| {
            attribute.namespace.is_none()
                && matches!(attribute.local_name.as_str(), "ref" | "descending")
        });
    let column = range.start.column;
    let mut node = DocumentNode::new(
        format!("{parent_path}/key[{index}]"),
        "sortkey",
        OfficeNodeType::SortKey,
    );
    node.text = format!("{} {}", column_name(column), direction.semantic_name());
    node.format.insert("column".into(), column_name(column));
    node.format.insert("ref".into(), range.a1());
    node.format
        .insert("direction".into(), direction.semantic_name().into());
    Ok(ParsedCondition {
        node,
        column,
        start_row: range.start.row,
        end_row: range.end.row,
        direction,
        mutable,
    })
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
        Some(value) => Err(sort_error(
            part_name,
            format!("contains invalid sort boolean {name}='{value}'"),
        )),
    }
}

fn sort_error(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(
        "use.office.spreadsheet_sort_invalid",
        format!(
            "Spreadsheet worksheet part '{part_name}' {}.",
            reason.into()
        ),
    )
    .with_detail("part", part_name)
}
