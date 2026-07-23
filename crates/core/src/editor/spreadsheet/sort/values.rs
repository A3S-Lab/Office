use std::cmp::Ordering;
use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::node_not_found;
use crate::semantic::{DocumentNode, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::NativeSpreadsheetSortDirection;

#[derive(Debug, Clone)]
struct SortRecord {
    row: u32,
    values: Vec<SortValue>,
}

#[derive(Debug, Clone)]
enum SortValue {
    Number(f64),
    Text(String),
    Blank,
}

pub(super) fn permutation(
    sheet: &DocumentNode,
    data_range: CellRange,
    keys: &[(u32, NativeSpreadsheetSortDirection)],
    case_sensitive: bool,
) -> UseResult<BTreeMap<u32, u32>> {
    let values = cell_values(sheet, case_sensitive)?;
    let mut records = (data_range.start.row..=data_range.end.row)
        .map(|row| SortRecord {
            row,
            values: keys
                .iter()
                .map(|(column, _)| {
                    values
                        .get(&CellReference {
                            column: *column,
                            row,
                        })
                        .cloned()
                        .unwrap_or(SortValue::Blank)
                })
                .collect(),
        })
        .collect::<Vec<_>>();
    records.sort_by(|left, right| compare_records(left, right, keys));
    Ok(records
        .iter()
        .zip(data_range.start.row..=data_range.end.row)
        .map(|(record, target)| (record.row, target))
        .collect())
}

fn cell_values(
    sheet: &DocumentNode,
    case_sensitive: bool,
) -> UseResult<BTreeMap<CellReference, SortValue>> {
    let mut values = BTreeMap::new();
    for cell in sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Row)
        .flat_map(|row| row.children.iter())
        .filter(|node| node.node_type == OfficeNodeType::Cell)
    {
        let reference = cell
            .path
            .rsplit_once('/')
            .ok_or_else(|| node_not_found(&cell.path))
            .and_then(|(_, reference)| CellReference::parse(reference))?;
        let value_type = cell.format.get("valueType").map(String::as_str);
        let value = if cell.text.is_empty() {
            SortValue::Blank
        } else {
            match value_type {
                Some("Number") => cell
                    .text
                    .parse::<f64>()
                    .ok()
                    .filter(|value| value.is_finite())
                    .map_or_else(|| text_value(&cell.text, case_sensitive), SortValue::Number),
                Some("Boolean") => match cell.text.as_str() {
                    "false" | "0" => SortValue::Number(0.0),
                    "true" | "1" => SortValue::Number(1.0),
                    _ => text_value(&cell.text, case_sensitive),
                },
                _ => text_value(&cell.text, case_sensitive),
            }
        };
        values.insert(reference, value);
    }
    Ok(values)
}

fn text_value(value: &str, case_sensitive: bool) -> SortValue {
    SortValue::Text(if case_sensitive {
        value.to_string()
    } else {
        value.to_lowercase()
    })
}

fn compare_records(
    left: &SortRecord,
    right: &SortRecord,
    keys: &[(u32, NativeSpreadsheetSortDirection)],
) -> Ordering {
    for (index, (_, direction)) in keys.iter().enumerate() {
        let ordering = compare_value(&left.values[index], &right.values[index], *direction);
        if ordering != Ordering::Equal {
            return ordering;
        }
    }
    Ordering::Equal
}

fn compare_value(
    left: &SortValue,
    right: &SortValue,
    direction: NativeSpreadsheetSortDirection,
) -> Ordering {
    let left_rank = value_rank(left);
    let right_rank = value_rank(right);
    if left_rank != right_rank {
        return left_rank.cmp(&right_rank);
    }
    let ordering = match (left, right) {
        (SortValue::Number(left), SortValue::Number(right)) => left.total_cmp(right),
        (SortValue::Text(left), SortValue::Text(right)) => left.cmp(right),
        _ => Ordering::Equal,
    };
    if direction == NativeSpreadsheetSortDirection::Descending {
        ordering.reverse()
    } else {
        ordering
    }
}

const fn value_rank(value: &SortValue) -> u8 {
    match value {
        SortValue::Number(_) => 0,
        SortValue::Text(_) => 1,
        SortValue::Blank => 2,
    }
}
