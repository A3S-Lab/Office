//! Rewrite merge anchors, chart references, and table ranges in the same
//! structural frame as cell formulas. Pixel boxes stay where they are.
//! A corrupt address or a three-dimensional reference plans nothing.

use std::collections::{BTreeMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::{json, Value as JsonValue};
use yrs::{Any, Array, Map, MapRef, Out, ReadTxn, Transact};

use super::json::{any_to_json, encode_flat_json_key, json_equal, json_to_any, FlatJsonEntryKind};
use super::state::{
    SpreadsheetArrayRemoval, SpreadsheetKeyRemoval, SpreadsheetReferenceWrite,
    SpreadsheetSheetState,
};
use super::structure::{map_index, structure_too_large};
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, MAX_SPREADSHEET_COLUMNS,
    MAX_SPREADSHEET_ROWS,
};
use crate::spreadsheet_formula::{rewrite_formula_references, ReferenceAxis, ReferenceEdit};

const CONFIG_KEY: &str = "config";
const CHARTS_KEY: &str = "charts";
const TABLES_KEY: &str = "tables";
const TABLE_ORDER_KEY: &str = "tableOrder";
const REFERENCE_FIELDS: &[&str] = &[
    "titleReference",
    "categoryReference",
    "nameReference",
    "valuesReference",
    "xValuesReference",
    "bubbleSizesReference",
    "calculatedFormula",
    "totalsFormula",
];

pub(super) struct RegionAddressEdits {
    pub(super) writes: Vec<SpreadsheetReferenceWrite>,
    pub(super) removals: Vec<SpreadsheetKeyRemoval>,
    pub(super) array_removals: Vec<SpreadsheetArrayRemoval>,
}

#[allow(clippy::too_many_arguments)]
pub(super) fn rewrite_region_addresses(
    doc: &yrs::Doc,
    loaded: &[(String, String, SpreadsheetSheetState)],
    target_sheet_id: &str,
    target_sheet_name: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    at: u32,
    count: u32,
    insert: bool,
) -> UseResult<RegionAddressEdits> {
    let mut edits = RegionAddressEdits {
        writes: Vec::new(),
        removals: Vec::new(),
        array_removals: Vec::new(),
    };
    let transaction = doc.transact();
    for (sheet_id, sheet_name, state) in loaded {
        let target = sheet_id == target_sheet_id;
        if target {
            if let Some(config) = optional_map(&state.record, &transaction, CONFIG_KEY)? {
                rewrite_merges(&config, &transaction, axis, at, count, insert, &mut edits)?;
            }
        }
        if let Some(charts) = optional_map(&state.record, &transaction, CHARTS_KEY)? {
            for (_, chart) in child_maps(&charts, &transaction, "chart")? {
                rewrite_formula_fields(
                    &chart,
                    &transaction,
                    sheet_name,
                    target_sheet_name,
                    axis,
                    edit,
                    &mut edits.writes,
                )?;
            }
        }
        if let Some(tables) = optional_map(&state.record, &transaction, TABLES_KEY)? {
            let order = optional_array(&state.record, &transaction, TABLE_ORDER_KEY)?;
            rewrite_tables(
                &tables,
                order.as_ref(),
                &transaction,
                sheet_name,
                target_sheet_name,
                target,
                axis,
                edit,
                at,
                count,
                insert,
                &mut edits,
            )?;
        }
    }
    Ok(edits)
}

fn rewrite_merges<T: ReadTxn>(
    config: &MapRef,
    transaction: &T,
    axis: ReferenceAxis,
    at: u32,
    count: u32,
    insert: bool,
    edits: &mut RegionAddressEdits,
) -> UseResult<()> {
    let entries = read_flat_map(config, transaction)?;
    let mut parent = false;
    let mut groups: BTreeMap<String, MergeGroup> = BTreeMap::new();
    for entry in &entries {
        if entry.path.first().map(String::as_str) != Some("merge") {
            continue;
        }
        match (entry.kind, entry.path.as_slice()) {
            (FlatJsonEntryKind::Object, [single]) if single == "merge" => {
                require_object_marker(&entry.value)?;
                parent = true;
            }
            (FlatJsonEntryKind::Object, [_, id]) => {
                require_object_marker(&entry.value)?;
                groups.entry(id.clone()).or_default().marker = true;
            }
            (FlatJsonEntryKind::Value, [_, id, field]) => {
                let group = groups.entry(id.clone()).or_default();
                let limit = if field == "r" || field == "rs" {
                    MAX_SPREADSHEET_ROWS
                } else if field == "c" || field == "cs" {
                    MAX_SPREADSHEET_COLUMNS
                } else {
                    return Err(invalid_shared_spreadsheet(
                        "A shared Spreadsheet merge contains an unknown field.",
                    ));
                };
                let number = if field == "rs" || field == "cs" {
                    any_span(&entry.value, limit)?
                } else {
                    any_index(&entry.value, limit)?
                };
                assign_merge_field(group, field, number)?;
            }
            _ => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet merge address is not a flat record.",
                ))
            }
        }
    }
    if groups.is_empty() {
        return Ok(());
    }
    if !parent {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet merge is missing its object marker.",
        ));
    }

    let mut desired: BTreeMap<String, (u32, u32, u32, u32)> = BTreeMap::new();
    for (id, group) in &groups {
        let (r, c, rs, cs) = complete_merge(id, group)?;
        let next = shift_merge(r, c, rs, cs, axis, at, count, insert)?;
        let Some((r, c, rs, cs)) = next else {
            continue;
        };
        let key = format!("{r}_{c}");
        if desired.insert(key, (r, c, rs, cs)).is_some() {
            return Err(invalid_spreadsheet_mutation(
                "A Spreadsheet structure edit would place two merges on the same anchor.",
            ));
        }
    }

    let mut keep = HashSet::new();
    if !desired.is_empty() {
        let parent_key = flat_key(FlatJsonEntryKind::Object, &["merge"])?;
        keep.insert(parent_key.clone());
        push_if_changed(
            &mut edits.writes,
            config,
            parent_key,
            Any::Bool(true),
            find_any(&entries, &["merge"], FlatJsonEntryKind::Object),
        )?;
    }
    for (id, (r, c, rs, cs)) in &desired {
        let marker = flat_key(FlatJsonEntryKind::Object, &["merge", id])?;
        keep.insert(marker.clone());
        push_if_changed(
            &mut edits.writes,
            config,
            marker,
            Any::Bool(true),
            find_any(&entries, &["merge", id], FlatJsonEntryKind::Object),
        )?;
        for (field, number) in [("r", *r), ("c", *c), ("rs", *rs), ("cs", *cs)] {
            let key = flat_key(FlatJsonEntryKind::Value, &["merge", id, field])?;
            keep.insert(key.clone());
            push_if_changed(
                &mut edits.writes,
                config,
                key,
                Any::Number(f64::from(number)),
                find_any(&entries, &["merge", id, field], FlatJsonEntryKind::Value),
            )?;
        }
    }
    for entry in &entries {
        if entry.path.first().map(String::as_str) == Some("merge") && !keep.contains(&entry.key) {
            edits.removals.push(SpreadsheetKeyRemoval {
                map: config.clone(),
                key: entry.key.clone(),
            });
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn rewrite_tables<T: ReadTxn>(
    tables: &MapRef,
    order: Option<&yrs::ArrayRef>,
    transaction: &T,
    current_sheet: &str,
    target_sheet: &str,
    target: bool,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    at: u32,
    count: u32,
    insert: bool,
    edits: &mut RegionAddressEdits,
) -> UseResult<()> {
    for (id, table) in child_maps(tables, transaction, "table")? {
        let removed = rewrite_table(
            &table,
            transaction,
            current_sheet,
            target_sheet,
            target,
            axis,
            edit,
            at,
            count,
            insert,
            &mut edits.writes,
        )?;
        if !removed {
            continue;
        }
        edits.removals.push(SpreadsheetKeyRemoval {
            map: tables.clone(),
            key: id.clone(),
        });
        if let Some(order) = order {
            let indexes = array_indexes_of(order, transaction, &id)?;
            if indexes.len() > 1 {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet table order lists the same table twice.",
                ));
            }
            if let Some(index) = indexes.first().copied() {
                edits.array_removals.push(SpreadsheetArrayRemoval {
                    array: order.clone(),
                    index,
                });
            }
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn rewrite_table<T: ReadTxn>(
    table: &MapRef,
    transaction: &T,
    current_sheet: &str,
    target_sheet: &str,
    target: bool,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    at: u32,
    count: u32,
    insert: bool,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<bool> {
    let entries = read_flat_map(table, transaction)?;
    let row = match find_json(&entries, &["range", "row"])? {
        Some(value) => Some(inclusive_pair(&value, MAX_SPREADSHEET_ROWS)?),
        None => None,
    };
    let column = match find_json(&entries, &["range", "column"])? {
        Some(value) => Some(inclusive_pair(&value, MAX_SPREADSHEET_COLUMNS)?),
        None => None,
    };
    if target && (row.is_none() || column.is_none()) {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table is missing its range.",
        ));
    }
    let (mut row_start, mut row_end) = row.unwrap_or((0, 0));
    let (mut column_start, mut column_end) = column.unwrap_or((0, 0));
    if target {
        let (next_row, next_column) = match axis {
            ReferenceAxis::Row => {
                let next =
                    shift_inclusive(row_start, row_end, at, count, insert, MAX_SPREADSHEET_ROWS)?;
                (next, Some((column_start, column_end)))
            }
            ReferenceAxis::Column => {
                let next = shift_inclusive(
                    column_start,
                    column_end,
                    at,
                    count,
                    insert,
                    MAX_SPREADSHEET_COLUMNS,
                )?;
                (Some((row_start, row_end)), next)
            }
        };
        let (Some(next_row), Some(next_column)) = (next_row, next_column) else {
            return Ok(true);
        };
        row_start = next_row.0;
        row_end = next_row.1;
        column_start = next_column.0;
        column_end = next_column.1;
        push_range(writes, table, &entries, "row", row_start, row_end)?;
        push_range(writes, table, &entries, "column", column_start, column_end)?;
    }

    for entry in &entries {
        if entry.kind != FlatJsonEntryKind::Value || path_is(&entry.path, &["columns"]) {
            continue;
        }
        if path_is(&entry.path, &["range", "row"]) || path_is(&entry.path, &["range", "column"]) {
            continue;
        }
        if let Some(next) = rewritten_leaf(
            &entry.value,
            &entry.path,
            current_sheet,
            target_sheet,
            axis,
            edit,
        )? {
            writes.push(SpreadsheetReferenceWrite {
                map: table.clone(),
                key: entry.key.clone(),
                value: next,
            });
        }
    }

    if let Some(columns) = find_json(&entries, &["columns"])? {
        let original = columns.clone();
        let mut next = columns;
        rewrite_formula_json(&mut next, current_sheet, target_sheet, axis, edit)?;
        if target && axis == ReferenceAxis::Column {
            let (old_start, old_end) = column.ok_or_else(|| {
                invalid_shared_spreadsheet("A shared Spreadsheet table is missing its range.")
            })?;
            next = realign_columns(
                next,
                old_start,
                old_end,
                column_start,
                column_end,
                at,
                count,
                insert,
            )?;
        }
        if !json_equal(&original, &next) {
            let key = flat_key(FlatJsonEntryKind::Value, &["columns"])?;
            writes.push(SpreadsheetReferenceWrite {
                map: table.clone(),
                key,
                value: json_to_any(&next, 0)?,
            });
        }
    } else if target && axis == ReferenceAxis::Column {
        let (old_start, old_end) = column.unwrap_or((column_start, column_end));
        if old_end - old_start != column_end - column_start {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet table is missing the columns for its range.",
            ));
        }
    }
    Ok(false)
}

fn rewrite_formula_fields<T: ReadTxn>(
    map: &MapRef,
    transaction: &T,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    for entry in read_flat_map(map, transaction)? {
        if entry.kind != FlatJsonEntryKind::Value {
            continue;
        }
        if let Some(next) = rewritten_leaf(
            &entry.value,
            &entry.path,
            current_sheet,
            target_sheet,
            axis,
            edit,
        )? {
            writes.push(SpreadsheetReferenceWrite {
                map: map.clone(),
                key: entry.key,
                value: next,
            });
        }
    }
    Ok(())
}

fn rewritten_leaf(
    value: &Any,
    path: &[String],
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Option<Any>> {
    let json = any_to_json(value.clone(), 0)?;
    if path
        .last()
        .is_some_and(|field| REFERENCE_FIELDS.contains(&field.as_str()))
    {
        let JsonValue::String(text) = json else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet reference field is not text.",
            ));
        };
        let next =
            rewrite_formula_references(&text, Some(current_sheet), target_sheet, axis, edit)?;
        if next == text {
            return Ok(None);
        }
        return Ok(Some(json_to_any(&JsonValue::String(next), 0)?));
    }
    if json.is_array() || json.is_object() {
        let mut next = json.clone();
        rewrite_formula_json(&mut next, current_sheet, target_sheet, axis, edit)?;
        if json_equal(&json, &next) {
            return Ok(None);
        }
        return Ok(Some(json_to_any(&next, 0)?));
    }
    Ok(None)
}

fn rewrite_formula_json(
    value: &mut JsonValue,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    match value {
        JsonValue::Object(object) => {
            let keys = object.keys().cloned().collect::<Vec<_>>();
            for key in keys {
                if REFERENCE_FIELDS.contains(&key.as_str()) {
                    let Some(child) = object.get_mut(&key) else {
                        continue;
                    };
                    match child {
                        JsonValue::String(text) => {
                            let next = rewrite_formula_references(
                                text,
                                Some(current_sheet),
                                target_sheet,
                                axis,
                                edit,
                            )?;
                            if next != *text {
                                *child = JsonValue::String(next);
                            }
                        }
                        JsonValue::Null => {}
                        _ => {
                            return Err(invalid_shared_spreadsheet(
                                "A shared Spreadsheet reference field is not text.",
                            ))
                        }
                    }
                    continue;
                }
                if let Some(child) = object.get_mut(&key) {
                    rewrite_formula_json(child, current_sheet, target_sheet, axis, edit)?;
                }
            }
        }
        JsonValue::Array(items) => {
            for item in items {
                rewrite_formula_json(item, current_sheet, target_sheet, axis, edit)?;
            }
        }
        _ => {}
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn realign_columns(
    columns: JsonValue,
    old_start: u32,
    old_end: u32,
    new_start: u32,
    new_end: u32,
    at: u32,
    count: u32,
    insert: bool,
) -> UseResult<JsonValue> {
    let JsonValue::Array(columns) = columns else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table columns field is not an array.",
        ));
    };
    let old_span = old_end
        .checked_sub(old_start)
        .and_then(|span| span.checked_add(1));
    if old_span != Some(columns.len() as u32) {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table columns field does not match its range.",
        ));
    }
    let new_span = new_end - new_start + 1;
    let mut slots = vec![None; new_span as usize];
    let mut names = HashSet::new();
    for (index, column) in columns.into_iter().enumerate() {
        let sheet_column = old_start + index as u32;
        let Some(mapped) = map_index(sheet_column, at, count, insert, MAX_SPREADSHEET_COLUMNS)?
        else {
            continue;
        };
        if mapped < new_start || mapped > new_end {
            return Err(invalid_spreadsheet_mutation(
                "A Spreadsheet table column moved outside its rewritten range.",
            ));
        }
        if let Some(name) = column.get("name").and_then(JsonValue::as_str) {
            names.insert(name.to_owned());
        }
        let offset = (mapped - new_start) as usize;
        if slots[offset].is_some() {
            return Err(invalid_spreadsheet_mutation(
                "A Spreadsheet structure edit would place two table columns on the same address.",
            ));
        }
        slots[offset] = Some(column);
    }
    let aligned = slots
        .into_iter()
        .enumerate()
        .map(|(offset, column)| {
            if let Some(column) = column {
                return Ok(column);
            }
            let sheet_column = new_start + offset as u32;
            Ok(placeholder_column(sheet_column, &mut names))
        })
        .collect::<UseResult<Vec<_>>>()?;
    Ok(JsonValue::Array(aligned))
}

fn placeholder_column(sheet_column: u32, names: &mut HashSet<String>) -> JsonValue {
    let mut name = format!("Column{}", sheet_column + 1);
    if names.contains(&name) {
        name = format!("{name}_{sheet_column}");
    }
    names.insert(name.clone());
    json!({ "name": name })
}

fn push_range(
    writes: &mut Vec<SpreadsheetReferenceWrite>,
    table: &MapRef,
    entries: &[FlatEntry],
    field: &str,
    start: u32,
    end: u32,
) -> UseResult<()> {
    let next = JsonValue::Array(vec![JsonValue::from(start), JsonValue::from(end)]);
    if find_json(entries, &["range", field])?.is_some_and(|current| json_equal(&current, &next)) {
        return Ok(());
    }
    writes.push(SpreadsheetReferenceWrite {
        map: table.clone(),
        key: flat_key(FlatJsonEntryKind::Value, &["range", field])?,
        value: json_to_any(&next, 0)?,
    });
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn shift_merge(
    row: u32,
    column: u32,
    row_span: u32,
    column_span: u32,
    axis: ReferenceAxis,
    at: u32,
    count: u32,
    insert: bool,
) -> UseResult<Option<(u32, u32, u32, u32)>> {
    let row_end = inclusive_end(row, row_span, MAX_SPREADSHEET_ROWS)?;
    let column_end = inclusive_end(column, column_span, MAX_SPREADSHEET_COLUMNS)?;
    let (next_row, next_row_end, next_column, next_column_end) = match axis {
        ReferenceAxis::Row => {
            let Some((start, end)) =
                shift_inclusive(row, row_end, at, count, insert, MAX_SPREADSHEET_ROWS)?
            else {
                return Ok(None);
            };
            (start, end, column, column_end)
        }
        ReferenceAxis::Column => {
            let Some((start, end)) = shift_inclusive(
                column,
                column_end,
                at,
                count,
                insert,
                MAX_SPREADSHEET_COLUMNS,
            )?
            else {
                return Ok(None);
            };
            (row, row_end, start, end)
        }
    };
    let next_row_span = next_row_end - next_row + 1;
    let next_column_span = next_column_end - next_column + 1;
    let unchanged = next_row == row
        && next_column == column
        && next_row_span == row_span
        && next_column_span == column_span;
    // A merge the edit reduced to one cell is no longer a merge. An address
    // that was already one cell, and did not move, stays where it is.
    if next_row_span == 1 && next_column_span == 1 && !unchanged {
        return Ok(None);
    }
    Ok(Some((
        next_row,
        next_column,
        next_row_span,
        next_column_span,
    )))
}

fn shift_inclusive(
    start: u32,
    end: u32,
    at: u32,
    count: u32,
    insert: bool,
    limit: u32,
) -> UseResult<Option<(u32, u32)>> {
    if end < start || end >= limit {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address range is outside the sheet.",
        ));
    }
    let mut mapped_start = None;
    let mut mapped_end = None;
    let mut index = start;
    loop {
        if let Some(mapped) = map_index(index, at, count, insert, limit)? {
            if mapped_start.is_none() {
                mapped_start = Some(mapped);
            }
            mapped_end = Some(mapped);
        }
        if index == end {
            break;
        }
        index = index.checked_add(1).ok_or_else(structure_too_large)?;
    }
    Ok(mapped_start.zip(mapped_end))
}

struct FlatEntry {
    key: String,
    kind: FlatJsonEntryKind,
    path: Vec<String>,
    value: Any,
}

#[derive(Default)]
struct MergeGroup {
    marker: bool,
    row: Option<u32>,
    column: Option<u32>,
    row_span: Option<u32>,
    column_span: Option<u32>,
}

fn read_flat_map<T: ReadTxn>(map: &MapRef, transaction: &T) -> UseResult<Vec<FlatEntry>> {
    let mut entries = Vec::new();
    for (key, value) in map.iter(transaction) {
        let (kind, path) = super::json::decode_flat_json_key(key).map_err(|_| {
            invalid_shared_spreadsheet("A shared Spreadsheet record field identity is invalid.")
        })?;
        let Out::Any(value) = value else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet record field is not a JSON value.",
            ));
        };
        entries.push(FlatEntry {
            key: key.to_owned(),
            kind,
            path,
            value,
        });
    }
    Ok(entries)
}

fn optional_map<T: ReadTxn>(
    record: &MapRef,
    transaction: &T,
    key: &str,
) -> UseResult<Option<MapRef>> {
    match record.get(transaction, key) {
        None => Ok(None),
        Some(Out::YMap(map)) => Ok(Some(map)),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet '{key}' is not a typed map."
        ))),
    }
}

fn optional_array<T: ReadTxn>(
    record: &MapRef,
    transaction: &T,
    key: &str,
) -> UseResult<Option<yrs::ArrayRef>> {
    match record.get(transaction, key) {
        None => Ok(None),
        Some(Out::YArray(array)) => Ok(Some(array)),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet '{key}' is not a typed array."
        ))),
    }
}

fn child_maps<T: ReadTxn>(
    parent: &MapRef,
    transaction: &T,
    label: &str,
) -> UseResult<Vec<(String, MapRef)>> {
    let mut children = Vec::new();
    for (id, value) in parent.iter(transaction) {
        let Out::YMap(map) = value else {
            return Err(invalid_shared_spreadsheet(format!(
                "A shared Spreadsheet {label} is not a typed map."
            )));
        };
        children.push((id.to_owned(), map));
    }
    Ok(children)
}

fn array_indexes_of<T: ReadTxn>(
    array: &yrs::ArrayRef,
    transaction: &T,
    id: &str,
) -> UseResult<Vec<u32>> {
    let mut indexes = Vec::new();
    for index in 0..array.len(transaction) {
        match array.get(transaction, index) {
            Some(Out::Any(Any::String(value))) if value.as_ref() == id => indexes.push(index),
            Some(Out::Any(Any::String(_))) => {}
            _ => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet table order entry is not text.",
                ))
            }
        }
    }
    Ok(indexes)
}

fn complete_merge(id: &str, group: &MergeGroup) -> UseResult<(u32, u32, u32, u32)> {
    if !group.marker {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet merge is missing its object marker.",
        ));
    }
    let (Some(row), Some(column), Some(row_span), Some(column_span)) =
        (group.row, group.column, group.row_span, group.column_span)
    else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet merge is missing an anchor or a span.",
        ));
    };
    inclusive_end(row, row_span, MAX_SPREADSHEET_ROWS)?;
    inclusive_end(column, column_span, MAX_SPREADSHEET_COLUMNS)?;
    if id != format!("{row}_{column}") {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet merge key does not match its anchor.",
        ));
    }
    Ok((row, column, row_span, column_span))
}

fn assign_merge_field(group: &mut MergeGroup, field: &str, number: u32) -> UseResult<()> {
    let slot = match field {
        "r" => &mut group.row,
        "c" => &mut group.column,
        "rs" => &mut group.row_span,
        "cs" => &mut group.column_span,
        _ => {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet merge contains an unknown field.",
            ))
        }
    };
    if slot.replace(number).is_some() {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet merge repeats an anchor field.",
        ));
    }
    Ok(())
}

fn inclusive_end(start: u32, span: u32, limit: u32) -> UseResult<u32> {
    let end = start
        .checked_add(span)
        .and_then(|end| end.checked_sub(1))
        .ok_or_else(structure_too_large)?;
    if end >= limit {
        return Err(structure_too_large());
    }
    Ok(end)
}

fn inclusive_pair(value: &JsonValue, limit: u32) -> UseResult<(u32, u32)> {
    let JsonValue::Array(items) = value else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is not a pair.",
        ));
    };
    if items.len() != 2 {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is not a pair.",
        ));
    }
    let start = json_index(&items[0], limit)?;
    let end = json_index(&items[1], limit)?;
    if end < start {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is reversed.",
        ));
    }
    Ok((start, end))
}

fn json_index(value: &JsonValue, limit: u32) -> UseResult<u32> {
    let Some(number) = value.as_f64() else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    };
    any_index(&Any::Number(number), limit)
}

fn any_index(value: &Any, limit: u32) -> UseResult<u32> {
    let integer = finite_u32(value)?;
    if integer >= limit {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is outside the sheet.",
        ));
    }
    Ok(integer)
}

fn any_span(value: &Any, limit: u32) -> UseResult<u32> {
    let integer = finite_u32(value)?;
    if integer == 0 || integer > limit {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet merge span is outside the sheet.",
        ));
    }
    Ok(integer)
}

fn finite_u32(value: &Any) -> UseResult<u32> {
    let Any::Number(number) = value else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    };
    if !number.is_finite() || number.fract() != 0.0 || *number < 0.0 {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    }
    let integer = *number as u32;
    if f64::from(integer) != *number {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    }
    Ok(integer)
}

fn require_object_marker(value: &Any) -> UseResult<()> {
    if matches!(value, Any::Bool(true)) {
        Ok(())
    } else {
        Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet object marker is not true.",
        ))
    }
}

fn push_if_changed(
    writes: &mut Vec<SpreadsheetReferenceWrite>,
    map: &MapRef,
    key: String,
    next: Any,
    current: Option<&Any>,
) -> UseResult<()> {
    if let Some(current) = current {
        let current = any_to_json(current.clone(), 0)?;
        let next_json = any_to_json(next.clone(), 0)?;
        if json_equal(&current, &next_json) {
            return Ok(());
        }
    }
    writes.push(SpreadsheetReferenceWrite {
        map: map.clone(),
        key,
        value: next,
    });
    Ok(())
}

fn find_any<'a>(
    entries: &'a [FlatEntry],
    path: &[&str],
    kind: FlatJsonEntryKind,
) -> Option<&'a Any> {
    entries.iter().find_map(|entry| {
        (entry.kind == kind && path_is(&entry.path, path)).then_some(&entry.value)
    })
}

fn find_json(entries: &[FlatEntry], path: &[&str]) -> UseResult<Option<JsonValue>> {
    let Some(value) = find_any(entries, path, FlatJsonEntryKind::Value) else {
        return Ok(None);
    };
    any_to_json(value.clone(), 0).map(Some)
}

fn path_is(path: &[String], expected: &[&str]) -> bool {
    path.len() == expected.len()
        && path
            .iter()
            .zip(expected)
            .all(|(part, expected)| part == expected)
}

fn flat_key(kind: FlatJsonEntryKind, path: &[&str]) -> UseResult<String> {
    let path = path.iter().copied().map(str::to_owned).collect::<Vec<_>>();
    encode_flat_json_key(kind, &path)
}
