//! Sort is a row permutation of one rectangle. Each cell, including its
//! formula text, moves to the destination row. References outside the
//! rectangle stay as written.

use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::{json, Value as JsonValue};
use yrs::{Any, GetString, Map, MapRef, Out, ReadTxn, Transact};

use super::json::{
    any_to_json, cell_field_patches, decode_flat_json_key, json_equal, validate_cell_json,
    FlatJsonEntryKind,
};
use super::state::{
    commit_sheet_structure, encode_coordinate, read_sheet_state, SheetStructureCommit,
    SpreadsheetCellMode, SpreadsheetCellPresenceChange, SpreadsheetCellWrite,
    SpreadsheetSheetState, SpreadsheetTextMove, CELL_TEXT_KEY,
};
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, spreadsheet_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation, MAX_SPREADSHEET_COLUMNS,
    MAX_SPREADSHEET_DENSE_CELLS, MAX_SPREADSHEET_ROWS,
};

const MAX_SORT_ROWS: u32 = 1_024;
const MAX_SORT_COLUMNS: u32 = 256;
const MAX_SORT_CELLS: u64 = 4_096;
const CONFIG_KEY: &str = "config";
const TABLES_KEY: &str = "tables";
const PIVOT_TABLES_KEY: &str = "pivotTables";

struct SortFrame<'a> {
    sheet_id: &'a str,
    row: u32,
    column: u32,
    row_count: u32,
    column_count: u32,
    source_rows: &'a [u32],
    expected_cells: &'a [Option<JsonValue>],
}

pub(super) fn is_sort_mutation(mutation: &NativeOfficeCollaborationMutation) -> bool {
    matches!(
        mutation,
        NativeOfficeCollaborationMutation::SpreadsheetSortRows { .. }
    )
}

pub(super) fn validate_sort_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let frame = decode(mutation)?;
    super::cell::validate_sheet_id(frame.sheet_id)?;
    for cell in frame.expected_cells.iter().flatten() {
        validate_cell_json(cell, "expected cell")?;
    }
    Ok(())
}

pub(super) fn apply_sort_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let frame = decode(mutation)?;
    let state = read_sheet_state(doc, manifest, frame.sheet_id)?;
    refuse_structural_overlap(doc, &state, &frame)?;
    let sources = observed_sources(&state, &frame)?;
    let changes = destination_writes(&state, &frame, &sources)?;
    let text_moves = text_moves(doc, &state, &frame)?;
    let next_row_lengths = planned_row_lengths(&state, &frame, &sources)?;
    if changes.is_empty() && text_moves.is_empty() && next_row_lengths.is_none() {
        return Ok(());
    }
    commit_sheet_structure(
        doc,
        vec![SheetStructureCommit {
            state,
            changes,
            next_row_lengths,
        }],
        text_moves,
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}

fn decode(mutation: &NativeOfficeCollaborationMutation) -> UseResult<SortFrame<'_>> {
    let NativeOfficeCollaborationMutation::SpreadsheetSortRows {
        sheet_id,
        row,
        column,
        row_count,
        column_count,
        source_rows,
        expected_cells,
    } = mutation
    else {
        return Err(invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet sort.",
        ));
    };
    let row_end = row.checked_add(*row_count).ok_or_else(sort_too_large)?;
    let column_end = column
        .checked_add(*column_count)
        .ok_or_else(sort_too_large)?;
    let cell_count = u64::from(*row_count) * u64::from(*column_count);
    if *row_count == 0
        || *column_count == 0
        || *row_count > MAX_SORT_ROWS
        || *column_count > MAX_SORT_COLUMNS
        || cell_count > MAX_SORT_CELLS
        || row_end > MAX_SPREADSHEET_ROWS
        || column_end > MAX_SPREADSHEET_COLUMNS
    {
        return Err(sort_too_large());
    }
    if source_rows.len() != *row_count as usize || expected_cells.len() != cell_count as usize {
        return Err(invalid_spreadsheet_mutation(
            "Spreadsheet sort source rows and observed cells must cover the rectangle.",
        ));
    }
    let mut seen = vec![false; *row_count as usize];
    for source in source_rows {
        if *source >= *row_count || seen[*source as usize] {
            return Err(invalid_spreadsheet_mutation(
                "Spreadsheet sort source rows must be a permutation of the rectangle.",
            ));
        }
        seen[*source as usize] = true;
    }
    Ok(SortFrame {
        sheet_id,
        row: *row,
        column: *column,
        row_count: *row_count,
        column_count: *column_count,
        source_rows,
        expected_cells,
    })
}

fn observed_sources(
    state: &SpreadsheetSheetState,
    frame: &SortFrame<'_>,
) -> UseResult<HashMap<(u32, u32), JsonValue>> {
    let mut sources = HashMap::new();
    for source_offset in 0..frame.row_count {
        for column_offset in 0..frame.column_count {
            let coordinate = (frame.row + source_offset, frame.column + column_offset);
            let expected = &frame.expected_cells
                [(source_offset * frame.column_count + column_offset) as usize];
            let live = state.cells.get(&coordinate);
            let matches = match (expected, live) {
                (None, None) => true,
                (Some(expected), Some(live)) => json_equal(expected, live),
                _ => false,
            };
            if !matches {
                return Err(spreadsheet_match_conflict(
                    "Spreadsheet sort observed a cell that has changed.",
                ));
            }
            if let Some(live) = live {
                sources.insert(coordinate, live.clone());
            }
        }
    }
    Ok(sources)
}

fn destination_writes(
    state: &SpreadsheetSheetState,
    frame: &SortFrame<'_>,
    sources: &HashMap<(u32, u32), JsonValue>,
) -> UseResult<Vec<SpreadsheetCellWrite>> {
    let mut changes = Vec::new();
    for (destination_offset, source_offset) in frame.source_rows.iter().copied().enumerate() {
        for column_offset in 0..frame.column_count {
            let column = frame.column + column_offset;
            let source = (frame.row + source_offset, column);
            let destination_row = frame.row + destination_offset as u32;
            let current = state.cells.get(&(destination_row, column));
            let next = sources.get(&source);
            if let Some(change) = cell_write(destination_row, column, current, next)? {
                changes.push(change);
            }
        }
    }
    Ok(changes)
}

fn cell_write(
    row: u32,
    column: u32,
    current: Option<&JsonValue>,
    next: Option<&JsonValue>,
) -> UseResult<Option<SpreadsheetCellWrite>> {
    if match (current, next) {
        (None, None) => true,
        (Some(current), Some(next)) => json_equal(current, next),
        _ => false,
    } {
        return Ok(None);
    }
    let empty = json!({});
    let patches = cell_field_patches(
        current.unwrap_or(&empty),
        current.unwrap_or(&empty),
        next.unwrap_or(&empty),
    )?;
    let presence = match (current.is_some(), next.is_some()) {
        (false, true) => SpreadsheetCellPresenceChange::Insert,
        (true, false) => SpreadsheetCellPresenceChange::Remove,
        _ => SpreadsheetCellPresenceChange::Keep,
    };
    if patches.is_empty() && matches!(presence, SpreadsheetCellPresenceChange::Keep) {
        return Ok(None);
    }
    Ok(Some(SpreadsheetCellWrite {
        row,
        column,
        patches,
        presence,
    }))
}

fn text_moves(
    doc: &yrs::Doc,
    state: &SpreadsheetSheetState,
    frame: &SortFrame<'_>,
) -> UseResult<Vec<SpreadsheetTextMove>> {
    let mut destination_of = vec![0_u32; frame.row_count as usize];
    for (destination, source) in frame.source_rows.iter().copied().enumerate() {
        destination_of[source as usize] = destination as u32;
    }
    let transaction = doc.transact();
    let Some(texts) = optional_map(&state.record, &transaction, CELL_TEXT_KEY)? else {
        return Ok(Vec::new());
    };
    let mut moves = Vec::new();
    for source_offset in 0..frame.row_count {
        let destination_offset = destination_of[source_offset as usize];
        if source_offset == destination_offset {
            continue;
        }
        for column_offset in 0..frame.column_count {
            let column = frame.column + column_offset;
            let from = encode_coordinate(frame.row + source_offset, column);
            match texts.get(&transaction, from.as_str()) {
                Some(Out::YText(text)) => moves.push(SpreadsheetTextMove {
                    texts: texts.clone(),
                    from,
                    to: Some(encode_coordinate(frame.row + destination_offset, column)),
                    content: text.get_string(&transaction),
                }),
                None => {}
                Some(_) => {
                    return Err(invalid_shared_spreadsheet(
                        "The spreadsheet cell text is not a collaborative text node.",
                    ));
                }
            }
        }
    }
    Ok(moves)
}

fn planned_row_lengths(
    state: &SpreadsheetSheetState,
    frame: &SortFrame<'_>,
    sources: &HashMap<(u32, u32), JsonValue>,
) -> UseResult<Option<Vec<u32>>> {
    if state.mode != Some(SpreadsheetCellMode::Data) {
        return Ok(None);
    }
    let mut occupied = state.cells.keys().copied().collect::<HashSet<_>>();
    let mut incoming = Vec::new();
    for (destination_offset, source_offset) in frame.source_rows.iter().copied().enumerate() {
        for column_offset in 0..frame.column_count {
            let column = frame.column + column_offset;
            let source = (frame.row + source_offset, column);
            let destination = (frame.row + destination_offset as u32, column);
            incoming.push((destination, sources.contains_key(&source)));
        }
    }
    for source_offset in 0..frame.row_count {
        for column_offset in 0..frame.column_count {
            occupied.remove(&(frame.row + source_offset, frame.column + column_offset));
        }
    }
    for (destination, present) in incoming {
        if present {
            occupied.insert(destination);
        }
    }
    let mut lengths = state.row_lengths.clone();
    let end = (frame.row + frame.row_count) as usize;
    if lengths.len() < end {
        lengths.resize(end, 0);
    }
    for offset in 0..frame.row_count {
        let sheet_row = frame.row + offset;
        let width = occupied
            .iter()
            .filter(|(row, _)| *row == sheet_row)
            .map(|(_, column)| column + 1)
            .max()
            .unwrap_or(0);
        lengths[sheet_row as usize] = width;
    }
    let materialized = lengths.iter().map(|value| u64::from(*value)).sum::<u64>();
    if materialized > MAX_SPREADSHEET_DENSE_CELLS {
        return Err(sort_too_large());
    }
    Ok((lengths != state.row_lengths).then_some(lengths))
}

fn refuse_structural_overlap(
    doc: &yrs::Doc,
    state: &SpreadsheetSheetState,
    frame: &SortFrame<'_>,
) -> UseResult<()> {
    let transaction = doc.transact();
    if optional_map(&state.record, &transaction, PIVOT_TABLES_KEY)?
        .is_some_and(|pivots| pivots.len(&transaction) > 0)
    {
        return Err(invalid_spreadsheet_mutation(
            "Spreadsheet sort cannot prove a pivot table is outside the rectangle.",
        ));
    }
    if let Some(config) = optional_map(&state.record, &transaction, CONFIG_KEY)? {
        refuse_overlapping_merges(&config, &transaction, frame)?;
    }
    if let Some(tables) = optional_map(&state.record, &transaction, TABLES_KEY)? {
        refuse_overlapping_tables(&tables, &transaction, frame)?;
    }
    Ok(())
}

fn refuse_overlapping_merges<T: ReadTxn>(
    config: &MapRef,
    transaction: &T,
    frame: &SortFrame<'_>,
) -> UseResult<()> {
    let mut groups: HashMap<String, MergeSpan> = HashMap::new();
    for (key, value) in config.iter(transaction) {
        let (kind, path) = decode_flat_json_key(key).map_err(|_| {
            invalid_shared_spreadsheet("A shared Spreadsheet merge field identity is invalid.")
        })?;
        if path.first().map(String::as_str) != Some("merge") || path.len() < 3 {
            continue;
        }
        if path.len() != 3 || kind != FlatJsonEntryKind::Value {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet merge address is not a flat record.",
            ));
        }
        let Out::Any(value) = value else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet merge field is not a JSON value.",
            ));
        };
        let group = groups.entry(path[1].clone()).or_default();
        let number = any_u32(&value)?;
        match path[2].as_str() {
            "r" => group.row = Some(number),
            "c" => group.column = Some(number),
            "rs" => group.row_span = Some(number),
            "cs" => group.column_span = Some(number),
            _ => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet merge contains an unknown field.",
                ));
            }
        }
    }
    for group in groups.values() {
        let (row, column, row_span, column_span) = match (
            group.row,
            group.column,
            group.row_span.filter(|span| *span > 0),
            group.column_span.filter(|span| *span > 0),
        ) {
            (Some(row), Some(column), Some(row_span), Some(column_span)) => {
                (row, column, row_span, column_span)
            }
            _ => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet merge is missing its span.",
                ));
            }
        };
        let row_end = row.checked_add(row_span).ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet merge span overflows.")
        })?;
        let column_end = column.checked_add(column_span).ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet merge span overflows.")
        })?;
        if ranges_overlap(frame.row, frame.row + frame.row_count, row, row_end)
            && ranges_overlap(
                frame.column,
                frame.column + frame.column_count,
                column,
                column_end,
            )
        {
            return Err(invalid_spreadsheet_mutation(
                "A merged range intersects the spreadsheet sort.",
            ));
        }
    }
    Ok(())
}

fn refuse_overlapping_tables<T: ReadTxn>(
    tables: &MapRef,
    transaction: &T,
    frame: &SortFrame<'_>,
) -> UseResult<()> {
    for (_id, value) in tables.iter(transaction) {
        let Out::YMap(table) = value else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet table is not a typed map.",
            ));
        };
        let mut row_range = None;
        let mut column_range = None;
        for (key, value) in table.iter(transaction) {
            let (kind, path) = decode_flat_json_key(key).map_err(|_| {
                invalid_shared_spreadsheet("A shared Spreadsheet table field identity is invalid.")
            })?;
            if kind != FlatJsonEntryKind::Value || path.len() != 2 || path[0] != "range" {
                continue;
            }
            let Out::Any(value) = value else {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet table range is not a JSON value.",
                ));
            };
            let pair = inclusive_pair(&value)?;
            match path[1].as_str() {
                "row" => row_range = Some(pair),
                "column" => column_range = Some(pair),
                _ => {}
            }
        }
        let Some((row_start, row_end)) = row_range else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet table is missing its row range.",
            ));
        };
        let Some((column_start, column_end)) = column_range else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet table is missing its column range.",
            ));
        };
        if ranges_overlap(
            frame.row,
            frame.row + frame.row_count,
            row_start,
            row_end + 1,
        ) && ranges_overlap(
            frame.column,
            frame.column + frame.column_count,
            column_start,
            column_end + 1,
        ) {
            return Err(invalid_spreadsheet_mutation(
                "A table intersects the spreadsheet sort.",
            ));
        }
    }
    Ok(())
}

fn inclusive_pair(value: &Any) -> UseResult<(u32, u32)> {
    let json = any_to_json(value.clone(), 0)?;
    let Some(values) = json.as_array() else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is not a pair.",
        ));
    };
    if values.len() != 2 {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is not a pair.",
        ));
    }
    let start = json_u32(&values[0])?;
    let end = json_u32(&values[1])?;
    if start > end {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is reversed.",
        ));
    }
    Ok((start, end))
}

fn ranges_overlap(left_start: u32, left_end: u32, right_start: u32, right_end: u32) -> bool {
    left_start < right_end && right_start < left_end
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
            "Shared Spreadsheet field '{key}' is not a typed map."
        ))),
    }
}

fn any_u32(value: &Any) -> UseResult<u32> {
    let Any::Number(number) = value else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    };
    json_u32(&JsonValue::from(*number))
}

fn json_u32(value: &JsonValue) -> UseResult<u32> {
    let Some(number) = value.as_f64() else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    };
    if !number.is_finite() || number.fract() != 0.0 || !(0.0..=u32::MAX as f64).contains(&number) {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet address is not an integer.",
        ));
    }
    Ok(number as u32)
}

fn sort_too_large() -> a3s_use_core::UseError {
    invalid_spreadsheet_mutation("Spreadsheet sort exceeds the supported rectangle.")
}

#[derive(Default)]
struct MergeSpan {
    row: Option<u32>,
    column: Option<u32>,
    row_span: Option<u32>,
    column_span: Option<u32>,
}
