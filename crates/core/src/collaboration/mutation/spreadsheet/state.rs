use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Any, Array, ArrayRef, Map, MapRef, Out, Transact};

use super::json::{
    any_to_json, decode_flat_json_key, reconstruct_cell, validate_shared_cell_json,
    DecodedFlatJsonEntry, FlatJsonEntryKind, FlatJsonPatch,
};
use super::{
    invalid_shared_spreadsheet, spreadsheet_match_conflict, NativeOfficeCollaborationManifest,
    MAX_SPREADSHEET_COLUMNS, MAX_SPREADSHEET_DENSE_CELLS, MAX_SPREADSHEET_POPULATED_CELLS,
    MAX_SPREADSHEET_ROWS,
};

const MAX_SPREADSHEET_CELL_FIELDS: u32 = 16_000_000;
const SHEETS_ROOT: &str = "spreadsheet.sheets";
const CELLS_KEY: &str = "cells";
const CELL_PRESENCE_KEY: &str = "cellPresence";
const CELL_MODE_KEY: &str = "cellMode";
const DATA_ROW_LENGTHS_KEY: &str = "dataRowLengths";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SpreadsheetCellMode {
    Data,
    CellData,
}

pub(super) struct SpreadsheetSheetState {
    pub(super) cells: HashMap<(u32, u32), JsonValue>,
    pub(super) fields: MapRef,
    pub(super) mode: Option<SpreadsheetCellMode>,
    pub(super) presence: MapRef,
    pub(super) record: MapRef,
    pub(super) row_lengths: Vec<u32>,
    pub(super) row_lengths_ref: ArrayRef,
}

pub(super) fn read_sheet_state(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
) -> UseResult<SpreadsheetSheetState> {
    let sheets = doc.get_or_insert_map(format!("{}.{}", manifest.namespace, SHEETS_ROOT));
    let transaction = doc.transact();
    let record = match sheets.get(&transaction, sheet_id) {
        Some(Out::YMap(record)) => record,
        Some(_) => {
            return Err(invalid_shared_spreadsheet(format!(
                "Shared Spreadsheet sheet '{sheet_id}' is not a typed map."
            )))
        }
        None => {
            return Err(spreadsheet_match_conflict(format!(
                "Spreadsheet sheet ID '{sheet_id}' does not exist."
            )))
        }
    };
    match record.get(&transaction, "id") {
        Some(Out::Any(Any::String(value))) if value.as_ref() == sheet_id => {}
        _ => {
            return Err(invalid_shared_spreadsheet(format!(
                "Shared Spreadsheet sheet '{sheet_id}' has an invalid identity field."
            )))
        }
    }
    let fields = required_map(&record, &transaction, CELLS_KEY, "cell field map")?;
    let presence = required_map(
        &record,
        &transaction,
        CELL_PRESENCE_KEY,
        "cell presence map",
    )?;
    let row_lengths_ref = required_array(
        &record,
        &transaction,
        DATA_ROW_LENGTHS_KEY,
        "data row-length array",
    )?;

    if presence.len(&transaction) > MAX_SPREADSHEET_POPULATED_CELLS {
        return Err(invalid_shared_spreadsheet(format!(
            "The shared Spreadsheet contains more than {MAX_SPREADSHEET_POPULATED_CELLS} populated cells in one sheet."
        )));
    }
    if fields.len(&transaction) > MAX_SPREADSHEET_CELL_FIELDS {
        return Err(invalid_shared_spreadsheet(
            "The shared Spreadsheet contains too many cell fields in one sheet.",
        ));
    }

    let mut coordinates = HashSet::with_capacity(presence.len(&transaction) as usize);
    for (encoded, value) in presence.iter(&transaction) {
        if !matches!(value, Out::Any(Any::Bool(true))) {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet cell presence marker is not true.",
            ));
        }
        coordinates.insert(decode_coordinate(encoded)?);
    }

    let mut grouped = HashMap::<(u32, u32), Vec<DecodedFlatJsonEntry>>::new();
    for (encoded, value) in fields.iter(&transaction) {
        let (row, column, flat_key) = decode_cell_field_key(encoded)?;
        if !coordinates.contains(&(row, column)) {
            return Err(invalid_shared_spreadsheet(format!(
                "The shared Spreadsheet contains an orphan cell field at '{row}:{column}'."
            )));
        }
        let (kind, path) = decode_flat_json_key(&flat_key)?;
        let value = match (kind, value) {
            (FlatJsonEntryKind::Object, Out::Any(Any::Bool(true))) => JsonValue::Bool(true),
            (FlatJsonEntryKind::Object, _) => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet cell contains an invalid object marker.",
                ))
            }
            (FlatJsonEntryKind::Value, Out::Any(value)) => any_to_json(value, 0)?,
            (FlatJsonEntryKind::Value, _) => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet cell contains a shared-type field.",
                ))
            }
        };
        grouped
            .entry((row, column))
            .or_default()
            .push(DecodedFlatJsonEntry { kind, path, value });
    }

    let mut cells = coordinates
        .iter()
        .copied()
        .map(|coordinate| (coordinate, JsonValue::Object(Default::default())))
        .collect::<HashMap<_, _>>();
    for (coordinate, entries) in grouped {
        let cell = reconstruct_cell(entries)?;
        validate_shared_cell_json(&cell)?;
        cells.insert(coordinate, cell);
    }

    let mode = match record.get(&transaction, CELL_MODE_KEY) {
        None => None,
        Some(Out::Any(Any::String(value))) if value.as_ref() == "data" => {
            Some(SpreadsheetCellMode::Data)
        }
        Some(Out::Any(Any::String(value))) if value.as_ref() == "celldata" => {
            Some(SpreadsheetCellMode::CellData)
        }
        _ => {
            return Err(invalid_shared_spreadsheet(
                "The shared Spreadsheet cell projection mode is invalid.",
            ))
        }
    };
    let row_lengths = read_row_lengths(&row_lengths_ref, &transaction)?;
    match mode {
        None if !coordinates.is_empty() || !row_lengths.is_empty() => {
            return Err(invalid_shared_spreadsheet(
                "The shared Spreadsheet has cells without a projection mode.",
            ))
        }
        Some(SpreadsheetCellMode::CellData) if !row_lengths.is_empty() => {
            return Err(invalid_shared_spreadsheet(
                "The shared Spreadsheet sparse projection contains dense row lengths.",
            ))
        }
        Some(SpreadsheetCellMode::Data) => {
            for (row, column) in &coordinates {
                if (*row as usize) >= row_lengths.len() || *column >= row_lengths[*row as usize] {
                    return Err(invalid_shared_spreadsheet(format!(
                        "Shared Spreadsheet cell '{row}:{column}' is outside its dense matrix."
                    )));
                }
            }
        }
        _ => {}
    }
    drop(transaction);

    Ok(SpreadsheetSheetState {
        cells,
        fields,
        mode,
        presence,
        record,
        row_lengths,
        row_lengths_ref,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SpreadsheetCellPresenceChange {
    Keep,
    Insert,
    Remove,
}

pub(super) struct SpreadsheetCellWrite {
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) patches: Vec<FlatJsonPatch>,
    pub(super) presence: SpreadsheetCellPresenceChange,
}

enum EncodedFieldPatch {
    Remove(String),
    Set(String, Any),
}

pub(super) fn write_cell_changes(
    doc: &yrs::Doc,
    state: &SpreadsheetSheetState,
    changes: Vec<SpreadsheetCellWrite>,
    initialize_sparse_mode: bool,
    next_row_lengths: Option<Vec<u32>>,
) -> UseResult<()> {
    let changes_row_lengths = next_row_lengths
        .as_ref()
        .is_some_and(|lengths| lengths != &state.row_lengths);
    let changes_content = changes.iter().any(|change| {
        !change.patches.is_empty() || change.presence != SpreadsheetCellPresenceChange::Keep
    });
    if !changes_content && !initialize_sparse_mode && !changes_row_lengths {
        return Ok(());
    }

    let encoded = changes
        .into_iter()
        .map(|change| {
            let patches = change
                .patches
                .into_iter()
                .map(|patch| match patch {
                    FlatJsonPatch::Remove(flat_key) => {
                        encode_cell_field_key(change.row, change.column, &flat_key)
                            .map(EncodedFieldPatch::Remove)
                    }
                    FlatJsonPatch::Set(flat_key, value) => {
                        encode_cell_field_key(change.row, change.column, &flat_key)
                            .map(|key| EncodedFieldPatch::Set(key, value))
                    }
                })
                .collect::<UseResult<Vec<_>>>()?;
            Ok((
                patches,
                encode_coordinate(change.row, change.column),
                change.presence,
            ))
        })
        .collect::<UseResult<Vec<_>>>()?;

    let mut transaction = doc.transact_mut();
    for (patches, coordinate, presence) in encoded {
        for patch in patches {
            match patch {
                EncodedFieldPatch::Remove(key) => {
                    state.fields.remove(&mut transaction, key.as_str());
                }
                EncodedFieldPatch::Set(key, value) => {
                    state.fields.insert(&mut transaction, key, value);
                }
            }
        }
        match presence {
            SpreadsheetCellPresenceChange::Keep => {}
            SpreadsheetCellPresenceChange::Insert => {
                state.presence.insert(&mut transaction, coordinate, true);
            }
            SpreadsheetCellPresenceChange::Remove => {
                state.presence.remove(&mut transaction, coordinate.as_str());
            }
        }
    }
    if initialize_sparse_mode {
        state
            .record
            .insert(&mut transaction, CELL_MODE_KEY, "celldata");
    }
    if let Some(lengths) = next_row_lengths.filter(|lengths| lengths != &state.row_lengths) {
        let current_len = state.row_lengths_ref.len(&transaction);
        if current_len > 0 {
            state
                .row_lengths_ref
                .remove_range(&mut transaction, 0, current_len);
        }
        state
            .row_lengths_ref
            .insert_range(&mut transaction, 0, lengths.into_iter().map(f64::from));
    }
    Ok(())
}

pub(super) fn extended_dense_row_lengths(
    current: &[u32],
    coordinates: impl IntoIterator<Item = (u32, u32)>,
) -> UseResult<Vec<u32>> {
    let mut result = current.to_vec();
    for (row, column) in coordinates {
        result.resize(result.len().max(row as usize + 1), 0);
        result[row as usize] = result[row as usize].max(column + 1);
    }
    let materialized = result.iter().map(|value| u64::from(*value)).sum::<u64>();
    if materialized > MAX_SPREADSHEET_DENSE_CELLS {
        return Err(super::invalid_spreadsheet_mutation(format!(
            "A dense Spreadsheet cell mutation would materialize {materialized} cells; the limit is {MAX_SPREADSHEET_DENSE_CELLS}."
        )));
    }
    Ok(result)
}

fn required_map<T: yrs::ReadTxn>(
    record: &MapRef,
    transaction: &T,
    key: &str,
    label: &str,
) -> UseResult<MapRef> {
    match record.get(transaction, key) {
        Some(Out::YMap(value)) => Ok(value),
        _ => Err(invalid_shared_spreadsheet(format!(
            "The shared Spreadsheet {label} is missing or invalid."
        ))),
    }
}

fn required_array<T: yrs::ReadTxn>(
    record: &MapRef,
    transaction: &T,
    key: &str,
    label: &str,
) -> UseResult<ArrayRef> {
    match record.get(transaction, key) {
        Some(Out::YArray(value)) => Ok(value),
        _ => Err(invalid_shared_spreadsheet(format!(
            "The shared Spreadsheet {label} is missing or invalid."
        ))),
    }
}

fn read_row_lengths<T: yrs::ReadTxn>(values: &ArrayRef, transaction: &T) -> UseResult<Vec<u32>> {
    if values.len(transaction) > MAX_SPREADSHEET_ROWS {
        return Err(invalid_shared_spreadsheet(
            "The shared Spreadsheet dense projection exceeds the row limit.",
        ));
    }
    let mut materialized = 0_u64;
    let mut result = Vec::with_capacity(values.len(transaction) as usize);
    for index in 0..values.len(transaction) {
        let value = match values.get(transaction, index) {
            Some(Out::Any(Any::Number(value)))
                if value.is_finite()
                    && value.fract() == 0.0
                    && value >= 0.0
                    && value <= f64::from(MAX_SPREADSHEET_COLUMNS) =>
            {
                value as u32
            }
            _ => {
                return Err(invalid_shared_spreadsheet(format!(
                    "The shared Spreadsheet data row length at row {index} is invalid."
                )))
            }
        };
        materialized += u64::from(value);
        if materialized > MAX_SPREADSHEET_DENSE_CELLS {
            return Err(invalid_shared_spreadsheet(
                "The shared Spreadsheet dense projection exceeds the materialized-cell limit.",
            ));
        }
        result.push(value);
    }
    Ok(result)
}

fn encode_coordinate(row: u32, column: u32) -> String {
    format!("{row}:{column}")
}

fn decode_coordinate(encoded: &str) -> UseResult<(u32, u32)> {
    let Some((row, column)) = encoded.split_once(':') else {
        return Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet cell coordinate '{encoded}' is invalid."
        )));
    };
    let row = row.parse::<u32>().ok();
    let column = column.parse::<u32>().ok();
    match (row, column) {
        (Some(row), Some(column))
            if row < MAX_SPREADSHEET_ROWS
                && column < MAX_SPREADSHEET_COLUMNS
                && encode_coordinate(row, column) == encoded =>
        {
            Ok((row, column))
        }
        _ => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet cell coordinate '{encoded}' is invalid."
        ))),
    }
}

fn encode_cell_field_key(row: u32, column: u32, flat_key: &str) -> UseResult<String> {
    serde_json::to_string(&(row, column, flat_key)).map_err(|error| {
        super::invalid_spreadsheet_mutation(format!(
            "Failed to encode a Spreadsheet cell field identity: {error}"
        ))
    })
}

fn decode_cell_field_key(encoded: &str) -> UseResult<(u32, u32, String)> {
    let (row, column, flat_key) =
        serde_json::from_str::<(u32, u32, String)>(encoded).map_err(|_| {
            invalid_shared_spreadsheet(format!(
                "Shared Spreadsheet cell field identity '{encoded}' is invalid."
            ))
        })?;
    if row >= MAX_SPREADSHEET_ROWS
        || column >= MAX_SPREADSHEET_COLUMNS
        || flat_key.is_empty()
        || encode_cell_field_key(row, column, &flat_key)? != encoded
    {
        return Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet cell field identity '{encoded}' is invalid."
        )));
    }
    Ok((row, column, flat_key))
}
