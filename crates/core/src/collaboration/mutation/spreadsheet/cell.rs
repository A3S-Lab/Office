use std::collections::HashSet;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;

use super::json::{
    assert_compatible_cell, cell_field_patches, flattened_cell, json_equal, merged_compatible_cell,
    validate_cell_json, FlatJsonPatch,
};
use super::state::{
    extended_dense_row_lengths, read_sheet_state, write_cell_changes, SpreadsheetCellMode,
    SpreadsheetCellPresenceChange, SpreadsheetCellWrite,
};
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, spreadsheet_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationSpreadsheetCellChange, MAX_SPREADSHEET_COLUMNS,
    MAX_SPREADSHEET_POPULATED_CELLS, MAX_SPREADSHEET_ROWS,
};

const MAX_SPREADSHEET_SHEET_ID_UTF16: usize = 256;
const MAX_SPREADSHEET_BATCH_CELL_CHANGES: usize = 4_096;

pub(super) fn validate_cell_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id,
            row,
            column,
            expected_cell,
            next_cell,
        } => {
            validate_sheet_id(sheet_id)?;
            validate_coordinate(*row, *column)?;
            if let Some(expected_cell) = expected_cell {
                validate_cell_json(expected_cell, "expected cell")?;
            }
            validate_cell_json(next_cell, "next cell")
        }
        NativeOfficeCollaborationMutation::SpreadsheetDeleteCell {
            sheet_id,
            row,
            column,
            expected_cell,
        } => {
            validate_sheet_id(sheet_id)?;
            validate_coordinate(*row, *column)?;
            validate_cell_json(expected_cell, "expected cell")
        }
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells { sheet_id, changes } => {
            validate_sheet_id(sheet_id)?;
            validate_batch_changes(changes)
        }
        _ => Err(invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet cell mutation.",
        )),
    }
}

pub(super) fn apply_cell_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id,
            row,
            column,
            expected_cell,
            next_cell,
        } => apply_cell_changes(
            doc,
            manifest,
            sheet_id,
            std::slice::from_ref(&NativeOfficeCollaborationSpreadsheetCellChange {
                row: *row,
                column: *column,
                expected_cell: expected_cell.clone(),
                next_cell: Some(next_cell.clone()),
            }),
        ),
        NativeOfficeCollaborationMutation::SpreadsheetDeleteCell {
            sheet_id,
            row,
            column,
            expected_cell,
        } => apply_cell_changes(
            doc,
            manifest,
            sheet_id,
            std::slice::from_ref(&NativeOfficeCollaborationSpreadsheetCellChange {
                row: *row,
                column: *column,
                expected_cell: Some(expected_cell.clone()),
                next_cell: None,
            }),
        ),
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells { sheet_id, changes } => {
            apply_cell_changes(doc, manifest, sheet_id, changes)
        }
        _ => Err(invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet cell mutation.",
        )),
    }
}

fn validate_batch_changes(
    changes: &[NativeOfficeCollaborationSpreadsheetCellChange],
) -> UseResult<()> {
    if changes.is_empty() || changes.len() > MAX_SPREADSHEET_BATCH_CELL_CHANGES {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet cell batch requires 1 to {MAX_SPREADSHEET_BATCH_CELL_CHANGES} distinct coordinates."
        ))
        .with_detail("changes", changes.len() as u64)
        .with_detail("maximum", MAX_SPREADSHEET_BATCH_CELL_CHANGES as u64));
    }
    let mut coordinates = HashSet::with_capacity(changes.len());
    for change in changes {
        validate_coordinate(change.row, change.column)?;
        if !coordinates.insert((change.row, change.column)) {
            return Err(invalid_spreadsheet_mutation(format!(
                "Spreadsheet cell '{}:{}' appears more than once in one batch.",
                change.row, change.column
            ))
            .with_detail("row", u64::from(change.row))
            .with_detail("column", u64::from(change.column)));
        }
        if let Some(expected_cell) = &change.expected_cell {
            validate_cell_json(expected_cell, "expected cell")?;
        }
        match &change.next_cell {
            Some(next_cell) => validate_cell_json(next_cell, "next cell")?,
            None if change.expected_cell.is_none() => {
                return Err(invalid_spreadsheet_mutation(format!(
                    "Deleting Spreadsheet cell '{}:{}' requires an expected cell value.",
                    change.row, change.column
                )))
            }
            None => {}
        }
    }
    Ok(())
}

fn apply_cell_changes(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    changes: &[NativeOfficeCollaborationSpreadsheetCellChange],
) -> UseResult<()> {
    let state = read_sheet_state(doc, manifest, sheet_id)?;
    let empty = JsonValue::Object(Default::default());
    let mut writes = Vec::with_capacity(changes.len());
    let mut expected_after = Vec::with_capacity(changes.len());
    let mut dense_coordinates = Vec::with_capacity(changes.len());
    let mut populated_cells = state.cells.len();
    let mut has_set = false;

    for change in changes {
        let coordinate = (change.row, change.column);
        let current = state.cells.get(&coordinate);
        let label = format!(
            "Spreadsheet cell '{}:{}' in sheet '{sheet_id}'",
            change.row, change.column
        );
        match &change.next_cell {
            Some(next_cell) => {
                match (&change.expected_cell, current) {
                    (None, Some(_)) => {
                        return Err(spreadsheet_match_conflict(format!(
                            "{label} is no longer blank."
                        )))
                    }
                    (Some(_), None) => {
                        return Err(spreadsheet_match_conflict(format!(
                            "{label} no longer exists."
                        )))
                    }
                    _ => {}
                }
                let expected = change.expected_cell.as_ref().unwrap_or(&empty);
                let shared = current.unwrap_or(&empty);
                assert_compatible_cell(expected, next_cell, shared, &label)?;
                let patches = cell_field_patches(expected, shared, next_cell)?;
                let presence = if current.is_some() {
                    SpreadsheetCellPresenceChange::Keep
                } else {
                    populated_cells += 1;
                    SpreadsheetCellPresenceChange::Insert
                };
                writes.push(SpreadsheetCellWrite {
                    row: change.row,
                    column: change.column,
                    patches,
                    presence,
                });
                expected_after.push((
                    coordinate,
                    Some(merged_compatible_cell(expected, next_cell, shared)),
                ));
                dense_coordinates.push(coordinate);
                has_set = true;
            }
            None => {
                let expected = change.expected_cell.as_ref().ok_or_else(|| {
                    invalid_spreadsheet_mutation(format!(
                        "Deleting {label} requires an expected cell value."
                    ))
                })?;
                let Some(current) = current else {
                    expected_after.push((coordinate, None));
                    continue;
                };
                if !json_equal(current, expected) {
                    return Err(spreadsheet_match_conflict(format!(
                        "{label} changed before it could be deleted."
                    )));
                }
                let patches = flattened_cell(current)?
                    .into_keys()
                    .map(FlatJsonPatch::Remove)
                    .collect();
                writes.push(SpreadsheetCellWrite {
                    row: change.row,
                    column: change.column,
                    patches,
                    presence: SpreadsheetCellPresenceChange::Remove,
                });
                populated_cells -= 1;
                expected_after.push((coordinate, None));
            }
        }
    }
    if populated_cells > MAX_SPREADSHEET_POPULATED_CELLS as usize {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet cell batch would exceed the {MAX_SPREADSHEET_POPULATED_CELLS} populated-cell limit."
        ))
        .with_detail("populatedCells", populated_cells as u64)
        .with_detail(
            "maximum",
            u64::from(MAX_SPREADSHEET_POPULATED_CELLS),
        ));
    }
    let next_row_lengths = match state.mode {
        Some(SpreadsheetCellMode::Data) => Some(extended_dense_row_lengths(
            &state.row_lengths,
            dense_coordinates,
        )?),
        _ => None,
    };
    write_cell_changes(
        doc,
        &state,
        writes,
        state.mode.is_none() && has_set,
        next_row_lengths,
    )?;

    let verified = read_sheet_state(doc, manifest, sheet_id)?;
    for (coordinate, expected) in expected_after {
        let current = verified.cells.get(&coordinate);
        let valid = match (expected.as_ref(), current) {
            (Some(expected), Some(current)) => json_equal(expected, current),
            (None, None) => true,
            _ => false,
        };
        if !valid {
            return Err(invalid_shared_spreadsheet(format!(
                "Spreadsheet cell '{}:{}' did not match the atomic batch result.",
                coordinate.0, coordinate.1
            )));
        }
    }
    Ok(())
}

fn validate_sheet_id(value: &str) -> UseResult<()> {
    let length = value.encode_utf16().count();
    let trimmed = value
        .chars()
        .next()
        .is_some_and(is_ecmascript_trim_character)
        || value
            .chars()
            .next_back()
            .is_some_and(is_ecmascript_trim_character);
    if (1..=MAX_SPREADSHEET_SHEET_ID_UTF16).contains(&length) && !trimmed {
        return Ok(());
    }
    Err(invalid_spreadsheet_mutation(
        "A Spreadsheet sheet ID must contain 1 to 256 UTF-16 code units without leading or trailing whitespace.",
    )
    .with_detail("sheetId", value.to_owned()))
}

fn validate_coordinate(row: u32, column: u32) -> UseResult<()> {
    if row < MAX_SPREADSHEET_ROWS && column < MAX_SPREADSHEET_COLUMNS {
        return Ok(());
    }
    Err(super::super::super::collaboration_error(
        "office.collaboration.mutation_range_invalid",
        "The Spreadsheet cell coordinate is outside the supported worksheet range.",
    )
    .with_detail("row", u64::from(row))
    .with_detail("column", u64::from(column))
    .with_detail("maxRows", u64::from(MAX_SPREADSHEET_ROWS))
    .with_detail("maxColumns", u64::from(MAX_SPREADSHEET_COLUMNS)))
}

fn is_ecmascript_trim_character(character: char) -> bool {
    matches!(
        character,
        '\u{0009}'
            | '\u{000A}'
            | '\u{000B}'
            | '\u{000C}'
            | '\u{000D}'
            | '\u{0020}'
            | '\u{00A0}'
            | '\u{1680}'
            | '\u{2000}'
            ..='\u{200A}'
                | '\u{2028}'
                | '\u{2029}'
                | '\u{202F}'
                | '\u{205F}'
                | '\u{3000}'
                | '\u{FEFF}'
    )
}
