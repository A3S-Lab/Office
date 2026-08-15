use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;

use super::json::{assert_compatible_cell, cell_field_patches, json_equal, validate_cell_json};
use super::state::{
    extended_dense_row_lengths, read_cell_state, write_delete_cell, write_set_cell,
    SpreadsheetCellMode,
};
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, spreadsheet_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation, MAX_SPREADSHEET_COLUMNS,
    MAX_SPREADSHEET_ROWS,
};

const MAX_SPREADSHEET_SHEET_ID_UTF16: usize = 256;

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
        } => set_cell(
            doc,
            manifest,
            sheet_id,
            *row,
            *column,
            expected_cell.as_ref(),
            next_cell,
        ),
        NativeOfficeCollaborationMutation::SpreadsheetDeleteCell {
            sheet_id,
            row,
            column,
            expected_cell,
        } => delete_cell(doc, manifest, sheet_id, *row, *column, expected_cell),
        _ => Err(invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet cell mutation.",
        )),
    }
}

#[allow(clippy::too_many_arguments)]
fn set_cell(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    row: u32,
    column: u32,
    expected_cell: Option<&JsonValue>,
    next_cell: &JsonValue,
) -> UseResult<()> {
    let state = read_cell_state(doc, manifest, sheet_id, row, column)?;
    match (expected_cell, state.cell.as_ref()) {
        (None, Some(_)) => {
            return Err(spreadsheet_match_conflict(format!(
                "Spreadsheet cell '{row}:{column}' in sheet '{sheet_id}' is no longer blank."
            )));
        }
        (Some(_), None) => {
            return Err(spreadsheet_match_conflict(format!(
                "Spreadsheet cell '{row}:{column}' in sheet '{sheet_id}' no longer exists."
            )));
        }
        _ => {}
    }
    let empty = JsonValue::Object(Default::default());
    let expected = expected_cell.unwrap_or(&empty);
    let shared = state.cell.as_ref().unwrap_or(&empty);
    let label = format!("Spreadsheet cell '{row}:{column}' in sheet '{sheet_id}'");
    assert_compatible_cell(expected, next_cell, shared, &label)?;
    let patches = cell_field_patches(expected, shared, next_cell)?;
    let next_row_lengths = match state.mode {
        Some(SpreadsheetCellMode::Data) => {
            Some(extended_dense_row_lengths(&state.row_lengths, row, column)?)
        }
        _ => None,
    };
    write_set_cell(doc, &state, row, column, patches, next_row_lengths)?;
    if read_cell_state(doc, manifest, sheet_id, row, column)?
        .cell
        .is_none()
    {
        return Err(invalid_shared_spreadsheet(
            "The Spreadsheet cell mutation did not produce a present cell.",
        ));
    }
    Ok(())
}

fn delete_cell(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    row: u32,
    column: u32,
    expected_cell: &JsonValue,
) -> UseResult<()> {
    let state = read_cell_state(doc, manifest, sheet_id, row, column)?;
    let Some(current) = state.cell.as_ref() else {
        return Ok(());
    };
    if !json_equal(current, expected_cell) {
        return Err(spreadsheet_match_conflict(format!(
            "Spreadsheet cell '{row}:{column}' in sheet '{sheet_id}' changed before it could be deleted."
        )));
    }
    write_delete_cell(doc, &state, row, column, current)?;
    if read_cell_state(doc, manifest, sheet_id, row, column)?
        .cell
        .is_some()
    {
        return Err(invalid_shared_spreadsheet(
            "The Spreadsheet cell deletion left a present cell behind.",
        ));
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
