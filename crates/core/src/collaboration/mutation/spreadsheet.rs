use a3s_use_core::UseResult;

use super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationProjectedContent, NativeOfficeCollaborationSpreadsheetCellAddress,
    NativeOfficeCollaborationSpreadsheetCellChange, NativeOfficeCollaborationSpreadsheetSheet,
};

mod cell;
mod find;
mod json;
mod state;

pub(in crate::collaboration) use find::find_spreadsheet_text;

const MAX_SPREADSHEET_ROWS: u32 = 1_048_576;
const MAX_SPREADSHEET_COLUMNS: u32 = 16_384;
const MAX_SPREADSHEET_POPULATED_CELLS: u32 = 1_000_000;
const MAX_SPREADSHEET_DENSE_CELLS: u64 = 1_000_000;

pub(super) fn validate_spreadsheet_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    cell::validate_cell_mutation(mutation)
}

pub(in crate::collaboration) fn project_spreadsheet_content(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<NativeOfficeCollaborationProjectedContent> {
    let sheet_ids = state::ordered_sheet_ids(doc, manifest)?;
    let mut sheets = Vec::with_capacity(sheet_ids.len());
    for sheet_id in sheet_ids {
        let sheet = state::read_sheet_state(doc, manifest, &sheet_id)?;
        let mut cells = sheet
            .cells
            .iter()
            .map(
                |((row, column), cell)| NativeOfficeCollaborationSpreadsheetCellAddress {
                    row: *row,
                    column: *column,
                    cell: cell.clone(),
                },
            )
            .collect::<Vec<_>>();
        cells.sort_by_key(|cell| (cell.row, cell.column));
        sheets.push(NativeOfficeCollaborationSpreadsheetSheet { sheet_id, cells });
    }
    Ok(NativeOfficeCollaborationProjectedContent::Spreadsheet { sheets })
}

pub(super) fn apply_spreadsheet_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    cell::apply_cell_mutation(doc, manifest, mutation)
}

fn invalid_spreadsheet_mutation(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}

fn spreadsheet_match_conflict(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message)
        .with_suggestion("Read the latest collaborative Spreadsheet cell and retry.")
}

fn invalid_shared_spreadsheet(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.content_invalid", message)
}
