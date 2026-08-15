use a3s_use_core::UseResult;

use super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
};

mod cell;
mod json;
mod state;

const MAX_SPREADSHEET_ROWS: u32 = 1_048_576;
const MAX_SPREADSHEET_COLUMNS: u32 = 16_384;
const MAX_SPREADSHEET_POPULATED_CELLS: u32 = 1_000_000;
const MAX_SPREADSHEET_DENSE_CELLS: u64 = 1_000_000;

pub(super) fn validate_spreadsheet_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    cell::validate_cell_mutation(mutation)
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
