use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationDocumentTextFindResult,
    NativeOfficeCollaborationDocumentTextMatch, NativeOfficeCollaborationManifest,
};
use super::super::utf16_len;
use super::state::{ordered_sheet_ids, read_sheet_state};

const MAX_SPREADSHEET_TEXT_FIND_MATCHES: usize = 4_096;

/// List spreadsheet cells whose display text contains `search`.
///
/// One hit per cell, in sheet-order then row/column. The edit that follows is
/// still `spreadsheet-set-cell` at the returned coordinate; this walk does not
/// invent a text-replace mutation.
pub(in crate::collaboration) fn find_spreadsheet_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    limit: usize,
) -> UseResult<NativeOfficeCollaborationDocumentTextFindResult> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Spreadsheet text find requires a non-empty search string.",
        ));
    }
    if limit == 0 {
        return Err(collaboration_error(
            "office.collaboration.find_limit_invalid",
            "Spreadsheet text find limit must be at least 1.",
        ));
    }

    let mut matches = Vec::new();
    let mut match_count = 0_usize;
    for sheet_id in ordered_sheet_ids(doc, manifest)? {
        let state = read_sheet_state(doc, manifest, &sheet_id)?;
        let mut coordinates = state.cells.keys().copied().collect::<Vec<_>>();
        coordinates.sort_unstable();
        for (row, column) in coordinates {
            let Some(cell) = state.cells.get(&(row, column)) else {
                continue;
            };
            let Some(display) = searchable_text(cell) else {
                continue;
            };
            let Some(byte_index) = display.find(search) else {
                continue;
            };
            match_count += 1;
            if match_count > MAX_SPREADSHEET_TEXT_FIND_MATCHES {
                break;
            }
            if matches.len() >= limit {
                continue;
            }
            let occurrence = u32::try_from(matches.len() + 1).map_err(|_| {
                collaboration_error(
                    "office.collaboration.mutation_too_large",
                    "Spreadsheet text find exceeded the supported occurrence range.",
                )
            })?;
            matches.push(NativeOfficeCollaborationDocumentTextMatch {
                occurrence,
                text: search.to_owned(),
                paragraph_id: None,
                text_id: None,
                sheet_id: Some(sheet_id.clone()),
                row: Some(row),
                column: Some(column),
                container_kind: None,
                container_id: None,
                element_id: None,
                field_id: None,
                annotation_id: None,
                page_index: None,
                annotation_type: None,
                index_utf16: utf16_len(&display[..byte_index])?,
            });
        }
        if match_count > MAX_SPREADSHEET_TEXT_FIND_MATCHES {
            break;
        }
    }
    if match_count > MAX_SPREADSHEET_TEXT_FIND_MATCHES {
        match_count = MAX_SPREADSHEET_TEXT_FIND_MATCHES;
    }
    Ok(NativeOfficeCollaborationDocumentTextFindResult {
        search: search.to_owned(),
        match_count,
        truncated: match_count > limit,
        matches,
    })
}

fn searchable_text(cell: &JsonValue) -> Option<&str> {
    match cell.get("m") {
        Some(JsonValue::String(value)) => Some(value.as_str()),
        _ => match cell.get("v") {
            Some(JsonValue::String(value)) => Some(value.as_str()),
            _ => None,
        },
    }
}
