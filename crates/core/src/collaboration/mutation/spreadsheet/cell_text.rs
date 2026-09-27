use std::sync::Arc;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Any, GetString, Map, MapPrelim, Out, Text, TextPrelim, Transact};

use super::json::{encode_flat_json_key, FlatJsonEntryKind};
use super::state::{encode_cell_field_key, encode_coordinate, read_sheet_state, CELL_TEXT_KEY};
use super::{
    invalid_shared_spreadsheet, spreadsheet_match_conflict, NativeOfficeCollaborationManifest,
};

pub(super) fn validate_spreadsheet_splice(
    index_utf16: u32,
    delete_utf16: u32,
    expected_slice: &str,
    insert: &str,
) -> UseResult<()> {
    if !expected_slice.is_empty() && delete_utf16 == 0 {
        return Err(super::invalid_spreadsheet_mutation(
            "A spreadsheet splice with an expected slice must delete that range.",
        ));
    }
    if utf16_len(expected_slice)? != delete_utf16 {
        return Err(super::invalid_spreadsheet_mutation(
            "A spreadsheet splice expected slice must cover deleteUtf16 code units.",
        ));
    }
    utf16_len(insert)?;
    let _ = index_utf16;
    Ok(())
}

pub(super) fn splice_spreadsheet_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    row: u32,
    column: u32,
    index_utf16: u32,
    delete_utf16: u32,
    expected_slice: &str,
    insert: &str,
) -> UseResult<()> {
    let state = read_sheet_state(doc, manifest, sheet_id)?;
    let cell = state.cells.get(&(row, column)).ok_or_else(|| {
        spreadsheet_match_conflict(format!(
            "Spreadsheet cell '{row}:{column}' in sheet '{sheet_id}' does not exist."
        ))
    })?;
    let initial = plain_cell_text(cell)?;
    let end = index_utf16.checked_add(delete_utf16).ok_or_else(|| {
        super::super::super::collaboration_error(
            "office.collaboration.mutation_too_large",
            "The spreadsheet splice range exceeds the supported UTF-16 offset range.",
        )
    })?;
    let coordinate = encode_coordinate(row, column);
    let current = {
        let transaction = doc.transact();
        match state.record.get(&transaction, CELL_TEXT_KEY) {
            Some(Out::YMap(texts)) => match texts.get(&transaction, coordinate.as_str()) {
                Some(Out::YText(text)) => text.get_string(&transaction),
                Some(_) => {
                    return Err(invalid_shared_spreadsheet(
                        "The spreadsheet cell text is not a collaborative text node.",
                    ));
                }
                None => initial.clone(),
            },
            Some(_) => {
                return Err(invalid_shared_spreadsheet(
                    "The shared Spreadsheet cell text map is not a map.",
                ));
            }
            None => initial.clone(),
        }
    };
    if !is_utf16_boundary(&current, index_utf16) || !is_utf16_boundary(&current, end) {
        return Err(super::super::super::collaboration_error(
            "office.collaboration.mutation_range_invalid",
            "The spreadsheet splice range splits a Unicode scalar or leaves the text.",
        ));
    }
    let Some(observed) = utf16_slice(&current, index_utf16, end) else {
        return Err(super::super::super::collaboration_error(
            "office.collaboration.mutation_range_invalid",
            "The spreadsheet splice range splits a Unicode scalar or leaves the text.",
        ));
    };
    if observed != expected_slice {
        return Err(spreadsheet_match_conflict(format!(
            "Spreadsheet cell '{row}:{column}' text slice drifted before the splice. Observed '{observed}'."
        )));
    }
    let next = splice_string(&current, index_utf16, delete_utf16, insert)?;
    let value_key = encode_cell_field_key(
        row,
        column,
        &encode_flat_json_key(FlatJsonEntryKind::Value, &["v".to_owned()])?,
    )?;
    let display_key = encode_cell_field_key(
        row,
        column,
        &encode_flat_json_key(FlatJsonEntryKind::Value, &["m".to_owned()])?,
    )?;
    let mut transaction = doc.transact_mut();
    let texts = match state.record.get(&transaction, CELL_TEXT_KEY) {
        Some(Out::YMap(texts)) => texts,
        None => state
            .record
            .insert(&mut transaction, CELL_TEXT_KEY, MapPrelim::default()),
        Some(_) => {
            return Err(invalid_shared_spreadsheet(
                "The shared Spreadsheet cell text map is not a map.",
            ));
        }
    };
    let text = match texts.get(&transaction, coordinate.as_str()) {
        Some(Out::YText(text)) => text,
        None => texts.insert(
            &mut transaction,
            coordinate.as_str(),
            TextPrelim::new(initial),
        ),
        Some(_) => {
            return Err(invalid_shared_spreadsheet(
                "The spreadsheet cell text is not a collaborative text node.",
            ));
        }
    };
    if delete_utf16 > 0 {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
    }
    if !insert.is_empty() {
        text.insert(&mut transaction, index_utf16, insert);
    }
    let stored = Any::String(Arc::from(next.as_str()));
    state
        .fields
        .insert(&mut transaction, value_key.as_str(), stored.clone());
    state
        .fields
        .insert(&mut transaction, display_key.as_str(), stored);
    Ok(())
}

pub(super) fn reconcile_cell_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    row: u32,
    column: u32,
    next: Option<&JsonValue>,
) -> UseResult<()> {
    let sheets = doc.get_or_insert_map(format!("{}.spreadsheet.sheets", manifest.namespace));
    let coordinate = encode_coordinate(row, column);
    let plain = next.and_then(|cell| plain_cell_text(cell).ok());
    let mut transaction = doc.transact_mut();
    let Some(Out::YMap(record)) = sheets.get(&transaction, sheet_id) else {
        return Ok(());
    };
    let Some(Out::YMap(texts)) = record.get(&transaction, CELL_TEXT_KEY) else {
        return Ok(());
    };
    let Some(Out::YText(text)) = texts.get(&transaction, coordinate.as_str()) else {
        return Ok(());
    };
    let Some(plain) = plain else {
        texts.remove(&mut transaction, coordinate.as_str());
        return Ok(());
    };
    if text.get_string(&transaction) == plain {
        return Ok(());
    }
    let length = u32::try_from(text.get_string(&transaction).encode_utf16().count()).unwrap_or(0);
    if length > 0 {
        text.remove_range(&mut transaction, 0, length);
    }
    if !plain.is_empty() {
        text.insert(&mut transaction, 0, plain.as_str());
    }
    Ok(())
}

fn plain_cell_text(cell: &JsonValue) -> UseResult<String> {
    if cell.get("f").is_some() {
        return Err(spreadsheet_match_conflict(
            "A formula cell is field-addressed. Use spreadsheet-set-cell instead of spreadsheet-splice.",
        ));
    }
    match (cell.get("v"), cell.get("m")) {
        (Some(JsonValue::String(value)), Some(JsonValue::String(display))) if value == display => {
            Ok(value.clone())
        }
        (Some(JsonValue::String(value)), None) => Ok(value.clone()),
        (None, Some(JsonValue::String(display))) => Ok(display.clone()),
        (Some(JsonValue::Number(_)), _) => Err(spreadsheet_match_conflict(
            "A numeric cell is field-addressed. Use spreadsheet-set-cell instead of spreadsheet-splice.",
        )),
        _ => Err(spreadsheet_match_conflict(
            "Spreadsheet splice requires one plain text value. Cached values and styles stay on spreadsheet-set-cell.",
        )),
    }
}

fn splice_string(
    current: &str,
    index_utf16: u32,
    delete_utf16: u32,
    insert: &str,
) -> UseResult<String> {
    let start = utf16_byte(current, index_utf16)?;
    let end = utf16_byte(current, index_utf16 + delete_utf16)?;
    let mut next = String::with_capacity(current.len() + insert.len());
    next.push_str(&current[..start]);
    next.push_str(insert);
    next.push_str(&current[end..]);
    Ok(next)
}

fn utf16_byte(value: &str, offset: u32) -> UseResult<usize> {
    let mut seen = 0u32;
    for (byte, character) in value.char_indices() {
        if seen == offset {
            return Ok(byte);
        }
        seen += u32::try_from(character.len_utf16()).unwrap_or(0);
    }
    if seen == offset {
        return Ok(value.len());
    }
    Err(super::super::super::collaboration_error(
        "office.collaboration.mutation_range_invalid",
        "The spreadsheet splice range splits a Unicode scalar or leaves the text.",
    ))
}

fn utf16_len(value: &str) -> UseResult<u32> {
    u32::try_from(value.encode_utf16().count()).map_err(|_| {
        super::super::super::collaboration_error(
            "office.collaboration.mutation_too_large",
            "The spreadsheet splice text exceeds the supported UTF-16 offset range.",
        )
    })
}

fn is_utf16_boundary(value: &str, offset: u32) -> bool {
    let mut seen = 0u32;
    for character in value.chars() {
        if seen == offset {
            return true;
        }
        if seen > offset {
            return false;
        }
        seen += u32::try_from(character.len_utf16()).unwrap_or(0);
    }
    seen == offset
}

fn utf16_slice(value: &str, start: u32, end: u32) -> Option<String> {
    let mut seen = 0u32;
    let mut start_byte = None;
    let mut end_byte = None;
    for (byte, character) in value.char_indices() {
        if seen == start {
            start_byte = Some(byte);
        }
        if seen == end {
            end_byte = Some(byte);
            break;
        }
        seen += u32::try_from(character.len_utf16()).unwrap_or(0);
    }
    if start_byte.is_none() && seen == start {
        start_byte = Some(value.len());
    }
    if end_byte.is_none() && seen == end {
        end_byte = Some(value.len());
    }
    Some(value[start_byte?..end_byte?].to_owned())
}
