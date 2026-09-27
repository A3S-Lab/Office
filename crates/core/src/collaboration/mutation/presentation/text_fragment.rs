use a3s_use_core::UseResult;
use yrs::{GetString, Map, Out, Text, TextPrelim, Transact};

use super::super::super::{collaboration_error, NativeOfficeCollaborationManifest};
use super::state::read_element_state;
use super::{presentation_match_conflict, NativeOfficeCollaborationPresentationContainerKind};

pub(super) const TEXT_FRAGMENT: &str = "textFragment";

pub(super) fn validate_presentation_splice(
    container_id: &str,
    element_id: &str,
    index_utf16: u32,
    delete_utf16: u32,
    expected_slice: &str,
    insert: &str,
) -> UseResult<()> {
    super::json::validate_presentation_identifier(container_id, "container", false)?;
    super::json::validate_presentation_identifier(element_id, "scene element", false)?;
    if !expected_slice.is_empty() && delete_utf16 == 0 {
        return Err(invalid_splice(
            "A presentation splice with an expected slice must delete that range.",
        ));
    }
    let expected_len = utf16_len(expected_slice)?;
    if expected_len != delete_utf16 {
        return Err(invalid_splice(
            "A presentation splice expected slice must cover deleteUtf16 code units.",
        ));
    }
    if !is_utf16_boundary(insert, 0) || utf16_len(insert).is_err() {
        return Err(invalid_splice(
            "A presentation splice insert must be valid Unicode.",
        ));
    }
    let _ = index_utf16;
    Ok(())
}

pub(super) fn splice_presentation_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    container_kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: &str,
    element_id: &str,
    index_utf16: u32,
    delete_utf16: u32,
    expected_slice: &str,
    insert: &str,
) -> UseResult<()> {
    let state = read_element_state(doc, manifest, container_kind, container_id)?;
    let record = state.records.get(element_id).ok_or_else(|| {
        presentation_match_conflict(format!(
            "Presentation scene element ID '{element_id}' does not exist."
        ))
    })?;
    if record.tombstoned {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element ID '{element_id}' was deleted."
        )));
    }
    let initial = record
        .value
        .get("text")
        .and_then(|value| value.as_str())
        .unwrap_or("")
        .to_owned();
    let end = index_utf16.checked_add(delete_utf16).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.mutation_too_large",
            "The presentation splice range exceeds the supported UTF-16 offset range.",
        )
    })?;
    let record_map = record.map.clone();
    let current = {
        let transaction = doc.transact();
        match record_map.get(&transaction, TEXT_FRAGMENT) {
            Some(Out::YText(text)) => text.get_string(&transaction),
            Some(_) => {
                return Err(super::invalid_shared_presentation(
                    "The presentation text fragment is not a collaborative text node.",
                ));
            }
            None => initial,
        }
    };
    if !is_utf16_boundary(&current, index_utf16) || !is_utf16_boundary(&current, end) {
        return Err(collaboration_error(
            "office.collaboration.mutation_range_invalid",
            "The presentation splice range splits a Unicode scalar or leaves the text.",
        ));
    }
    let Some(observed) = utf16_slice(&current, index_utf16, end) else {
        return Err(collaboration_error(
            "office.collaboration.mutation_range_invalid",
            "The presentation splice range splits a Unicode scalar or leaves the text.",
        ));
    };
    if observed != expected_slice {
        return Err(presentation_match_conflict(format!(
            "Presentation scene element '{element_id}' text slice drifted before the splice. Observed '{observed}'."
        )));
    }
    let mut transaction = doc.transact_mut();
    let text = match record_map.get(&transaction, TEXT_FRAGMENT) {
        Some(Out::YText(text)) => text,
        None => record_map.insert(
            &mut transaction,
            TEXT_FRAGMENT,
            TextPrelim::new(current.clone()),
        ),
        Some(_) => {
            return Err(super::invalid_shared_presentation(
                "The presentation text fragment is not a collaborative text node.",
            ));
        }
    };
    if delete_utf16 > 0 {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
    }
    if !insert.is_empty() {
        text.insert(&mut transaction, index_utf16, insert);
    }
    Ok(())
}

pub(super) fn sync_text_fragment_to_string(
    map: &yrs::MapRef,
    transaction: &mut yrs::TransactionMut,
    next: &str,
) {
    let Some(Out::YText(text)) = map.get(transaction, TEXT_FRAGMENT) else {
        return;
    };
    let current = text.get_string(transaction);
    if current == next {
        return;
    }
    let length = u32::try_from(current.encode_utf16().count()).unwrap_or(u32::MAX);
    if length > 0 {
        text.remove_range(transaction, 0, length);
    }
    if !next.is_empty() {
        text.insert(transaction, 0, next);
    }
}

fn utf16_len(value: &str) -> UseResult<u32> {
    u32::try_from(value.encode_utf16().count()).map_err(|_| {
        collaboration_error(
            "office.collaboration.mutation_too_large",
            "The presentation splice text exceeds the supported UTF-16 offset range.",
        )
    })
}

fn is_utf16_boundary(value: &str, offset: u32) -> bool {
    let mut seen = 0u32;
    for character in value.chars() {
        if seen == offset {
            return true;
        }
        let width = u32::try_from(character.encode_utf16(&mut [0; 2]).len()).unwrap_or(0);
        if seen.saturating_add(width) > offset && seen < offset {
            return false;
        }
        seen = seen.saturating_add(width);
    }
    seen == offset
}

fn utf16_slice(value: &str, start: u32, end: u32) -> Option<&str> {
    if start > end || !is_utf16_boundary(value, start) || !is_utf16_boundary(value, end) {
        return None;
    }
    let mut start_byte = None;
    let mut end_byte = None;
    let mut utf16 = 0u32;
    for (byte, character) in value.char_indices() {
        if start_byte.is_none() && utf16 == start {
            start_byte = Some(byte);
        }
        if end_byte.is_none() && utf16 == end {
            end_byte = Some(byte);
        }
        utf16 = utf16.saturating_add(u32::try_from(character.len_utf16()).unwrap_or(0));
    }
    if start_byte.is_none() && utf16 == start {
        start_byte = Some(value.len());
    }
    if end_byte.is_none() && utf16 == end {
        end_byte = Some(value.len());
    }
    Some(&value[start_byte?..end_byte?])
}

fn invalid_splice(message: &str) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}
