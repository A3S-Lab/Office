use a3s_use_core::UseResult;
use yrs::{Any, Map, Out, Transact};

use super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
};

mod identity;
mod paragraph;
mod text;

pub(super) fn validate_document_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentReplaceText {
            search,
            expected_matches,
            ..
        } => text::validate_text_replacement(search, *expected_matches),
        NativeOfficeCollaborationMutation::DocumentInsertParagraph {
            anchor_paragraph_id,
            paragraph_id,
            text_id,
            text,
            ..
        } => paragraph::validate_insert_paragraph(anchor_paragraph_id, paragraph_id, text_id, text),
        NativeOfficeCollaborationMutation::DocumentDeleteParagraph {
            paragraph_id,
            expected_text_id,
            expected_text,
        } => paragraph::validate_delete_paragraph(paragraph_id, expected_text_id, expected_text),
        NativeOfficeCollaborationMutation::DocumentSetPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentClearPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentSetTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentClearTrackChanges { .. } => Ok(()),
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a Document mutation.",
        )),
    }
}

pub(super) fn apply_document_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentReplaceText {
            search,
            replacement,
            expected_matches,
        } => text::replace_document_text(doc, manifest, search, replacement, *expected_matches),
        NativeOfficeCollaborationMutation::DocumentInsertParagraph {
            anchor_paragraph_id,
            position,
            paragraph_id,
            text_id,
            text,
        } => paragraph::insert_paragraph(
            doc,
            manifest,
            anchor_paragraph_id,
            *position,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteParagraph {
            paragraph_id,
            expected_text_id,
            expected_text,
        } => paragraph::delete_paragraph(
            doc,
            manifest,
            paragraph_id,
            expected_text_id,
            expected_text,
        ),
        NativeOfficeCollaborationMutation::DocumentSetPageColor { page_color } => {
            set_page_color(doc, manifest, Some(page_color))
        }
        NativeOfficeCollaborationMutation::DocumentClearPageColor { .. } => {
            set_page_color(doc, manifest, None)
        }
        NativeOfficeCollaborationMutation::DocumentSetTrackChanges { track_changes } => {
            set_track_changes(doc, manifest, Some(*track_changes))
        }
        NativeOfficeCollaborationMutation::DocumentClearTrackChanges { .. } => {
            set_track_changes(doc, manifest, None)
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a Document mutation.",
        )),
    }
}

fn set_page_color(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    page_color: Option<&str>,
) -> UseResult<()> {
    let root = format!("{}.document.options", manifest.namespace);
    let options = doc.get_or_insert_map(root);
    let transaction = doc.transact();
    let current = match options.get(&transaction, "pageColor") {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        Some(_) => return Err(invalid_document_option("page color")),
        None => None,
    };
    if current.as_deref() == page_color {
        return Ok(());
    }
    drop(transaction);
    let mut transaction = doc.transact_mut();
    match page_color {
        Some(value) => {
            options.insert(&mut transaction, "pageColor", value);
        }
        None => {
            options.remove(&mut transaction, "pageColor");
        }
    }
    Ok(())
}

fn set_track_changes(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    track_changes: Option<bool>,
) -> UseResult<()> {
    let root = format!("{}.document.options", manifest.namespace);
    let options = doc.get_or_insert_map(root);
    let transaction = doc.transact();
    let current = match options.get(&transaction, "trackChanges") {
        Some(Out::Any(Any::Bool(value))) => Some(value),
        Some(_) => return Err(invalid_document_option("track-changes setting")),
        None => None,
    };
    if current == track_changes {
        return Ok(());
    }
    drop(transaction);
    let mut transaction = doc.transact_mut();
    match track_changes {
        Some(value) => {
            options.insert(&mut transaction, "trackChanges", value);
        }
        None => {
            options.remove(&mut transaction, "trackChanges");
        }
    }
    Ok(())
}

fn invalid_document_option(label: &str) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.content_invalid",
        format!("The shared Document collaboration {label} is invalid."),
    )
}
