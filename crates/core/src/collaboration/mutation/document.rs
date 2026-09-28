use a3s_use_core::UseResult;
use yrs::{Any, Map, Out, Transact};

use super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
};

pub(in crate::collaboration) mod comment;
pub(in crate::collaboration) mod identity;
mod list_item;
mod paragraph;
mod section;
mod structure;
pub(in crate::collaboration) mod suggestion;
mod table_row;
pub(in crate::collaboration) mod text;

pub(super) fn validate_document_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentReplaceText {
            search,
            expected_matches,
            occurrence,
            paragraph_id,
            text_id,
            index_utf16,
            ..
        } => text::validate_text_replacement(
            search,
            *expected_matches,
            *occurrence,
            paragraph_id.as_deref(),
            text_id.as_deref(),
            *index_utf16,
        ),
        NativeOfficeCollaborationMutation::DocumentSplice {
            paragraph_id,
            text_id,
            ..
        } => text::validate_document_splice(paragraph_id, text_id),
        NativeOfficeCollaborationMutation::DocumentReplaceParagraph {
            paragraph_id,
            expected_text_id,
            expected_text,
            replacement,
        } => paragraph::validate_replace_paragraph(
            paragraph_id,
            expected_text_id,
            expected_text,
            replacement,
        ),
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
        NativeOfficeCollaborationMutation::DocumentInsertTableRow {
            anchor_row_id,
            expected_row_text_id,
            row_id,
            row_text_id,
            paragraph_id,
            text_id,
            text,
            ..
        } => table_row::validate_insert_table_row(
            anchor_row_id,
            expected_row_text_id,
            row_id,
            row_text_id,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteTableRow {
            row_id,
            expected_row_text_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        } => table_row::validate_delete_table_row(
            row_id,
            expected_row_text_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        ),
        NativeOfficeCollaborationMutation::DocumentInsertListItem {
            anchor_paragraph_id,
            expected_text_id,
            paragraph_id,
            text_id,
            text,
            ..
        } => list_item::validate_insert_list_item(
            anchor_paragraph_id,
            expected_text_id,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteListItem {
            paragraph_id,
            expected_text_id,
            expected_text,
        } => list_item::validate_delete_list_item(paragraph_id, expected_text_id, expected_text),
        NativeOfficeCollaborationMutation::DocumentInsertSection {
            anchor_section_id,
            expected_paragraph_id,
            expected_text_id,
            section_id,
            paragraph_id,
            text_id,
            text,
            ..
        } => section::validate_insert_section(
            anchor_section_id,
            expected_paragraph_id,
            expected_text_id,
            section_id,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteSection {
            section_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        } => section::validate_delete_section(
            section_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        ),
        NativeOfficeCollaborationMutation::DocumentCommentCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentReply { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentSetResolved { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentDelete { .. } => {
            comment::validate_comment_mutation(mutation)
        }
        NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentSuggestionDecide { .. } => {
            suggestion::validate_suggestion_mutation(mutation)
        }
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
            occurrence,
            paragraph_id,
            text_id,
            index_utf16,
        } => text::replace_document_text(
            doc,
            manifest,
            search,
            replacement,
            *expected_matches,
            *occurrence,
            paragraph_id.as_deref(),
            text_id.as_deref(),
            *index_utf16,
        ),
        NativeOfficeCollaborationMutation::DocumentSplice {
            paragraph_id,
            text_id,
            index_utf16,
            delete_utf16,
            expected_slice,
            insert,
        } => text::splice_document_text(
            doc,
            manifest,
            paragraph_id,
            text_id,
            *index_utf16,
            *delete_utf16,
            expected_slice,
            insert,
        ),
        NativeOfficeCollaborationMutation::DocumentReplaceParagraph {
            paragraph_id,
            expected_text_id,
            expected_text,
            replacement,
        } => paragraph::replace_paragraph(
            doc,
            manifest,
            paragraph_id,
            expected_text_id,
            expected_text,
            replacement,
        ),
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
        NativeOfficeCollaborationMutation::DocumentInsertTableRow {
            anchor_row_id,
            expected_row_text_id,
            position,
            row_id,
            row_text_id,
            paragraph_id,
            text_id,
            text,
        } => table_row::insert_table_row(
            doc,
            manifest,
            anchor_row_id,
            expected_row_text_id,
            *position,
            row_id,
            row_text_id,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteTableRow {
            row_id,
            expected_row_text_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        } => table_row::delete_table_row(
            doc,
            manifest,
            row_id,
            expected_row_text_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        ),
        NativeOfficeCollaborationMutation::DocumentInsertListItem {
            anchor_paragraph_id,
            expected_text_id,
            position,
            paragraph_id,
            text_id,
            text,
        } => list_item::insert_list_item(
            doc,
            manifest,
            anchor_paragraph_id,
            expected_text_id,
            *position,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteListItem {
            paragraph_id,
            expected_text_id,
            expected_text,
        } => list_item::delete_list_item(
            doc,
            manifest,
            paragraph_id,
            expected_text_id,
            expected_text,
        ),
        NativeOfficeCollaborationMutation::DocumentInsertSection {
            anchor_section_id,
            expected_paragraph_id,
            expected_text_id,
            position,
            section_id,
            paragraph_id,
            text_id,
            text,
        } => section::insert_section(
            doc,
            manifest,
            anchor_section_id,
            expected_paragraph_id,
            expected_text_id,
            *position,
            section_id,
            paragraph_id,
            text_id,
            text,
        ),
        NativeOfficeCollaborationMutation::DocumentDeleteSection {
            section_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        } => section::delete_section(
            doc,
            manifest,
            section_id,
            paragraph_id,
            expected_text_id,
            expected_text,
        ),
        NativeOfficeCollaborationMutation::DocumentCommentCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentReply { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentSetResolved { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentDelete { .. } => {
            comment::apply_comment_mutation(doc, manifest, mutation)
        }
        NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentSuggestionDecide { .. } => {
            suggestion::apply_suggestion_mutation(doc, manifest, mutation)
        }
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
