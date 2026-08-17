use a3s_use_core::UseResult;
use yrs::{GetString, ReadTxn, StateVector, Text, Transact};

use super::document::inspect_document;
use super::{
    collaboration_error, NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationMutation,
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
};

pub(in crate::collaboration) mod document;
mod pdf;
mod presentation;
mod spreadsheet;

use document::{apply_document_mutation, validate_document_mutation};
use pdf::{apply_pdf_mutation, validate_pdf_mutation};
use presentation::{apply_presentation_mutation, validate_presentation_mutation};
use spreadsheet::{apply_spreadsheet_mutation, validate_spreadsheet_mutation};

pub(super) fn validate_mutation_contract(
    manifest: &NativeOfficeCollaborationManifest,
    mode: NativeOfficeCollaborationMode,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let document_comment_mutation = matches!(
        mutation,
        NativeOfficeCollaborationMutation::DocumentCommentCreate { .. }
            | NativeOfficeCollaborationMutation::DocumentCommentReply { .. }
            | NativeOfficeCollaborationMutation::DocumentCommentSetResolved { .. }
            | NativeOfficeCollaborationMutation::DocumentCommentDelete { .. }
    );
    let document_suggestion_create = matches!(
        mutation,
        NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. }
    );
    let mutation_allowed = match mode {
        NativeOfficeCollaborationMode::Edit => !document_suggestion_create,
        NativeOfficeCollaborationMode::Comment => document_comment_mutation,
        NativeOfficeCollaborationMode::Suggest => document_suggestion_create,
        NativeOfficeCollaborationMode::View => false,
    };
    if !mutation_allowed {
        return Err(collaboration_error(
            "office.collaboration.mutation_forbidden",
            format!(
                "The '{}' collaboration mode cannot apply this Office mutation.",
                mode.as_str()
            ),
        )
        .with_suggestion(
            "Use edit mode for canonical content and tracked-change decisions, comment mode for Document comments, or suggest mode for Document suggestion creation.",
        )
        .with_detail("mode", mode.as_str()));
    }

    let mutation_kind = match mutation {
        NativeOfficeCollaborationMutation::MarkdownReplace { .. }
        | NativeOfficeCollaborationMutation::MarkdownSplice { .. } => {
            NativeOfficeCollaborationArtifactKind::Markdown
        }
        NativeOfficeCollaborationMutation::DocumentReplaceText { .. }
        | NativeOfficeCollaborationMutation::DocumentReplaceParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentSetPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentClearPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentSetTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentClearTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentReply { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentSetResolved { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentDelete { .. }
        | NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentSuggestionDecide { .. } => {
            NativeOfficeCollaborationArtifactKind::Document
        }
        NativeOfficeCollaborationMutation::SpreadsheetSetCell { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteCell { .. } => {
            NativeOfficeCollaborationArtifactKind::Spreadsheet
        }
        NativeOfficeCollaborationMutation::PresentationCreateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationDeleteElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveElement { .. } => {
            NativeOfficeCollaborationArtifactKind::Presentation
        }
        NativeOfficeCollaborationMutation::PdfCreateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfUpdateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfDeleteAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfSetFormValue { .. }
        | NativeOfficeCollaborationMutation::PdfProposeRedaction { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageRotation { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageDeletion { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageReorder { .. }
        | NativeOfficeCollaborationMutation::PdfDecideReview { .. } => {
            NativeOfficeCollaborationArtifactKind::Pdf
        }
    };
    if manifest.kind != mutation_kind {
        return Err(collaboration_error(
            "office.collaboration.mutation_kind_mismatch",
            format!(
                "The typed mutation targets '{}' content, but the replica contains '{}'.",
                mutation_kind.as_str(),
                manifest.kind.as_str()
            ),
        )
        .with_detail("mutationKind", mutation_kind.as_str())
        .with_detail("artifactKind", manifest.kind.as_str()));
    }

    let encoded_len = serde_json::to_vec(mutation)
        .map_err(|error| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                format!("Failed to encode the typed collaboration mutation: {error}"),
            )
        })?
        .len();
    if encoded_len > MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            format!(
                "The typed collaboration mutation is {encoded_len} bytes; the limit is {MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES} bytes."
            ),
        )
        .with_detail("bytes", encoded_len as u64)
        .with_detail(
            "maxBytes",
            MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES as u64,
        ));
    }
    if mutation_kind == NativeOfficeCollaborationArtifactKind::Document {
        validate_document_mutation(mutation)?;
    } else if mutation_kind == NativeOfficeCollaborationArtifactKind::Spreadsheet {
        validate_spreadsheet_mutation(mutation)?;
    } else if mutation_kind == NativeOfficeCollaborationArtifactKind::Presentation {
        validate_presentation_mutation(mutation)?;
    } else if mutation_kind == NativeOfficeCollaborationArtifactKind::Pdf {
        validate_pdf_mutation(mutation)?;
    }
    Ok(())
}

pub(super) fn apply_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
    before_state_vector: &StateVector,
) -> UseResult<Vec<u8>> {
    let inspection = inspect_document(doc, manifest)?;
    if !inspection
        .metadata
        .as_ref()
        .is_some_and(|metadata| metadata.initialized)
    {
        return Err(collaboration_error(
            "office.collaboration.mutation_uninitialized",
            "Typed collaboration mutations require initialized shared Office metadata.",
        )
        .with_suggestion(
            "Join or synchronize the browser bootstrap update before mutating canonical content.",
        ));
    }

    match mutation {
        NativeOfficeCollaborationMutation::MarkdownReplace { markdown } => {
            apply_markdown_replace(doc, manifest, markdown)?;
        }
        NativeOfficeCollaborationMutation::MarkdownSplice {
            index_utf16,
            delete_utf16,
            insert,
        } => {
            apply_markdown_splice(doc, manifest, *index_utf16, *delete_utf16, insert)?;
        }
        NativeOfficeCollaborationMutation::DocumentReplaceText { .. }
        | NativeOfficeCollaborationMutation::DocumentReplaceParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentSetPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentClearPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentSetTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentClearTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentReply { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentSetResolved { .. }
        | NativeOfficeCollaborationMutation::DocumentCommentDelete { .. }
        | NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. }
        | NativeOfficeCollaborationMutation::DocumentSuggestionDecide { .. } => {
            apply_document_mutation(doc, manifest, mutation)?;
        }
        NativeOfficeCollaborationMutation::SpreadsheetSetCell { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteCell { .. } => {
            apply_spreadsheet_mutation(doc, manifest, mutation)?;
        }
        NativeOfficeCollaborationMutation::PresentationCreateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationDeleteElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveElement { .. } => {
            apply_presentation_mutation(doc, manifest, mutation)?;
        }
        NativeOfficeCollaborationMutation::PdfCreateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfUpdateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfDeleteAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfSetFormValue { .. }
        | NativeOfficeCollaborationMutation::PdfProposeRedaction { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageRotation { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageDeletion { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageReorder { .. }
        | NativeOfficeCollaborationMutation::PdfDecideReview { .. } => {
            apply_pdf_mutation(doc, manifest, mutation)?;
        }
    }
    inspect_document(doc, manifest)?;
    Ok(doc
        .transact()
        .encode_state_as_update_v1(before_state_vector))
}

fn apply_markdown_replace(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    markdown: &str,
) -> UseResult<()> {
    let root = format!("{}.markdown.source", manifest.namespace);
    let text = doc.get_or_insert_text(root);
    let current = text.get_string(&doc.transact());
    let (index_utf16, delete_utf16, insert) = minimal_text_replacement(&current, markdown);
    if delete_utf16 == 0 && insert.is_empty() {
        return Ok(());
    }
    let mut transaction = doc.transact_mut();
    if delete_utf16 > 0 {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
    }
    if !insert.is_empty() {
        text.insert(&mut transaction, index_utf16, &insert);
    }
    Ok(())
}

fn apply_markdown_splice(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    index_utf16: u32,
    delete_utf16: u32,
    insert: &str,
) -> UseResult<()> {
    let root = format!("{}.markdown.source", manifest.namespace);
    let text = doc.get_or_insert_text(root);
    let current = text.get_string(&doc.transact());
    let current_len = utf16_len(&current)?;
    let end_utf16 = index_utf16
        .checked_add(delete_utf16)
        .ok_or_else(|| invalid_markdown_range(index_utf16, delete_utf16, current_len))?;
    if index_utf16 > current_len
        || end_utf16 > current_len
        || !is_utf16_boundary(&current, index_utf16)
        || !is_utf16_boundary(&current, end_utf16)
    {
        return Err(invalid_markdown_range(
            index_utf16,
            delete_utf16,
            current_len,
        ));
    }
    if delete_utf16 == 0 && insert.is_empty() {
        return Ok(());
    }
    let mut transaction = doc.transact_mut();
    if delete_utf16 > 0 {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
    }
    if !insert.is_empty() {
        text.insert(&mut transaction, index_utf16, insert);
    }
    Ok(())
}

fn minimal_text_replacement(current: &str, replacement: &str) -> (u32, u32, String) {
    let mut current_prefix_bytes = 0;
    let mut replacement_prefix_bytes = 0;
    let mut index_utf16 = 0;
    for (current_character, replacement_character) in current.chars().zip(replacement.chars()) {
        if current_character != replacement_character {
            break;
        }
        current_prefix_bytes += current_character.len_utf8();
        replacement_prefix_bytes += replacement_character.len_utf8();
        index_utf16 += current_character.len_utf16() as u32;
    }

    let current_tail = &current[current_prefix_bytes..];
    let replacement_tail = &replacement[replacement_prefix_bytes..];
    let mut current_suffix_bytes = 0;
    let mut replacement_suffix_bytes = 0;
    for (current_character, replacement_character) in current_tail
        .chars()
        .rev()
        .zip(replacement_tail.chars().rev())
    {
        if current_character != replacement_character {
            break;
        }
        current_suffix_bytes += current_character.len_utf8();
        replacement_suffix_bytes += replacement_character.len_utf8();
    }

    let current_middle = &current_tail[..current_tail.len() - current_suffix_bytes];
    let replacement_middle = &replacement_tail[..replacement_tail.len() - replacement_suffix_bytes];
    let delete_utf16 = current_middle.encode_utf16().count() as u32;
    let insert = replacement_middle.to_owned();
    (index_utf16, delete_utf16, insert)
}

fn utf16_len(value: &str) -> UseResult<u32> {
    u32::try_from(value.encode_utf16().count()).map_err(|_| {
        collaboration_error(
            "office.collaboration.mutation_too_large",
            "The collaboration text exceeds the supported UTF-16 offset range.",
        )
    })
}

fn is_utf16_boundary(value: &str, offset: u32) -> bool {
    if offset == 0 {
        return true;
    }
    let mut cursor = 0_u32;
    for character in value.chars() {
        cursor = cursor.saturating_add(character.len_utf16() as u32);
        if cursor == offset {
            return true;
        }
        if cursor > offset {
            return false;
        }
    }
    cursor == offset
}

fn invalid_markdown_range(
    index_utf16: u32,
    delete_utf16: u32,
    length_utf16: u32,
) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_range_invalid",
        "The Markdown splice range is outside the current source or splits a UTF-16 surrogate pair.",
    )
    .with_suggestion("Inspect the latest state vector and retry with UTF-16 code-unit offsets.")
    .with_detail("indexUtf16", index_utf16 as u64)
    .with_detail("deleteUtf16", delete_utf16 as u64)
    .with_detail("lengthUtf16", length_utf16 as u64)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::collaboration::document::new_replica_document;

    #[test]
    fn replacement_retains_unicode_scalar_boundaries() {
        assert_eq!(
            minimal_text_replacement("A😀B", "A🦀B"),
            (1, 2, "🦀".to_owned())
        );
        assert!(is_utf16_boundary("A😀B", 1));
        assert!(!is_utf16_boundary("A😀B", 2));
        assert!(is_utf16_boundary("A😀B", 3));
    }

    #[test]
    fn replica_documents_use_browser_utf16_offsets() {
        let doc = new_replica_document(
            7,
            "a3s.office",
            NativeOfficeCollaborationArtifactKind::Markdown,
        );
        let text = doc.get_or_insert_text("a3s.office.markdown.source");
        text.insert(&mut doc.transact_mut(), 0, "A😀B");
        text.remove_range(&mut doc.transact_mut(), 1, 2);
        assert_eq!(text.get_string(&doc.transact()), "AB");
    }
}
