use a3s_use_core::UseResult;
use yrs::{ReadTxn, StateVector, Transact};

use super::document::inspect_document;
use super::{
    collaboration_error, NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationMutation,
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
};

pub(in crate::collaboration) mod document;
pub(in crate::collaboration) mod markdown;
pub(in crate::collaboration) mod pdf;
pub(in crate::collaboration) mod presentation;
pub(in crate::collaboration) mod spreadsheet;

use document::{apply_document_mutation, validate_document_mutation};
use markdown::{
    apply_markdown_replace, apply_markdown_splice, replace_markdown_text,
    validate_text_replacement as validate_markdown_text_replacement,
};
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
        | NativeOfficeCollaborationMutation::MarkdownSplice { .. }
        | NativeOfficeCollaborationMutation::MarkdownReplaceText { .. } => {
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
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteCell { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetBatchCells { .. } => {
            NativeOfficeCollaborationArtifactKind::Spreadsheet
        }
        NativeOfficeCollaborationMutation::PresentationCreateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationDeleteElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveElement { .. }
        | NativeOfficeCollaborationMutation::PresentationReplaceText { .. } => {
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
    } else if mutation_kind == NativeOfficeCollaborationArtifactKind::Markdown {
        if let NativeOfficeCollaborationMutation::MarkdownReplaceText {
            search,
            expected_matches,
            occurrence,
            index_utf16,
            ..
        } = mutation
        {
            validate_markdown_text_replacement(
                search,
                *expected_matches,
                *occurrence,
                *index_utf16,
            )?;
        }
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
        NativeOfficeCollaborationMutation::MarkdownReplace {
            expected_markdown,
            markdown,
        } => {
            apply_markdown_replace(doc, manifest, expected_markdown, markdown)?;
        }
        NativeOfficeCollaborationMutation::MarkdownSplice {
            index_utf16,
            delete_utf16,
            insert,
        } => {
            apply_markdown_splice(doc, manifest, *index_utf16, *delete_utf16, insert)?;
        }
        NativeOfficeCollaborationMutation::MarkdownReplaceText {
            search,
            replacement,
            expected_matches,
            occurrence,
            index_utf16,
        } => {
            replace_markdown_text(
                doc,
                manifest,
                search,
                replacement,
                *expected_matches,
                *occurrence,
                *index_utf16,
            )?;
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
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteCell { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetBatchCells { .. } => {
            apply_spreadsheet_mutation(doc, manifest, mutation)?;
        }
        NativeOfficeCollaborationMutation::PresentationCreateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationDeleteElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveElement { .. }
        | NativeOfficeCollaborationMutation::PresentationReplaceText { .. } => {
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

pub(super) fn utf16_len(value: &str) -> UseResult<u32> {
    u32::try_from(value.encode_utf16().count()).map_err(|_| {
        collaboration_error(
            "office.collaboration.mutation_too_large",
            "The collaboration text exceeds the supported UTF-16 offset range.",
        )
    })
}

pub(super) fn is_utf16_boundary(value: &str, offset: u32) -> bool {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::collaboration::document::new_replica_document;
    use yrs::{GetString, Text};

    #[test]
    fn utf16_boundaries_reject_surrogate_splits() {
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
