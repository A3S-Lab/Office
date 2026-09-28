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

pub(in crate::collaboration) use pdf::project_pdf_content;
pub(in crate::collaboration) use presentation::project_presentation_content;
pub(in crate::collaboration) use spreadsheet::project_spreadsheet_content;

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
        | NativeOfficeCollaborationMutation::DocumentSplice { .. }
        | NativeOfficeCollaborationMutation::DocumentReplaceParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentSetPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentClearPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentSetTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentClearTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertTableRow { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteTableRow { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertListItem { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteListItem { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertSection { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteSection { .. }
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
        | NativeOfficeCollaborationMutation::SpreadsheetBatchCells { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetSplice { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetInsertRows { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteRows { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetInsertColumns { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteColumns { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetSortRows { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetCreateTable { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetUpdateTable { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteTable { .. } => {
            NativeOfficeCollaborationArtifactKind::Spreadsheet
        }
        NativeOfficeCollaborationMutation::PresentationCreateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationDeleteElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveSlide { .. }
        | NativeOfficeCollaborationMutation::PresentationSetGroup { .. }
        | NativeOfficeCollaborationMutation::PresentationSetBackground { .. }
        | NativeOfficeCollaborationMutation::PresentationSplice { .. }
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
            expected_slice,
            insert,
        } => {
            apply_markdown_splice(
                doc,
                manifest,
                *index_utf16,
                *delete_utf16,
                expected_slice,
                insert,
            )?;
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
        | NativeOfficeCollaborationMutation::DocumentSplice { .. }
        | NativeOfficeCollaborationMutation::DocumentReplaceParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentSetPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentClearPageColor { .. }
        | NativeOfficeCollaborationMutation::DocumentSetTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentClearTrackChanges { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteParagraph { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertTableRow { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteTableRow { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertListItem { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteListItem { .. }
        | NativeOfficeCollaborationMutation::DocumentInsertSection { .. }
        | NativeOfficeCollaborationMutation::DocumentDeleteSection { .. }
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
        | NativeOfficeCollaborationMutation::SpreadsheetBatchCells { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetSplice { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetInsertRows { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteRows { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetInsertColumns { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteColumns { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetSortRows { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetCreateTable { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetUpdateTable { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteTable { .. } => {
            apply_spreadsheet_mutation(doc, manifest, mutation)?;
        }
        NativeOfficeCollaborationMutation::PresentationCreateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement { .. }
        | NativeOfficeCollaborationMutation::PresentationDeleteElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveElement { .. }
        | NativeOfficeCollaborationMutation::PresentationMoveSlide { .. }
        | NativeOfficeCollaborationMutation::PresentationSetGroup { .. }
        | NativeOfficeCollaborationMutation::PresentationSetBackground { .. }
        | NativeOfficeCollaborationMutation::PresentationSplice { .. }
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

/// Caret at the end of an applied splice. Insert-only and no-op splices still
/// return this position. Other mutations have no caret.
pub(super) fn frame_caret(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<Option<super::NativeOfficeCollaborationFrameCaret>> {
    match mutation {
        NativeOfficeCollaborationMutation::MarkdownSplice {
            index_utf16,
            insert,
            ..
        } => {
            let next = splice_caret_index(*index_utf16, insert)?;
            Ok(Some(super::NativeOfficeCollaborationFrameCaret::Markdown {
                index_utf16: next,
            }))
        }
        NativeOfficeCollaborationMutation::DocumentSplice {
            paragraph_id,
            text_id,
            index_utf16,
            insert,
            ..
        } => {
            let next = splice_caret_index(*index_utf16, insert)?;
            Ok(Some(super::NativeOfficeCollaborationFrameCaret::Document {
                paragraph_id: paragraph_id.clone(),
                text_id: text_id.clone(),
                index_utf16: next,
            }))
        }
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id,
            row,
            column,
            ..
        } => Ok(Some(
            super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                sheet_id: sheet_id.clone(),
                row: *row,
                column: *column,
                index_utf16: None,
            },
        )),
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells { sheet_id, changes } => {
            let Some(last) = changes.last() else {
                return Ok(None);
            };
            Ok(Some(
                super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                    sheet_id: sheet_id.clone(),
                    row: last.row,
                    column: last.column,
                    index_utf16: None,
                },
            ))
        }
        NativeOfficeCollaborationMutation::SpreadsheetSplice {
            sheet_id,
            row,
            column,
            index_utf16,
            insert,
            ..
        } => {
            let next = splice_caret_index(*index_utf16, insert)?;
            Ok(Some(
                super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                    sheet_id: sheet_id.clone(),
                    row: *row,
                    column: *column,
                    index_utf16: Some(next),
                },
            ))
        }
        NativeOfficeCollaborationMutation::SpreadsheetInsertRows { sheet_id, at, .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteRows { sheet_id, at, .. } => Ok(
            Some(super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                sheet_id: sheet_id.clone(),
                row: *at,
                column: 0,
                index_utf16: None,
            }),
        ),
        NativeOfficeCollaborationMutation::SpreadsheetInsertColumns { sheet_id, at, .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteColumns { sheet_id, at, .. } => Ok(
            Some(super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                sheet_id: sheet_id.clone(),
                row: 0,
                column: *at,
                index_utf16: None,
            }),
        ),
        NativeOfficeCollaborationMutation::SpreadsheetSortRows {
            sheet_id,
            row,
            column,
            ..
        } => Ok(Some(
            super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                sheet_id: sheet_id.clone(),
                row: *row,
                column: *column,
                index_utf16: None,
            },
        )),
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetUpdateTable { .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteTable { .. } => {
            let (sheet_id, row, column) = spreadsheet::table_origin(mutation)?;
            Ok(Some(
                super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                    sheet_id,
                    row,
                    column,
                    index_utf16: None,
                },
            ))
        }
        NativeOfficeCollaborationMutation::PresentationSplice {
            container_kind,
            container_id,
            element_id,
            index_utf16,
            insert,
            ..
        } => {
            let next = splice_caret_index(*index_utf16, insert)?;
            Ok(Some(
                super::NativeOfficeCollaborationFrameCaret::Presentation {
                    container_kind: *container_kind,
                    container_id: container_id.clone(),
                    element_id: element_id.clone(),
                    index_utf16: Some(next),
                },
            ))
        }
        NativeOfficeCollaborationMutation::PresentationReplaceText {
            container_kind: Some(container_kind),
            container_id: Some(container_id),
            element_id: Some(element_id),
            ..
        }
        | NativeOfficeCollaborationMutation::PresentationUpdateElement {
            container_kind,
            container_id,
            element_id,
            ..
        } => Ok(Some(
            super::NativeOfficeCollaborationFrameCaret::Presentation {
                container_kind: *container_kind,
                container_id: container_id.clone(),
                element_id: element_id.clone(),
                index_utf16: None,
            },
        )),
        NativeOfficeCollaborationMutation::PdfSetFormValue { field_id, .. } => {
            Ok(Some(super::NativeOfficeCollaborationFrameCaret::PdfField {
                field_id: field_id.clone(),
            }))
        }
        NativeOfficeCollaborationMutation::PdfCreateAnnotation { annotation_id, .. }
        | NativeOfficeCollaborationMutation::PdfUpdateAnnotation { annotation_id, .. } => Ok(Some(
            super::NativeOfficeCollaborationFrameCaret::PdfAnnotation {
                annotation_id: annotation_id.clone(),
            },
        )),
        _ => Ok(None),
    }
}

fn splice_caret_index(index_utf16: u32, insert: &str) -> UseResult<u32> {
    let inserted = utf16_len(insert)?;
    index_utf16.checked_add(inserted).ok_or_else(|| {
        collaboration_error(
            "office.collaboration.mutation_too_large",
            "The splice caret exceeds the supported UTF-16 offset range.",
        )
    })
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
    use crate::NativeOfficeCollaborationSpreadsheetCellChange;
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

    #[test]
    fn frame_caret_names_the_written_target_without_a_text_offset() {
        let cell = frame_caret(&NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: "sheet".to_owned(),
            row: 2,
            column: 3,
            expected_cell: None,
            next_cell: serde_json::json!({ "value": "中" }),
        })
        .unwrap()
        .unwrap();
        let cell_json = serde_json::to_value(&cell).unwrap();
        assert_eq!(cell_json["kind"], "spreadsheet");
        assert_eq!(cell_json["row"], 2);
        assert!(cell_json.get("indexUtf16").is_none());

        let batch = frame_caret(&NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet".to_owned(),
            changes: vec![
                NativeOfficeCollaborationSpreadsheetCellChange {
                    row: 0,
                    column: 0,
                    expected_cell: None,
                    next_cell: Some(serde_json::json!("a")),
                },
                NativeOfficeCollaborationSpreadsheetCellChange {
                    row: 4,
                    column: 1,
                    expected_cell: None,
                    next_cell: Some(serde_json::json!("b")),
                },
            ],
        })
        .unwrap()
        .unwrap();
        assert_eq!(
            batch,
            super::super::NativeOfficeCollaborationFrameCaret::Spreadsheet {
                sheet_id: "sheet".to_owned(),
                row: 4,
                column: 1,
                index_utf16: None,
            }
        );

        let form = frame_caret(&NativeOfficeCollaborationMutation::PdfSetFormValue {
            field_id: "name".to_owned(),
            expected_value: String::new(),
            value: "中".to_owned(),
            search: Some("中".to_owned()),
            index_utf16: Some(0),
        })
        .unwrap()
        .unwrap();
        let form_json = serde_json::to_value(&form).unwrap();
        assert_eq!(form_json["kind"], "pdf-field");
        assert_eq!(form_json["fieldId"], "name");
        assert!(form_json.get("indexUtf16").is_none());

        let note = frame_caret(&NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
            annotation_id: "note-1".to_owned(),
            expected_annotation: serde_json::json!({}),
            next_annotation: serde_json::json!({ "contents": "中" }),
            search: Some("a".to_owned()),
            index_utf16: Some(4),
        })
        .unwrap()
        .unwrap();
        let note_json = serde_json::to_value(&note).unwrap();
        assert_eq!(note_json["kind"], "pdf-annotation");
        assert_eq!(note_json["annotationId"], "note-1");
        assert!(note_json.get("indexUtf16").is_none());
        assert!(note_json.get("pageIndex").is_none());

        let unanchored = frame_caret(
            &NativeOfficeCollaborationMutation::PresentationReplaceText {
                search: "old".to_owned(),
                replacement: "new".to_owned(),
                expected_matches: 1,
                occurrence: None,
                container_kind: None,
                container_id: None,
                element_id: None,
                index_utf16: Some(2),
            },
        )
        .unwrap();
        assert_eq!(unanchored, None);
    }
}
