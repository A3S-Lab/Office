use a3s_office::{
    NativeOfficeCollaborationMutation, NativeOfficeCollaborationParagraphPosition,
    NativeOfficeCollaborationPdfAnnotationSource, NativeOfficeCollaborationPdfRect,
    NativeOfficeCollaborationPdfReviewDecision, NativeOfficeCollaborationPdfReviewTargetKind,
    NativeOfficeCollaborationPresentationContainerKind,
};
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeCollaborationParagraphPosition {
    Before,
    After,
}

impl From<OfficeCollaborationParagraphPosition> for NativeOfficeCollaborationParagraphPosition {
    fn from(value: OfficeCollaborationParagraphPosition) -> Self {
        match value {
            OfficeCollaborationParagraphPosition::Before => Self::Before,
            OfficeCollaborationParagraphPosition::After => Self::After,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeCollaborationPresentationContainerKind {
    Slide,
    Master,
    Layout,
}

impl From<OfficeCollaborationPresentationContainerKind>
    for NativeOfficeCollaborationPresentationContainerKind
{
    fn from(value: OfficeCollaborationPresentationContainerKind) -> Self {
        match value {
            OfficeCollaborationPresentationContainerKind::Slide => Self::Slide,
            OfficeCollaborationPresentationContainerKind::Master => Self::Master,
            OfficeCollaborationPresentationContainerKind::Layout => Self::Layout,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Deserialize, Serialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeCollaborationPdfRect {
    left: f64,
    top: f64,
    right: f64,
    bottom: f64,
}

impl From<OfficeCollaborationPdfRect> for NativeOfficeCollaborationPdfRect {
    fn from(value: OfficeCollaborationPdfRect) -> Self {
        Self {
            left: value.left,
            top: value.top,
            right: value.right,
            bottom: value.bottom,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeCollaborationPdfReviewTargetKind {
    Redaction,
    PageOperation,
}

impl From<OfficeCollaborationPdfReviewTargetKind> for NativeOfficeCollaborationPdfReviewTargetKind {
    fn from(value: OfficeCollaborationPdfReviewTargetKind) -> Self {
        match value {
            OfficeCollaborationPdfReviewTargetKind::Redaction => Self::Redaction,
            OfficeCollaborationPdfReviewTargetKind::PageOperation => Self::PageOperation,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeCollaborationPdfReviewDecision {
    Approve,
    Reject,
}

impl From<OfficeCollaborationPdfReviewDecision> for NativeOfficeCollaborationPdfReviewDecision {
    fn from(value: OfficeCollaborationPdfReviewDecision) -> Self {
        match value {
            OfficeCollaborationPdfReviewDecision::Approve => Self::Approve,
            OfficeCollaborationPdfReviewDecision::Reject => Self::Reject,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeCollaborationPdfAnnotationSource {
    Base,
    Created,
}

impl From<OfficeCollaborationPdfAnnotationSource> for NativeOfficeCollaborationPdfAnnotationSource {
    fn from(value: OfficeCollaborationPdfAnnotationSource) -> Self {
        match value {
            OfficeCollaborationPdfAnnotationSource::Base => Self::Base,
            OfficeCollaborationPdfAnnotationSource::Created => Self::Created,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Deserialize, Serialize, JsonSchema)]
#[serde(
    tag = "type",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub(in crate::office_mcp) enum OfficeCollaborationMutation {
    /// Replace the canonical Markdown source using a minimal Y.Text edit.
    MarkdownReplace { markdown: String },
    /// Splice Markdown using browser-compatible UTF-16 code-unit offsets.
    MarkdownSplice {
        index_utf16: u32,
        delete_utf16: u32,
        insert: String,
    },
    /// Replace an exact number of non-overlapping Document Y.XmlText matches.
    DocumentReplaceText {
        search: String,
        replacement: String,
        expected_matches: u32,
    },
    /// Set the conflict-local Document page-color option.
    DocumentSetPageColor { page_color: String },
    /// Explicitly clear the conflict-local Document page-color option.
    DocumentClearPageColor {},
    /// Set the conflict-local Document track-changes option.
    DocumentSetTrackChanges { track_changes: bool },
    /// Explicitly clear the conflict-local Document track-changes option.
    DocumentClearTrackChanges {},
    /// Insert a plain paragraph beside a stable paragraph in a supported block container.
    DocumentInsertParagraph {
        anchor_paragraph_id: String,
        position: OfficeCollaborationParagraphPosition,
        paragraph_id: String,
        text_id: String,
        text: String,
    },
    /// Delete a plain paragraph from a supported block container with exact guards.
    DocumentDeleteParagraph {
        paragraph_id: String,
        expected_text_id: String,
        expected_text: String,
    },
    /// Create or recursively patch one conflict-local Spreadsheet cell.
    SpreadsheetSetCell {
        sheet_id: String,
        row: u32,
        column: u32,
        expected_cell: Option<JsonValue>,
        next_cell: JsonValue,
    },
    /// Delete one Spreadsheet cell after matching its complete current value.
    SpreadsheetDeleteCell {
        sheet_id: String,
        row: u32,
        column: u32,
        expected_cell: JsonValue,
    },
    /// Create one scene element in a slide, master, or layout with a stable ID.
    PresentationCreateElement {
        container_kind: OfficeCollaborationPresentationContainerKind,
        container_id: String,
        element: JsonValue,
        after_element_id: Option<String>,
    },
    /// Merge top-level scene-element fields using optimistic expectations.
    PresentationUpdateElement {
        container_kind: OfficeCollaborationPresentationContainerKind,
        container_id: String,
        element_id: String,
        expected_element: JsonValue,
        next_element: JsonValue,
    },
    /// Tombstone one scene element after matching its complete current value.
    PresentationDeleteElement {
        container_kind: OfficeCollaborationPresentationContainerKind,
        container_id: String,
        expected_element: JsonValue,
    },
    /// Move one active scene element after its stable observed predecessor.
    PresentationMoveElement {
        container_kind: OfficeCollaborationPresentationContainerKind,
        container_id: String,
        element_id: String,
        expected_after_element_id: Option<String>,
        after_element_id: Option<String>,
    },
    /// Create one supported portable PDF annotation with a stable ID.
    PdfCreateAnnotation {
        annotation_id: String,
        page_index: u32,
        annotation: JsonValue,
    },
    /// Merge mutable annotation leaves using an exact recursive expectation.
    PdfUpdateAnnotation {
        annotation_id: String,
        expected_annotation: JsonValue,
        next_annotation: JsonValue,
    },
    /// Irreversibly tombstone one annotation after matching immutable identity.
    PdfDeleteAnnotation {
        annotation_id: String,
        expected_source: OfficeCollaborationPdfAnnotationSource,
        expected_page_index: u32,
        expected_type: u32,
    },
    /// Set one PDF form value by its stable fully-qualified field name.
    PdfSetFormValue { field_id: String, value: String },
    /// Append one attributable PDF redaction proposal with immutable geometry.
    PdfProposeRedaction {
        proposal_id: String,
        page_index: u32,
        rects: Vec<OfficeCollaborationPdfRect>,
        proposed_at: String,
        reason: Option<String>,
        text: Option<String>,
    },
    /// Append an attributable request to rotate selected source pages.
    PdfProposePageRotation {
        page_operation_id: String,
        page_indices: Vec<u32>,
        degrees: u16,
        proposed_at: String,
    },
    /// Append an attributable request to delete a proper subset of source pages.
    PdfProposePageDeletion {
        page_operation_id: String,
        page_indices: Vec<u32>,
        proposed_at: String,
    },
    /// Append an attributable request containing a complete source-page permutation.
    PdfProposePageReorder {
        page_operation_id: String,
        page_order: Vec<u32>,
        proposed_at: String,
    },
    /// Append the single final review decision for a PDF redaction or page operation.
    PdfDecideReview {
        decision_id: String,
        target_kind: OfficeCollaborationPdfReviewTargetKind,
        target_id: String,
        decision: OfficeCollaborationPdfReviewDecision,
        created_at: String,
    },
}

impl From<OfficeCollaborationMutation> for NativeOfficeCollaborationMutation {
    fn from(value: OfficeCollaborationMutation) -> Self {
        match value {
            OfficeCollaborationMutation::MarkdownReplace { markdown } => {
                Self::MarkdownReplace { markdown }
            }
            OfficeCollaborationMutation::MarkdownSplice {
                index_utf16,
                delete_utf16,
                insert,
            } => Self::MarkdownSplice {
                index_utf16,
                delete_utf16,
                insert,
            },
            OfficeCollaborationMutation::DocumentReplaceText {
                search,
                replacement,
                expected_matches,
            } => Self::DocumentReplaceText {
                search,
                replacement,
                expected_matches,
            },
            OfficeCollaborationMutation::DocumentSetPageColor { page_color } => {
                Self::DocumentSetPageColor { page_color }
            }
            OfficeCollaborationMutation::DocumentClearPageColor { .. } => {
                Self::DocumentClearPageColor {}
            }
            OfficeCollaborationMutation::DocumentSetTrackChanges { track_changes } => {
                Self::DocumentSetTrackChanges { track_changes }
            }
            OfficeCollaborationMutation::DocumentClearTrackChanges { .. } => {
                Self::DocumentClearTrackChanges {}
            }
            OfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id,
                position,
                paragraph_id,
                text_id,
                text,
            } => Self::DocumentInsertParagraph {
                anchor_paragraph_id,
                position: position.into(),
                paragraph_id,
                text_id,
                text,
            },
            OfficeCollaborationMutation::DocumentDeleteParagraph {
                paragraph_id,
                expected_text_id,
                expected_text,
            } => Self::DocumentDeleteParagraph {
                paragraph_id,
                expected_text_id,
                expected_text,
            },
            OfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id,
                row,
                column,
                expected_cell,
                next_cell,
            } => Self::SpreadsheetSetCell {
                sheet_id,
                row,
                column,
                expected_cell,
                next_cell,
            },
            OfficeCollaborationMutation::SpreadsheetDeleteCell {
                sheet_id,
                row,
                column,
                expected_cell,
            } => Self::SpreadsheetDeleteCell {
                sheet_id,
                row,
                column,
                expected_cell,
            },
            OfficeCollaborationMutation::PresentationCreateElement {
                container_kind,
                container_id,
                element,
                after_element_id,
            } => Self::PresentationCreateElement {
                container_kind: container_kind.into(),
                container_id,
                element,
                after_element_id,
            },
            OfficeCollaborationMutation::PresentationUpdateElement {
                container_kind,
                container_id,
                element_id,
                expected_element,
                next_element,
            } => Self::PresentationUpdateElement {
                container_kind: container_kind.into(),
                container_id,
                element_id,
                expected_element,
                next_element,
            },
            OfficeCollaborationMutation::PresentationDeleteElement {
                container_kind,
                container_id,
                expected_element,
            } => Self::PresentationDeleteElement {
                container_kind: container_kind.into(),
                container_id,
                expected_element,
            },
            OfficeCollaborationMutation::PresentationMoveElement {
                container_kind,
                container_id,
                element_id,
                expected_after_element_id,
                after_element_id,
            } => Self::PresentationMoveElement {
                container_kind: container_kind.into(),
                container_id,
                element_id,
                expected_after_element_id,
                after_element_id,
            },
            OfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id,
                page_index,
                annotation,
            } => Self::PdfCreateAnnotation {
                annotation_id,
                page_index,
                annotation,
            },
            OfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id,
                expected_annotation,
                next_annotation,
            } => Self::PdfUpdateAnnotation {
                annotation_id,
                expected_annotation,
                next_annotation,
            },
            OfficeCollaborationMutation::PdfDeleteAnnotation {
                annotation_id,
                expected_source,
                expected_page_index,
                expected_type,
            } => Self::PdfDeleteAnnotation {
                annotation_id,
                expected_source: expected_source.into(),
                expected_page_index,
                expected_type,
            },
            OfficeCollaborationMutation::PdfSetFormValue { field_id, value } => {
                Self::PdfSetFormValue { field_id, value }
            }
            OfficeCollaborationMutation::PdfProposeRedaction {
                proposal_id,
                page_index,
                rects,
                proposed_at,
                reason,
                text,
            } => Self::PdfProposeRedaction {
                proposal_id,
                page_index,
                rects: rects.into_iter().map(Into::into).collect(),
                proposed_at,
                reason,
                text,
            },
            OfficeCollaborationMutation::PdfProposePageRotation {
                page_operation_id,
                page_indices,
                degrees,
                proposed_at,
            } => Self::PdfProposePageRotation {
                page_operation_id,
                page_indices,
                degrees,
                proposed_at,
            },
            OfficeCollaborationMutation::PdfProposePageDeletion {
                page_operation_id,
                page_indices,
                proposed_at,
            } => Self::PdfProposePageDeletion {
                page_operation_id,
                page_indices,
                proposed_at,
            },
            OfficeCollaborationMutation::PdfProposePageReorder {
                page_operation_id,
                page_order,
                proposed_at,
            } => Self::PdfProposePageReorder {
                page_operation_id,
                page_order,
                proposed_at,
            },
            OfficeCollaborationMutation::PdfDecideReview {
                decision_id,
                target_kind,
                target_id,
                decision,
                created_at,
            } => Self::PdfDecideReview {
                decision_id,
                target_kind: target_kind.into(),
                target_id,
                decision: decision.into(),
                created_at,
            },
        }
    }
}
