use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

use super::{NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationMode};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationParagraphPosition {
    Before,
    After,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationPresentationContainerKind {
    Slide,
    Master,
    Layout,
}

impl NativeOfficeCollaborationPresentationContainerKind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Slide => "slide",
            Self::Master => "master",
            Self::Layout => "layout",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeCollaborationPdfRect {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationPdfAnnotationSource {
    Base,
    Created,
}

impl NativeOfficeCollaborationPdfAnnotationSource {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Base => "base",
            Self::Created => "created",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationPdfReviewTargetKind {
    Redaction,
    PageOperation,
}

impl NativeOfficeCollaborationPdfReviewTargetKind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Redaction => "redaction",
            Self::PageOperation => "page-operation",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationPdfReviewDecision {
    Approve,
    Reject,
}

impl NativeOfficeCollaborationPdfReviewDecision {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Approve => "approve",
            Self::Reject => "reject",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum NativeOfficeCollaborationMutation {
    /// Replace the canonical Markdown source while retaining the longest
    /// common prefix and suffix in the underlying Y.Text.
    MarkdownReplace { markdown: String },
    /// Splice the canonical Markdown Y.Text using browser-compatible UTF-16
    /// offsets. Surrogate pairs may not be split.
    MarkdownSplice {
        index_utf16: u32,
        delete_utf16: u32,
        insert: String,
    },
    /// Replace an exact number of non-overlapping text matches inside the
    /// canonical ProseMirror `Y.XmlFragment`. Matches may span formatting
    /// runs inside one `Y.XmlText`, but never cross an embedded object or XML
    /// text-node boundary.
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
    /// Insert one plain paragraph next to a uniquely identified paragraph in
    /// a supported section, list-item, table-cell, or blockquote container.
    /// New Word identities are explicit so independent native replicas encode
    /// the same logical operation.
    DocumentInsertParagraph {
        anchor_paragraph_id: String,
        position: NativeOfficeCollaborationParagraphPosition,
        paragraph_id: String,
        text_id: String,
        text: String,
    },
    /// Delete one plain paragraph from a supported structural container after
    /// verifying its current Word text identity and complete visible text.
    DocumentDeleteParagraph {
        paragraph_id: String,
        expected_text_id: String,
        expected_text: String,
    },
    /// Create or update one Spreadsheet cell through the browser-compatible
    /// field-addressed cell map. `expectedCell` is absent when the caller
    /// observed a blank coordinate. Recursive optimistic guards allow
    /// unrelated concurrent leaf edits to merge while same-leaf conflicts
    /// fail closed.
    SpreadsheetSetCell {
        sheet_id: String,
        row: u32,
        column: u32,
        expected_cell: Option<JsonValue>,
        next_cell: JsonValue,
    },
    /// Delete one Spreadsheet cell after matching its complete current JSON
    /// value. Dense projection dimensions are retained; sparse and empty
    /// sheets remain sparse.
    SpreadsheetDeleteCell {
        sheet_id: String,
        row: u32,
        column: u32,
        expected_cell: JsonValue,
    },
    /// Create one scene element inside a slide, master, or layout. The full
    /// element is fingerprinted in the browser-compatible record-claims root,
    /// preventing different records from converging under one stable ID.
    PresentationCreateElement {
        container_kind: NativeOfficeCollaborationPresentationContainerKind,
        container_id: String,
        element: JsonValue,
        after_element_id: Option<String>,
    },
    /// Update top-level fields of one scene element with optimistic guards.
    /// Independent field edits merge, while stale edits to the same field
    /// fail closed. Element IDs and types are immutable.
    PresentationUpdateElement {
        container_kind: NativeOfficeCollaborationPresentationContainerKind,
        container_id: String,
        element_id: String,
        expected_element: JsonValue,
        next_element: JsonValue,
    },
    /// Tombstone one scene element after matching its complete current JSON
    /// value. Tombstoned IDs remain reserved and cannot be reused.
    PresentationDeleteElement {
        container_kind: NativeOfficeCollaborationPresentationContainerKind,
        container_id: String,
        expected_element: JsonValue,
    },
    /// Create one portable PDF annotation with a caller-owned stable ID.
    /// Native creation always records `source: "created"` and accepts only
    /// annotation types supported by the browser projection.
    PdfCreateAnnotation {
        annotation_id: String,
        page_index: u32,
        annotation: JsonValue,
    },
    /// Update mutable leaves of one portable PDF annotation. The expected
    /// value provides a recursive optimistic guard, so unrelated concurrent
    /// leaf edits merge while conflicting edits fail closed.
    PdfUpdateAnnotation {
        annotation_id: String,
        expected_annotation: JsonValue,
        next_annotation: JsonValue,
    },
    /// Irreversibly tombstone one PDF annotation after matching its immutable
    /// source, page, and annotation-type identity.
    PdfDeleteAnnotation {
        annotation_id: String,
        expected_source: NativeOfficeCollaborationPdfAnnotationSource,
        expected_page_index: u32,
        expected_type: u32,
    },
    /// Set one PDF form value by its stable fully-qualified field name. New
    /// fields are added to the typed presence/fields/order collection; an
    /// existing field changes only its conflict-local value leaf.
    PdfSetFormValue { field_id: String, value: String },
    /// Append one attributable PDF redaction proposal. The replica actor is
    /// recorded as `proposedBy`; callers provide stable identity, geometry,
    /// and a deterministic UTC timestamp.
    PdfProposeRedaction {
        proposal_id: String,
        page_index: u32,
        rects: Vec<NativeOfficeCollaborationPdfRect>,
        proposed_at: String,
        reason: Option<String>,
        text: Option<String>,
    },
    /// Append one attributable request to rotate a non-empty set of source
    /// pages clockwise by 90, 180, or 270 degrees.
    PdfProposePageRotation {
        page_operation_id: String,
        page_indices: Vec<u32>,
        degrees: u16,
        proposed_at: String,
    },
    /// Append one attributable request to delete a non-empty proper subset of
    /// source pages. At least one source page must remain.
    PdfProposePageDeletion {
        page_operation_id: String,
        page_indices: Vec<u32>,
        proposed_at: String,
    },
    /// Append one attributable request to reorder every source page using a
    /// complete zero-based permutation.
    PdfProposePageReorder {
        page_operation_id: String,
        page_order: Vec<u32>,
        proposed_at: String,
    },
    /// Append the single final decision for a redaction or page operation.
    /// The replica actor is recorded as `actorId`.
    PdfDecideReview {
        decision_id: String,
        target_kind: NativeOfficeCollaborationPdfReviewTargetKind,
        target_id: String,
        decision: NativeOfficeCollaborationPdfReviewDecision,
        created_at: String,
    },
}

#[derive(Debug, Clone)]
pub struct NativeOfficeCollaborationMutationRequest {
    pub operation_id: String,
    pub actor_id: String,
    pub mode: NativeOfficeCollaborationMode,
    pub expected_artifact_id: String,
    pub expected_kind: NativeOfficeCollaborationArtifactKind,
    pub mutation: NativeOfficeCollaborationMutation,
    pub if_state_vector: Option<Vec<u8>>,
}
