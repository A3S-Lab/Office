use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

use super::{
    NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationDocumentChangeKind,
    NativeOfficeCollaborationDocumentSuggestionDecision,
    NativeOfficeCollaborationDocumentSuggestionKind,
    NativeOfficeCollaborationPresentationContainerKind,
};

pub const NATIVE_OFFICE_COLLABORATION_PROJECTION_SCHEMA: &str =
    "a3s.office.collaboration.projection";
pub const NATIVE_OFFICE_COLLABORATION_PROJECTION_VERSION: u32 = 4;

/// A bounded, Office-owned view of the current collaborative document.
///
/// The Yjs update log remains canonical. This projection gives native agents
/// enough current content and stable identity to choose typed mutations
/// without interpreting private browser schema inside a product host.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationProjection {
    pub schema: String,
    pub version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub sequence: u64,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
    pub content: NativeOfficeCollaborationProjectedContent,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase"
)]
pub enum NativeOfficeCollaborationProjectedContent {
    /// Canonical Markdown source plus the line slices an agent patches.
    /// `source` is context for `expectedMarkdown`. Agents splice one slice;
    /// they do not write `source` back as a whole-document replacement.
    Markdown {
        source: String,
        slices: Vec<NativeOfficeCollaborationMarkdownSlice>,
    },
    /// A structure-aware, agent-readable view of the Document collaboration
    /// root. Paragraph records retain the Word identities and UTF-16 range
    /// required by typed mutations; `plainText` is subordinate search/context
    /// text and is not a whole-document writeback payload.
    Document {
        plain_text: String,
        paragraphs: Vec<NativeOfficeCollaborationDocumentParagraph>,
        comments: Vec<NativeOfficeCollaborationDocumentComment>,
        suggestions: Vec<NativeOfficeCollaborationDocumentSuggestion>,
        change_decisions: Vec<NativeOfficeCollaborationDocumentChangeDecision>,
        page_color: Option<String>,
        track_changes: Option<bool>,
    },
    /// Sheet and cell addresses from the current replica. Each cell record is
    /// the guarded value for `spreadsheet-set-cell`, not a workbook blob.
    Spreadsheet {
        sheets: Vec<NativeOfficeCollaborationSpreadsheetSheet>,
    },
    /// Container and scene-element addresses from the current replica.
    Presentation {
        containers: Vec<NativeOfficeCollaborationPresentationContainer>,
    },
    /// Page, annotation, and form-field addresses. This is not a PDF byte payload.
    Pdf {
        page_count: u32,
        annotations: Vec<NativeOfficeCollaborationPdfAnnotationAddress>,
        form_fields: Vec<NativeOfficeCollaborationPdfFormField>,
    },
}

/// One non-overlapping Markdown line slice. Offsets are UTF-16 code units in
/// the canonical source and do not include the separating newline.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationMarkdownSlice {
    pub index: u32,
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub text: String,
}

/// One sheet and the populated cell coordinates an agent can patch.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSpreadsheetSheet {
    pub sheet_id: String,
    pub cells: Vec<NativeOfficeCollaborationSpreadsheetCellAddress>,
}

/// A stable spreadsheet address: sheet identity is on the parent, and `cell`
/// is the current JSON value guarded by `spreadsheet-set-cell`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSpreadsheetCellAddress {
    pub row: u32,
    pub column: u32,
    pub cell: JsonValue,
}

/// One slide, master, or layout and its active scene elements.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresentationContainer {
    pub container_kind: NativeOfficeCollaborationPresentationContainerKind,
    pub container_id: String,
    pub elements: Vec<NativeOfficeCollaborationPresentationElementAddress>,
}

/// A stable presentation address. `element` is the current guarded object.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresentationElementAddress {
    pub element_id: String,
    pub element: JsonValue,
}

/// A stable PDF annotation address. `annotation` is the portable annotation
/// object, not the PDF file bytes.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPdfAnnotationAddress {
    pub annotation_id: String,
    pub page_index: u32,
    pub annotation_type: u32,
    pub source: String,
    pub annotation: JsonValue,
}

/// A stable PDF form-field address.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPdfFormField {
    pub field_id: String,
    pub value: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentParagraph {
    /// One-based traversal order in the converged Office document.
    pub ordinal: u32,
    /// Office/ProseMirror node type such as `paragraph` or `heading`.
    pub node_type: String,
    /// Stable Word paragraph identity when the shared node owns one.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_id: Option<String>,
    /// Current optimistic text identity paired with `paragraphId`.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_id: Option<String>,
    /// Outermost-to-innermost XML container types for structural context.
    pub container_path: Vec<String>,
    /// Current visible string content, excluding embedded objects.
    pub text: String,
    /// UTF-16 start of `text` inside this paragraph. Whole-paragraph addresses
    /// use `0`.
    pub start_utf16: u32,
    /// UTF-16 end of `text` inside this paragraph, exclusive.
    pub end_utf16: u32,
    /// True when stable identity plus the current node shape permit the
    /// fail-closed `document-replace-paragraph` mutation.
    pub replaceable: bool,
    pub has_inline_objects: bool,
    pub has_review_marks: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentComment {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actor_id: Option<String>,
    pub author: String,
    pub date: String,
    pub text: String,
    pub resolved: bool,
    pub replies: Vec<NativeOfficeCollaborationDocumentCommentReply>,
    pub anchors: Vec<NativeOfficeCollaborationDocumentCommentAnchor>,
    pub detached: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentCommentReply {
    pub id: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actor_id: Option<String>,
    pub author: String,
    pub date: String,
    pub text: String,
}

/// One exact paragraph-local span carrying a Document comment mark. Offsets
/// use browser-compatible UTF-16 code units and `text` is the current anchor
/// text for optimistic agent decisions.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentCommentAnchor {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_id: Option<String>,
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentSuggestion {
    pub id: String,
    pub kind: NativeOfficeCollaborationDocumentSuggestionKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actor_id: Option<String>,
    pub author: String,
    pub created_at: String,
    pub text: String,
    pub placements: Vec<NativeOfficeCollaborationDocumentSuggestionPlacement>,
}

/// One exact Y.XmlText span carrying a tracked-change mark. Offsets use the
/// current paragraph text in browser-compatible UTF-16 code units.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentSuggestionPlacement {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub paragraph_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text_id: Option<String>,
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationDocumentChangeDecision {
    pub id: String,
    pub change_id: String,
    pub change_kind: NativeOfficeCollaborationDocumentChangeKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suggested_by_actor_id: Option<String>,
    pub suggested_by: String,
    pub suggested_at: String,
    pub text: String,
    pub decision: NativeOfficeCollaborationDocumentSuggestionDecision,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub decided_by_actor_id: Option<String>,
    pub decided_by: String,
    pub decided_at: String,
}
