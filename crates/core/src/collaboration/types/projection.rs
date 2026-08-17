use serde::{Deserialize, Serialize};

use super::{
    NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationDocumentSuggestionDecision,
    NativeOfficeCollaborationDocumentSuggestionKind,
};

pub const NATIVE_OFFICE_COLLABORATION_PROJECTION_SCHEMA: &str =
    "a3s.office.collaboration.projection";
pub const NATIVE_OFFICE_COLLABORATION_PROJECTION_VERSION: u32 = 3;

/// A bounded, Office-owned view of the current collaborative document.
///
/// The Yjs update log remains canonical. This projection gives native agents
/// enough current content and stable identity to choose typed mutations
/// without interpreting private browser schema inside a product host.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
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

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase"
)]
pub enum NativeOfficeCollaborationProjectedContent {
    /// The exact canonical Markdown `Y.Text` source.
    Markdown { source: String },
    /// A structure-aware, agent-readable view of the Document collaboration
    /// root. Paragraph records retain the Word identities required by typed
    /// mutations; `plainText` is subordinate search/context text.
    Document {
        plain_text: String,
        paragraphs: Vec<NativeOfficeCollaborationDocumentParagraph>,
        comments: Vec<NativeOfficeCollaborationDocumentComment>,
        suggestions: Vec<NativeOfficeCollaborationDocumentSuggestion>,
        change_decisions: Vec<NativeOfficeCollaborationDocumentChangeDecision>,
        page_color: Option<String>,
        track_changes: Option<bool>,
    },
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
    pub change_kind: NativeOfficeCollaborationDocumentSuggestionKind,
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
