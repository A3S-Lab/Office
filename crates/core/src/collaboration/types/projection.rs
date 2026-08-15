use serde::{Deserialize, Serialize};

use super::NativeOfficeCollaborationArtifactKind;

pub const NATIVE_OFFICE_COLLABORATION_PROJECTION_SCHEMA: &str =
    "a3s.office.collaboration.projection";
pub const NATIVE_OFFICE_COLLABORATION_PROJECTION_VERSION: u32 = 1;

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
