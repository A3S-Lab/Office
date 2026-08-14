use serde::{Deserialize, Serialize};

use super::{NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationMode};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationParagraphPosition {
    Before,
    After,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
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
