use a3s_office::{NativeOfficeCollaborationMutation, NativeOfficeCollaborationParagraphPosition};
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};

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

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize, JsonSchema)]
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
        }
    }
}
