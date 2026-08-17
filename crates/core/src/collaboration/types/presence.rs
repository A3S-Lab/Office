use serde::{Deserialize, Serialize};

use super::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationMode,
};

pub const MAX_NATIVE_OFFICE_COLLABORATION_AWARENESS_BYTES: usize = 256 * 1024;
pub const MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_RANGES: usize = 64;
pub const MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_ELEMENT_IDS: usize = 128;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationPresenceActivity {
    Active,
    Idle,
    Away,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationMarkdownSurface {
    Source,
    Visual,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSpreadsheetPresenceCell {
    pub row: u64,
    pub column: u64,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSpreadsheetPresenceRange {
    pub start_row: u64,
    pub start_column: u64,
    pub end_row: u64,
    pub end_column: u64,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase"
)]
pub enum NativeOfficeCollaborationPresenceLocation {
    Document {
        anchor: u64,
        head: u64,
    },
    Markdown {
        anchor: u64,
        head: u64,
        #[serde(skip_serializing_if = "Option::is_none")]
        surface: Option<NativeOfficeCollaborationMarkdownSurface>,
    },
    Spreadsheet {
        sheet_id: String,
        ranges: Vec<NativeOfficeCollaborationSpreadsheetPresenceRange>,
        #[serde(skip_serializing_if = "Option::is_none")]
        active_cell: Option<NativeOfficeCollaborationSpreadsheetPresenceCell>,
    },
    Presentation {
        slide_id: String,
        element_ids: Vec<String>,
    },
    Pdf {
        page_index: u64,
        #[serde(skip_serializing_if = "Option::is_none")]
        annotation_id: Option<String>,
    },
}

impl NativeOfficeCollaborationPresenceLocation {
    pub const fn artifact_kind(&self) -> NativeOfficeCollaborationArtifactKind {
        match self {
            Self::Document { .. } => NativeOfficeCollaborationArtifactKind::Document,
            Self::Markdown { .. } => NativeOfficeCollaborationArtifactKind::Markdown,
            Self::Spreadsheet { .. } => NativeOfficeCollaborationArtifactKind::Spreadsheet,
            Self::Presentation { .. } => NativeOfficeCollaborationArtifactKind::Presentation,
            Self::Pdf { .. } => NativeOfficeCollaborationArtifactKind::Pdf,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresenceProfile {
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub avatar_url: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresenceActor {
    pub id: String,
    pub name: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub avatar_url: Option<String>,
    pub kind: NativeOfficeCollaborationActorKind,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresenceState {
    pub protocol: String,
    pub version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
    pub presence_id: String,
    pub actor: NativeOfficeCollaborationPresenceActor,
    pub mode: NativeOfficeCollaborationMode,
    pub activity: NativeOfficeCollaborationPresenceActivity,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<NativeOfficeCollaborationPresenceLocation>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativeOfficeCollaborationPresenceUpdate {
    pub activity: NativeOfficeCollaborationPresenceActivity,
    pub location: Option<NativeOfficeCollaborationPresenceLocation>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationParticipant {
    pub client_id: u64,
    pub local: bool,
    #[serde(flatten)]
    pub state: NativeOfficeCollaborationPresenceState,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresenceSnapshot {
    pub local_client_id: u64,
    pub participants: Vec<NativeOfficeCollaborationParticipant>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationAwarenessMessage {
    pub protocol: String,
    pub version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
    pub sender_client_id: u64,
    pub payload: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationPresenceReceiveResult {
    pub ignored: bool,
    pub changed: bool,
    pub snapshot: NativeOfficeCollaborationPresenceSnapshot,
}
