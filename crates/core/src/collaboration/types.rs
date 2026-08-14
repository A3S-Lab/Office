use std::path::PathBuf;
use std::str::FromStr;

use a3s_use_core::UseError;
use serde::{Deserialize, Serialize};

use super::collaboration_error;

pub const NATIVE_OFFICE_COLLABORATION_PROTOCOL: &str = "a3s.office.collaboration";
pub const NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION: u32 = 1;
pub const NATIVE_OFFICE_COLLABORATION_NAMESPACE: &str = "a3s.office";
pub const NATIVE_OFFICE_COLLABORATION_STORE_FORMAT: &str = "a3s.office.collaboration-replica";
pub const NATIVE_OFFICE_COLLABORATION_STORE_SCHEMA_VERSION: u32 = 1;
pub const MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES: usize = 64 * 1024 * 1024;
pub const MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES: usize = 1024 * 1024;
pub const MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationArtifactKind {
    Document,
    Markdown,
    Spreadsheet,
    Presentation,
    Pdf,
}

impl NativeOfficeCollaborationArtifactKind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Document => "document",
            Self::Markdown => "markdown",
            Self::Spreadsheet => "spreadsheet",
            Self::Presentation => "presentation",
            Self::Pdf => "pdf",
        }
    }
}

impl FromStr for NativeOfficeCollaborationArtifactKind {
    type Err = UseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "document" => Ok(Self::Document),
            "markdown" => Ok(Self::Markdown),
            "spreadsheet" => Ok(Self::Spreadsheet),
            "presentation" => Ok(Self::Presentation),
            "pdf" => Ok(Self::Pdf),
            _ => Err(collaboration_error(
                "office.collaboration.kind_invalid",
                format!("The collaboration artifact kind '{value}' is invalid."),
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationMode {
    View,
    Comment,
    Suggest,
    Edit,
}

impl NativeOfficeCollaborationMode {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::View => "view",
            Self::Comment => "comment",
            Self::Suggest => "suggest",
            Self::Edit => "edit",
        }
    }
}

impl FromStr for NativeOfficeCollaborationMode {
    type Err = UseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "view" => Ok(Self::View),
            "comment" => Ok(Self::Comment),
            "suggest" => Ok(Self::Suggest),
            "edit" => Ok(Self::Edit),
            _ => Err(collaboration_error(
                "office.collaboration.mode_invalid",
                format!("The collaboration mode '{value}' is invalid."),
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationActorKind {
    Human,
    Agent,
    System,
}

impl NativeOfficeCollaborationActorKind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Human => "human",
            Self::Agent => "agent",
            Self::System => "system",
        }
    }
}

impl FromStr for NativeOfficeCollaborationActorKind {
    type Err = UseError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "human" => Ok(Self::Human),
            "agent" => Ok(Self::Agent),
            "system" => Ok(Self::System),
            _ => Err(collaboration_error(
                "office.collaboration.actor_invalid",
                format!("The collaboration actor kind '{value}' is invalid."),
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationOperationKind {
    Create,
    Join,
    Synchronize,
    Mutate,
    Checkpoint,
    Leave,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationOriginKind {
    Bootstrap,
    Editor,
    Agent,
    Import,
    System,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationOrigin {
    pub protocol: String,
    pub kind: NativeOfficeCollaborationOriginKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actor_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub operation_id: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum NativeOfficeCollaborationTransportMessageType {
    #[serde(rename = "sync-step-1")]
    SyncStep1,
    #[serde(rename = "sync-step-2")]
    SyncStep2,
    #[serde(rename = "update")]
    Update,
}

/// A native representation of the browser host-channel envelope. Payloads are
/// raw Yjs v1 state vectors or updates; the live session translates them
/// through the standard y-sync framing implemented by the durable store.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationTransportMessage {
    pub protocol: String,
    pub version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
    pub sender_client_id: u64,
    #[serde(rename = "type")]
    pub message_type: NativeOfficeCollaborationTransportMessageType,
    pub payload: Vec<u8>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub origin: Option<NativeOfficeCollaborationOrigin>,
}

#[derive(Debug, Clone)]
pub struct NativeOfficeCollaborationTransportReceiveRequest {
    pub message: NativeOfficeCollaborationTransportMessage,
    /// Stable host delivery identity. Required for SyncStep2 and Update.
    pub operation_id: Option<String>,
    pub if_state_vector: Option<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationTransportReceiveResult {
    pub kind: NativeOfficeCollaborationTransportMessageType,
    pub ignored: bool,
    pub response: Option<NativeOfficeCollaborationTransportMessage>,
    pub apply: Option<NativeOfficeCollaborationApplyResult>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationTransportPollResult {
    pub starting_sequence: u64,
    pub cursor_sequence: u64,
    pub current_sequence: u64,
    pub has_more: bool,
    /// True when compaction made incremental replay unsafe and `messages`
    /// contains a complete Update followed by a fresh SyncStep1 handshake.
    pub resynchronized: bool,
    pub messages: Vec<NativeOfficeCollaborationTransportMessage>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationManifest {
    pub format: String,
    pub schema_version: u32,
    pub protocol: String,
    pub protocol_version: u32,
    pub namespace: String,
    pub artifact_id: String,
    pub kind: NativeOfficeCollaborationArtifactKind,
    pub actor_id: String,
    pub actor_kind: NativeOfficeCollaborationActorKind,
    pub mode: NativeOfficeCollaborationMode,
    pub client_id: u64,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationMetadata {
    pub protocol: String,
    pub version: u32,
    pub artifact_id: String,
    pub kind: NativeOfficeCollaborationArtifactKind,
    pub initialized: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationStateVectorEntry {
    pub client_id: u64,
    pub clock: u32,
}

#[derive(Debug, Clone)]
pub struct NativeOfficeCollaborationCreateRequest {
    pub store: PathBuf,
    pub artifact_id: String,
    pub kind: NativeOfficeCollaborationArtifactKind,
    pub actor_id: String,
    pub actor_kind: NativeOfficeCollaborationActorKind,
    pub mode: NativeOfficeCollaborationMode,
    pub operation_id: String,
    pub namespace: Option<String>,
    pub client_id: Option<u64>,
    pub initial_update: Option<Vec<u8>>,
}

#[derive(Debug, Clone)]
pub struct NativeOfficeCollaborationApplyRequest {
    pub operation_id: String,
    pub actor_id: String,
    pub mode: NativeOfficeCollaborationMode,
    pub expected_artifact_id: String,
    pub expected_kind: NativeOfficeCollaborationArtifactKind,
    pub update: Vec<u8>,
    pub if_state_vector: Option<Vec<u8>>,
    /// Optional authenticated source attribution supplied by a host
    /// transport. This is audit metadata, not an authorization token.
    pub origin: Option<NativeOfficeCollaborationOrigin>,
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

#[derive(Debug, Clone)]
pub struct NativeOfficeCollaborationCheckpointRequest {
    pub operation_id: String,
    pub actor_id: String,
    pub mode: NativeOfficeCollaborationMode,
    pub expected_artifact_id: String,
    pub expected_kind: NativeOfficeCollaborationArtifactKind,
    pub if_state_vector: Option<Vec<u8>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NativeOfficeCollaborationEventsRequest {
    /// `None` starts at the current durable sequence and observes new events.
    pub after_sequence: Option<u64>,
    pub limit: usize,
}

impl Default for NativeOfficeCollaborationEventsRequest {
    fn default() -> Self {
        Self {
            after_sequence: None,
            limit: MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationApplyResult {
    pub operation_id: String,
    pub duplicate: bool,
    pub state_changed: bool,
    pub sequence: Option<u64>,
    pub update_sha256: String,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
    pub checkpointed: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationCheckpointResult {
    pub operation_id: String,
    pub duplicate: bool,
    pub sequence: u64,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
    pub compacted_updates: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSyncResult {
    pub update: Vec<u8>,
    pub update_sha256: String,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSyncStepResult {
    pub message: Vec<u8>,
    pub message_sha256: String,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeCollaborationSyncMessageKind {
    SyncStep1,
    SyncStep2,
    Update,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationSyncMessageResult {
    pub kind: NativeOfficeCollaborationSyncMessageKind,
    pub response: Option<Vec<u8>>,
    pub response_sha256: Option<String>,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
    pub apply: Option<NativeOfficeCollaborationApplyResult>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationUpdateEvent {
    pub sequence: u64,
    pub operation_id: String,
    pub operation_kind: NativeOfficeCollaborationOperationKind,
    pub actor_id: String,
    pub actor_kind: NativeOfficeCollaborationActorKind,
    pub mode: NativeOfficeCollaborationMode,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub payload_sha256: String,
    pub update: Vec<u8>,
    pub update_bytes: u64,
    pub update_sha256: String,
    pub before_state_vector_sha256: String,
    pub after_state_vector_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub origin: Option<NativeOfficeCollaborationOrigin>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationResetEvent {
    pub sequence: u64,
    pub update: Vec<u8>,
    pub update_bytes: u64,
    pub update_sha256: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationEventBatch {
    pub starting_sequence: u64,
    pub cursor_sequence: u64,
    pub checkpoint_sequence: u64,
    pub current_sequence: u64,
    pub has_more: bool,
    pub current_state_vector: Vec<u8>,
    pub current_state_vector_sha256: String,
    pub reset: Option<NativeOfficeCollaborationResetEvent>,
    pub updates: Vec<NativeOfficeCollaborationUpdateEvent>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeCollaborationInspection {
    pub store: PathBuf,
    pub manifest: NativeOfficeCollaborationManifest,
    pub metadata: Option<NativeOfficeCollaborationMetadata>,
    pub metadata_fields_present: usize,
    pub metadata_complete: bool,
    pub bootstrap_initializer_count: Option<u32>,
    pub bootstrap_valid: Option<bool>,
    pub root_names: Vec<String>,
    pub state_vector: Vec<u8>,
    pub state_vector_sha256: String,
    pub state_vector_entries: Vec<NativeOfficeCollaborationStateVectorEntry>,
    pub document_state_sha256: String,
    pub pending_updates: bool,
    pub checkpoint_sequence: u64,
    pub current_sequence: u64,
    pub update_count: usize,
    pub update_bytes: u64,
    pub operation_count: usize,
}
