use std::str::FromStr;

use a3s_boot::{BootError, Result, WebSocketMessage};
use a3s_office::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationOrigin,
    NativeOfficeCollaborationTransportMessage, NativeOfficeCollaborationTransportMessageType,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde::{Deserialize, Serialize};

pub const DOCUMENT_EVENT: &str = "collaboration.document";
pub const AWARENESS_EVENT: &str = "collaboration.awareness";
pub const HELLO_EVENT: &str = "collaboration.hello";
pub const READY_EVENT: &str = "collaboration.ready";
pub const PEER_LEFT_EVENT: &str = "collaboration.peer-left";
pub const ACK_EVENT: &str = "collaboration.ack";
pub const ERROR_EVENT: &str = "collaboration.error";

const MAX_IDENTIFIER_BYTES: usize = 256;
const MAX_YJS_CLIENT_ID: u64 = (1_u64 << 53) - 1;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RoomIdentity {
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
}

impl RoomIdentity {
    pub fn new(
        artifact_id: impl Into<String>,
        artifact_kind: NativeOfficeCollaborationArtifactKind,
        namespace: impl Into<String>,
    ) -> Result<Self> {
        let artifact_id = validated_identifier(artifact_id.into(), "artifact ID")?;
        let namespace = validated_namespace(namespace.into())?;
        Ok(Self {
            artifact_id,
            artifact_kind,
            namespace,
        })
    }

    pub fn key_material(&self) -> String {
        format!(
            "{}\0{}\0{}",
            self.namespace,
            self.artifact_kind.as_str(),
            self.artifact_id
        )
    }

    pub fn matches_wire(
        &self,
        artifact_id: &str,
        artifact_kind: NativeOfficeCollaborationArtifactKind,
        namespace: &str,
    ) -> bool {
        self.artifact_id == artifact_id
            && self.artifact_kind == artifact_kind
            && self.namespace == namespace
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TicketRequest {
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub actor_id: String,
    pub actor_name: String,
    #[serde(default = "default_actor_kind")]
    pub actor_kind: NativeOfficeCollaborationActorKind,
    pub mode: NativeOfficeCollaborationMode,
}

impl TicketRequest {
    pub fn validate(&self) -> Result<()> {
        validated_identifier(self.artifact_id.clone(), "artifact ID")?;
        validated_identifier(self.actor_id.clone(), "actor ID")?;
        validated_actor_name(&self.actor_name)?;
        if self.actor_kind == NativeOfficeCollaborationActorKind::System {
            return Err(BootError::BadRequest(
                "system actors cannot receive browser collaboration tickets".to_string(),
            ));
        }
        Ok(())
    }
}

fn default_actor_kind() -> NativeOfficeCollaborationActorKind {
    NativeOfficeCollaborationActorKind::Human
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TicketResponse {
    pub ticket: String,
    pub web_socket_url: String,
    pub expires_at: u64,
    pub protocol: &'static str,
    pub version: u32,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TicketClaims {
    pub ticket_version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
    pub actor_id: String,
    pub actor_name: String,
    pub actor_kind: NativeOfficeCollaborationActorKind,
    pub mode: NativeOfficeCollaborationMode,
    pub issued_at: u64,
    pub expires_at: u64,
}

impl TicketClaims {
    pub fn room_identity(&self) -> Result<RoomIdentity> {
        RoomIdentity::new(
            self.artifact_id.clone(),
            self.artifact_kind,
            self.namespace.clone(),
        )
    }

    pub fn can_publish_document_updates(&self) -> bool {
        matches!(
            self.mode,
            NativeOfficeCollaborationMode::Edit | NativeOfficeCollaborationMode::Comment
        )
    }
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct HelloMessage {
    pub protocol: String,
    pub version: u32,
    pub sender_client_id: u64,
}

impl HelloMessage {
    pub fn from_message(message: WebSocketMessage) -> Result<Self> {
        let hello: Self = message.data_as()?;
        if hello.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
            || hello.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
        {
            return Err(BootError::BadRequest(
                "unsupported Office collaboration protocol or version".to_string(),
            ));
        }
        validate_client_id(hello.sender_client_id)?;
        Ok(hello)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DocumentWireMessage {
    pub protocol: String,
    pub version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
    pub sender_client_id: u64,
    #[serde(rename = "type")]
    pub message_type: NativeOfficeCollaborationTransportMessageType,
    pub payload_base64: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub origin: Option<NativeOfficeCollaborationOrigin>,
}

impl DocumentWireMessage {
    pub fn from_message(message: WebSocketMessage) -> Result<Self> {
        message.data_as()
    }

    pub fn decode(
        &self,
        room: &RoomIdentity,
        maximum_payload_bytes: usize,
    ) -> Result<NativeOfficeCollaborationTransportMessage> {
        self.validate_identity(room)?;
        validate_client_id(self.sender_client_id)?;
        let payload = decode_bounded_base64(&self.payload_base64, maximum_payload_bytes)?;
        Ok(NativeOfficeCollaborationTransportMessage {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_string(),
            version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
            artifact_id: room.artifact_id.clone(),
            artifact_kind: room.artifact_kind,
            namespace: room.namespace.clone(),
            sender_client_id: self.sender_client_id,
            message_type: self.message_type,
            payload,
            // Ticket claims, not client JSON, are the authorization and audit source.
            origin: None,
        })
    }

    fn validate_identity(&self, room: &RoomIdentity) -> Result<()> {
        if self.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
            || self.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
        {
            return Err(BootError::BadRequest(
                "unsupported Office collaboration protocol or version".to_string(),
            ));
        }
        if !room.matches_wire(&self.artifact_id, self.artifact_kind, &self.namespace) {
            return Err(BootError::Forbidden(
                "the collaboration message does not belong to the authenticated room".to_string(),
            ));
        }
        Ok(())
    }
}

impl From<&NativeOfficeCollaborationTransportMessage> for DocumentWireMessage {
    fn from(message: &NativeOfficeCollaborationTransportMessage) -> Self {
        Self {
            protocol: message.protocol.clone(),
            version: message.version,
            artifact_id: message.artifact_id.clone(),
            artifact_kind: message.artifact_kind,
            namespace: message.namespace.clone(),
            sender_client_id: message.sender_client_id,
            message_type: message.message_type,
            payload_base64: STANDARD.encode(&message.payload),
            origin: message.origin.clone(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct AwarenessWireMessage {
    pub protocol: String,
    pub version: u32,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub namespace: String,
    pub sender_client_id: u64,
    pub payload_base64: String,
}

impl AwarenessWireMessage {
    pub fn from_message(message: WebSocketMessage) -> Result<Self> {
        message.data_as()
    }

    pub fn decode(&self, room: &RoomIdentity, maximum_payload_bytes: usize) -> Result<Vec<u8>> {
        if self.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
            || self.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
            || !room.matches_wire(&self.artifact_id, self.artifact_kind, &self.namespace)
        {
            return Err(BootError::Forbidden(
                "the awareness message does not belong to the authenticated room".to_string(),
            ));
        }
        validate_client_id(self.sender_client_id)?;
        decode_bounded_base64(&self.payload_base64, maximum_payload_bytes)
    }
}

pub fn websocket_document_message(
    message: &NativeOfficeCollaborationTransportMessage,
) -> Result<WebSocketMessage> {
    WebSocketMessage::json(DOCUMENT_EVENT, &DocumentWireMessage::from(message))
}

pub fn trusted_origin(
    claims: &TicketClaims,
    operation_id: String,
) -> NativeOfficeCollaborationOrigin {
    let kind = match claims.actor_kind {
        NativeOfficeCollaborationActorKind::Human => {
            a3s_office::NativeOfficeCollaborationOriginKind::Editor
        }
        NativeOfficeCollaborationActorKind::Agent => {
            a3s_office::NativeOfficeCollaborationOriginKind::Agent
        }
        NativeOfficeCollaborationActorKind::System => {
            a3s_office::NativeOfficeCollaborationOriginKind::System
        }
    };
    NativeOfficeCollaborationOrigin {
        protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_string(),
        kind,
        actor_id: Some(claims.actor_id.clone()),
        operation_id: Some(operation_id),
    }
}

pub fn parse_artifact_kind(value: &str) -> Result<NativeOfficeCollaborationArtifactKind> {
    NativeOfficeCollaborationArtifactKind::from_str(value)
        .map_err(|error| BootError::BadRequest(error.message))
}

pub fn validate_client_id(client_id: u64) -> Result<()> {
    if client_id <= MAX_YJS_CLIENT_ID {
        Ok(())
    } else {
        Err(BootError::BadRequest(
            "Yjs client IDs must fit in JavaScript's 53-bit safe integer range".to_string(),
        ))
    }
}

fn validated_identifier(value: String, label: &str) -> Result<String> {
    let value = value.trim().to_string();
    if value.is_empty()
        || value.as_bytes().len() > MAX_IDENTIFIER_BYTES
        || !value
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || "._-".contains(character))
    {
        return Err(BootError::BadRequest(format!(
            "{label} must contain only ASCII letters, numbers, dot, underscore, or hyphen"
        )));
    }
    Ok(value)
}

fn validated_namespace(value: String) -> Result<String> {
    let value = validated_identifier(value, "namespace")?;
    if value.starts_with('.') || value.ends_with('.') {
        return Err(BootError::BadRequest(
            "namespace cannot start or end with a dot".to_string(),
        ));
    }
    Ok(value)
}

fn validated_actor_name(value: &str) -> Result<()> {
    if !value.is_empty() && value == value.trim() && value.chars().count() <= 256 {
        return Ok(());
    }
    Err(BootError::BadRequest(
        "actor name must contain 1 to 256 non-padded characters".to_string(),
    ))
}

fn decode_bounded_base64(value: &str, maximum_bytes: usize) -> Result<Vec<u8>> {
    let maximum_encoded = maximum_bytes.saturating_add(2) / 3 * 4;
    if value.len() > maximum_encoded.saturating_add(4) {
        return Err(BootError::PayloadTooLarge(format!(
            "encoded collaboration payload exceeds {maximum_bytes} bytes"
        )));
    }
    let payload = STANDARD
        .decode(value)
        .map_err(|error| BootError::BadRequest(format!("invalid base64 payload: {error}")))?;
    if payload.len() > maximum_bytes {
        return Err(BootError::PayloadTooLarge(format!(
            "collaboration payload exceeds {maximum_bytes} bytes"
        )));
    }
    Ok(payload)
}
