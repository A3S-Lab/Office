use std::collections::BTreeSet;

use a3s_use_core::UseResult;
use yrs::encoding::read::Cursor;
use yrs::sync::MessageReader;
use yrs::sync::{Message, SyncMessage};
use yrs::updates::decoder::{Decode, DecoderV1};
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};

use super::validation::{decode_state_vector, normalized_identifier, validate_update_size};
use super::{
    collaboration_error, NativeOfficeCollaborationActorKind, NativeOfficeCollaborationApplyRequest,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationOperationKind,
    NativeOfficeCollaborationOrigin, NativeOfficeCollaborationOriginKind,
    NativeOfficeCollaborationStore, NativeOfficeCollaborationTransportMessage,
    NativeOfficeCollaborationTransportMessageType, NativeOfficeCollaborationTransportPollResult,
    NativeOfficeCollaborationTransportReceiveRequest,
    NativeOfficeCollaborationTransportReceiveResult, NativeOfficeCollaborationUpdateEvent,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};

const MAX_YJS_CLIENT_ID: u64 = (1_u64 << 53) - 1;

/// A host-driven live transport state machine around one durable native
/// replica. The host injects connectivity and delivers browser host-channel
/// envelopes; this type never opens a socket, chooses a room, or authorizes a
/// participant.
#[derive(Debug)]
pub struct NativeOfficeCollaborationTransportSession {
    store: NativeOfficeCollaborationStore,
    manifest: NativeOfficeCollaborationManifest,
    cursor_sequence: u64,
    suppressed_sequences: BTreeSet<u64>,
}

impl NativeOfficeCollaborationTransportSession {
    pub fn attach(store: NativeOfficeCollaborationStore) -> UseResult<Self> {
        let inspection = store.inspect()?;
        Ok(Self {
            store,
            manifest: inspection.manifest,
            cursor_sequence: inspection.current_sequence,
            suppressed_sequences: BTreeSet::new(),
        })
    }

    pub fn open(path: impl AsRef<std::path::Path>) -> UseResult<Self> {
        Self::attach(NativeOfficeCollaborationStore::open(path)?)
    }

    pub fn manifest(&self) -> &NativeOfficeCollaborationManifest {
        &self.manifest
    }

    pub fn cursor_sequence(&self) -> u64 {
        self.cursor_sequence
    }

    /// Produce a fresh browser-compatible SyncStep1 envelope. Call this after
    /// each authenticated transport connect or reconnect.
    pub fn synchronize(&self) -> UseResult<NativeOfficeCollaborationTransportMessage> {
        let step = self.store.sync_step1()?;
        Ok(self.message(
            NativeOfficeCollaborationTransportMessageType::SyncStep1,
            step.state_vector,
            None,
        ))
    }

    /// Handle one browser host-channel envelope. Mutating messages require a
    /// stable host delivery operation ID so redelivery remains idempotent.
    pub fn receive(
        &mut self,
        request: NativeOfficeCollaborationTransportReceiveRequest,
    ) -> UseResult<NativeOfficeCollaborationTransportReceiveResult> {
        let message = self.validate_message(request.message)?;
        let kind = message.message_type;
        if message.sender_client_id == self.manifest.client_id {
            return Ok(NativeOfficeCollaborationTransportReceiveResult {
                kind,
                ignored: true,
                response: None,
                apply: None,
            });
        }

        if kind == NativeOfficeCollaborationTransportMessageType::SyncStep1 {
            if request.operation_id.is_some() || request.if_state_vector.is_some() {
                return Err(collaboration_error(
                    "office.collaboration.transport_message_invalid",
                    "SyncStep1 is read-only and must not include mutation identity or a state-vector precondition.",
                ));
            }
            let encoded = encode_y_sync_message(kind, &message.payload)?;
            let handled = self.store.handle_sync_message(&encoded, None)?;
            let response = handled.response.ok_or_else(|| {
                collaboration_error(
                    "office.collaboration.transport_message_invalid",
                    "SyncStep1 did not produce the required SyncStep2 response.",
                )
            })?;
            let (response_kind, payload) = decode_one_y_sync_message(&response)?;
            if response_kind != NativeOfficeCollaborationTransportMessageType::SyncStep2 {
                return Err(collaboration_error(
                    "office.collaboration.transport_message_invalid",
                    "SyncStep1 produced an unexpected y-sync response.",
                ));
            }
            return Ok(NativeOfficeCollaborationTransportReceiveResult {
                kind,
                ignored: false,
                response: Some(self.message(response_kind, payload, None)),
                apply: None,
            });
        }

        let operation_id = request.operation_id.ok_or_else(|| {
            collaboration_error(
                "office.collaboration.sync_identity_required",
                "Applying a transport SyncStep2 or Update requires a stable host delivery operation ID.",
            )
        })?;
        let encoded = encode_y_sync_message(kind, &message.payload)?;
        let handled = self.store.handle_sync_message(
            &encoded,
            Some(NativeOfficeCollaborationApplyRequest {
                operation_id,
                actor_id: self.manifest.actor_id.clone(),
                mode: self.manifest.mode,
                expected_artifact_id: self.manifest.artifact_id.clone(),
                expected_kind: self.manifest.kind,
                update: Vec::new(),
                if_state_vector: request.if_state_vector,
            }),
        )?;
        let apply = handled.apply.ok_or_else(|| {
            collaboration_error(
                "office.collaboration.transport_message_invalid",
                "A mutating transport message completed without a durable apply receipt.",
            )
        })?;
        if let Some(sequence) = apply.sequence {
            if sequence > self.cursor_sequence {
                self.suppressed_sequences.insert(sequence);
            }
        }
        Ok(NativeOfficeCollaborationTransportReceiveResult {
            kind,
            ignored: false,
            response: None,
            apply: Some(apply),
        })
    }

    /// Project durable changes made by other CLI/MCP processes into outbound
    /// Update envelopes. Updates received through this session are consumed
    /// but suppressed, preventing a room echo without persisting Awareness.
    pub fn poll(
        &mut self,
        limit: usize,
    ) -> UseResult<NativeOfficeCollaborationTransportPollResult> {
        let batch = self
            .store
            .events(super::NativeOfficeCollaborationEventsRequest {
                after_sequence: Some(self.cursor_sequence),
                limit,
            })?;
        let starting_sequence = batch.starting_sequence;
        if let Some(reset) = &batch.reset {
            self.cursor_sequence = batch.current_sequence;
            self.suppressed_sequences
                .retain(|sequence| *sequence > self.cursor_sequence);
            let full_state = self.message(
                NativeOfficeCollaborationTransportMessageType::Update,
                reset.update.clone(),
                Some(NativeOfficeCollaborationOrigin {
                    protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
                    kind: NativeOfficeCollaborationOriginKind::System,
                    actor_id: Some(self.manifest.actor_id.clone()),
                    operation_id: None,
                }),
            );
            let handshake = self.message(
                NativeOfficeCollaborationTransportMessageType::SyncStep1,
                batch.current_state_vector,
                None,
            );
            return Ok(NativeOfficeCollaborationTransportPollResult {
                starting_sequence,
                cursor_sequence: self.cursor_sequence,
                current_sequence: batch.current_sequence,
                has_more: false,
                resynchronized: true,
                messages: vec![full_state, handshake],
            });
        }

        let mut messages = Vec::new();
        for event in &batch.updates {
            if self.suppressed_sequences.remove(&event.sequence) {
                continue;
            }
            messages.push(self.update_message(event));
        }
        self.cursor_sequence = batch.cursor_sequence;
        Ok(NativeOfficeCollaborationTransportPollResult {
            starting_sequence,
            cursor_sequence: self.cursor_sequence,
            current_sequence: batch.current_sequence,
            has_more: batch.has_more,
            resynchronized: false,
            messages,
        })
    }

    fn validate_message(
        &self,
        message: NativeOfficeCollaborationTransportMessage,
    ) -> UseResult<NativeOfficeCollaborationTransportMessage> {
        if message.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
            || message.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
        {
            return Err(collaboration_error(
                "office.collaboration.transport_message_invalid",
                "The Office collaboration transport protocol or version is unsupported.",
            ));
        }
        if message.artifact_id != self.manifest.artifact_id
            || message.artifact_kind != self.manifest.kind
            || message.namespace != self.manifest.namespace
        {
            return Err(collaboration_error(
                "office.collaboration.transport_identity_mismatch",
                "The Office collaboration transport message belongs to another artifact, kind, or namespace.",
            ));
        }
        if message.sender_client_id > MAX_YJS_CLIENT_ID {
            return Err(collaboration_error(
                "office.collaboration.transport_message_invalid",
                "The Office collaboration transport sender client ID must fit in 53 bits.",
            ));
        }
        if message.message_type != NativeOfficeCollaborationTransportMessageType::Update
            && message.origin.is_some()
        {
            return Err(collaboration_error(
                "office.collaboration.transport_message_invalid",
                "Only incremental Office collaboration Update messages may carry a typed origin.",
            ));
        }
        if let Some(origin) = &message.origin {
            validate_origin(origin)?;
        }
        match message.message_type {
            NativeOfficeCollaborationTransportMessageType::SyncStep1 => {
                decode_state_vector(&message.payload)?;
            }
            NativeOfficeCollaborationTransportMessageType::SyncStep2
            | NativeOfficeCollaborationTransportMessageType::Update => {
                validate_update_size(&message.payload)?;
                yrs::Update::decode_v1(&message.payload).map_err(|error| {
                    collaboration_error(
                        "office.collaboration.transport_message_invalid",
                        format!("The transport payload is not a valid Yjs v1 update: {error}"),
                    )
                })?;
            }
        }
        Ok(message)
    }

    fn update_message(
        &self,
        event: &NativeOfficeCollaborationUpdateEvent,
    ) -> NativeOfficeCollaborationTransportMessage {
        let kind = match event.actor_kind {
            NativeOfficeCollaborationActorKind::Human => {
                NativeOfficeCollaborationOriginKind::Editor
            }
            NativeOfficeCollaborationActorKind::Agent => NativeOfficeCollaborationOriginKind::Agent,
            NativeOfficeCollaborationActorKind::System => {
                NativeOfficeCollaborationOriginKind::System
            }
        };
        let kind = match event.operation_kind {
            NativeOfficeCollaborationOperationKind::Create => {
                NativeOfficeCollaborationOriginKind::Bootstrap
            }
            NativeOfficeCollaborationOperationKind::Join => {
                NativeOfficeCollaborationOriginKind::Import
            }
            _ => kind,
        };
        self.message(
            NativeOfficeCollaborationTransportMessageType::Update,
            event.update.clone(),
            Some(NativeOfficeCollaborationOrigin {
                protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
                kind,
                actor_id: Some(event.actor_id.clone()),
                operation_id: Some(event.operation_id.clone()),
            }),
        )
    }

    fn message(
        &self,
        message_type: NativeOfficeCollaborationTransportMessageType,
        payload: Vec<u8>,
        origin: Option<NativeOfficeCollaborationOrigin>,
    ) -> NativeOfficeCollaborationTransportMessage {
        NativeOfficeCollaborationTransportMessage {
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
            artifact_id: self.manifest.artifact_id.clone(),
            artifact_kind: self.manifest.kind,
            namespace: self.manifest.namespace.clone(),
            sender_client_id: self.manifest.client_id,
            message_type,
            payload,
            origin,
        }
    }
}

fn validate_origin(origin: &NativeOfficeCollaborationOrigin) -> UseResult<()> {
    if origin.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL {
        return Err(collaboration_error(
            "office.collaboration.origin_invalid",
            "Office collaboration transport origins require the versioned collaboration protocol.",
        ));
    }
    if let Some(actor_id) = &origin.actor_id {
        normalized_identifier(actor_id, "origin actor ID")?;
    }
    if let Some(operation_id) = &origin.operation_id {
        normalized_identifier(operation_id, "origin operation ID")?;
    }
    Ok(())
}

fn encode_y_sync_message(
    kind: NativeOfficeCollaborationTransportMessageType,
    payload: &[u8],
) -> UseResult<Vec<u8>> {
    let sync = match kind {
        NativeOfficeCollaborationTransportMessageType::SyncStep1 => {
            SyncMessage::SyncStep1(decode_state_vector(payload)?)
        }
        NativeOfficeCollaborationTransportMessageType::SyncStep2 => {
            validate_update_size(payload)?;
            SyncMessage::SyncStep2(payload.to_vec())
        }
        NativeOfficeCollaborationTransportMessageType::Update => {
            validate_update_size(payload)?;
            SyncMessage::Update(payload.to_vec())
        }
    };
    let mut encoder = EncoderV1::new();
    Message::Sync(sync).encode(&mut encoder);
    Ok(encoder.to_vec())
}

fn decode_one_y_sync_message(
    bytes: &[u8],
) -> UseResult<(NativeOfficeCollaborationTransportMessageType, Vec<u8>)> {
    let mut decoder = DecoderV1::new(Cursor::new(bytes));
    let mut reader = MessageReader::new(&mut decoder);
    let message = reader
        .next()
        .transpose()
        .map_err(|error| {
            collaboration_error(
                "office.collaboration.sync_message_invalid",
                format!("The y-sync response is invalid: {error}"),
            )
        })?
        .ok_or_else(|| {
            collaboration_error(
                "office.collaboration.sync_message_invalid",
                "The y-sync response is empty.",
            )
        })?;
    if reader.next().is_some() {
        return Err(collaboration_error(
            "office.collaboration.sync_message_invalid",
            "The y-sync response contains more than one message.",
        ));
    }
    match message {
        Message::Sync(SyncMessage::SyncStep1(vector)) => Ok((
            NativeOfficeCollaborationTransportMessageType::SyncStep1,
            vector.encode_v1(),
        )),
        Message::Sync(SyncMessage::SyncStep2(update)) => Ok((
            NativeOfficeCollaborationTransportMessageType::SyncStep2,
            update,
        )),
        Message::Sync(SyncMessage::Update(update)) => Ok((
            NativeOfficeCollaborationTransportMessageType::Update,
            update,
        )),
        _ => Err(collaboration_error(
            "office.collaboration.sync_message_unsupported",
            "The y-sync response is not a supported document message.",
        )),
    }
}
