//! Crash-safe Yjs-compatible replica storage for native Office clients.
//!
//! The browser package owns format-specific collaborative models. This module
//! owns the native transport boundary: standard Yjs v1 updates and state
//! vectors, stable replica identity, durable operation receipts, and bounded
//! append-only persistence. It intentionally does not invent a second Office
//! collaboration schema.

mod document;
mod persistence;
mod types;
mod validation;

use std::path::{Path, PathBuf};

use a3s_use_core::UseResult;
use yrs::encoding::read::Cursor;
use yrs::sync::MessageReader;
use yrs::sync::{Message, SyncMessage};
use yrs::updates::decoder::{Decode, DecoderV1};
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};
use yrs::{ReadTxn, StateVector, Transact};

use document::{
    canonical_state_vector, document_state_sha256, inspect_document, new_replica_document,
    state_vector_sha256, validate_and_apply_update,
};
use persistence::{
    compact, create_store, load_store, open_store, write_archived_operation, write_checkpoint,
    write_update_entry, LoadedStore, OperationRecord, StoreLock,
};
pub use types::*;
use validation::{
    assert_operation_replay, assert_state_vector_precondition, creation_payload_sha256,
    decode_state_vector, normalized_identifier, normalized_namespace, operation_payload_sha256,
    validate_client_id, validate_expected_replica, validate_update_size,
};
pub(crate) use validation::{collaboration_error, sha256_hex};

const AUTO_CHECKPOINT_UPDATE_COUNT: usize = 64;
const AUTO_CHECKPOINT_UPDATE_BYTES: u64 = 64 * 1024 * 1024;

/// A local native replica. Every method reopens and locks the durable state so
/// separate CLI, MCP, and coding-agent processes can safely share one replica.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NativeOfficeCollaborationStore {
    root: PathBuf,
}

impl NativeOfficeCollaborationStore {
    pub fn create(request: NativeOfficeCollaborationCreateRequest) -> UseResult<Self> {
        let artifact_id = normalized_identifier(&request.artifact_id, "artifact ID")?;
        let actor_id = normalized_identifier(&request.actor_id, "actor ID")?;
        let operation_id = normalized_identifier(&request.operation_id, "operation ID")?;
        let namespace = normalized_namespace(
            request
                .namespace
                .as_deref()
                .unwrap_or(NATIVE_OFFICE_COLLABORATION_NAMESPACE),
        )?;
        let client_id = request
            .client_id
            .unwrap_or_else(|| yrs::Doc::new().client_id().get());
        validate_client_id(client_id)?;
        if let Some(update) = &request.initial_update {
            validate_update_size(update)?;
        }
        let operation_kind = if request.initial_update.is_some() {
            NativeOfficeCollaborationOperationKind::Join
        } else {
            NativeOfficeCollaborationOperationKind::Create
        };
        if let Some(store) = open_idempotent_create(
            &request.store,
            &artifact_id,
            request.kind,
            &actor_id,
            request.actor_kind,
            request.mode,
            &operation_id,
            &namespace,
            request.client_id,
            request.initial_update.as_deref(),
            operation_kind,
        )? {
            return Ok(store);
        }

        let manifest = NativeOfficeCollaborationManifest {
            format: NATIVE_OFFICE_COLLABORATION_STORE_FORMAT.to_owned(),
            schema_version: NATIVE_OFFICE_COLLABORATION_STORE_SCHEMA_VERSION,
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            protocol_version: NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
            namespace,
            artifact_id,
            kind: request.kind,
            actor_id,
            actor_kind: request.actor_kind,
            mode: request.mode,
            client_id,
        };
        let doc = new_replica_document(client_id, &manifest.namespace, manifest.kind);
        if let Some(update) = &request.initial_update {
            validate_and_apply_update(&doc, update, &manifest)?;
        }
        let document = inspect_document(&doc, &manifest)?;
        let state_vector = canonical_state_vector(&doc.transact().state_vector());
        let payload_sha256 = creation_payload_sha256(
            &manifest,
            request.initial_update.as_deref(),
            &operation_id,
            operation_kind,
        )?;
        let operation = OperationRecord {
            schema_version: 1,
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            operation_id,
            actor_id: manifest.actor_id.clone(),
            actor_kind: manifest.actor_kind,
            mode: manifest.mode,
            kind: operation_kind,
            artifact_id: manifest.artifact_id.clone(),
            artifact_kind: manifest.kind,
            payload_sha256,
            update_sha256: request.initial_update.as_deref().map(sha256_hex),
            before_state_vector_sha256: state_vector_sha256(&StateVector::default()),
            after_state_vector_sha256: sha256_hex(&state_vector),
            sequence: Some(0),
            state_changed: request.initial_update.is_some(),
        };
        let root = match create_store(&request.store, &manifest, &doc, &operation) {
            Ok(root) => root,
            Err(error) if error.code == "office.collaboration.store_exists" => {
                if let Some(store) = open_idempotent_create(
                    &request.store,
                    &manifest.artifact_id,
                    manifest.kind,
                    &manifest.actor_id,
                    manifest.actor_kind,
                    manifest.mode,
                    &operation.operation_id,
                    &manifest.namespace,
                    request.client_id,
                    request.initial_update.as_deref(),
                    operation_kind,
                )? {
                    return Ok(store);
                }
                return Err(error);
            }
            Err(error) => return Err(error),
        };
        let store = Self { root };
        // Reopen once so a successfully returned store is guaranteed to pass
        // the same validation path used after a process restart.
        let inspection = store.inspect()?;
        debug_assert_eq!(inspection.document_state_sha256, document.state_sha256);
        Ok(store)
    }

    pub fn open(path: impl AsRef<Path>) -> UseResult<Self> {
        Ok(Self {
            root: open_store(path.as_ref())?,
        })
    }

    pub fn path(&self) -> &Path {
        &self.root
    }

    pub fn inspect(&self) -> UseResult<NativeOfficeCollaborationInspection> {
        let (_lock, loaded) = self.lock_and_load()?;
        inspection_from_loaded(&self.root, &loaded)
    }

    /// Read a bounded, resumable batch of durable collaboration updates.
    ///
    /// A missing cursor starts at the current sequence so callers can begin
    /// watching without racing a separate inspection. If compaction has
    /// removed the requested history, `reset` contains a complete Yjs v1
    /// update and advances the cursor to the current durable sequence.
    pub fn events(
        &self,
        request: NativeOfficeCollaborationEventsRequest,
    ) -> UseResult<NativeOfficeCollaborationEventBatch> {
        if request.limit == 0 || request.limit > MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH {
            return Err(collaboration_error(
                "office.collaboration.event_limit_invalid",
                format!(
                    "The collaboration event batch limit must be between 1 and {MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH}."
                ),
            )
            .with_detail("limit", request.limit)
            .with_detail(
                "maximum",
                MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
            ));
        }

        let (_lock, loaded) = self.lock_and_load()?;
        let starting_sequence = request.after_sequence.unwrap_or(loaded.current_sequence);
        if starting_sequence > loaded.current_sequence {
            return Err(collaboration_error(
                "office.collaboration.sequence_ahead",
                format!(
                    "The requested collaboration cursor {starting_sequence} is ahead of the current durable sequence {}.",
                    loaded.current_sequence
                ),
            )
            .with_suggestion(
                "Inspect the replica and resume from a cursor at or before its current sequence.",
            )
            .with_detail("afterSequence", starting_sequence)
            .with_detail("currentSequence", loaded.current_sequence));
        }

        let transaction = loaded.doc.transact();
        let current_state_vector = canonical_state_vector(&transaction.state_vector());
        let current_state_vector_sha256 = sha256_hex(&current_state_vector);
        if starting_sequence < loaded.checkpoint_sequence {
            let update = transaction.encode_state_as_update_v1(&StateVector::default());
            validate_update_size(&update)?;
            return Ok(NativeOfficeCollaborationEventBatch {
                starting_sequence,
                cursor_sequence: loaded.current_sequence,
                checkpoint_sequence: loaded.checkpoint_sequence,
                current_sequence: loaded.current_sequence,
                has_more: false,
                current_state_vector,
                current_state_vector_sha256,
                reset: Some(NativeOfficeCollaborationResetEvent {
                    sequence: loaded.current_sequence,
                    update_bytes: update.len() as u64,
                    update_sha256: sha256_hex(&update),
                    update,
                }),
                updates: Vec::new(),
            });
        }
        drop(transaction);

        let updates = loaded
            .update_entries
            .iter()
            .filter(|entry| entry.sequence > starting_sequence)
            .take(request.limit)
            .map(|entry| NativeOfficeCollaborationUpdateEvent {
                sequence: entry.sequence,
                operation_id: entry.operation.operation_id.clone(),
                operation_kind: entry.operation.kind,
                actor_id: entry.operation.actor_id.clone(),
                actor_kind: entry.operation.actor_kind,
                mode: entry.operation.mode,
                artifact_id: entry.operation.artifact_id.clone(),
                artifact_kind: entry.operation.artifact_kind,
                payload_sha256: entry.operation.payload_sha256.clone(),
                update: entry.update.clone(),
                update_bytes: entry.update_bytes,
                update_sha256: sha256_hex(&entry.update),
                before_state_vector_sha256: entry.operation.before_state_vector_sha256.clone(),
                after_state_vector_sha256: entry.operation.after_state_vector_sha256.clone(),
            })
            .collect::<Vec<_>>();
        let cursor_sequence = updates
            .last()
            .map_or(starting_sequence, |event| event.sequence);

        Ok(NativeOfficeCollaborationEventBatch {
            starting_sequence,
            cursor_sequence,
            checkpoint_sequence: loaded.checkpoint_sequence,
            current_sequence: loaded.current_sequence,
            has_more: cursor_sequence < loaded.current_sequence,
            current_state_vector,
            current_state_vector_sha256,
            reset: None,
            updates,
        })
    }

    /// Encode the local state missing from `remote_state_vector`. Supplying
    /// `None` produces a complete standard Yjs v1 update.
    pub fn synchronize(
        &self,
        remote_state_vector: Option<&[u8]>,
    ) -> UseResult<NativeOfficeCollaborationSyncResult> {
        let (_lock, loaded) = self.lock_and_load()?;
        let remote = match remote_state_vector {
            Some(bytes) => decode_state_vector(bytes)?,
            None => StateVector::default(),
        };
        let transaction = loaded.doc.transact();
        let update = transaction.encode_state_as_update_v1(&remote);
        let local_state_vector = canonical_state_vector(&transaction.state_vector());
        Ok(NativeOfficeCollaborationSyncResult {
            update_sha256: sha256_hex(&update),
            update,
            state_vector_sha256: sha256_hex(&local_state_vector),
            state_vector: local_state_vector,
        })
    }

    /// Encode a standard y-sync SyncStep1 message for host-owned transports.
    /// Awareness and room identity stay outside the durable replica.
    pub fn sync_step1(&self) -> UseResult<NativeOfficeCollaborationSyncStepResult> {
        let (_lock, loaded) = self.lock_and_load()?;
        let state_vector_value = loaded.doc.transact().state_vector();
        let state_vector = canonical_state_vector(&state_vector_value);
        let mut encoder = EncoderV1::new();
        Message::Sync(SyncMessage::SyncStep1(state_vector_value)).encode(&mut encoder);
        let message = encoder.to_vec();
        Ok(NativeOfficeCollaborationSyncStepResult {
            message_sha256: sha256_hex(&message),
            message,
            state_vector_sha256: sha256_hex(&state_vector),
            state_vector,
        })
    }

    /// Wrap a standard Yjs v1 update in a y-sync Update message. The caller
    /// remains responsible for transport framing, rooms, and authorization.
    pub fn encode_sync_update(update: &[u8]) -> UseResult<Vec<u8>> {
        validate_update_size(update)?;
        yrs::Update::decode_v1(update).map_err(|error| {
            collaboration_error(
                "office.collaboration.update_invalid",
                format!("The input is not a valid Yjs v1 update: {error}"),
            )
        })?;
        let mut encoder = EncoderV1::new();
        Message::Sync(SyncMessage::Update(update.to_vec())).encode(&mut encoder);
        Ok(encoder.to_vec())
    }

    /// Handle one standard y-sync document message. Awareness, auth, room
    /// selection, and multi-message transport packets remain host concerns.
    /// SyncStep2 and Update messages are persisted through the same
    /// actor/operation/precondition contract as raw v1 updates.
    pub fn handle_sync_message(
        &self,
        message: &[u8],
        mutation: Option<NativeOfficeCollaborationApplyRequest>,
    ) -> UseResult<NativeOfficeCollaborationSyncMessageResult> {
        if message.len() > MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES + 32 {
            return Err(collaboration_error(
                "office.collaboration.sync_message_too_large",
                "The y-sync message exceeds the bounded collaboration message size.",
            ));
        }
        let mut decoder = DecoderV1::new(Cursor::new(message));
        let mut messages = MessageReader::new(&mut decoder);
        let decoded = messages
            .next()
            .transpose()
            .map_err(|error| {
                collaboration_error(
                    "office.collaboration.sync_message_invalid",
                    format!("The input is not one complete y-sync v1 message: {error}"),
                )
            })?
            .ok_or_else(|| {
                collaboration_error(
                    "office.collaboration.sync_message_invalid",
                    "The y-sync v1 message is empty.",
                )
            })?;
        if messages.next().is_some() {
            return Err(collaboration_error(
                "office.collaboration.sync_message_invalid",
                "The input contains more than one y-sync v1 message.",
            ));
        }
        match decoded {
            Message::Sync(SyncMessage::SyncStep1(remote)) => {
                if mutation.is_some() {
                    return Err(collaboration_error(
                        "office.collaboration.sync_message_invalid",
                        "SyncStep1 is read-only and must not include mutation identity.",
                    ));
                }
                let (_lock, loaded) = self.lock_and_load()?;
                let transaction = loaded.doc.transact();
                let update = transaction.encode_state_as_update_v1(&remote);
                let state_vector = canonical_state_vector(&transaction.state_vector());
                let mut encoder = EncoderV1::new();
                Message::Sync(SyncMessage::SyncStep2(update)).encode(&mut encoder);
                let response = encoder.to_vec();
                Ok(NativeOfficeCollaborationSyncMessageResult {
                    kind: NativeOfficeCollaborationSyncMessageKind::SyncStep1,
                    response_sha256: Some(sha256_hex(&response)),
                    response: Some(response),
                    state_vector_sha256: sha256_hex(&state_vector),
                    state_vector,
                    apply: None,
                })
            }
            Message::Sync(SyncMessage::SyncStep2(update)) => {
                self.handle_sync_update(
                    update,
                    mutation,
                    NativeOfficeCollaborationSyncMessageKind::SyncStep2,
                )
            }
            Message::Sync(SyncMessage::Update(update)) => self.handle_sync_update(
                update,
                mutation,
                NativeOfficeCollaborationSyncMessageKind::Update,
            ),
            Message::Awareness(_) | Message::AwarenessQuery | Message::Auth(_) => {
                Err(collaboration_error(
                    "office.collaboration.sync_message_unsupported",
                    "Durable Office replicas accept y-sync document messages only; awareness and authentication belong to the host transport.",
                ))
            }
            Message::Custom(tag, _) => Err(collaboration_error(
                "office.collaboration.sync_message_unsupported",
                format!("The y-sync custom message tag {tag} is unsupported."),
            )),
        }
    }

    fn handle_sync_update(
        &self,
        update: Vec<u8>,
        mutation: Option<NativeOfficeCollaborationApplyRequest>,
        kind: NativeOfficeCollaborationSyncMessageKind,
    ) -> UseResult<NativeOfficeCollaborationSyncMessageResult> {
        let Some(mut mutation) = mutation else {
            return Err(collaboration_error(
                "office.collaboration.sync_identity_required",
                "Applying a y-sync document message requires actor, operation, artifact, kind, and mode identity.",
            ));
        };
        if !mutation.update.is_empty() && mutation.update != update {
            return Err(collaboration_error(
                "office.collaboration.sync_message_invalid",
                "The explicit update does not match the y-sync message payload.",
            ));
        }
        mutation.update = update;
        let apply = self.apply(mutation)?;
        let state_vector = apply.state_vector.clone();
        Ok(NativeOfficeCollaborationSyncMessageResult {
            kind,
            response: None,
            response_sha256: None,
            state_vector_sha256: apply.state_vector_sha256.clone(),
            state_vector,
            apply: Some(apply),
        })
    }

    pub fn apply(
        &self,
        request: NativeOfficeCollaborationApplyRequest,
    ) -> UseResult<NativeOfficeCollaborationApplyResult> {
        validate_update_size(&request.update)?;
        let operation_id = normalized_identifier(&request.operation_id, "operation ID")?;
        let actor_id = normalized_identifier(&request.actor_id, "actor ID")?;
        let expected_artifact_id =
            normalized_identifier(&request.expected_artifact_id, "expected artifact ID")?;
        let (_lock, mut loaded) = self.lock_and_load()?;
        validate_expected_replica(
            &loaded.manifest,
            &actor_id,
            request.mode,
            &expected_artifact_id,
            request.expected_kind,
        )?;
        let update_sha256 = sha256_hex(&request.update);
        let payload_sha256 = operation_payload_sha256(
            NativeOfficeCollaborationOperationKind::Synchronize,
            &loaded.manifest,
            &operation_id,
            Some(&update_sha256),
            request.if_state_vector.as_deref(),
        )?;
        if let Some(existing) = loaded.find_operation(&operation_id)? {
            assert_operation_replay(&existing, &payload_sha256)?;
            let state_vector = canonical_state_vector(&loaded.doc.transact().state_vector());
            return Ok(NativeOfficeCollaborationApplyResult {
                operation_id,
                duplicate: true,
                state_changed: existing.state_changed,
                sequence: existing.sequence,
                update_sha256,
                state_vector_sha256: sha256_hex(&state_vector),
                state_vector,
                checkpointed: false,
            });
        }
        assert_state_vector_precondition(&loaded, request.if_state_vector.as_deref())?;

        let before_state_vector = loaded.doc.transact().state_vector();
        let before_state_sha256 = document_state_sha256(&loaded.doc);
        validate_and_apply_update(&loaded.doc, &request.update, &loaded.manifest)?;
        let after_state_sha256 = document_state_sha256(&loaded.doc);
        let state_changed = before_state_sha256 != after_state_sha256;
        let state_vector = canonical_state_vector(&loaded.doc.transact().state_vector());
        let sequence = state_changed.then_some(loaded.next_sequence);
        let operation = OperationRecord {
            schema_version: 1,
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            operation_id: operation_id.clone(),
            actor_id,
            actor_kind: loaded.manifest.actor_kind,
            mode: request.mode,
            kind: NativeOfficeCollaborationOperationKind::Synchronize,
            artifact_id: expected_artifact_id,
            artifact_kind: request.expected_kind,
            payload_sha256,
            update_sha256: Some(update_sha256.clone()),
            before_state_vector_sha256: state_vector_sha256(&before_state_vector),
            after_state_vector_sha256: sha256_hex(&state_vector),
            sequence,
            state_changed,
        };
        if state_changed {
            write_update_entry(
                &self.root,
                loaded.next_sequence,
                &request.update,
                &operation,
            )?;
            loaded = load_store(&self.root)?;
        } else {
            write_archived_operation(&self.root, &operation)?;
        }

        let should_checkpoint = loaded.update_entries.len() >= AUTO_CHECKPOINT_UPDATE_COUNT
            || loaded.update_bytes >= AUTO_CHECKPOINT_UPDATE_BYTES;
        if should_checkpoint {
            let checkpoint_sequence = loaded.current_sequence;
            write_checkpoint(&self.root, checkpoint_sequence, &loaded.doc)?;
            compact(&self.root, &loaded, checkpoint_sequence)?;
        }
        Ok(NativeOfficeCollaborationApplyResult {
            operation_id,
            duplicate: false,
            state_changed,
            sequence,
            update_sha256,
            state_vector_sha256: sha256_hex(&state_vector),
            state_vector,
            checkpointed: should_checkpoint,
        })
    }

    pub fn checkpoint(
        &self,
        request: NativeOfficeCollaborationCheckpointRequest,
    ) -> UseResult<NativeOfficeCollaborationCheckpointResult> {
        self.checkpoint_with_kind(request, NativeOfficeCollaborationOperationKind::Checkpoint)
    }

    /// Persist a final checkpoint and release the replica lock. The store is
    /// deliberately retained so a coding agent can reconnect without losing
    /// offline state or operation receipts.
    pub fn leave(
        &self,
        request: NativeOfficeCollaborationCheckpointRequest,
    ) -> UseResult<NativeOfficeCollaborationCheckpointResult> {
        self.checkpoint_with_kind(request, NativeOfficeCollaborationOperationKind::Leave)
    }

    fn checkpoint_with_kind(
        &self,
        request: NativeOfficeCollaborationCheckpointRequest,
        kind: NativeOfficeCollaborationOperationKind,
    ) -> UseResult<NativeOfficeCollaborationCheckpointResult> {
        let operation_id = normalized_identifier(&request.operation_id, "operation ID")?;
        let actor_id = normalized_identifier(&request.actor_id, "actor ID")?;
        let expected_artifact_id =
            normalized_identifier(&request.expected_artifact_id, "expected artifact ID")?;
        let (_lock, loaded) = self.lock_and_load()?;
        validate_expected_replica(
            &loaded.manifest,
            &actor_id,
            request.mode,
            &expected_artifact_id,
            request.expected_kind,
        )?;
        let payload_sha256 = operation_payload_sha256(
            kind,
            &loaded.manifest,
            &operation_id,
            None,
            request.if_state_vector.as_deref(),
        )?;
        let state_vector = canonical_state_vector(&loaded.doc.transact().state_vector());
        if let Some(existing) = loaded.find_operation(&operation_id)? {
            assert_operation_replay(&existing, &payload_sha256)?;
            return Ok(NativeOfficeCollaborationCheckpointResult {
                operation_id,
                duplicate: true,
                sequence: existing.sequence.unwrap_or(loaded.current_sequence),
                state_vector_sha256: sha256_hex(&state_vector),
                state_vector,
                compacted_updates: 0,
            });
        }
        assert_state_vector_precondition(&loaded, request.if_state_vector.as_deref())?;

        let sequence = loaded.current_sequence;
        write_checkpoint(&self.root, sequence, &loaded.doc)?;
        let compacted_updates = compact(&self.root, &loaded, sequence)?;
        let state_vector_sha256 = sha256_hex(&state_vector);
        let operation = OperationRecord {
            schema_version: 1,
            protocol: NATIVE_OFFICE_COLLABORATION_PROTOCOL.to_owned(),
            operation_id: operation_id.clone(),
            actor_id,
            actor_kind: loaded.manifest.actor_kind,
            mode: request.mode,
            kind,
            artifact_id: expected_artifact_id,
            artifact_kind: request.expected_kind,
            payload_sha256,
            update_sha256: None,
            before_state_vector_sha256: state_vector_sha256.clone(),
            after_state_vector_sha256: state_vector_sha256.clone(),
            sequence: Some(sequence),
            state_changed: false,
        };
        write_archived_operation(&self.root, &operation)?;
        Ok(NativeOfficeCollaborationCheckpointResult {
            operation_id,
            duplicate: false,
            sequence,
            state_vector,
            state_vector_sha256,
            compacted_updates,
        })
    }

    fn lock_and_load(&self) -> UseResult<(StoreLock, LoadedStore)> {
        let lock = StoreLock::acquire(&self.root)?;
        let loaded = load_store(&self.root)?;
        Ok((lock, loaded))
    }
}

#[allow(clippy::too_many_arguments)]
fn open_idempotent_create(
    path: &Path,
    artifact_id: &str,
    artifact_kind: NativeOfficeCollaborationArtifactKind,
    actor_id: &str,
    actor_kind: NativeOfficeCollaborationActorKind,
    mode: NativeOfficeCollaborationMode,
    operation_id: &str,
    namespace: &str,
    explicit_client_id: Option<u64>,
    initial_update: Option<&[u8]>,
    operation_kind: NativeOfficeCollaborationOperationKind,
) -> UseResult<Option<NativeOfficeCollaborationStore>> {
    if !path.exists() {
        return Ok(None);
    }
    let store = match NativeOfficeCollaborationStore::open(path) {
        Ok(store) => store,
        Err(error)
            if matches!(
                error.code.as_str(),
                "office.collaboration.store_invalid"
                    | "office.collaboration.store_unavailable"
                    | "office.collaboration.store_io_failed"
            ) =>
        {
            return Err(collaboration_error(
                "office.collaboration.store_exists",
                format!(
                    "Path '{}' already exists and is not the requested collaboration replica.",
                    path.display()
                ),
            )
            .with_detail("store", path.display().to_string()));
        }
        Err(error) => return Err(error),
    };
    let (_lock, loaded) = store.lock_and_load()?;
    validate_expected_replica(&loaded.manifest, actor_id, mode, artifact_id, artifact_kind)?;
    if loaded.manifest.actor_kind != actor_kind {
        return Err(collaboration_error(
            "office.collaboration.actor_mismatch",
            format!(
                "The replica uses '{}' actor kind, not '{}'.",
                loaded.manifest.actor_kind.as_str(),
                actor_kind.as_str()
            ),
        )
        .with_detail("expectedActorKind", actor_kind.as_str())
        .with_detail("actualActorKind", loaded.manifest.actor_kind.as_str()));
    }
    if loaded.manifest.namespace != namespace {
        return Err(collaboration_error(
            "office.collaboration.namespace_mismatch",
            format!(
                "The replica uses namespace '{}', not '{}'.",
                loaded.manifest.namespace, namespace
            ),
        )
        .with_detail("expectedNamespace", namespace.to_owned())
        .with_detail("actualNamespace", loaded.manifest.namespace.clone()));
    }
    if explicit_client_id.is_some_and(|client_id| client_id != loaded.manifest.client_id) {
        return Err(collaboration_error(
            "office.collaboration.client_id_mismatch",
            "The existing replica uses a different Yjs client ID.",
        )
        .with_detail("expectedClientId", explicit_client_id.unwrap_or_default())
        .with_detail("actualClientId", loaded.manifest.client_id));
    }
    let payload_sha256 = creation_payload_sha256(
        &loaded.manifest,
        initial_update,
        operation_id,
        operation_kind,
    )?;
    let Some(existing) = loaded.find_operation(operation_id)? else {
        return Err(collaboration_error(
            "office.collaboration.store_exists",
            format!(
                "Collaboration replica '{}' already exists and was created by another operation.",
                store.path().display()
            ),
        )
        .with_detail("store", store.path().display().to_string()));
    };
    if existing.kind != operation_kind {
        return Err(collaboration_error(
            "office.collaboration.operation_conflict",
            format!(
                "Operation ID '{}' was already used for a different collaboration action.",
                operation_id
            ),
        )
        .with_detail("operationId", operation_id.to_owned()));
    }
    assert_operation_replay(&existing, &payload_sha256)?;
    drop(loaded);
    drop(_lock);
    Ok(Some(store))
}

fn inspection_from_loaded(
    root: &Path,
    loaded: &LoadedStore,
) -> UseResult<NativeOfficeCollaborationInspection> {
    let document = inspect_document(&loaded.doc, &loaded.manifest)?;
    let state = loaded.doc.transact().state_vector();
    let state_vector = canonical_state_vector(&state);
    let mut state_vector_entries = state
        .iter()
        .map(
            |(client_id, clock)| NativeOfficeCollaborationStateVectorEntry {
                client_id: client_id.get(),
                clock: *clock,
            },
        )
        .collect::<Vec<_>>();
    state_vector_entries.sort_by_key(|entry| entry.client_id);
    Ok(NativeOfficeCollaborationInspection {
        store: root.to_path_buf(),
        manifest: loaded.manifest.clone(),
        metadata: document.metadata,
        metadata_fields_present: document.metadata_fields_present,
        metadata_complete: document.metadata_complete,
        bootstrap_initializer_count: document.bootstrap_initializer_count,
        bootstrap_valid: document.bootstrap_valid,
        root_names: document.root_names,
        state_vector_sha256: sha256_hex(&state_vector),
        state_vector,
        state_vector_entries,
        document_state_sha256: document.state_sha256,
        pending_updates: document.pending_updates,
        checkpoint_sequence: loaded.checkpoint_sequence,
        current_sequence: loaded.current_sequence,
        update_count: loaded.update_entries.len(),
        update_bytes: loaded.update_bytes,
        operation_count: loaded.operation_count()?,
    })
}

#[cfg(test)]
mod tests;
