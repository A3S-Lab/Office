use std::path::PathBuf;

use a3s_office::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationApplyRequest,
    NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationCheckpointRequest,
    NativeOfficeCollaborationCreateRequest, NativeOfficeCollaborationEventBatch,
    NativeOfficeCollaborationEventsRequest, NativeOfficeCollaborationInspection,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationMutationRequest,
    NativeOfficeCollaborationStore, MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH,
    MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
};
use a3s_use_core::{UseError, UseResult};
use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};

use super::support::output_encoding_error;

mod mutation;

use mutation::OfficeCollaborationMutation;

const DEFAULT_EVENT_LIMIT: usize = 64;
const MAX_MCP_COLLABORATION_UPDATE_BYTES: usize = 4 * 1024 * 1024;
const _: () = assert!(DEFAULT_EVENT_LIMIT <= MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(super) enum OfficeCollaborationArtifactKind {
    Document,
    Markdown,
    Spreadsheet,
    Presentation,
    Pdf,
}

impl From<OfficeCollaborationArtifactKind> for NativeOfficeCollaborationArtifactKind {
    fn from(value: OfficeCollaborationArtifactKind) -> Self {
        match value {
            OfficeCollaborationArtifactKind::Document => Self::Document,
            OfficeCollaborationArtifactKind::Markdown => Self::Markdown,
            OfficeCollaborationArtifactKind::Spreadsheet => Self::Spreadsheet,
            OfficeCollaborationArtifactKind::Presentation => Self::Presentation,
            OfficeCollaborationArtifactKind::Pdf => Self::Pdf,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(super) enum OfficeCollaborationActorKind {
    Human,
    Agent,
    System,
}

impl From<OfficeCollaborationActorKind> for NativeOfficeCollaborationActorKind {
    fn from(value: OfficeCollaborationActorKind) -> Self {
        match value {
            OfficeCollaborationActorKind::Human => Self::Human,
            OfficeCollaborationActorKind::Agent => Self::Agent,
            OfficeCollaborationActorKind::System => Self::System,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(super) enum OfficeCollaborationMode {
    View,
    Comment,
    Suggest,
    Edit,
}

impl From<OfficeCollaborationMode> for NativeOfficeCollaborationMode {
    fn from(value: OfficeCollaborationMode) -> Self {
        match value {
            OfficeCollaborationMode::View => Self::View,
            OfficeCollaborationMode::Comment => Self::Comment,
            OfficeCollaborationMode::Suggest => Self::Suggest,
            OfficeCollaborationMode::Edit => Self::Edit,
        }
    }
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationCreateInput {
    /// New durable collaboration replica directory. Existing unrelated paths are never replaced.
    pub(super) store: String,
    /// Stable host-owned artifact identity.
    pub(super) artifact_id: String,
    /// Canonical collaborative model family.
    pub(super) kind: OfficeCollaborationArtifactKind,
    /// Stable human, agent, or system identity recorded on every mutation.
    pub(super) actor_id: String,
    /// Actor category. Defaults to `agent`.
    pub(super) actor_kind: Option<OfficeCollaborationActorKind>,
    /// Authorized collaboration mode. Defaults to `edit`.
    pub(super) mode: Option<OfficeCollaborationMode>,
    /// Stable idempotency key for this create/join operation.
    pub(super) operation_id: String,
    /// Optional Yjs namespace override.
    pub(super) namespace: Option<String>,
    /// Optional explicit non-zero 53-bit Yjs client ID.
    pub(super) client_id: Option<u64>,
    /// Optional standard Yjs v1 update used to join an existing shared state.
    pub(super) initial_update_base64: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationStoreInput {
    /// Existing durable collaboration replica directory.
    pub(super) store: String,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationDiffInput {
    /// Existing durable collaboration replica directory.
    pub(super) store: String,
    /// Optional remote standard Yjs state vector. Omit for a complete update.
    pub(super) state_vector_base64: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationEventsInput {
    /// Existing durable collaboration replica directory.
    pub(super) store: String,
    /// Last durably consumed sequence. Omit to start at the current sequence.
    pub(super) after_sequence: Option<u64>,
    /// Maximum returned updates, from 1 through 256. Defaults to 64.
    pub(super) limit: Option<usize>,
    /// Include each standard Yjs v1 update as base64. Defaults to metadata only.
    #[serde(default)]
    pub(super) include_updates: bool,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationApplyInput {
    /// Existing durable collaboration replica directory.
    pub(super) store: String,
    /// Stable idempotency key for this delivered update.
    pub(super) operation_id: String,
    /// Actor identity, which must match the replica manifest.
    pub(super) actor_id: String,
    /// Authorized collaboration mode, which must match the replica manifest.
    pub(super) mode: OfficeCollaborationMode,
    /// Expected host-owned artifact identity.
    pub(super) artifact_id: String,
    /// Expected canonical collaborative model family.
    pub(super) kind: OfficeCollaborationArtifactKind,
    /// Standard Yjs v1 update, limited to 4 MiB on the MCP JSON transport.
    pub(super) update_base64: String,
    /// Optional exact state-vector precondition for fail-closed agent decisions.
    pub(super) if_state_vector_base64: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationMutationInput {
    /// Existing durable collaboration replica directory.
    pub(super) store: String,
    /// Stable idempotency key for this typed mutation.
    pub(super) operation_id: String,
    /// Actor identity, which must match the replica manifest.
    pub(super) actor_id: String,
    /// Authorized collaboration mode. Canonical mutations require `edit`.
    pub(super) mode: OfficeCollaborationMode,
    /// Expected host-owned artifact identity.
    pub(super) artifact_id: String,
    /// Expected canonical collaborative model family.
    pub(super) kind: OfficeCollaborationArtifactKind,
    /// Closed, format-aware Markdown or Document mutation.
    pub(super) mutation: OfficeCollaborationMutation,
    /// Optional exact state-vector precondition for fail-closed agent decisions.
    pub(super) if_state_vector_base64: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct OfficeCollaborationCheckpointInput {
    /// Existing durable collaboration replica directory.
    pub(super) store: String,
    /// Stable idempotency key for this checkpoint or leave operation.
    pub(super) operation_id: String,
    /// Actor identity, which must match the replica manifest.
    pub(super) actor_id: String,
    /// Authorized collaboration mode, which must match the replica manifest.
    pub(super) mode: OfficeCollaborationMode,
    /// Expected host-owned artifact identity.
    pub(super) artifact_id: String,
    /// Expected canonical collaborative model family.
    pub(super) kind: OfficeCollaborationArtifactKind,
    /// Optional exact state-vector precondition.
    pub(super) if_state_vector_base64: Option<String>,
    /// Record a durable leave operation instead of an ordinary checkpoint.
    #[serde(default)]
    pub(super) leave: bool,
}

pub(super) async fn create(input: OfficeCollaborationCreateInput) -> UseResult<serde_json::Value> {
    let joined = input.initial_update_base64.is_some();
    let initial_update = decode_optional_binary(
        input.initial_update_base64.as_deref(),
        MAX_MCP_COLLABORATION_UPDATE_BYTES,
        "initial update",
    )?;
    let request = NativeOfficeCollaborationCreateRequest {
        store: PathBuf::from(input.store),
        artifact_id: input.artifact_id,
        kind: input.kind.into(),
        actor_id: input.actor_id,
        actor_kind: input
            .actor_kind
            .unwrap_or(OfficeCollaborationActorKind::Agent)
            .into(),
        mode: input.mode.unwrap_or(OfficeCollaborationMode::Edit).into(),
        operation_id: input.operation_id,
        namespace: input.namespace,
        client_id: input.client_id,
        initial_update,
    };
    let inspection = run_blocking(move || {
        let store = NativeOfficeCollaborationStore::create(request)?;
        store.inspect()
    })
    .await?;
    Ok(json!({
        "action": if joined { "joined" } else { "created" },
        "replica": inspection_value(inspection)?,
    }))
}

pub(super) async fn inspect(input: OfficeCollaborationStoreInput) -> UseResult<serde_json::Value> {
    let inspection =
        run_blocking(move || NativeOfficeCollaborationStore::open(input.store)?.inspect()).await?;
    inspection_value(inspection)
}

pub(super) async fn diff(input: OfficeCollaborationDiffInput) -> UseResult<serde_json::Value> {
    let remote_state_vector = decode_optional_binary(
        input.state_vector_base64.as_deref(),
        MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
        "state vector",
    )?;
    let result = run_blocking(move || {
        NativeOfficeCollaborationStore::open(input.store)?
            .synchronize(remote_state_vector.as_deref())
    })
    .await?;
    Ok(json!({
        "updateBase64": STANDARD.encode(&result.update),
        "updateBytes": result.update.len(),
        "updateSha256": result.update_sha256,
        "stateVectorBase64": STANDARD.encode(&result.state_vector),
        "stateVectorSha256": result.state_vector_sha256,
    }))
}

pub(super) async fn events(input: OfficeCollaborationEventsInput) -> UseResult<serde_json::Value> {
    let include_updates = input.include_updates;
    let limit = input.limit.unwrap_or(DEFAULT_EVENT_LIMIT);
    if !(1..=MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH).contains(&limit) {
        return Err(UseError::new(
            "office.collaboration.event_limit_invalid",
            format!(
                "The MCP collaboration event limit must be between 1 and {MAX_NATIVE_OFFICE_COLLABORATION_EVENT_BATCH}."
            ),
        )
        .with_detail("limit", limit));
    }
    let request = NativeOfficeCollaborationEventsRequest {
        after_sequence: input.after_sequence,
        limit,
    };
    let batch =
        run_blocking(move || NativeOfficeCollaborationStore::open(input.store)?.events(request))
            .await?;
    event_batch_value(batch, include_updates)
}

pub(super) async fn apply(input: OfficeCollaborationApplyInput) -> UseResult<serde_json::Value> {
    let update = decode_binary(
        &input.update_base64,
        MAX_MCP_COLLABORATION_UPDATE_BYTES,
        "update",
    )?;
    let if_state_vector = decode_optional_binary(
        input.if_state_vector_base64.as_deref(),
        MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
        "state-vector precondition",
    )?;
    let request = NativeOfficeCollaborationApplyRequest {
        operation_id: input.operation_id,
        actor_id: input.actor_id,
        mode: input.mode.into(),
        expected_artifact_id: input.artifact_id,
        expected_kind: input.kind.into(),
        update,
        if_state_vector,
        origin: None,
    };
    let result =
        run_blocking(move || NativeOfficeCollaborationStore::open(input.store)?.apply(request))
            .await?;
    let state_vector_base64 = STANDARD.encode(&result.state_vector);
    result_with_state_vector(result, state_vector_base64)
}

pub(super) async fn mutate(
    input: OfficeCollaborationMutationInput,
) -> UseResult<serde_json::Value> {
    let mutation_bytes = serde_json::to_vec(&input.mutation).map_err(output_encoding_error)?;
    if mutation_bytes.len() > MAX_MCP_COLLABORATION_UPDATE_BYTES {
        return Err(UseError::new(
            "office.collaboration.input_too_large",
            format!(
                "The MCP typed collaboration mutation is {} bytes; the limit is {} bytes.",
                mutation_bytes.len(),
                MAX_MCP_COLLABORATION_UPDATE_BYTES
            ),
        )
        .with_suggestion("Use the collaboration CLI with a bounded mutation input file."));
    }
    let if_state_vector = decode_optional_binary(
        input.if_state_vector_base64.as_deref(),
        MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
        "state-vector precondition",
    )?;
    let output_mutation = input.mutation.clone();
    let request = NativeOfficeCollaborationMutationRequest {
        operation_id: input.operation_id,
        actor_id: input.actor_id,
        mode: input.mode.into(),
        expected_artifact_id: input.artifact_id,
        expected_kind: input.kind.into(),
        mutation: input.mutation.into(),
        if_state_vector,
    };
    let result =
        run_blocking(move || NativeOfficeCollaborationStore::open(input.store)?.mutate(request))
            .await?;
    let state_vector_base64 = STANDARD.encode(&result.state_vector);
    let mut value = result_with_state_vector(result, state_vector_base64)?;
    value["action"] = Value::String("mutated".to_owned());
    value["mutation"] = serde_json::to_value(output_mutation).map_err(output_encoding_error)?;
    Ok(value)
}

pub(super) async fn checkpoint(
    input: OfficeCollaborationCheckpointInput,
) -> UseResult<serde_json::Value> {
    let if_state_vector = decode_optional_binary(
        input.if_state_vector_base64.as_deref(),
        MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
        "state-vector precondition",
    )?;
    let leave = input.leave;
    let request = NativeOfficeCollaborationCheckpointRequest {
        operation_id: input.operation_id,
        actor_id: input.actor_id,
        mode: input.mode.into(),
        expected_artifact_id: input.artifact_id,
        expected_kind: input.kind.into(),
        if_state_vector,
    };
    let result = run_blocking(move || {
        let store = NativeOfficeCollaborationStore::open(input.store)?;
        if leave {
            store.leave(request)
        } else {
            store.checkpoint(request)
        }
    })
    .await?;
    let state_vector_base64 = STANDARD.encode(&result.state_vector);
    let mut value = result_with_state_vector(result, state_vector_base64)?;
    value["action"] = Value::String(if leave { "left" } else { "checkpointed" }.to_owned());
    Ok(value)
}

fn inspection_value(
    inspection: NativeOfficeCollaborationInspection,
) -> UseResult<serde_json::Value> {
    let state_vector_base64 = STANDARD.encode(&inspection.state_vector);
    let mut value = serde_json::to_value(inspection).map_err(output_encoding_error)?;
    replace_binary_field(
        &mut value,
        "stateVector",
        "stateVectorBase64",
        state_vector_base64,
    )?;
    Ok(value)
}

fn result_with_state_vector<T: Serialize>(
    result: T,
    state_vector_base64: String,
) -> UseResult<serde_json::Value> {
    let mut value = serde_json::to_value(result).map_err(output_encoding_error)?;
    replace_binary_field(
        &mut value,
        "stateVector",
        "stateVectorBase64",
        state_vector_base64,
    )?;
    Ok(value)
}

fn replace_binary_field(
    value: &mut Value,
    original: &str,
    encoded: &str,
    base64: String,
) -> UseResult<()> {
    let object = value.as_object_mut().ok_or_else(|| {
        UseError::new(
            "use.office.output_invalid",
            "Native collaboration output is not a JSON object.",
        )
    })?;
    object.remove(original);
    object.insert(encoded.to_owned(), Value::String(base64));
    Ok(())
}

fn event_batch_value(
    batch: NativeOfficeCollaborationEventBatch,
    include_updates: bool,
) -> UseResult<serde_json::Value> {
    let reset = batch.reset.map(|event| {
        let mut value = json!({
            "reason": "history-compacted",
            "sequence": event.sequence,
            "cursorSequence": event.sequence,
            "updateBytes": event.update_bytes,
            "updateSha256": event.update_sha256,
        });
        if include_updates {
            value["updateBase64"] = Value::String(STANDARD.encode(&event.update));
        }
        value
    });
    let updates = batch
        .updates
        .into_iter()
        .map(|event| {
            let mut value = json!({
                "sequence": event.sequence,
                "cursorSequence": event.sequence,
                "operationId": event.operation_id,
                "operationKind": event.operation_kind,
                "actorId": event.actor_id,
                "actorKind": event.actor_kind,
                "mode": event.mode,
                "artifactId": event.artifact_id,
                "artifactKind": event.artifact_kind,
                "payloadSha256": event.payload_sha256,
                "updateBytes": event.update_bytes,
                "updateSha256": event.update_sha256,
                "beforeStateVectorSha256": event.before_state_vector_sha256,
                "afterStateVectorSha256": event.after_state_vector_sha256,
            });
            if let Some(origin) = &event.origin {
                value["origin"] = json!(origin);
            }
            if include_updates {
                value["updateBase64"] = Value::String(STANDARD.encode(&event.update));
            }
            value
        })
        .collect::<Vec<_>>();
    let event_count = updates.len() + usize::from(reset.is_some());
    Ok(json!({
        "startingSequence": batch.starting_sequence,
        "cursorSequence": batch.cursor_sequence,
        "checkpointSequence": batch.checkpoint_sequence,
        "currentSequence": batch.current_sequence,
        "hasMore": batch.has_more,
        "eventCount": event_count,
        "includeUpdates": include_updates,
        "stateVectorBase64": STANDARD.encode(&batch.current_state_vector),
        "stateVectorSha256": batch.current_state_vector_sha256,
        "reset": reset,
        "updates": updates,
    }))
}

fn decode_optional_binary(
    encoded: Option<&str>,
    max_bytes: usize,
    label: &str,
) -> UseResult<Option<Vec<u8>>> {
    encoded
        .map(|encoded| decode_binary(encoded, max_bytes, label))
        .transpose()
}

fn decode_binary(encoded: &str, max_bytes: usize, label: &str) -> UseResult<Vec<u8>> {
    let max_encoded_bytes = max_bytes.saturating_add(2) / 3 * 4;
    if encoded.len() > max_encoded_bytes {
        return Err(UseError::new(
            "office.collaboration.input_too_large",
            format!(
                "The MCP collaboration {label} exceeds the encoded limit of {max_encoded_bytes} bytes."
            ),
        )
        .with_suggestion(
            "Use the collaboration CLI with a bounded binary input file for larger updates.",
        ));
    }
    let bytes = STANDARD.decode(encoded).map_err(|error| {
        UseError::new(
            "office.collaboration.input_invalid",
            format!("The MCP collaboration {label} is not valid base64: {error}"),
        )
    })?;
    if bytes.len() > max_bytes {
        return Err(UseError::new(
            "office.collaboration.input_too_large",
            format!(
                "The MCP collaboration {label} is {} bytes; the limit is {max_bytes} bytes.",
                bytes.len()
            ),
        )
        .with_suggestion(
            "Use the collaboration CLI with a bounded binary input file for larger updates.",
        ));
    }
    Ok(bytes)
}

async fn run_blocking<T, F>(operation: F) -> UseResult<T>
where
    T: Send + 'static,
    F: FnOnce() -> UseResult<T> + Send + 'static,
{
    tokio::task::spawn_blocking(operation)
        .await
        .map_err(|error| {
            UseError::new(
                "use.office.mcp_task_failed",
                format!("The native Office MCP collaboration task failed: {error}"),
            )
        })?
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn collaboration_inputs_are_closed_and_typed() {
        let schema = schemars::schema_for!(OfficeCollaborationCreateInput);
        let encoded = serde_json::to_string(&schema).unwrap();
        for expected in [
            "artifactId",
            "initialUpdateBase64",
            "document",
            "markdown",
            "spreadsheet",
            "presentation",
            "pdf",
            "human",
            "agent",
            "system",
            "comment",
            "suggest",
            "edit",
        ] {
            assert!(encoded.contains(expected), "missing {expected}");
        }

        let unknown = serde_json::from_value::<OfficeCollaborationEventsInput>(json!({
            "store": "agent.replica",
            "subscribe": true
        }));
        assert!(unknown.is_err());

        let mutation_schema = schemars::schema_for!(OfficeCollaborationMutationInput);
        let encoded = serde_json::to_string(&mutation_schema).unwrap();
        for expected in [
            "markdown-replace",
            "markdown-splice",
            "indexUtf16",
            "deleteUtf16",
            "document-replace-text",
            "expectedMatches",
            "document-set-page-color",
            "pageColor",
            "document-clear-page-color",
            "document-set-track-changes",
            "trackChanges",
            "document-clear-track-changes",
            "document-insert-paragraph",
            "anchorParagraphId",
            "position",
            "before",
            "after",
            "paragraphId",
            "textId",
            "document-delete-paragraph",
            "expectedTextId",
            "expectedText",
            "ifStateVectorBase64",
        ] {
            assert!(encoded.contains(expected), "missing {expected}");
        }
        let unknown_mutation = serde_json::from_value::<OfficeCollaborationMutation>(json!({
            "type": "markdown-splice",
            "indexUtf16": 0,
            "deleteUtf16": 0,
            "insert": "text",
            "force": true
        }));
        assert!(unknown_mutation.is_err());
        assert!(
            serde_json::from_value::<OfficeCollaborationMutation>(json!({
                "type": "document-set-page-color"
            }))
            .is_err()
        );
        assert!(
            serde_json::from_value::<OfficeCollaborationMutation>(json!({
                "type": "document-set-track-changes"
            }))
            .is_err()
        );
        assert!(
            serde_json::from_value::<OfficeCollaborationMutation>(json!({
                "type": "document-clear-page-color",
                "pageColor": "#FFFFFF"
            }))
            .is_err()
        );
        assert!(
            serde_json::from_value::<OfficeCollaborationMutation>(json!({
                "type": "document-insert-paragraph",
                "anchorParagraphId": "00000001",
                "paragraphId": "00000002",
                "textId": "00000003",
                "text": "missing position"
            }))
            .is_err()
        );
        assert!(
            serde_json::from_value::<OfficeCollaborationMutation>(json!({
                "type": "document-delete-paragraph",
                "paragraphId": "00000002",
                "expectedTextId": "00000003",
                "expectedText": "value",
                "force": true
            }))
            .is_err()
        );
    }

    #[test]
    fn collaboration_base64_inputs_are_bounded_before_decoding() {
        let oversized =
            "A".repeat(MAX_MCP_COLLABORATION_UPDATE_BYTES.saturating_add(2) / 3 * 4 + 1);
        let error =
            decode_binary(&oversized, MAX_MCP_COLLABORATION_UPDATE_BYTES, "update").unwrap_err();
        assert_eq!(error.code, "office.collaboration.input_too_large");
    }
}
