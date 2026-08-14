use a3s_use_core::{UseError, UseResult};
use sha2::{Digest, Sha256};
use yrs::updates::decoder::Decode;
use yrs::{ReadTxn, StateVector, Transact};

use super::document::state_vector_sha256;
use super::persistence::{LoadedStore, OperationRecord};
use super::{
    NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationOperationKind,
    MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES,
};

const MAX_IDENTIFIER_CHARACTERS: usize = 256;
const MAX_NAMESPACE_CHARACTERS: usize = 256;
const MAX_YJS_CLIENT_ID: u64 = (1_u64 << 53) - 1;

pub(super) fn validate_expected_replica(
    manifest: &NativeOfficeCollaborationManifest,
    actor_id: &str,
    mode: NativeOfficeCollaborationMode,
    artifact_id: &str,
    kind: NativeOfficeCollaborationArtifactKind,
) -> UseResult<()> {
    if manifest.actor_id != actor_id {
        return Err(collaboration_error(
            "office.collaboration.actor_mismatch",
            format!(
                "The replica belongs to actor '{}', not '{}'.",
                manifest.actor_id, actor_id
            ),
        )
        .with_detail("expectedActorId", actor_id.to_owned())
        .with_detail("actualActorId", manifest.actor_id.clone()));
    }
    if manifest.mode != mode {
        return Err(collaboration_error(
            "office.collaboration.mode_mismatch",
            format!(
                "The replica uses '{}' mode, not '{}'.",
                manifest.mode.as_str(),
                mode.as_str()
            ),
        )
        .with_detail("expectedMode", mode.as_str())
        .with_detail("actualMode", manifest.mode.as_str()));
    }
    if manifest.artifact_id != artifact_id {
        return Err(collaboration_error(
            "office.collaboration.artifact_mismatch",
            format!(
                "The replica belongs to artifact '{}', not '{}'.",
                manifest.artifact_id, artifact_id
            ),
        )
        .with_detail("expectedArtifactId", artifact_id.to_owned())
        .with_detail("actualArtifactId", manifest.artifact_id.clone()));
    }
    if manifest.kind != kind {
        return Err(collaboration_error(
            "office.collaboration.kind_mismatch",
            format!(
                "The replica contains '{}' content, not '{}'.",
                manifest.kind.as_str(),
                kind.as_str()
            ),
        )
        .with_detail("expectedKind", kind.as_str())
        .with_detail("actualKind", manifest.kind.as_str()));
    }
    Ok(())
}

pub(super) fn assert_state_vector_precondition(
    loaded: &LoadedStore,
    precondition: Option<&[u8]>,
) -> UseResult<()> {
    let Some(precondition) = precondition else {
        return Ok(());
    };
    let expected = decode_state_vector(precondition)?;
    let actual = loaded.doc.transact().state_vector();
    if expected == actual {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.stale_state",
        "The collaboration replica changed after the supplied state-vector precondition.",
    )
    .with_suggestion("Inspect the replica and decide whether to rebase the operation.")
    .with_detail("expectedStateVectorSha256", state_vector_sha256(&expected))
    .with_detail("actualStateVectorSha256", state_vector_sha256(&actual)))
}

pub(super) fn assert_operation_replay(
    existing: &OperationRecord,
    payload_sha256: &str,
) -> UseResult<()> {
    if existing.payload_sha256 == payload_sha256 {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.operation_conflict",
        format!(
            "Operation ID '{}' was already used for a different collaboration request.",
            existing.operation_id
        ),
    )
    .with_suggestion("Use a new stable operation ID after inspecting the shared state.")
    .with_detail("operationId", existing.operation_id.clone())
    .with_detail("existingPayloadSha256", existing.payload_sha256.clone())
    .with_detail("receivedPayloadSha256", payload_sha256.to_owned()))
}

pub(super) fn validate_update_size(update: &[u8]) -> UseResult<()> {
    if update.len() <= MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.update_too_large",
        format!(
            "The Yjs v1 update is {} bytes; the limit is {} bytes.",
            update.len(),
            MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES
        ),
    )
    .with_detail("bytes", update.len() as u64)
    .with_detail(
        "maxBytes",
        MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES as u64,
    ))
}

pub(super) fn decode_state_vector(bytes: &[u8]) -> UseResult<StateVector> {
    if bytes.len() > MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES {
        return Err(collaboration_error(
            "office.collaboration.state_vector_too_large",
            format!(
                "The Yjs state vector is {} bytes; the limit is {} bytes.",
                bytes.len(),
                MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES
            ),
        ));
    }
    StateVector::decode_v1(bytes).map_err(|error| {
        collaboration_error(
            "office.collaboration.state_vector_invalid",
            format!("The input is not a valid Yjs v1 state vector: {error}"),
        )
    })
}

pub(super) fn normalized_identifier(value: &str, label: &str) -> UseResult<String> {
    let value = value.trim();
    let length = value.chars().count();
    if (1..=MAX_IDENTIFIER_CHARACTERS).contains(&length) {
        return Ok(value.to_owned());
    }
    Err(collaboration_error(
        "office.collaboration.identifier_invalid",
        format!("The {label} must contain between 1 and 256 characters."),
    ))
}

pub(super) fn normalized_namespace(value: &str) -> UseResult<String> {
    let value = value.trim();
    if value.is_empty()
        || value.chars().count() > MAX_NAMESPACE_CHARACTERS
        || value.starts_with('.')
        || value.ends_with('.')
        || !value
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || "._-".contains(character))
    {
        return Err(collaboration_error(
            "office.collaboration.namespace_invalid",
            "The collaboration namespace may contain only letters, digits, dots, underscores, and hyphens.",
        ));
    }
    Ok(value.to_owned())
}

pub(super) fn validate_client_id(client_id: u64) -> UseResult<()> {
    if (1..=MAX_YJS_CLIENT_ID).contains(&client_id) {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.client_id_invalid",
        "The native collaboration client ID must be a non-zero 53-bit integer.",
    )
    .with_detail("clientId", client_id))
}

pub(super) fn creation_payload_sha256(
    manifest: &NativeOfficeCollaborationManifest,
    update: Option<&[u8]>,
    operation_id: &str,
    operation_kind: NativeOfficeCollaborationOperationKind,
) -> UseResult<String> {
    let payload = serde_json::json!({
        "action": operation_kind,
        "actorId": manifest.actor_id,
        "actorKind": manifest.actor_kind,
        "artifactId": manifest.artifact_id,
        "clientId": manifest.client_id,
        "kind": manifest.kind,
        "mode": manifest.mode,
        "namespace": manifest.namespace,
        "operationId": operation_id,
        "updateSha256": update.map(sha256_hex),
    });
    json_sha256(&payload)
}

pub(super) fn operation_payload_sha256(
    kind: NativeOfficeCollaborationOperationKind,
    manifest: &NativeOfficeCollaborationManifest,
    operation_id: &str,
    update_sha256: Option<&str>,
    precondition: Option<&[u8]>,
) -> UseResult<String> {
    let precondition_sha256 = precondition
        .map(decode_state_vector)
        .transpose()?
        .as_ref()
        .map(state_vector_sha256);
    let payload = serde_json::json!({
        "actorId": manifest.actor_id,
        "artifactId": manifest.artifact_id,
        "artifactKind": manifest.kind,
        "kind": kind,
        "mode": manifest.mode,
        "operationId": operation_id,
        "preconditionStateVectorSha256": precondition_sha256,
        "updateSha256": update_sha256,
    });
    json_sha256(&payload)
}

fn json_sha256(value: &serde_json::Value) -> UseResult<String> {
    let bytes = serde_json::to_vec(value).map_err(|error| {
        collaboration_error(
            "office.collaboration.operation_invalid",
            format!("Failed to encode the collaboration operation: {error}"),
        )
    })?;
    Ok(sha256_hex(&bytes))
}

pub(crate) fn sha256_hex(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}

pub(crate) fn collaboration_error(code: &'static str, message: impl Into<String>) -> UseError {
    UseError::new(code, message)
}
