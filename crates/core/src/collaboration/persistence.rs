use std::collections::BTreeMap;
use std::fs::{self, File, OpenOptions};
use std::io::{ErrorKind, Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};

use a3s_use_core::UseResult;
use fs2::FileExt;
use serde::{Deserialize, Serialize};
use yrs::updates::decoder::Decode;
use yrs::{Doc, ReadTxn, StateVector, Transact, Update};

use super::document::{
    canonical_replay_document, canonical_state_vector, inspect_document, new_replica_document,
};
use super::{
    collaboration_error, sha256_hex, validate_client_id, NativeOfficeCollaborationActorKind,
    NativeOfficeCollaborationArtifactKind, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMode, NativeOfficeCollaborationOperationKind,
    NativeOfficeCollaborationOrigin, MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES,
    MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES, NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION, NATIVE_OFFICE_COLLABORATION_STORE_FORMAT,
    NATIVE_OFFICE_COLLABORATION_STORE_SCHEMA_VERSION,
};

mod paths;
use paths::*;

const MANIFEST_FILE: &str = "manifest.json";
const LOCK_FILE: &str = ".lock";
const CHECKPOINT_DIRECTORY: &str = "checkpoints";
const UPDATE_DIRECTORY: &str = "updates";
const OPERATION_DIRECTORY: &str = "operations";
const ENTRY_MAGIC: &[u8; 8] = b"A3SOCU1\n";
const MAX_JSON_BYTES: u64 = 64 * 1024;
const MAX_ENTRY_BYTES: u64 = MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES as u64
    + MAX_JSON_BYTES
    + ENTRY_MAGIC.len() as u64
    + 4;

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CheckpointCommit {
    schema_version: u32,
    sequence: u64,
    update_sha256: String,
    state_vector_sha256: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OperationRecord {
    pub schema_version: u32,
    pub protocol: String,
    pub operation_id: String,
    pub actor_id: String,
    pub actor_kind: NativeOfficeCollaborationActorKind,
    pub mode: NativeOfficeCollaborationMode,
    pub kind: NativeOfficeCollaborationOperationKind,
    pub artifact_id: String,
    pub artifact_kind: NativeOfficeCollaborationArtifactKind,
    pub payload_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub update_sha256: Option<String>,
    pub before_state_vector_sha256: String,
    pub after_state_vector_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sequence: Option<u64>,
    pub state_changed: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub origin: Option<NativeOfficeCollaborationOrigin>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub host_authorization: Option<HostAuthorizationRecord>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(super) struct HostAuthorizationRecord {
    pub version: u32,
    pub actor_name: String,
}

#[derive(Debug, Clone)]
pub(super) struct UpdateEntry {
    pub path: PathBuf,
    pub sequence: u64,
    pub update_bytes: u64,
    pub update: Vec<u8>,
    pub operation: OperationRecord,
}

pub(super) struct LoadedStore {
    pub manifest: NativeOfficeCollaborationManifest,
    pub doc: Doc,
    pub checkpoint_update: Vec<u8>,
    pub checkpoint_sequence: u64,
    pub current_sequence: u64,
    pub next_sequence: u64,
    pub update_entries: Vec<UpdateEntry>,
    pub update_bytes: u64,
    covered_update_entries: Vec<UpdateEntry>,
    operations: BTreeMap<String, OperationRecord>,
}

impl LoadedStore {
    pub fn find_operation(&self, operation_id: &str) -> UseResult<Option<OperationRecord>> {
        Ok(self.operations.get(operation_id).cloned())
    }

    pub fn operation_count(&self) -> UseResult<usize> {
        Ok(self.operations.len())
    }
}

pub(super) struct StoreLock {
    file: File,
}

impl StoreLock {
    pub fn acquire(root: &Path) -> UseResult<Self> {
        let path = root.join(LOCK_FILE);
        // Revalidate immediately before opening so a swapped symlink is not
        // followed between replica discovery and each cross-process lock.
        assert_real_file(&path, "replica lock")?;
        let file = OpenOptions::new()
            .read(true)
            .write(true)
            .create(false)
            .open(&path)
            .map_err(|error| io_error("lock", &path, error))?;
        file.lock_exclusive().map_err(|error| {
            collaboration_error(
                "office.collaboration.lock_failed",
                format!(
                    "Failed to lock collaboration replica '{}': {error}",
                    root.display()
                ),
            )
            .with_detail("store", root.display().to_string())
        })?;
        Ok(Self { file })
    }
}

impl Drop for StoreLock {
    fn drop(&mut self) {
        let _ = FileExt::unlock(&self.file);
    }
}

pub(super) fn create_store(
    requested_root: &Path,
    manifest: &NativeOfficeCollaborationManifest,
    doc: &Doc,
    operation: &OperationRecord,
) -> UseResult<PathBuf> {
    let root = normalized_new_root(requested_root)?;
    let parent = root.parent().expect("validated new root has a parent");
    let temporary = tempfile::Builder::new()
        .prefix(".a3s-office-collaboration-")
        .tempdir_in(parent)
        .map_err(|error| io_error("create staged replica", parent, error))?;
    let staged_root = temporary.path();
    fs::create_dir(staged_root.join(CHECKPOINT_DIRECTORY))
        .map_err(|error| io_error("create checkpoint directory", staged_root, error))?;
    fs::create_dir(staged_root.join(UPDATE_DIRECTORY))
        .map_err(|error| io_error("create update directory", staged_root, error))?;
    fs::create_dir(staged_root.join(OPERATION_DIRECTORY))
        .map_err(|error| io_error("create operation directory", staged_root, error))?;
    write_new_bytes(
        &staged_root.join(MANIFEST_FILE),
        &json_bytes(manifest, "manifest")?,
    )?;
    write_new_bytes(&staged_root.join(LOCK_FILE), b"")?;
    write_checkpoint(staged_root, 0, doc)?;
    write_archived_operation(staged_root, operation)?;
    sync_directory_best_effort(staged_root)?;

    let staged_path = temporary.keep();
    match fs::rename(&staged_path, &root) {
        Ok(()) => {}
        Err(error) => {
            let _ = fs::remove_dir_all(&staged_path);
            if error.kind() == ErrorKind::AlreadyExists || root.exists() {
                return Err(collaboration_error(
                    "office.collaboration.store_exists",
                    format!(
                        "Collaboration replica '{}' already exists; refusing to replace it.",
                        root.display()
                    ),
                )
                .with_detail("store", root.display().to_string()));
            }
            return Err(io_error("publish collaboration replica", &root, error));
        }
    }
    sync_directory_best_effort(parent)?;
    fs::canonicalize(&root).map_err(|error| io_error("resolve collaboration replica", &root, error))
}

pub(super) fn open_store(requested_root: &Path) -> UseResult<PathBuf> {
    let root = fs::canonicalize(requested_root).map_err(|error| {
        collaboration_error(
            "office.collaboration.store_unavailable",
            format!(
                "Collaboration replica '{}' is unavailable: {error}",
                requested_root.display()
            ),
        )
        .with_detail("store", requested_root.display().to_string())
    })?;
    let metadata = fs::symlink_metadata(&root)
        .map_err(|error| io_error("inspect collaboration replica", &root, error))?;
    if !metadata.is_dir() || metadata.file_type().is_symlink() {
        return Err(collaboration_error(
            "office.collaboration.store_invalid",
            format!(
                "Collaboration replica '{}' must be a real directory.",
                root.display()
            ),
        ));
    }
    assert_real_file(&root.join(MANIFEST_FILE), "replica manifest")?;
    assert_real_file(&root.join(LOCK_FILE), "replica lock")?;
    for name in [CHECKPOINT_DIRECTORY, UPDATE_DIRECTORY, OPERATION_DIRECTORY] {
        assert_real_directory(&root.join(name), "replica directory")?;
    }
    Ok(root)
}

pub(super) fn load_store(root: &Path) -> UseResult<LoadedStore> {
    let manifest: NativeOfficeCollaborationManifest =
        read_json(&root.join(MANIFEST_FILE), "manifest")?;
    validate_manifest(&manifest)?;
    let checkpoint =
        load_latest_checkpoint(root, manifest.client_id, &manifest.namespace, manifest.kind)?;
    let LoadedCheckpoint {
        sequence: checkpoint_sequence,
        doc: checkpoint_doc,
        update: checkpoint_update,
    } = checkpoint;
    let mut doc = checkpoint_doc;
    let mut operations = load_archived_operations(root, &manifest)?;
    let mut current_sequence = checkpoint_sequence;
    let mut expected_sequence = checkpoint_sequence.saturating_add(1);
    let mut update_entries = Vec::new();
    let mut covered_update_entries = Vec::new();
    let mut update_bytes = 0_u64;
    let entries = load_update_entries(root, &manifest)?;
    for entry in entries {
        insert_operation(&mut operations, &entry.operation)?;
        if entry.sequence <= checkpoint_sequence {
            covered_update_entries.push(entry);
            continue;
        }
        if entry.sequence != expected_sequence {
            return Err(collaboration_error(
                "office.collaboration.log_incomplete",
                format!(
                    "Collaboration update sequence {} is missing after checkpoint {}.",
                    expected_sequence, checkpoint_sequence
                ),
            )
            .with_suggestion(
                "Restore the missing update entry from durable storage before retrying.",
            )
            .with_detail("expectedSequence", expected_sequence)
            .with_detail("actualSequence", entry.sequence));
        }
        let update = Update::decode_v1(&entry.update).map_err(|error| {
            collaboration_error(
                "office.collaboration.log_corrupt",
                format!(
                    "Collaboration update entry '{}' is invalid: {error}",
                    entry.path.display()
                ),
            )
        })?;
        doc.transact_mut().apply_update(update).map_err(|error| {
            collaboration_error(
                "office.collaboration.log_corrupt",
                format!(
                    "Collaboration update entry '{}' cannot be replayed: {error}",
                    entry.path.display()
                ),
            )
        })?;
        update_bytes = update_bytes.saturating_add(entry.update_bytes);
        current_sequence = entry.sequence;
        expected_sequence = expected_sequence.saturating_add(1);
        update_entries.push(entry);
    }
    if doc.transact().has_missing_updates() {
        doc = canonical_replay_document(&doc, &manifest)?;
    }
    inspect_document(&doc, &manifest)?;
    Ok(LoadedStore {
        manifest,
        doc,
        checkpoint_update,
        checkpoint_sequence,
        current_sequence,
        next_sequence: current_sequence.checked_add(1).ok_or_else(|| {
            collaboration_error(
                "office.collaboration.sequence_exhausted",
                "The collaboration replica exhausted its update sequence space.",
            )
        })?,
        update_entries,
        update_bytes,
        covered_update_entries,
        operations,
    })
}

pub(super) fn write_update_entry(
    root: &Path,
    sequence: u64,
    update: &[u8],
    operation: &OperationRecord,
) -> UseResult<()> {
    if operation.sequence != Some(sequence)
        || operation.update_sha256.as_deref() != Some(&sha256_hex(update))
    {
        return Err(collaboration_error(
            "office.collaboration.operation_invalid",
            "The collaboration operation receipt does not match its update entry.",
        ));
    }
    let header = json_bytes(operation, "operation receipt")?;
    if header.len() as u64 > MAX_JSON_BYTES {
        return Err(collaboration_error(
            "office.collaboration.operation_invalid",
            "The collaboration operation receipt exceeds the bounded JSON size.",
        ));
    }
    let update_hash = sha256_hex(update);
    let path = root
        .join(UPDATE_DIRECTORY)
        .join(format!("{sequence:020}-{update_hash}.entry"));
    let mut bytes = Vec::with_capacity(ENTRY_MAGIC.len() + 4 + header.len() + update.len());
    bytes.extend_from_slice(ENTRY_MAGIC);
    bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&header);
    bytes.extend_from_slice(update);
    write_new_bytes(&path, &bytes)?;
    sync_directory_best_effort(path.parent().expect("entry has parent"))
}

pub(super) fn write_archived_operation(root: &Path, operation: &OperationRecord) -> UseResult<()> {
    let path = operation_path(root, &operation.operation_id);
    let bytes = json_bytes(operation, "operation receipt")?;
    if path.exists() {
        let existing: OperationRecord = read_json(&path, "operation receipt")?;
        if existing == *operation {
            return Ok(());
        }
        return Err(collaboration_error(
            "office.collaboration.operation_conflict",
            format!(
                "Operation ID '{}' already has a different durable receipt.",
                operation.operation_id
            ),
        )
        .with_detail("operationId", operation.operation_id.clone()));
    }
    write_new_bytes(&path, &bytes)?;
    sync_directory_best_effort(path.parent().expect("operation has parent"))
}

pub(super) fn write_checkpoint(root: &Path, sequence: u64, doc: &Doc) -> UseResult<()> {
    let transaction = doc.transact();
    let update = transaction.encode_state_as_update_v1(&StateVector::default());
    if update.len() > MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES {
        return Err(collaboration_error(
            "office.collaboration.checkpoint_too_large",
            format!(
                "The collaboration checkpoint is {} bytes; the limit is {} bytes.",
                update.len(),
                MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES
            ),
        ));
    }
    let state_vector = canonical_state_vector(&transaction.state_vector());
    drop(transaction);
    let state_hash = sha256_hex(&update);
    let stem = format!("{sequence:020}-{state_hash}");
    let directory = root.join(CHECKPOINT_DIRECTORY);
    write_new_or_same(&directory.join(format!("{stem}.update")), &update)?;
    write_new_or_same(
        &directory.join(format!("{stem}.state-vector")),
        &state_vector,
    )?;
    let commit = CheckpointCommit {
        schema_version: 1,
        sequence,
        update_sha256: state_hash,
        state_vector_sha256: sha256_hex(&state_vector),
    };
    write_new_or_same(
        &directory.join(format!("{stem}.complete")),
        &json_bytes(&commit, "checkpoint commit")?,
    )?;
    sync_directory_best_effort(&directory)
}

pub(super) fn compact(
    root: &Path,
    loaded: &LoadedStore,
    checkpoint_sequence: u64,
) -> UseResult<usize> {
    let mut compacted = 0;
    for entry in loaded
        .covered_update_entries
        .iter()
        .chain(loaded.update_entries.iter())
        .filter(|entry| entry.sequence <= checkpoint_sequence)
    {
        write_archived_operation(root, &entry.operation)?;
        remove_scoped_file(root, UPDATE_DIRECTORY, &entry.path)?;
        compacted += 1;
    }
    let checkpoint_directory = root.join(CHECKPOINT_DIRECTORY);
    for item in read_real_files(&checkpoint_directory)? {
        let Some((sequence, _hash, _extension)) = parse_checkpoint_name(&item)? else {
            continue;
        };
        if sequence < checkpoint_sequence {
            remove_scoped_file(root, CHECKPOINT_DIRECTORY, &item)?;
        }
    }
    sync_directory_best_effort(&root.join(UPDATE_DIRECTORY))?;
    sync_directory_best_effort(&checkpoint_directory)?;
    Ok(compacted)
}

struct LoadedCheckpoint {
    sequence: u64,
    doc: Doc,
    update: Vec<u8>,
}

fn load_latest_checkpoint(
    root: &Path,
    client_id: u64,
    namespace: &str,
    artifact_kind: NativeOfficeCollaborationArtifactKind,
) -> UseResult<LoadedCheckpoint> {
    let directory = root.join(CHECKPOINT_DIRECTORY);
    let mut candidates = BTreeMap::<u64, (String, PathBuf)>::new();
    for path in read_real_files(&directory)? {
        let Some((sequence, hash, extension)) = parse_checkpoint_name(&path)? else {
            return Err(log_corrupt(&path, "has an unsupported checkpoint filename"));
        };
        // Update/vector files without a final marker are an interrupted
        // checkpoint publication. They are ignored until an identical retry
        // publishes the immutable marker last.
        if extension != "complete" {
            continue;
        }
        if candidates.insert(sequence, (hash, path.clone())).is_some() {
            return Err(collaboration_error(
                "office.collaboration.checkpoint_ambiguous",
                format!("More than one committed checkpoint exists at sequence {sequence}."),
            ));
        }
    }
    let Some((sequence, (expected_hash, commit_path))) = candidates.pop_last() else {
        return Err(collaboration_error(
            "office.collaboration.checkpoint_missing",
            "The collaboration replica has no complete checkpoint.",
        ));
    };
    let commit: CheckpointCommit = read_json(&commit_path, "checkpoint commit")?;
    if commit.schema_version != 1
        || commit.sequence != sequence
        || commit.update_sha256 != expected_hash
        || commit.state_vector_sha256.len() != 64
    {
        return Err(collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint commit '{}' is invalid.",
                commit_path.display()
            ),
        ));
    }
    let update_path = directory.join(format!("{sequence:020}-{expected_hash}.update"));
    let vector_path = directory.join(format!("{sequence:020}-{expected_hash}.state-vector"));
    let update_bytes = read_bounded_file(
        &update_path,
        MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES as u64,
        "checkpoint",
    )?;
    if sha256_hex(&update_bytes) != expected_hash {
        return Err(collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint '{}' does not match its SHA-256 name.",
                update_path.display()
            ),
        ));
    }
    let persisted_vector = read_bounded_file(
        &vector_path,
        MAX_NATIVE_OFFICE_COLLABORATION_STATE_VECTOR_BYTES as u64,
        "state vector",
    )?;
    if sha256_hex(&persisted_vector) != commit.state_vector_sha256 {
        return Err(collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint state vector '{}' does not match its commit.",
                vector_path.display()
            ),
        ));
    }
    StateVector::decode_v1(&persisted_vector).map_err(|error| {
        collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint state vector '{}' is invalid: {error}",
                vector_path.display()
            ),
        )
    })?;
    let update = Update::decode_v1(&update_bytes).map_err(|error| {
        collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint '{}' is invalid: {error}",
                update_path.display()
            ),
        )
    })?;
    let doc = new_replica_document(client_id, namespace, artifact_kind);
    doc.transact_mut().apply_update(update).map_err(|error| {
        collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint '{}' cannot be loaded: {error}",
                update_path.display()
            ),
        )
    })?;
    let actual_vector = canonical_state_vector(&doc.transact().state_vector());
    if actual_vector != persisted_vector {
        return Err(collaboration_error(
            "office.collaboration.checkpoint_corrupt",
            format!(
                "Collaboration checkpoint '{}' has a mismatched state vector.",
                update_path.display()
            ),
        ));
    }
    Ok(LoadedCheckpoint {
        sequence,
        doc,
        update: update_bytes,
    })
}

fn load_update_entries(
    root: &Path,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<Vec<UpdateEntry>> {
    let directory = root.join(UPDATE_DIRECTORY);
    let mut entries = Vec::new();
    for path in read_real_files(&directory)? {
        if path.extension().and_then(|value| value.to_str()) != Some("entry") {
            return Err(log_corrupt(&path, "has an unsupported update-log filename"));
        }
        let (sequence, expected_update_hash) = parse_update_entry_name(&path)?;
        let mut file = open_bounded_file(&path, MAX_ENTRY_BYTES, "update entry")?;
        let mut magic = [0_u8; ENTRY_MAGIC.len()];
        file.read_exact(&mut magic)
            .map_err(|error| io_error("read update entry", &path, error))?;
        if &magic != ENTRY_MAGIC {
            return Err(log_corrupt(&path, "has an invalid entry header"));
        }
        let mut header_length = [0_u8; 4];
        file.read_exact(&mut header_length)
            .map_err(|error| io_error("read update entry", &path, error))?;
        let header_length = u32::from_le_bytes(header_length) as u64;
        if header_length == 0 || header_length > MAX_JSON_BYTES {
            return Err(log_corrupt(
                &path,
                "has an invalid operation receipt length",
            ));
        }
        let total_length = file
            .metadata()
            .map_err(|error| io_error("inspect update entry", &path, error))?
            .len();
        let update_offset = ENTRY_MAGIC.len() as u64 + 4 + header_length;
        if update_offset > total_length {
            return Err(log_corrupt(&path, "is truncated before its update payload"));
        }
        let update_length = total_length - update_offset;
        if update_length > MAX_NATIVE_OFFICE_COLLABORATION_UPDATE_BYTES as u64 {
            return Err(log_corrupt(&path, "contains an oversized update payload"));
        }
        file.seek(SeekFrom::Start(ENTRY_MAGIC.len() as u64 + 4))
            .map_err(|error| io_error("seek update entry", &path, error))?;
        let mut header = vec![0_u8; header_length as usize];
        file.read_exact(&mut header)
            .map_err(|error| io_error("read update entry", &path, error))?;
        let operation: OperationRecord = serde_json::from_slice(&header).map_err(|error| {
            log_corrupt(
                &path,
                format!("contains an invalid operation receipt: {error}"),
            )
        })?;
        let mut update = vec![0_u8; update_length as usize];
        file.read_exact(&mut update)
            .map_err(|error| io_error("read update entry", &path, error))?;
        let actual_hash = sha256_hex(&update);
        validate_operation(&operation, manifest)?;
        if operation.sequence != Some(sequence)
            || operation.update_sha256.as_deref() != Some(actual_hash.as_str())
            || expected_update_hash != actual_hash
        {
            return Err(log_corrupt(
                &path,
                "does not match its sequence, update hash, and operation receipt",
            ));
        }
        entries.push(UpdateEntry {
            path,
            sequence,
            update_bytes: update_length,
            update,
            operation,
        });
    }
    entries.sort_by_key(|entry| entry.sequence);
    for pair in entries.windows(2) {
        if pair[0].sequence == pair[1].sequence {
            return Err(collaboration_error(
                "office.collaboration.log_ambiguous",
                format!(
                    "More than one collaboration update exists at sequence {}.",
                    pair[0].sequence
                ),
            ));
        }
    }
    Ok(entries)
}

fn load_archived_operations(
    root: &Path,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<BTreeMap<String, OperationRecord>> {
    let mut operations = BTreeMap::new();
    for path in read_real_files(&root.join(OPERATION_DIRECTORY))? {
        if path.extension().and_then(|value| value.to_str()) != Some("json") {
            return Err(log_corrupt(
                &path,
                "has an unsupported operation-receipt filename",
            ));
        }
        let operation: OperationRecord = read_json(&path, "operation receipt")?;
        validate_operation(&operation, manifest)?;
        let expected_name = format!("{}.json", sha256_hex(operation.operation_id.as_bytes()));
        if path.file_name().and_then(|value| value.to_str()) != Some(&expected_name) {
            return Err(log_corrupt(
                &path,
                "does not match the SHA-256 of its operation ID",
            ));
        }
        insert_operation(&mut operations, &operation)?;
    }
    Ok(operations)
}

fn insert_operation(
    operations: &mut BTreeMap<String, OperationRecord>,
    operation: &OperationRecord,
) -> UseResult<()> {
    match operations.get(&operation.operation_id) {
        Some(existing) if existing == operation => Ok(()),
        Some(existing) => Err(collaboration_error(
            "office.collaboration.operation_conflict",
            format!(
                "Operation ID '{}' has conflicting durable receipts.",
                operation.operation_id
            ),
        )
        .with_detail("operationId", operation.operation_id.clone())
        .with_detail("existingPayloadSha256", existing.payload_sha256.clone())
        .with_detail("receivedPayloadSha256", operation.payload_sha256.clone())),
        None => {
            operations.insert(operation.operation_id.clone(), operation.clone());
            Ok(())
        }
    }
}

fn validate_manifest(manifest: &NativeOfficeCollaborationManifest) -> UseResult<()> {
    if manifest.format != NATIVE_OFFICE_COLLABORATION_STORE_FORMAT
        || manifest.schema_version != NATIVE_OFFICE_COLLABORATION_STORE_SCHEMA_VERSION
        || manifest.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
        || manifest.protocol_version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
    {
        return Err(collaboration_error(
            "office.collaboration.store_version_unsupported",
            "The collaboration replica manifest uses an unsupported format or protocol version.",
        ));
    }
    validate_client_id(manifest.client_id)?;
    if manifest.artifact_id.trim().is_empty()
        || manifest.artifact_id.chars().count() > 256
        || manifest.actor_id.trim().is_empty()
        || manifest.actor_id.chars().count() > 256
        || manifest.namespace.trim().is_empty()
        || manifest.namespace.chars().count() > 256
        || manifest.namespace.starts_with('.')
        || manifest.namespace.ends_with('.')
        || !manifest
            .namespace
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || "._-".contains(character))
    {
        return Err(collaboration_error(
            "office.collaboration.store_invalid",
            "The collaboration replica manifest contains an empty identity field.",
        ));
    }
    Ok(())
}

fn validate_operation(
    operation: &OperationRecord,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<()> {
    let identity_valid = match &operation.host_authorization {
        None => {
            operation.actor_id == manifest.actor_id
                && operation.actor_kind == manifest.actor_kind
                && operation.mode == manifest.mode
        }
        Some(authorization) => {
            let expected_origin_kind = match operation.actor_kind {
                NativeOfficeCollaborationActorKind::Human => {
                    super::NativeOfficeCollaborationOriginKind::Editor
                }
                NativeOfficeCollaborationActorKind::Agent => {
                    super::NativeOfficeCollaborationOriginKind::Agent
                }
                NativeOfficeCollaborationActorKind::System => {
                    super::NativeOfficeCollaborationOriginKind::System
                }
            };
            authorization.version == 1
                && !authorization.actor_name.trim().is_empty()
                && authorization.actor_name == authorization.actor_name.trim()
                && authorization.actor_name.chars().count() <= 256
                && !operation.actor_id.trim().is_empty()
                && operation.actor_id == operation.actor_id.trim()
                && operation.actor_id.chars().count() <= 256
                && operation.kind == NativeOfficeCollaborationOperationKind::Synchronize
                && matches!(
                    operation.mode,
                    NativeOfficeCollaborationMode::Edit
                        | NativeOfficeCollaborationMode::Comment
                        | NativeOfficeCollaborationMode::Suggest
                )
                && operation.origin.as_ref().is_none_or(|origin| {
                    origin.actor_id.as_deref() == Some(operation.actor_id.as_str())
                        && origin.operation_id.as_deref() == Some(operation.operation_id.as_str())
                        && origin.kind == expected_origin_kind
                })
        }
    };
    if operation.schema_version != 1
        || operation.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
        || operation.operation_id.trim().is_empty()
        || !identity_valid
        || operation.artifact_id != manifest.artifact_id
        || operation.artifact_kind != manifest.kind
        || operation.payload_sha256.len() != 64
        || operation.before_state_vector_sha256.len() != 64
        || operation.after_state_vector_sha256.len() != 64
    {
        return Err(collaboration_error(
            "office.collaboration.operation_invalid",
            format!(
                "Operation receipt '{}' does not match the collaboration replica.",
                operation.operation_id
            ),
        ));
    }
    Ok(())
}
