use a3s_use_core::UseResult;
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Array, ClientID, Doc, Map, OffsetKind, Options, Out, ReadTxn, StateVector, Transact,
    Update,
};

use super::{
    collaboration_error, sha256_hex, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMetadata,
    NativeOfficeCollaborationStateVectorEntry, NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};

pub(super) struct DocumentInspection {
    pub metadata: Option<NativeOfficeCollaborationMetadata>,
    pub metadata_fields_present: usize,
    pub metadata_complete: bool,
    pub bootstrap_initializer_count: Option<u32>,
    pub bootstrap_valid: Option<bool>,
    pub root_names: Vec<String>,
    pub state_sha256: String,
    pub pending_updates: bool,
}

pub(super) fn new_replica_document(
    client_id: u64,
    namespace: &str,
    kind: NativeOfficeCollaborationArtifactKind,
) -> Doc {
    let mut options = Options::with_client_id(ClientID::new(client_id));
    // Browser Yjs indexes text by UTF-16 code units. Matching that convention
    // is essential for native typed text operations around astral characters.
    options.offset_kind = OffsetKind::Utf16;
    let doc = Doc::with_options(options);
    // Named Yjs roots do not carry a standalone root-type declaration in an
    // update. Register the protocol roots before replay so Yrs interprets the
    // incoming named branches with the same types used by browser clients.
    doc.get_or_insert_map(format!("{namespace}.metadata"));
    doc.get_or_insert_array(format!("{namespace}.bootstrap.initializers"));
    register_format_roots(&doc, namespace, kind);
    doc
}

fn register_format_roots(doc: &Doc, namespace: &str, kind: NativeOfficeCollaborationArtifactKind) {
    let root = |suffix: &str| format!("{namespace}.{suffix}");
    match kind {
        NativeOfficeCollaborationArtifactKind::Markdown => {
            doc.get_or_insert_text(root("markdown.source"));
        }
        NativeOfficeCollaborationArtifactKind::Document => {
            doc.get_or_insert_xml_fragment(root("document.content"));
            for suffix in [
                "document.options",
                "document.comments",
                "document.bibliography",
                "document.bibliography.sources",
            ] {
                doc.get_or_insert_map(root(suffix));
            }
            for suffix in [
                "document.comment-order",
                "document.bibliography.source-order",
                "document.record-claims",
            ] {
                doc.get_or_insert_array(root(suffix));
            }
        }
        NativeOfficeCollaborationArtifactKind::Spreadsheet => {
            for suffix in [
                "spreadsheet.options",
                "spreadsheet.sheets",
                "spreadsheet.named-ranges",
                "spreadsheet.print-areas",
                "spreadsheet.print-titles",
                "spreadsheet.page-breaks",
                "spreadsheet.page-setups",
            ] {
                doc.get_or_insert_map(root(suffix));
            }
            for suffix in [
                "spreadsheet.sheet-order",
                "spreadsheet.named-range-order",
                "spreadsheet.record-claims",
            ] {
                doc.get_or_insert_array(root(suffix));
            }
        }
        NativeOfficeCollaborationArtifactKind::Presentation => {
            for suffix in [
                "presentation.options",
                "presentation.slides",
                "presentation.masters",
                "presentation.layouts",
            ] {
                doc.get_or_insert_map(root(suffix));
            }
            for suffix in [
                "presentation.slide-order",
                "presentation.master-order",
                "presentation.layout-order",
            ] {
                doc.get_or_insert_array(root(suffix));
            }
        }
        NativeOfficeCollaborationArtifactKind::Pdf => {
            doc.get_or_insert_map(root("pdf.source"));
            doc.get_or_insert_array(root("pdf.source-identities"));
            doc.get_or_insert_array(root("pdf.record-claims"));
            for collection in [
                "annotations",
                "form-values",
                "signature-placements",
                "redaction-proposals",
                "page-operations",
                "review-decisions",
            ] {
                doc.get_or_insert_map(root(&format!("pdf.{collection}.presence")));
                doc.get_or_insert_map(root(&format!("pdf.{collection}.fields")));
                doc.get_or_insert_array(root(&format!("pdf.{collection}.order")));
            }
        }
    }
}

pub(super) fn validate_and_apply_update(
    doc: &Doc,
    bytes: &[u8],
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<()> {
    let update = Update::decode_v1(bytes).map_err(|error| {
        collaboration_error(
            "office.collaboration.update_invalid",
            format!("The input is not a valid Yjs v1 update: {error}"),
        )
    })?;
    doc.transact_mut().apply_update(update).map_err(|error| {
        collaboration_error(
            "office.collaboration.update_invalid",
            format!("The Yjs v1 update could not be integrated: {error}"),
        )
    })?;
    inspect_document(doc, manifest)?;
    Ok(())
}

pub(super) fn inspect_document(
    doc: &Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<DocumentInspection> {
    let transaction = doc.transact();
    let mut root_names = transaction
        .root_refs()
        .map(|(name, _)| name.to_owned())
        .collect::<Vec<_>>();
    root_names.sort();

    let metadata_root = format!("{}.metadata", manifest.namespace);
    let initializer_root = format!("{}.bootstrap.initializers", manifest.namespace);
    let initializer_count = match transaction.get(&initializer_root) {
        Some(Out::YArray(value)) => Some(value.len(&transaction)),
        Some(_) => {
            return Err(collaboration_error(
                "office.collaboration.root_invalid",
                format!("The collaboration root '{initializer_root}' must be a Y.Array."),
            ))
        }
        None => None,
    };
    let (metadata, metadata_fields_present, metadata_complete) = match transaction
        .get(&metadata_root)
    {
        Some(Out::YMap(map)) => {
            let fields_present = ["protocol", "version", "artifactId", "kind", "initialized"]
                .into_iter()
                .filter(|key| map.get(&transaction, key).is_some())
                .count();
            if map.len(&transaction) == 0 {
                (None, 0, false)
            } else {
                let protocol = map_string(&map, &transaction, "protocol");
                let version = map_u32(&map, &transaction, "version");
                let artifact_id = map_string(&map, &transaction, "artifactId");
                let kind =
                    map_string(&map, &transaction, "kind").and_then(|value| value.parse().ok());
                let initialized = map_bool(&map, &transaction, "initialized");
                let complete = protocol.is_some()
                    && version.is_some()
                    && artifact_id.is_some()
                    && kind.is_some()
                    && initialized.is_some();
                if !complete {
                    return Err(collaboration_error(
                            "office.collaboration.metadata_invalid",
                            "The shared document contains incomplete or unsupported Office collaboration metadata.",
                        )
                        .with_detail("fieldsPresent", fields_present as u64));
                }
                let metadata = NativeOfficeCollaborationMetadata {
                    protocol: protocol.expect("checked above"),
                    version: version.expect("checked above"),
                    artifact_id: artifact_id.expect("checked above"),
                    kind: kind.expect("checked above"),
                    initialized: initialized.expect("checked above"),
                };
                validate_metadata(&metadata, manifest, initializer_count)?;
                (Some(metadata), fields_present, true)
            }
        }
        Some(_) => {
            return Err(collaboration_error(
                "office.collaboration.root_invalid",
                format!("The collaboration root '{metadata_root}' must be a Y.Map."),
            ))
        }
        None => (None, 0, false),
    };

    if metadata.is_none() && initializer_count.unwrap_or(0) > 0 {
        return Err(collaboration_error(
            "office.collaboration.metadata_missing",
            "The shared document contains bootstrap state without Office collaboration metadata.",
        ));
    }
    let bootstrap_valid = metadata.as_ref().map(|metadata| {
        initializer_count.unwrap_or(0) <= 1
            && (!metadata.initialized || initializer_count == Some(1))
    });
    Ok(DocumentInspection {
        metadata,
        metadata_fields_present,
        metadata_complete,
        bootstrap_initializer_count: initializer_count,
        bootstrap_valid,
        root_names,
        state_sha256: sha256_hex(&transaction.encode_state_as_update_v1(&StateVector::default())),
        pending_updates: transaction.has_missing_updates(),
    })
}

fn validate_metadata(
    metadata: &NativeOfficeCollaborationMetadata,
    manifest: &NativeOfficeCollaborationManifest,
    initializer_count: Option<u32>,
) -> UseResult<()> {
    if metadata.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
        || metadata.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
    {
        return Err(collaboration_error(
            "office.collaboration.metadata_invalid",
            "The shared document contains invalid or unsupported Office collaboration metadata.",
        )
        .with_detail("protocol", metadata.protocol.clone())
        .with_detail("version", metadata.version as u64));
    }
    if metadata.artifact_id != manifest.artifact_id {
        return Err(collaboration_error(
            "office.collaboration.artifact_mismatch",
            format!(
                "The shared document belongs to artifact '{}', not '{}'.",
                metadata.artifact_id, manifest.artifact_id
            ),
        )
        .with_detail("expectedArtifactId", manifest.artifact_id.clone())
        .with_detail("actualArtifactId", metadata.artifact_id.clone()));
    }
    if metadata.kind != manifest.kind {
        return Err(collaboration_error(
            "office.collaboration.kind_mismatch",
            format!(
                "The shared document contains '{}' content, not '{}'.",
                metadata.kind.as_str(),
                manifest.kind.as_str()
            ),
        )
        .with_detail("expectedKind", manifest.kind.as_str())
        .with_detail("actualKind", metadata.kind.as_str()));
    }
    let initializer_count = initializer_count.unwrap_or(0);
    if initializer_count > 1 {
        return Err(collaboration_error(
            "office.collaboration.bootstrap_ambiguous",
            "Multiple clients initialized the shared document before synchronization completed.",
        )
        .with_detail("initializerCount", initializer_count as u64));
    }
    if metadata.initialized && initializer_count != 1 {
        return Err(collaboration_error(
            "office.collaboration.bootstrap_invalid",
            "Initialized Office collaboration metadata must have exactly one initializer.",
        )
        .with_detail("initializerCount", initializer_count as u64));
    }
    Ok(())
}

fn map_string<T: ReadTxn>(map: &yrs::MapRef, transaction: &T, key: &str) -> Option<String> {
    match map.get(transaction, key)? {
        Out::Any(Any::String(value)) => Some(value.to_string()),
        _ => None,
    }
}

fn map_bool<T: ReadTxn>(map: &yrs::MapRef, transaction: &T, key: &str) -> Option<bool> {
    match map.get(transaction, key)? {
        Out::Any(Any::Bool(value)) => Some(value),
        _ => None,
    }
}

fn map_u32<T: ReadTxn>(map: &yrs::MapRef, transaction: &T, key: &str) -> Option<u32> {
    match map.get(transaction, key)? {
        Out::Any(Any::Number(value))
            if value.fract() == 0.0 && value >= 0.0 && value <= u32::MAX as f64 =>
        {
            Some(value as u32)
        }
        Out::Any(Any::BigInt(value)) if (0..=u32::MAX as i64).contains(&value) => {
            Some(value as u32)
        }
        _ => None,
    }
}

pub(super) fn canonical_state_vector(state_vector: &StateVector) -> Vec<u8> {
    let mut entries = state_vector
        .iter()
        .map(
            |(client_id, clock)| NativeOfficeCollaborationStateVectorEntry {
                client_id: client_id.get(),
                clock: *clock,
            },
        )
        .collect::<Vec<_>>();
    entries.sort_by_key(|entry| entry.client_id);
    let mut bytes = Vec::with_capacity(1 + entries.len() * 8);
    write_var_uint(&mut bytes, entries.len() as u64);
    for entry in entries {
        write_var_uint(&mut bytes, entry.client_id);
        write_var_uint(&mut bytes, entry.clock as u64);
    }
    bytes
}

fn write_var_uint(output: &mut Vec<u8>, mut value: u64) {
    while value > 0x7f {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

pub(super) fn state_vector_sha256(state_vector: &StateVector) -> String {
    sha256_hex(&canonical_state_vector(state_vector))
}

pub(super) fn document_state_sha256(doc: &Doc) -> String {
    let transaction = doc.transact();
    sha256_hex(&transaction.encode_state_as_update_v1(&StateVector::default()))
}
