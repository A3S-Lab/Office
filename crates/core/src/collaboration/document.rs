use a3s_use_core::UseResult;
use yrs::types::ToJson;
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Array, ClientID, Doc, Map, OffsetKind, Options, Out, ReadTxn, StateVector, Text, Transact,
    Update, Xml, XmlFragment, XmlOut,
};

use super::{
    collaboration_error, sha256_hex, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMetadata,
    NativeOfficeCollaborationStateVectorEntry, NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};

const MAX_CANONICAL_REPLAY_PASSES: usize = 32;

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
                "document.change-decisions",
                "document.bibliography",
                "document.bibliography.sources",
            ] {
                doc.get_or_insert_map(root(suffix));
            }
            for suffix in [
                "document.comment-order",
                "document.change-decision-order",
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
                "presentation.record-claims",
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

/// Rebuild a replica from the complete Yjs state, including pending updates.
///
/// Yrs can retain causally delayed array items in a pending traversal state
/// after otherwise complete updates arrive out of order. Encoding the full
/// state includes those pending structs; replaying that canonical update into
/// a fresh document integrates every dependency that is now available while
/// preserving genuinely incomplete updates for later delivery.
pub(super) fn canonical_replay_document(
    doc: &Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<Doc> {
    let mut update = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let mut update_sha256 = sha256_hex(&update);
    for _ in 0..MAX_CANONICAL_REPLAY_PASSES {
        let replayed = new_replica_document(manifest.client_id, &manifest.namespace, manifest.kind);
        validate_and_apply_update(&replayed, &update, manifest)?;
        let transaction = replayed.transact();
        if !transaction.has_missing_updates() {
            drop(transaction);
            return Ok(replayed);
        }
        let next_update = transaction.encode_state_as_update_v1(&StateVector::default());
        drop(transaction);
        let next_update_sha256 = sha256_hex(&next_update);
        if next_update_sha256 == update_sha256 {
            return Ok(replayed);
        }
        update = next_update;
        update_sha256 = next_update_sha256;
    }
    Err(collaboration_error(
        "office.collaboration.update_invalid",
        "The Yjs v1 update did not reach a stable causal state after bounded canonical replay.",
    )
    .with_detail("maximumReplayPasses", MAX_CANONICAL_REPLAY_PASSES as u64))
}

pub(super) fn replay_update_sequence<I, B>(
    updates: I,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<Doc>
where
    I: IntoIterator<Item = B>,
    B: AsRef<[u8]>,
{
    let mut replayed = new_replica_document(manifest.client_id, &manifest.namespace, manifest.kind);
    for bytes in updates {
        let update = Update::decode_v1(bytes.as_ref()).map_err(|error| {
            collaboration_error(
                "office.collaboration.update_invalid",
                format!("The Yjs v1 update sequence contains invalid data: {error}"),
            )
        })?;
        replayed
            .transact_mut()
            .apply_update(update)
            .map_err(|error| {
                collaboration_error(
                    "office.collaboration.update_invalid",
                    format!("The Yjs v1 update sequence could not be integrated: {error}"),
                )
            })?;
    }
    if replayed.transact().has_missing_updates() {
        replayed = canonical_replay_document(&replayed, manifest)?;
    }
    inspect_document(&replayed, manifest)?;
    Ok(replayed)
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
        state_sha256: canonical_document_state_sha256(&transaction),
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

pub(super) fn document_update_sha256(doc: &Doc) -> String {
    let transaction = doc.transact();
    sha256_hex(&transaction.encode_state_as_update_v1(&StateVector::default()))
}

/// Hash the converged logical document state rather than the incidental byte
/// order used by Yrs when it re-encodes JSON object values.
///
/// `Any::Map` is backed by a randomized `HashMap`. Two replicas can therefore
/// encode byte-distinct full-state updates after applying the same Yjs structs,
/// especially for portable PDF arrays such as `segmentRects`. The canonical
/// digest combines the sorted state vector, sorted named roots, root types, and
/// recursively key-sorted logical values. This keeps inspection, restart, and
/// delivery-order comparisons stable without changing the standard Yjs v1
/// updates exchanged with browser peers.
fn canonical_document_state_sha256<T: ReadTxn>(transaction: &T) -> String {
    let mut bytes = b"a3s.office.document-state.v1\0".to_vec();
    let state_vector = canonical_state_vector(&transaction.state_vector());
    write_length_prefixed(&mut bytes, &state_vector);

    let mut roots = transaction
        .root_refs()
        .map(|(name, value)| (name.to_owned(), value))
        .collect::<Vec<_>>();
    roots.sort_by(|left, right| left.0.cmp(&right.0));
    write_var_uint(&mut bytes, roots.len() as u64);
    for (name, value) in roots {
        write_length_prefixed(&mut bytes, name.as_bytes());
        write_canonical_out(&mut bytes, transaction, &value);
    }
    sha256_hex(&bytes)
}

pub(super) fn canonical_visible_root_sha256<T: ReadTxn>(transaction: &T, value: &Out) -> String {
    let mut bytes = b"a3s.office.visible-root.v1\0".to_vec();
    write_canonical_out(&mut bytes, transaction, value);
    sha256_hex(&bytes)
}

pub(super) fn canonical_map_without_key_sha256<T: ReadTxn>(
    transaction: &T,
    map: &yrs::MapRef,
    omitted_key: &str,
) -> String {
    let mut bytes = b"a3s.office.filtered-map.v1\0".to_vec();
    let mut entries = map
        .iter(transaction)
        .filter(|(key, _)| *key != omitted_key)
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| left.0.cmp(&right.0));
    write_var_uint(&mut bytes, entries.len() as u64);
    for (key, value) in entries {
        write_length_prefixed(&mut bytes, key.as_bytes());
        write_canonical_out(&mut bytes, transaction, &value);
    }
    sha256_hex(&bytes)
}

pub(super) fn canonical_content_without_comment_marks_sha256<T: ReadTxn>(
    transaction: &T,
    value: &Out,
) -> String {
    let mut bytes = b"a3s.office.document-content-without-comments.v1\0".to_vec();
    write_canonical_out_without_comment_marks(&mut bytes, transaction, value);
    sha256_hex(&bytes)
}

pub(super) fn canonical_content_without_suggestion_effects_sha256<T: ReadTxn>(
    transaction: &T,
    value: &Out,
) -> String {
    let mut bytes = b"a3s.office.document-content-without-suggestions.v1\0".to_vec();
    write_canonical_out_without_suggestion_effects(&mut bytes, transaction, value);
    sha256_hex(&bytes)
}

fn write_canonical_out<T: ReadTxn>(output: &mut Vec<u8>, transaction: &T, value: &Out) {
    output.push(out_type_tag(value));
    match value {
        Out::YXmlElement(element) => {
            write_length_prefixed(output, element.tag().as_bytes());
            write_canonical_xml_attributes(output, transaction, element);
            write_canonical_xml_children(output, transaction, element);
        }
        Out::YXmlFragment(fragment) => {
            write_canonical_xml_children(output, transaction, fragment);
        }
        Out::YXmlText(text) => {
            write_canonical_xml_attributes(output, transaction, text);
            let chunks = text.diff(transaction, |_| ());
            write_var_uint(output, chunks.len() as u64);
            for chunk in chunks {
                write_canonical_out(output, transaction, &chunk.insert);
                match chunk.attributes {
                    Some(attributes) => {
                        output.push(1);
                        write_canonical_attributes(output, &attributes);
                    }
                    None => output.push(0),
                }
            }
        }
        _ => write_canonical_any(output, &value.to_json(transaction)),
    }
}

fn write_canonical_out_without_comment_marks<T: ReadTxn>(
    output: &mut Vec<u8>,
    transaction: &T,
    value: &Out,
) {
    output.push(out_type_tag(value));
    match value {
        Out::YXmlElement(element) => {
            write_length_prefixed(output, element.tag().as_bytes());
            write_canonical_xml_attributes(output, transaction, element);
            write_canonical_xml_children_without_comment_marks(output, transaction, element);
        }
        Out::YXmlFragment(fragment) => {
            write_canonical_xml_children_without_comment_marks(output, transaction, fragment);
        }
        Out::YXmlText(text) => {
            write_canonical_xml_attributes(output, transaction, text);
            let mut chunks = Vec::<FilteredXmlTextChunk>::new();
            for chunk in text.diff(transaction, |_| ()) {
                let attributes = filtered_text_attributes(
                    chunk.attributes.as_ref().map(|value| &**value),
                    is_document_comment_attribute,
                );
                match chunk.insert {
                    Out::Any(Any::String(value)) => {
                        if let Some(FilteredXmlTextChunk {
                            insert: FilteredXmlTextInsert::Text(previous),
                            attributes: previous_attributes,
                        }) = chunks.last_mut()
                        {
                            if *previous_attributes == attributes {
                                previous.push_str(&value);
                                continue;
                            }
                        }
                        chunks.push(FilteredXmlTextChunk {
                            insert: FilteredXmlTextInsert::Text(value.to_string()),
                            attributes,
                        });
                    }
                    value => chunks.push(FilteredXmlTextChunk {
                        insert: FilteredXmlTextInsert::Other(value),
                        attributes,
                    }),
                }
            }
            write_var_uint(output, chunks.len() as u64);
            for chunk in chunks {
                match chunk.insert {
                    FilteredXmlTextInsert::Text(value) => write_canonical_out(
                        output,
                        transaction,
                        &Out::Any(Any::String(value.into())),
                    ),
                    FilteredXmlTextInsert::Other(value) => {
                        write_canonical_out(output, transaction, &value)
                    }
                }
                match chunk.attributes {
                    Some(attributes) => {
                        output.push(1);
                        output.extend_from_slice(&attributes);
                    }
                    None => output.push(0),
                }
            }
        }
        _ => write_canonical_any(output, &value.to_json(transaction)),
    }
}

fn write_canonical_out_without_suggestion_effects<T: ReadTxn>(
    output: &mut Vec<u8>,
    transaction: &T,
    value: &Out,
) {
    output.push(out_type_tag(value));
    match value {
        Out::YXmlElement(element) => {
            write_length_prefixed(output, element.tag().as_bytes());
            write_canonical_xml_attributes(output, transaction, element);
            write_canonical_xml_children_without_suggestion_effects(output, transaction, element);
        }
        Out::YXmlFragment(fragment) => {
            write_canonical_xml_children_without_suggestion_effects(output, transaction, fragment);
        }
        Out::YXmlText(text) => {
            write_canonical_xml_attributes(output, transaction, text);
            let mut chunks = Vec::<FilteredXmlTextChunk>::new();
            for chunk in text.diff(transaction, |_| ()) {
                let source_attributes = chunk.attributes.as_ref().map(|value| &**value);
                if document_change_chunk_kind(source_attributes)
                    == Some(DocumentChangeChunkKind::Insertion)
                {
                    continue;
                }
                let attributes =
                    filtered_text_attributes(source_attributes, is_document_change_attribute);
                match chunk.insert {
                    Out::Any(Any::String(value)) => {
                        if let Some(FilteredXmlTextChunk {
                            insert: FilteredXmlTextInsert::Text(previous),
                            attributes: previous_attributes,
                        }) = chunks.last_mut()
                        {
                            if *previous_attributes == attributes {
                                previous.push_str(&value);
                                continue;
                            }
                        }
                        chunks.push(FilteredXmlTextChunk {
                            insert: FilteredXmlTextInsert::Text(value.to_string()),
                            attributes,
                        });
                    }
                    value => chunks.push(FilteredXmlTextChunk {
                        insert: FilteredXmlTextInsert::Other(value),
                        attributes,
                    }),
                }
            }
            write_var_uint(output, chunks.len() as u64);
            for chunk in chunks {
                match chunk.insert {
                    FilteredXmlTextInsert::Text(value) => write_canonical_out(
                        output,
                        transaction,
                        &Out::Any(Any::String(value.into())),
                    ),
                    FilteredXmlTextInsert::Other(value) => {
                        write_canonical_out(output, transaction, &value)
                    }
                }
                match chunk.attributes {
                    Some(attributes) => {
                        output.push(1);
                        output.extend_from_slice(&attributes);
                    }
                    None => output.push(0),
                }
            }
        }
        _ => write_canonical_any(output, &value.to_json(transaction)),
    }
}

enum FilteredXmlTextInsert {
    Text(String),
    Other(Out),
}

struct FilteredXmlTextChunk {
    insert: FilteredXmlTextInsert,
    attributes: Option<Vec<u8>>,
}

fn filtered_text_attributes(
    attributes: Option<&yrs::types::Attrs>,
    omitted: fn(&str) -> bool,
) -> Option<Vec<u8>> {
    let mut entries = attributes?
        .iter()
        .filter(|(key, _)| !omitted(key))
        .collect::<Vec<_>>();
    if entries.is_empty() {
        return None;
    }
    entries.sort_by(|left, right| left.0.cmp(right.0));
    let mut output = Vec::new();
    write_var_uint(&mut output, entries.len() as u64);
    for (key, value) in entries {
        write_length_prefixed(&mut output, key.as_bytes());
        write_canonical_any(&mut output, value);
    }
    Some(output)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum DocumentChangeChunkKind {
    Insertion,
    Deletion,
}

fn document_change_chunk_kind(
    attributes: Option<&yrs::types::Attrs>,
) -> Option<DocumentChangeChunkKind> {
    attributes?.iter().find_map(|(key, value)| {
        if !is_document_change_attribute(key) {
            return None;
        }
        let Any::Map(fields) = value else {
            return None;
        };
        match fields.get("kind") {
            Some(Any::String(value)) if value.as_ref() == "insertion" => {
                Some(DocumentChangeChunkKind::Insertion)
            }
            Some(Any::String(value)) if value.as_ref() == "deletion" => {
                Some(DocumentChangeChunkKind::Deletion)
            }
            _ => None,
        }
    })
}

fn is_document_comment_attribute(value: &str) -> bool {
    if value == "documentComment" {
        return true;
    }
    value
        .strip_prefix("documentComment--")
        .is_some_and(|suffix| {
            suffix.len() == 8
                && suffix.bytes().all(|value| {
                    value.is_ascii_alphanumeric() || matches!(value, b'+' | b'/' | b'=')
                })
        })
}

pub(super) fn is_document_change_attribute(value: &str) -> bool {
    if value == "documentChange" {
        return true;
    }
    value
        .strip_prefix("documentChange--")
        .is_some_and(|suffix| {
            suffix.len() == 8
                && suffix.bytes().all(|value| {
                    value.is_ascii_alphanumeric() || matches!(value, b'+' | b'/' | b'=')
                })
        })
}

fn write_canonical_xml_attributes<T: ReadTxn>(
    output: &mut Vec<u8>,
    transaction: &T,
    value: &impl Xml,
) {
    let mut attributes = value.attributes(transaction).collect::<Vec<_>>();
    attributes.sort_by(|left, right| left.0.cmp(right.0));
    write_var_uint(output, attributes.len() as u64);
    for (key, value) in attributes {
        write_length_prefixed(output, key.as_bytes());
        write_canonical_out(output, transaction, &value);
    }
}

fn write_canonical_xml_children<T: ReadTxn>(
    output: &mut Vec<u8>,
    transaction: &T,
    value: &impl XmlFragment,
) {
    let children = value.children(transaction).collect::<Vec<_>>();
    write_var_uint(output, children.len() as u64);
    for child in children {
        match child {
            XmlOut::Element(element) => {
                write_canonical_out(output, transaction, &Out::YXmlElement(element));
            }
            XmlOut::Fragment(fragment) => {
                write_canonical_out(output, transaction, &Out::YXmlFragment(fragment));
            }
            XmlOut::Text(text) => {
                write_canonical_out(output, transaction, &Out::YXmlText(text));
            }
        }
    }
}

fn write_canonical_xml_children_without_comment_marks<T: ReadTxn>(
    output: &mut Vec<u8>,
    transaction: &T,
    value: &impl XmlFragment,
) {
    let children = value.children(transaction).collect::<Vec<_>>();
    write_var_uint(output, children.len() as u64);
    for child in children {
        match child {
            XmlOut::Element(element) => write_canonical_out_without_comment_marks(
                output,
                transaction,
                &Out::YXmlElement(element),
            ),
            XmlOut::Fragment(fragment) => write_canonical_out_without_comment_marks(
                output,
                transaction,
                &Out::YXmlFragment(fragment),
            ),
            XmlOut::Text(text) => {
                write_canonical_out_without_comment_marks(output, transaction, &Out::YXmlText(text))
            }
        }
    }
}

fn write_canonical_xml_children_without_suggestion_effects<T: ReadTxn>(
    output: &mut Vec<u8>,
    transaction: &T,
    value: &impl XmlFragment,
) {
    let children = value.children(transaction).collect::<Vec<_>>();
    write_var_uint(output, children.len() as u64);
    for child in children {
        match child {
            XmlOut::Element(element) => write_canonical_out_without_suggestion_effects(
                output,
                transaction,
                &Out::YXmlElement(element),
            ),
            XmlOut::Fragment(fragment) => write_canonical_out_without_suggestion_effects(
                output,
                transaction,
                &Out::YXmlFragment(fragment),
            ),
            XmlOut::Text(text) => write_canonical_out_without_suggestion_effects(
                output,
                transaction,
                &Out::YXmlText(text),
            ),
        }
    }
}

fn write_canonical_attributes(output: &mut Vec<u8>, attributes: &yrs::types::Attrs) {
    let mut entries = attributes.iter().collect::<Vec<_>>();
    entries.sort_by(|left, right| left.0.cmp(right.0));
    write_var_uint(output, entries.len() as u64);
    for (key, value) in entries {
        write_length_prefixed(output, key.as_bytes());
        write_canonical_any(output, value);
    }
}

fn out_type_tag(value: &Out) -> u8 {
    match value {
        Out::Any(_) => 0,
        Out::YText(_) => 1,
        Out::YArray(_) => 2,
        Out::YMap(_) => 3,
        Out::YXmlElement(_) => 4,
        Out::YXmlFragment(_) => 5,
        Out::YXmlText(_) => 6,
        Out::YDoc(_) => 7,
        Out::UndefinedRef(_) => 9,
    }
}

fn write_canonical_any(output: &mut Vec<u8>, value: &Any) {
    match value {
        Any::Null => output.push(0),
        Any::Undefined => output.push(1),
        Any::Bool(false) => output.push(2),
        Any::Bool(true) => output.push(3),
        Any::Number(value) => {
            output.push(4);
            let normalized = if *value == 0.0 { 0.0 } else { *value };
            output.extend_from_slice(&normalized.to_bits().to_le_bytes());
        }
        Any::BigInt(value) => {
            output.push(5);
            output.extend_from_slice(&value.to_le_bytes());
        }
        Any::String(value) => {
            output.push(6);
            write_length_prefixed(output, value.as_bytes());
        }
        Any::Buffer(value) => {
            output.push(7);
            write_length_prefixed(output, value);
        }
        Any::Array(values) => {
            output.push(8);
            write_var_uint(output, values.len() as u64);
            for value in values.iter() {
                write_canonical_any(output, value);
            }
        }
        Any::Map(values) => {
            output.push(9);
            let mut entries = values.iter().collect::<Vec<_>>();
            entries.sort_by(|left, right| left.0.cmp(right.0));
            write_var_uint(output, entries.len() as u64);
            for (key, value) in entries {
                write_length_prefixed(output, key.as_bytes());
                write_canonical_any(output, value);
            }
        }
    }
}

fn write_length_prefixed(output: &mut Vec<u8>, value: &[u8]) {
    write_var_uint(output, value.len() as u64);
    output.extend_from_slice(value);
}
