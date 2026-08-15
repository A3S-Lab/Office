use std::cmp::Ordering;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::sync::Arc;

use a3s_use_core::UseResult;
use serde_json::{Map as JsonMap, Number as JsonNumber, Value as JsonValue};
use yrs::{Any, Array, ArrayRef, Map, Out, Transact};

use super::super::super::{collaboration_error, NativeOfficeCollaborationManifest};
use super::{
    decoded_record_field_identity, encoded_record_entry_key, invalid_shared_pdf,
    validate_shared_pdf_identifier, PdfRecordCollectionRoots, PdfRecordFieldKind, MAX_PDF_RECORDS,
};

const MAX_PDF_BYTES: u64 = 4 * 1024 * 1024 * 1024;
pub(super) const MAX_PDF_PAGES: u32 = 1_000_000;
const MAX_JSON_DEPTH: usize = 128;

#[derive(Debug)]
pub(super) struct PdfRecords {
    pub(super) order: Vec<String>,
    pub(super) by_id: HashMap<String, JsonValue>,
}

pub(super) struct PdfClaims {
    by_identity: HashMap<(String, String), String>,
}

impl PdfClaims {
    pub(super) fn claim_for(&self, kind: &str, id: &str) -> Option<&str> {
        self.by_identity
            .get(&(kind.to_owned(), id.to_owned()))
            .map(String::as_str)
    }

    pub(super) fn assert_exact_records(&self, kind: &str, records: &PdfRecords) -> UseResult<()> {
        for id in &records.order {
            let expected = canonical_json(&records.by_id[id])?;
            match self.claim_for(kind, id) {
                Some(actual) if actual == expected => {}
                Some(_) => {
                    return Err(invalid_shared_pdf(format!(
                        "The shared PDF {kind} record claim does not match ID '{id}'."
                    )))
                }
                None => {
                    return Err(invalid_shared_pdf(format!(
                        "The shared PDF {kind} record ID '{id}' is missing its immutable claim."
                    )))
                }
            }
        }
        Ok(())
    }
}

pub(super) fn append_record(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    collection: &str,
    record: &JsonValue,
    claim_kind: &str,
    append_claim: bool,
) -> UseResult<()> {
    append_record_with_fingerprint(
        doc,
        manifest,
        collection,
        record,
        claim_kind,
        record,
        append_claim,
    )
}

pub(super) fn append_record_with_fingerprint(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    collection: &str,
    record: &JsonValue,
    claim_kind: &str,
    fingerprint: &JsonValue,
    append_claim: bool,
) -> UseResult<()> {
    let object = record.as_object().ok_or_else(|| {
        collaboration_error(
            "office.collaboration.mutation_invalid",
            "A native PDF record must be a JSON object.",
        )
    })?;
    let record_id = required_json_string(object.get("id"), "native PDF record ID")?;
    let encoded_fields = flattened_record_fields(record_id, object)?;
    let claim = canonical_json(&serde_json::json!({
        "fingerprint": canonical_json(fingerprint)?,
        "id": record_id,
        "kind": claim_kind,
    }))?;
    let roots = PdfRecordCollectionRoots::new(doc, manifest, collection);
    let claims = pdf_record_claims(doc, manifest);
    let mut transaction = doc.transact_mut();
    roots.presence.insert(&mut transaction, record_id, true);
    for (key, value) in encoded_fields {
        roots.fields.insert(&mut transaction, key, value);
    }
    roots.order.push_back(&mut transaction, record_id);
    if append_claim {
        claims.push_back(&mut transaction, claim);
    }
    Ok(())
}

pub(super) fn patch_record(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    collection: &str,
    record_id: &str,
    previous: &JsonValue,
    shared: &JsonValue,
    next: &JsonValue,
) -> UseResult<()> {
    let previous = previous.as_object().ok_or_else(|| {
        collaboration_error(
            "office.collaboration.mutation_invalid",
            "A native PDF record patch expectation must be a JSON object.",
        )
    })?;
    let shared = shared.as_object().ok_or_else(|| {
        invalid_shared_pdf("The shared PDF record being patched is not an object.")
    })?;
    let next = next.as_object().ok_or_else(|| {
        collaboration_error(
            "office.collaboration.mutation_invalid",
            "A native PDF record patch must produce a JSON object.",
        )
    })?;
    let before = flattened_json_fields(record_id, previous)?;
    let current = flattened_json_fields(record_id, shared)?;
    let after = flattened_json_fields(record_id, next)?;
    let keys = before
        .keys()
        .chain(after.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let mut changes = Vec::new();
    for key in keys {
        if optional_json_equal(before.get(&key), after.get(&key))?
            || optional_json_equal(current.get(&key), after.get(&key))?
        {
            continue;
        }
        changes.push(match after.get(&key) {
            Some(value) => RecordFieldPatch::Set(key, json_to_any(value, 0)?),
            None => RecordFieldPatch::Remove(key),
        });
    }
    if changes.is_empty() {
        return Ok(());
    }

    let roots = PdfRecordCollectionRoots::new(doc, manifest, collection);
    let mut transaction = doc.transact_mut();
    for change in changes {
        match change {
            RecordFieldPatch::Remove(key) => {
                roots.fields.remove(&mut transaction, key.as_str());
            }
            RecordFieldPatch::Set(key, value) => {
                roots.fields.insert(&mut transaction, key, value);
            }
        }
    }
    Ok(())
}

enum RecordFieldPatch {
    Remove(String),
    Set(String, Any),
}

pub(super) fn read_pdf_records(
    doc: &yrs::Doc,
    roots: &PdfRecordCollectionRoots,
    label: &str,
) -> UseResult<PdfRecords> {
    let transaction = doc.transact();
    if roots.presence.len(&transaction) > MAX_PDF_RECORDS {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF contains too many {label} records."
        )));
    }
    let mut seen = HashSet::new();
    let mut order = Vec::new();
    for index in 0..roots.order.len(&transaction) {
        let id = match roots.order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_pdf(format!(
                    "The shared PDF {label} order contains a non-string identity."
                )))
            }
        };
        validate_shared_pdf_identifier(&id, label)?;
        if seen.insert(id.clone()) {
            order.push(id);
        }
    }

    let mut present = HashSet::new();
    for (id, value) in roots.presence.iter(&transaction) {
        validate_shared_pdf_identifier(id, label)?;
        if !matches!(value, Out::Any(Any::Bool(true))) || !seen.contains(id) {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF {label} presence and order roots disagree."
            )));
        }
        present.insert(id.to_owned());
    }
    if seen != present {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} presence and order roots disagree."
        )));
    }

    let mut entries = present
        .iter()
        .map(|id| (id.clone(), Vec::new()))
        .collect::<HashMap<_, Vec<DecodedField>>>();
    for (encoded, value) in roots.fields.iter(&transaction) {
        let identity = decoded_record_field_identity(encoded, label)?;
        if !present.contains(&identity.record_id) {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF contains an orphan {label} field."
            )));
        }
        let value = match (identity.kind, value) {
            (PdfRecordFieldKind::Object, Out::Any(Any::Bool(true))) => JsonValue::Bool(true),
            (PdfRecordFieldKind::Object, _) => {
                return Err(invalid_shared_pdf(format!(
                    "The shared PDF {label} contains an invalid object marker."
                )))
            }
            (PdfRecordFieldKind::Value, Out::Any(value)) => any_to_json(value, 0, label)?,
            (PdfRecordFieldKind::Value, _) => {
                return Err(invalid_shared_pdf(format!(
                    "The shared PDF {label} contains a shared-type field."
                )))
            }
        };
        entries
            .get_mut(&identity.record_id)
            .expect("presence initialized above")
            .push(DecodedField {
                kind: identity.kind,
                path: identity.path,
                value,
            });
    }
    drop(transaction);

    let mut by_id = HashMap::with_capacity(entries.len());
    for (id, fields) in entries {
        let object = reconstructed_record(fields, label)?;
        if object.get("id").and_then(JsonValue::as_str) != Some(id.as_str()) {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF {label} identity field does not match its record."
            )));
        }
        by_id.insert(id, JsonValue::Object(object));
    }
    Ok(PdfRecords { order, by_id })
}

pub(super) fn read_pdf_source_page_count(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<u32> {
    let source = doc.get_or_insert_map(format!("{}.pdf.source", manifest.namespace));
    let identities =
        doc.get_or_insert_array(format!("{}.pdf.source-identities", manifest.namespace));
    let transaction = doc.transact();
    if source.len(&transaction) != 3 {
        return Err(invalid_shared_pdf(
            "The shared PDF source identity is incomplete.",
        ));
    }
    let sha256 = match source.get(&transaction, "sha256") {
        Some(Out::Any(Any::String(value)))
            if value.len() == 64
                && value
                    .bytes()
                    .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte)) =>
        {
            value.to_string()
        }
        _ => {
            return Err(invalid_shared_pdf(
                "The shared PDF source SHA-256 is invalid.",
            ))
        }
    };
    let byte_length = source
        .get(&transaction, "byteLength")
        .and_then(out_json_number)
        .and_then(|value| exact_u64(value, 1, MAX_PDF_BYTES))
        .ok_or_else(|| invalid_shared_pdf("The shared PDF source byte length is invalid."))?;
    let page_count = source
        .get(&transaction, "pageCount")
        .and_then(out_json_number)
        .and_then(|value| exact_u64(value, 1, MAX_PDF_PAGES as u64))
        .map(|value| value as u32)
        .ok_or_else(|| invalid_shared_pdf("The shared PDF source page count is invalid."))?;
    let expected_identity = canonical_json(&serde_json::json!({
        "sha256": sha256,
        "byteLength": byte_length,
        "pageCount": page_count,
    }))?;
    match identities.len(&transaction) {
        0 => {}
        1 => match identities.get(&transaction, 0) {
            Some(Out::Any(Any::String(value))) if value.as_ref() == expected_identity => {}
            _ => {
                return Err(invalid_shared_pdf(
                    "The shared PDF immutable source identity is invalid.",
                ))
            }
        },
        _ => {
            return Err(invalid_shared_pdf(
                "The shared PDF contains conflicting source identities.",
            ))
        }
    }
    Ok(page_count)
}

pub(super) fn read_pdf_claims(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<PdfClaims> {
    let claims = pdf_record_claims(doc, manifest);
    let transaction = doc.transact();
    let mut by_identity = HashMap::new();
    for index in 0..claims.len(&transaction) {
        let raw = match claims.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_pdf(
                    "The shared PDF record claim is not a string.",
                ))
            }
        };
        let value: JsonValue = serde_json::from_str(&raw)
            .map_err(|_| invalid_shared_pdf("The shared PDF record claim is not valid JSON."))?;
        if canonical_json(&value)? != raw {
            return Err(invalid_shared_pdf(
                "The shared PDF record claim is not canonical JSON.",
            ));
        }
        let object = value
            .as_object()
            .ok_or_else(|| invalid_shared_pdf("The shared PDF record claim is not an object."))?;
        assert_exact_keys(object, &["fingerprint", "id", "kind"], "record claim")?;
        let kind = required_json_string(object.get("kind"), "record claim kind")?;
        if !matches!(
            kind,
            "annotation"
                | "page-operation"
                | "redaction"
                | "review-decision"
                | "signature-placement"
        ) {
            return Err(invalid_shared_pdf(
                "The shared PDF record claim kind is unsupported.",
            ));
        }
        let id = required_json_string(object.get("id"), "record claim ID")?;
        validate_shared_pdf_identifier(id, "record claim")?;
        let fingerprint =
            required_json_string(object.get("fingerprint"), "record claim fingerprint")?;
        let fingerprint_value: JsonValue = serde_json::from_str(fingerprint).map_err(|_| {
            invalid_shared_pdf("The shared PDF record claim fingerprint is not valid JSON.")
        })?;
        if canonical_json(&fingerprint_value)? != fingerprint {
            return Err(invalid_shared_pdf(
                "The shared PDF record claim fingerprint is not canonical JSON.",
            ));
        }
        let identity = (kind.to_owned(), id.to_owned());
        if let Some(existing) = by_identity.insert(identity, fingerprint.to_owned()) {
            if existing != fingerprint {
                return Err(invalid_shared_pdf(format!(
                    "The shared PDF {kind} ID '{id}' was claimed by different records."
                )));
            }
        }
    }
    Ok(PdfClaims { by_identity })
}

pub(super) fn canonical_json(value: &JsonValue) -> UseResult<String> {
    let mut result = String::new();
    write_canonical_json(value, &mut result, 0)?;
    Ok(result)
}

pub(super) fn required_json_object<'a>(
    value: &'a JsonValue,
    label: &str,
) -> UseResult<&'a JsonMap<String, JsonValue>> {
    value
        .as_object()
        .ok_or_else(|| invalid_shared_pdf(format!("The shared PDF {label} is not an object.")))
}

pub(super) fn required_json_string<'a>(
    value: Option<&'a JsonValue>,
    label: &str,
) -> UseResult<&'a str> {
    value
        .and_then(JsonValue::as_str)
        .ok_or_else(|| invalid_shared_pdf(format!("The shared PDF {label} is not a string.")))
}

pub(super) fn assert_exact_keys(
    object: &JsonMap<String, JsonValue>,
    expected: &[&str],
    label: &str,
) -> UseResult<()> {
    assert_allowed_keys(object, expected, expected, label)
}

pub(super) fn assert_allowed_keys(
    object: &JsonMap<String, JsonValue>,
    allowed: &[&str],
    required: &[&str],
    label: &str,
) -> UseResult<()> {
    if object.len() > allowed.len()
        || object.keys().any(|key| !allowed.contains(&key.as_str()))
        || required.iter().any(|key| !object.contains_key(*key))
    {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} contains missing or unsupported fields."
        )));
    }
    Ok(())
}

fn flattened_record_fields(
    record_id: &str,
    object: &JsonMap<String, JsonValue>,
) -> UseResult<Vec<(String, Any)>> {
    flattened_json_fields(record_id, object)?
        .into_iter()
        .map(|(key, value)| Ok((key, json_to_any(&value, 0)?)))
        .collect()
}

fn flattened_json_fields(
    record_id: &str,
    object: &JsonMap<String, JsonValue>,
) -> UseResult<BTreeMap<String, JsonValue>> {
    let mut fields = BTreeMap::new();
    for (key, value) in object {
        flatten_json_value(record_id, &mut fields, &mut vec![key.clone()], value, 0)?;
    }
    Ok(fields)
}

fn flatten_json_value(
    record_id: &str,
    fields: &mut BTreeMap<String, JsonValue>,
    path: &mut Vec<String>,
    value: &JsonValue,
    depth: usize,
) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A native PDF record exceeds the JSON nesting limit.",
        ));
    }
    if let JsonValue::Object(object) = value {
        fields.insert(
            encoded_record_entry_key(record_id, PdfRecordFieldKind::Object, path)?,
            JsonValue::Bool(true),
        );
        for (key, child) in object {
            path.push(key.clone());
            flatten_json_value(record_id, fields, path, child, depth + 1)?;
            path.pop();
        }
    } else {
        fields.insert(
            encoded_record_entry_key(record_id, PdfRecordFieldKind::Value, path)?,
            value.clone(),
        );
    }
    Ok(())
}

struct DecodedField {
    kind: PdfRecordFieldKind,
    path: Vec<String>,
    value: JsonValue,
}

fn reconstructed_record(
    mut fields: Vec<DecodedField>,
    label: &str,
) -> UseResult<JsonMap<String, JsonValue>> {
    fields.sort_by(|left, right| {
        left.path.len().cmp(&right.path.len()).then_with(|| {
            if left.kind == right.kind {
                Ordering::Equal
            } else if left.kind == PdfRecordFieldKind::Object {
                Ordering::Less
            } else {
                Ordering::Greater
            }
        })
    });
    let object_paths = fields
        .iter()
        .filter(|field| field.kind == PdfRecordFieldKind::Object)
        .map(|field| field.path.clone())
        .collect::<HashSet<_>>();
    for field in &fields {
        for depth in 1..field.path.len() {
            if !object_paths.contains(&field.path[..depth]) {
                return Err(invalid_shared_pdf(format!(
                    "The shared PDF {label} contains a value without an object marker."
                )));
            }
        }
    }

    let mut result = JsonMap::new();
    for field in fields {
        match field.kind {
            PdfRecordFieldKind::Object => ensure_object_path(&mut result, &field.path, label)?,
            PdfRecordFieldKind::Value => {
                set_value_at_path(&mut result, &field.path, field.value, label)?
            }
        }
    }
    Ok(result)
}

fn ensure_object_path(
    object: &mut JsonMap<String, JsonValue>,
    path: &[String],
    label: &str,
) -> UseResult<()> {
    let Some((key, rest)) = path.split_first() else {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} contains an empty object path."
        )));
    };
    let value = object
        .entry(key.clone())
        .or_insert_with(|| JsonValue::Object(JsonMap::new()));
    let child = value.as_object_mut().ok_or_else(|| {
        invalid_shared_pdf(format!(
            "The shared PDF {label} contains overlapping object and value fields."
        ))
    })?;
    if rest.is_empty() {
        Ok(())
    } else {
        ensure_object_path(child, rest, label)
    }
}

fn set_value_at_path(
    object: &mut JsonMap<String, JsonValue>,
    path: &[String],
    value: JsonValue,
    label: &str,
) -> UseResult<()> {
    let Some((key, rest)) = path.split_first() else {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} contains an empty value path."
        )));
    };
    if rest.is_empty() {
        if object.insert(key.clone(), value).is_some() {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF {label} contains duplicate object/value fields."
            )));
        }
        return Ok(());
    }
    let child = object
        .get_mut(key)
        .and_then(JsonValue::as_object_mut)
        .ok_or_else(|| {
            invalid_shared_pdf(format!(
                "The shared PDF {label} contains a value without an object parent."
            ))
        })?;
    set_value_at_path(child, rest, value, label)
}

fn pdf_record_claims(doc: &yrs::Doc, manifest: &NativeOfficeCollaborationManifest) -> ArrayRef {
    doc.get_or_insert_array(format!("{}.pdf.record-claims", manifest.namespace))
}

fn out_json_number(value: Out) -> Option<f64> {
    match value {
        Out::Any(Any::Number(value)) if value.is_finite() => Some(value),
        _ => None,
    }
}

fn exact_u64(value: f64, minimum: u64, maximum: u64) -> Option<u64> {
    if value.fract() != 0.0 || value < minimum as f64 || value > maximum as f64 {
        return None;
    }
    Some(value as u64)
}

fn any_to_json(value: Any, depth: usize, label: &str) -> UseResult<JsonValue> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} exceeds the JSON nesting limit."
        )));
    }
    match value {
        Any::Null => Ok(JsonValue::Null),
        Any::Bool(value) => Ok(JsonValue::Bool(value)),
        Any::Number(value) => JsonNumber::from_f64(value)
            .map(JsonValue::Number)
            .ok_or_else(|| invalid_shared_pdf(format!("The shared PDF {label} is not finite."))),
        Any::String(value) => Ok(JsonValue::String(value.to_string())),
        Any::Array(values) => values
            .iter()
            .cloned()
            .map(|value| any_to_json(value, depth + 1, label))
            .collect::<UseResult<Vec<_>>>()
            .map(JsonValue::Array),
        Any::Map(values) => {
            let mut object = JsonMap::new();
            for (key, value) in values.iter() {
                if matches!(key.as_str(), "__proto__" | "constructor" | "prototype") {
                    return Err(invalid_shared_pdf(format!(
                        "The shared PDF {label} contains an unsafe JSON key."
                    )));
                }
                object.insert(key.clone(), any_to_json(value.clone(), depth + 1, label)?);
            }
            Ok(JsonValue::Object(object))
        }
        Any::Undefined | Any::BigInt(_) | Any::Buffer(_) => Err(invalid_shared_pdf(format!(
            "The shared PDF {label} contains a non-JSON value."
        ))),
    }
}

fn json_to_any(value: &JsonValue, depth: usize) -> UseResult<Any> {
    if depth > MAX_JSON_DEPTH {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A native PDF record exceeds the JSON nesting limit.",
        ));
    }
    match value {
        JsonValue::Null => Ok(Any::Null),
        JsonValue::Bool(value) => Ok(Any::Bool(*value)),
        JsonValue::Number(value) => value.as_f64().map(Any::Number).ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                "A native PDF record contains a non-finite number.",
            )
        }),
        JsonValue::String(value) => Ok(Any::String(Arc::from(value.as_str()))),
        JsonValue::Array(values) => values
            .iter()
            .map(|value| json_to_any(value, depth + 1))
            .collect::<UseResult<Vec<_>>>()
            .map(|values| Any::Array(Arc::from(values))),
        JsonValue::Object(values) => {
            let mut object = HashMap::with_capacity(values.len());
            for (key, value) in values {
                object.insert(key.clone(), json_to_any(value, depth + 1)?);
            }
            Ok(Any::Map(Arc::new(object)))
        }
    }
}

pub(super) fn json_equal(left: &JsonValue, right: &JsonValue) -> UseResult<bool> {
    Ok(canonical_json(left)? == canonical_json(right)?)
}

fn optional_json_equal(left: Option<&JsonValue>, right: Option<&JsonValue>) -> UseResult<bool> {
    match (left, right) {
        (Some(left), Some(right)) => json_equal(left, right),
        (None, None) => Ok(true),
        _ => Ok(false),
    }
}

fn write_canonical_json(value: &JsonValue, output: &mut String, depth: usize) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_shared_pdf(
            "The shared PDF canonical JSON exceeds the nesting limit.",
        ));
    }
    match value {
        JsonValue::Null => output.push_str("null"),
        JsonValue::Bool(value) => output.push_str(if *value { "true" } else { "false" }),
        JsonValue::Number(value) => {
            let value = value.as_f64().ok_or_else(|| {
                invalid_shared_pdf("The shared PDF canonical JSON number is unsupported.")
            })?;
            if !value.is_finite() {
                return Err(invalid_shared_pdf(
                    "The shared PDF canonical JSON number is not finite.",
                ));
            }
            output.push_str(ryu_js::Buffer::new().format_finite(value));
        }
        JsonValue::String(value) => {
            output.push_str(&serde_json::to_string(value).map_err(|error| {
                invalid_shared_pdf(format!("Failed to encode shared PDF JSON: {error}"))
            })?)
        }
        JsonValue::Array(values) => {
            output.push('[');
            for (index, value) in values.iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                write_canonical_json(value, output, depth + 1)?;
            }
            output.push(']');
        }
        JsonValue::Object(values) => {
            output.push('{');
            let mut keys = values.keys().collect::<Vec<_>>();
            keys.sort_by(|left, right| utf16_cmp(left, right));
            for (index, key) in keys.into_iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                output.push_str(&serde_json::to_string(key).map_err(|error| {
                    invalid_shared_pdf(format!("Failed to encode shared PDF JSON: {error}"))
                })?);
                output.push(':');
                write_canonical_json(&values[key], output, depth + 1)?;
            }
            output.push('}');
        }
    }
    Ok(())
}

fn utf16_cmp(left: &str, right: &str) -> Ordering {
    let mut left = left.encode_utf16();
    let mut right = right.encode_utf16();
    loop {
        match (left.next(), right.next()) {
            (Some(left), Some(right)) => match left.cmp(&right) {
                Ordering::Equal => {}
                ordering => return ordering,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::collaboration::{
        NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
        NativeOfficeCollaborationMode,
    };

    #[test]
    fn recursive_record_fields_round_trip_with_atomic_arrays_and_empty_objects() {
        let doc = yrs::Doc::with_client_id(7);
        let manifest = manifest();
        let record = serde_json::json!({
            "id": "annotation-1",
            "pageIndex": 1,
            "source": "created",
            "annotation": {
                "id": "annotation-1",
                "pageIndex": 1,
                "type": 3,
                "appearance": {
                    "color": "#ffd400",
                    "metadata": {},
                },
                "points": [{ "x": 1, "y": 2 }],
            },
        });

        append_record(&doc, &manifest, "annotations", &record, "annotation", false).unwrap();

        let roots = PdfRecordCollectionRoots::new(&doc, &manifest, "annotations");
        let fields = &roots.fields;
        let transaction = doc.transact();
        let object_key = encoded_record_entry_key(
            "annotation-1",
            PdfRecordFieldKind::Object,
            &["annotation".to_owned()],
        )
        .unwrap();
        assert!(matches!(
            fields.get(&transaction, object_key.as_str()),
            Some(Out::Any(Any::Bool(true)))
        ));
        let empty_object_key = encoded_record_entry_key(
            "annotation-1",
            PdfRecordFieldKind::Object,
            &[
                "annotation".to_owned(),
                "appearance".to_owned(),
                "metadata".to_owned(),
            ],
        )
        .unwrap();
        assert!(matches!(
            fields.get(&transaction, empty_object_key.as_str()),
            Some(Out::Any(Any::Bool(true)))
        ));
        let points_key = encoded_record_entry_key(
            "annotation-1",
            PdfRecordFieldKind::Value,
            &["annotation".to_owned(), "points".to_owned()],
        )
        .unwrap();
        assert!(matches!(
            fields.get(&transaction, points_key.as_str()),
            Some(Out::Any(Any::Array(_)))
        ));
        drop(transaction);

        let records = read_pdf_records(&doc, &roots, "annotation").unwrap();
        assert_eq!(records.order, ["annotation-1"]);
        assert_eq!(
            canonical_json(&records.by_id["annotation-1"]).unwrap(),
            canonical_json(&record).unwrap()
        );
    }

    #[test]
    fn recursive_record_reader_rejects_a_missing_object_marker() {
        let doc = yrs::Doc::with_client_id(7);
        let manifest = manifest();
        let record = serde_json::json!({
            "id": "annotation-1",
            "annotation": {
                "id": "annotation-1",
                "pageIndex": 0,
                "type": 3,
            },
            "pageIndex": 0,
            "source": "created",
        });
        append_record(&doc, &manifest, "annotations", &record, "annotation", false).unwrap();
        let roots = PdfRecordCollectionRoots::new(&doc, &manifest, "annotations");
        let object_key = encoded_record_entry_key(
            "annotation-1",
            PdfRecordFieldKind::Object,
            &["annotation".to_owned()],
        )
        .unwrap();
        roots
            .fields
            .remove(&mut doc.transact_mut(), object_key.as_str());

        let error = read_pdf_records(&doc, &roots, "annotation").unwrap_err();
        assert_eq!(error.code, "office.collaboration.content_invalid");
        assert!(error.message.contains("without an object marker"));
    }

    fn manifest() -> NativeOfficeCollaborationManifest {
        NativeOfficeCollaborationManifest {
            format: String::new(),
            schema_version: 1,
            protocol: String::new(),
            protocol_version: 1,
            namespace: "a3s.office".to_owned(),
            artifact_id: "fixture-pdf".to_owned(),
            kind: NativeOfficeCollaborationArtifactKind::Pdf,
            actor_id: "agent-alpha".to_owned(),
            actor_kind: NativeOfficeCollaborationActorKind::Agent,
            mode: NativeOfficeCollaborationMode::Edit,
            client_id: 7,
        }
    }
}
