use std::collections::BTreeSet;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationPdfAnnotationSource,
};
use super::records::{
    append_record_with_fingerprint, assert_allowed_keys, canonical_json, json_equal, patch_record,
    read_pdf_claims, read_pdf_records, read_pdf_source_page_count, required_json_object,
    required_json_string, PdfClaims, PdfRecords,
};
use super::{
    invalid_shared_pdf, validate_pdf_identifier, validate_shared_pdf_identifier,
    PdfRecordCollectionRoots,
};

const MAX_JSON_DEPTH: usize = 128;
const SUPPORTED_ANNOTATION_TYPES: [u32; 5] = [3, 9, 10, 12, 15];

pub(super) fn validate_pdf_annotation_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PdfCreateAnnotation {
            annotation_id,
            page_index,
            annotation,
        } => {
            validate_pdf_identifier(annotation_id, "annotationId", "PDF annotation")?;
            let identity = validate_annotation_input(annotation, "annotation")?;
            assert_input_identity(annotation_id, *page_index, identity, "created annotation")
        }
        NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
            annotation_id,
            expected_annotation,
            next_annotation,
        } => {
            validate_pdf_identifier(annotation_id, "annotationId", "PDF annotation")?;
            let expected = validate_annotation_input(expected_annotation, "expected annotation")?;
            let next = validate_annotation_input(next_annotation, "next annotation")?;
            if expected.id != annotation_id || next.id != annotation_id {
                return Err(invalid_annotation_mutation(
                    "A PDF annotation update must retain its requested annotation ID.",
                ));
            }
            if expected.page_index != next.page_index
                || expected.annotation_type != next.annotation_type
            {
                return Err(invalid_annotation_mutation(
                    "A PDF annotation update cannot change its page or type identity.",
                ));
            }
            Ok(())
        }
        NativeOfficeCollaborationMutation::PdfDeleteAnnotation {
            annotation_id,
            expected_type,
            ..
        } => {
            validate_pdf_identifier(annotation_id, "annotationId", "PDF annotation")?;
            validate_supported_annotation_type(*expected_type, "expectedType")
        }
        _ => Err(invalid_annotation_mutation(
            "The supplied mutation is not a PDF annotation mutation.",
        )),
    }
}

pub(super) fn apply_pdf_annotation_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let state = read_pdf_annotation_state(doc, manifest)?;
    match mutation {
        NativeOfficeCollaborationMutation::PdfCreateAnnotation {
            annotation_id,
            page_index,
            annotation,
        } => create_annotation(doc, manifest, state, annotation_id, *page_index, annotation)?,
        NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
            annotation_id,
            expected_annotation,
            next_annotation,
        } => update_annotation(
            doc,
            manifest,
            state,
            annotation_id,
            expected_annotation,
            next_annotation,
        )?,
        NativeOfficeCollaborationMutation::PdfDeleteAnnotation {
            annotation_id,
            expected_source,
            expected_page_index,
            expected_type,
        } => delete_annotation(
            doc,
            manifest,
            state,
            annotation_id,
            *expected_source,
            *expected_page_index,
            *expected_type,
        )?,
        _ => {
            return Err(invalid_annotation_mutation(
                "The supplied mutation is not a PDF annotation mutation.",
            ))
        }
    }
    read_pdf_annotation_state(doc, manifest)?;
    Ok(())
}

fn create_annotation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    state: PdfAnnotationState,
    annotation_id: &str,
    page_index: u32,
    annotation: &JsonValue,
) -> UseResult<()> {
    if page_index >= state.page_count {
        return Err(annotation_range_error(page_index, state.page_count));
    }
    let record = serde_json::json!({
        "id": annotation_id,
        "pageIndex": page_index,
        "source": "created",
        "annotation": annotation,
    });
    let fingerprint = canonical_json(&record)?;
    if state.annotations.by_id.contains_key(annotation_id) {
        if state.claims.claim_for("annotation", annotation_id) == Some(fingerprint.as_str()) {
            return Ok(());
        }
        return Err(annotation_match_conflict(format!(
            "PDF annotation ID '{annotation_id}' already belongs to a different record."
        )));
    }
    let claim = state.claims.claim_for("annotation", annotation_id);
    if claim.is_some_and(|existing| existing != fingerprint) {
        return Err(annotation_match_conflict(format!(
            "PDF annotation ID '{annotation_id}' was already claimed by a different record."
        )));
    }
    append_record_with_fingerprint(
        doc,
        manifest,
        "annotations",
        &record,
        "annotation",
        &record,
        claim.is_none(),
    )
}

fn update_annotation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    state: PdfAnnotationState,
    annotation_id: &str,
    expected_annotation: &JsonValue,
    next_annotation: &JsonValue,
) -> UseResult<()> {
    let shared_record = state.annotations.by_id.get(annotation_id).ok_or_else(|| {
        annotation_match_conflict(format!(
            "PDF annotation ID '{annotation_id}' does not exist."
        ))
    })?;
    let shared_identity =
        shared_annotation_identity(shared_record, annotation_id, state.page_count)?;
    let expected_identity = validate_annotation_input(expected_annotation, "expected annotation")?;
    if expected_identity.page_index != shared_identity.page_index
        || expected_identity.annotation_type != shared_identity.annotation_type
    {
        return Err(annotation_match_conflict(format!(
            "PDF annotation ID '{annotation_id}' no longer matches the expected page and type identity."
        )));
    }
    let shared_annotation = required_json_object(shared_record, "annotation record")?
        .get("annotation")
        .expect("validated annotation record");
    assert_compatible_value(
        expected_annotation,
        next_annotation,
        shared_annotation,
        &format!("PDF annotation '{annotation_id}'"),
    )?;
    if json_equal(shared_annotation, next_annotation)? {
        return Ok(());
    }

    let mut previous_record = shared_record.clone();
    previous_record
        .as_object_mut()
        .expect("validated annotation record")
        .insert("annotation".to_owned(), expected_annotation.clone());
    let mut next_record = shared_record.clone();
    next_record
        .as_object_mut()
        .expect("validated annotation record")
        .insert("annotation".to_owned(), next_annotation.clone());
    patch_record(
        doc,
        manifest,
        "annotations",
        annotation_id,
        &previous_record,
        shared_record,
        &next_record,
    )
}

#[allow(clippy::too_many_arguments)]
fn delete_annotation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    state: PdfAnnotationState,
    annotation_id: &str,
    expected_source: NativeOfficeCollaborationPdfAnnotationSource,
    expected_page_index: u32,
    expected_type: u32,
) -> UseResult<()> {
    let shared_record = state.annotations.by_id.get(annotation_id).ok_or_else(|| {
        annotation_match_conflict(format!(
            "PDF annotation ID '{annotation_id}' does not exist."
        ))
    })?;
    let identity = shared_annotation_identity(shared_record, annotation_id, state.page_count)?;
    if identity.source != expected_source.as_str()
        || identity.page_index != expected_page_index
        || identity.annotation_type != expected_type
    {
        return Err(annotation_match_conflict(format!(
            "PDF annotation ID '{annotation_id}' no longer matches its expected immutable identity."
        )));
    }
    if identity.deleted {
        return Ok(());
    }
    let mut next = shared_record.clone();
    next.as_object_mut()
        .expect("validated annotation record")
        .insert("deleted".to_owned(), JsonValue::Bool(true));
    patch_record(
        doc,
        manifest,
        "annotations",
        annotation_id,
        shared_record,
        shared_record,
        &next,
    )
}

struct PdfAnnotationState {
    page_count: u32,
    annotations: PdfRecords,
    claims: PdfClaims,
}

fn read_pdf_annotation_state(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<PdfAnnotationState> {
    let page_count = read_pdf_source_page_count(doc, manifest)?;
    let annotations = read_pdf_records(
        doc,
        &PdfRecordCollectionRoots::new(doc, manifest, "annotations"),
        "annotation",
    )?;
    for id in &annotations.order {
        shared_annotation_identity(&annotations.by_id[id], id, page_count)?;
    }
    let claims = read_pdf_claims(doc, manifest)?;
    assert_annotation_claims(&claims, &annotations, page_count)?;
    Ok(PdfAnnotationState {
        page_count,
        annotations,
        claims,
    })
}

#[derive(Debug)]
struct AnnotationIdentity<'a> {
    id: &'a str,
    page_index: u32,
    annotation_type: u32,
}

#[derive(Debug)]
struct SharedAnnotationIdentity<'a> {
    source: &'a str,
    page_index: u32,
    annotation_type: u32,
    deleted: bool,
}

fn validate_annotation_input<'a>(
    annotation: &'a JsonValue,
    label: &str,
) -> UseResult<AnnotationIdentity<'a>> {
    validate_input_json(annotation, label, 0)?;
    let object = annotation.as_object().ok_or_else(|| {
        invalid_annotation_mutation(format!("A PDF {label} must be a JSON object."))
    })?;
    let id = object
        .get("id")
        .and_then(JsonValue::as_str)
        .ok_or_else(|| {
            invalid_annotation_mutation(format!("A PDF {label} must contain a string ID."))
        })?;
    validate_pdf_identifier(id, "annotation.id", "PDF annotation")?;
    let page_index = input_exact_u32(object.get("pageIndex")).ok_or_else(|| {
        invalid_annotation_mutation(format!(
            "A PDF {label} must contain a non-negative integer page index."
        ))
    })?;
    let annotation_type = input_exact_u32(object.get("type")).ok_or_else(|| {
        invalid_annotation_mutation(format!(
            "A PDF {label} must contain an integer annotation type."
        ))
    })?;
    validate_supported_annotation_type(annotation_type, "annotation.type")?;
    Ok(AnnotationIdentity {
        id,
        page_index,
        annotation_type,
    })
}

fn assert_input_identity(
    annotation_id: &str,
    page_index: u32,
    identity: AnnotationIdentity<'_>,
    label: &str,
) -> UseResult<()> {
    if identity.id == annotation_id && identity.page_index == page_index {
        return Ok(());
    }
    Err(invalid_annotation_mutation(format!(
        "A PDF {label} must repeat its requested ID and page index inside the portable annotation value."
    )))
}

fn shared_annotation_identity<'a>(
    record: &'a JsonValue,
    expected_id: &str,
    page_count: u32,
) -> UseResult<SharedAnnotationIdentity<'a>> {
    let object = required_json_object(record, "annotation record")?;
    assert_allowed_keys(
        object,
        &["annotation", "deleted", "id", "pageIndex", "source"],
        &["annotation", "id", "pageIndex", "source"],
        "annotation record",
    )?;
    let id = required_json_string(object.get("id"), "annotation ID")?;
    validate_shared_pdf_identifier(id, "annotation")?;
    if id != expected_id {
        return Err(invalid_shared_pdf(
            "The shared PDF annotation identity does not match its collection key.",
        ));
    }
    let page_index = shared_exact_u32(object.get("pageIndex"), "annotation page index")?;
    if page_index >= page_count {
        return Err(invalid_shared_pdf(
            "The shared PDF annotation page index is outside the source page range.",
        ));
    }
    let source = required_json_string(object.get("source"), "annotation source")?;
    if !matches!(source, "base" | "created") {
        return Err(invalid_shared_pdf(
            "The shared PDF annotation source is unsupported.",
        ));
    }
    let annotation = required_json_object(
        object
            .get("annotation")
            .expect("required annotation field checked above"),
        "annotation value",
    )?;
    let inner_id = required_json_string(annotation.get("id"), "annotation value ID")?;
    let inner_page = shared_exact_u32(annotation.get("pageIndex"), "annotation value page index")?;
    if inner_id != id || inner_page != page_index {
        return Err(invalid_shared_pdf(
            "The shared PDF annotation value identity does not match its record.",
        ));
    }
    let annotation_type = shared_exact_u32(annotation.get("type"), "annotation type")?;
    if !SUPPORTED_ANNOTATION_TYPES.contains(&annotation_type) {
        return Err(invalid_shared_pdf(
            "The shared PDF annotation type is unsupported.",
        ));
    }
    let deleted = match object.get("deleted") {
        None => false,
        Some(JsonValue::Bool(true)) => true,
        Some(_) => {
            return Err(invalid_shared_pdf(
                "The shared PDF annotation deletion tombstone is invalid.",
            ))
        }
    };
    Ok(SharedAnnotationIdentity {
        source,
        page_index,
        annotation_type,
        deleted,
    })
}

fn assert_annotation_claims(
    claims: &PdfClaims,
    annotations: &PdfRecords,
    page_count: u32,
) -> UseResult<()> {
    for id in &annotations.order {
        let record = &annotations.by_id[id];
        let identity = shared_annotation_identity(record, id, page_count)?;
        let fingerprint = claims.claim_for("annotation", id).ok_or_else(|| {
            invalid_shared_pdf(format!(
                "The shared PDF annotation record ID '{id}' is missing its immutable claim."
            ))
        })?;
        let claimed: JsonValue = serde_json::from_str(fingerprint).map_err(|_| {
            invalid_shared_pdf("The shared PDF annotation claim fingerprint is not valid JSON.")
        })?;
        let claimed = required_json_object(&claimed, "annotation claim fingerprint")?;
        let claimed_id = required_json_string(claimed.get("id"), "annotation claim ID")?;
        let claimed_page =
            shared_exact_u32(claimed.get("pageIndex"), "annotation claim page index")?;
        let claimed_source =
            required_json_string(claimed.get("source"), "annotation claim source")?;
        if claimed_id != id
            || claimed_page != identity.page_index
            || claimed_source != identity.source
        {
            return Err(invalid_shared_pdf(
                "The shared PDF annotation immutable identity claim is invalid.",
            ));
        }
        let claimed_type = claimed.get("type").or_else(|| {
            claimed
                .get("annotation")
                .and_then(JsonValue::as_object)
                .and_then(|annotation| annotation.get("type"))
        });
        if identity.source == "created" && claimed_type.is_none() {
            return Err(invalid_shared_pdf(
                "The shared PDF created annotation claim omits its type identity.",
            ));
        }
        if let Some(claimed_type) = claimed_type {
            let claimed_type = shared_exact_u32(Some(claimed_type), "annotation claim type")?;
            if claimed_type != identity.annotation_type {
                return Err(invalid_shared_pdf(
                    "The shared PDF annotation immutable identity claim is invalid.",
                ));
            }
        }
    }
    Ok(())
}

fn assert_compatible_value(
    previous: &JsonValue,
    next: &JsonValue,
    shared: &JsonValue,
    label: &str,
) -> UseResult<()> {
    if json_equal(previous, next)? {
        return Ok(());
    }
    if let (Some(previous), Some(next), Some(shared)) =
        (previous.as_object(), next.as_object(), shared.as_object())
    {
        let keys = previous
            .keys()
            .chain(next.keys())
            .cloned()
            .collect::<BTreeSet<_>>();
        for key in keys {
            let before = previous.get(&key);
            let after = next.get(&key);
            let current = shared.get(&key);
            let field_label = format!("{label} field '{key}'");
            match (before, after, current) {
                (Some(before), None, Some(current)) if !json_equal(before, current)? => {
                    return Err(annotation_match_conflict(format!(
                        "The {field_label} changed before it could be removed."
                    )))
                }
                (_, None, _) => {}
                (None, Some(after), Some(current)) if !json_equal(after, current)? => {
                    return Err(annotation_match_conflict(format!(
                        "The {field_label} was added concurrently with a different value."
                    )))
                }
                (None, Some(_), _) => {}
                (Some(before), Some(after), _) if json_equal(before, after)? => {}
                (Some(_), Some(_), None) => {
                    return Err(annotation_match_conflict(format!(
                        "The {field_label} was removed before this change could be applied."
                    )))
                }
                (Some(before), Some(after), Some(current)) => {
                    assert_compatible_value(before, after, current, &field_label)?;
                }
            }
        }
        return Ok(());
    }
    if !json_equal(previous, shared)? && !json_equal(next, shared)? {
        return Err(annotation_match_conflict(format!(
            "The {label} changed concurrently."
        )));
    }
    Ok(())
}

fn validate_input_json(value: &JsonValue, label: &str, depth: usize) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_annotation_mutation(format!(
            "A PDF {label} exceeds the JSON nesting limit."
        )));
    }
    match value {
        JsonValue::Array(values) => {
            for value in values {
                validate_input_json(value, label, depth + 1)?;
            }
        }
        JsonValue::Object(values) => {
            for (key, value) in values {
                if key.is_empty()
                    || matches!(key.as_str(), "__proto__" | "constructor" | "prototype")
                {
                    return Err(invalid_annotation_mutation(format!(
                        "A PDF {label} contains an empty or unsafe JSON key."
                    )));
                }
                validate_input_json(value, label, depth + 1)?;
            }
        }
        JsonValue::Number(value) if value.as_f64().is_none_or(|value| !value.is_finite()) => {
            return Err(invalid_annotation_mutation(format!(
                "A PDF {label} contains a non-finite number."
            )))
        }
        _ => {}
    }
    Ok(())
}

fn input_exact_u32(value: Option<&JsonValue>) -> Option<u32> {
    let value = value?.as_f64()?;
    if !value.is_finite() || value.fract() != 0.0 || !(0.0..=u32::MAX as f64).contains(&value) {
        return None;
    }
    Some(value as u32)
}

fn shared_exact_u32(value: Option<&JsonValue>, label: &str) -> UseResult<u32> {
    input_exact_u32(value)
        .ok_or_else(|| invalid_shared_pdf(format!("The shared PDF {label} is invalid.")))
}

fn validate_supported_annotation_type(annotation_type: u32, field: &str) -> UseResult<()> {
    if SUPPORTED_ANNOTATION_TYPES.contains(&annotation_type) {
        return Ok(());
    }
    Err(invalid_annotation_mutation(
        "PDF collaboration supports FreeText (3), Highlight (9), Underline (10), StrikeOut (12), and Ink (15) annotations.",
    )
    .with_detail(field, annotation_type as u64))
}

fn annotation_range_error(page_index: u32, page_count: u32) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_range_invalid",
        format!(
            "PDF annotation page index {page_index} is outside the source page range 0..{page_count}."
        ),
    )
    .with_detail("pageIndex", page_index as u64)
    .with_detail("pageCount", page_count as u64)
}

fn invalid_annotation_mutation(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}

fn annotation_match_conflict(message: String) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message).with_suggestion(
        "Read the latest PDF annotation record and retry with a fresh expected value or stable ID.",
    )
}
