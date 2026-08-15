use std::collections::HashSet;

use a3s_use_core::UseResult;
use serde_json::{Map as JsonMap, Number as JsonNumber, Value as JsonValue};

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationPdfRect, NativeOfficeCollaborationPdfReviewDecision,
    NativeOfficeCollaborationPdfReviewTargetKind,
};
use super::records::{
    append_record, assert_allowed_keys, assert_exact_keys, canonical_json, read_pdf_claims,
    read_pdf_records, read_pdf_source_page_count, required_json_object, required_json_string,
    PdfClaims, PdfRecords, MAX_PDF_PAGES,
};
use super::{
    invalid_shared_pdf, validate_pdf_identifier, validate_shared_pdf_identifier,
    PdfRecordCollectionRoots,
};

const MAX_PDF_REDACTION_RECTS: usize = 10_000;

pub(super) fn validate_pdf_review_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PdfProposeRedaction {
            proposal_id,
            rects,
            proposed_at,
            ..
        } => {
            validate_pdf_identifier(proposal_id, "proposalId", "PDF redaction proposal")?;
            if !(1..=MAX_PDF_REDACTION_RECTS).contains(&rects.len()) {
                return Err(collaboration_error(
                    "office.collaboration.mutation_invalid",
                    "A PDF redaction proposal must contain between 1 and 10,000 rectangles.",
                )
                .with_detail("rectCount", rects.len() as u64));
            }
            for (index, rect) in rects.iter().enumerate() {
                if !valid_pdf_rect(rect) {
                    return Err(collaboration_error(
                        "office.collaboration.mutation_invalid",
                        "Every PDF redaction rectangle must contain finite coordinates with positive width and height.",
                    )
                    .with_detail("rectIndex", index as u64));
                }
            }
            validate_native_timestamp(proposed_at, "proposedAt", "redaction proposal")
        }
        NativeOfficeCollaborationMutation::PdfProposePageRotation {
            page_operation_id,
            page_indices,
            degrees,
            proposed_at,
        } => {
            validate_page_operation_input(page_operation_id, page_indices, proposed_at)?;
            if !matches!(degrees, 90 | 180 | 270) {
                return Err(collaboration_error(
                    "office.collaboration.mutation_invalid",
                    "A PDF page rotation must be 90, 180, or 270 degrees clockwise.",
                )
                .with_detail("degrees", u64::from(*degrees)));
            }
            Ok(())
        }
        NativeOfficeCollaborationMutation::PdfProposePageDeletion {
            page_operation_id,
            page_indices,
            proposed_at,
        } => validate_page_operation_input(page_operation_id, page_indices, proposed_at),
        NativeOfficeCollaborationMutation::PdfProposePageReorder {
            page_operation_id,
            page_order,
            proposed_at,
        } => validate_page_operation_input(page_operation_id, page_order, proposed_at),
        NativeOfficeCollaborationMutation::PdfDecideReview {
            decision_id,
            target_id,
            created_at,
            ..
        } => {
            validate_pdf_identifier(decision_id, "decisionId", "PDF review decision")?;
            validate_pdf_identifier(target_id, "targetId", "PDF review target")?;
            validate_native_timestamp(created_at, "createdAt", "review decision")
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a PDF review mutation.",
        )),
    }
}

pub(super) fn apply_pdf_review_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let state = read_pdf_review_state(doc, manifest)?;
    match mutation {
        NativeOfficeCollaborationMutation::PdfProposeRedaction {
            proposal_id,
            page_index,
            rects,
            proposed_at,
            reason,
            text,
        } => propose_redaction(
            doc,
            manifest,
            state,
            proposal_id,
            *page_index,
            rects,
            proposed_at,
            reason.as_deref(),
            text.as_deref(),
        )?,
        NativeOfficeCollaborationMutation::PdfProposePageRotation {
            page_operation_id,
            page_indices,
            degrees,
            proposed_at,
        } => propose_page_operation(
            doc,
            manifest,
            state,
            page_operation_id,
            NativePageOperation::Rotate {
                page_indices,
                degrees: *degrees,
            },
            proposed_at,
        )?,
        NativeOfficeCollaborationMutation::PdfProposePageDeletion {
            page_operation_id,
            page_indices,
            proposed_at,
        } => propose_page_operation(
            doc,
            manifest,
            state,
            page_operation_id,
            NativePageOperation::Delete { page_indices },
            proposed_at,
        )?,
        NativeOfficeCollaborationMutation::PdfProposePageReorder {
            page_operation_id,
            page_order,
            proposed_at,
        } => propose_page_operation(
            doc,
            manifest,
            state,
            page_operation_id,
            NativePageOperation::Reorder { page_order },
            proposed_at,
        )?,
        NativeOfficeCollaborationMutation::PdfDecideReview {
            decision_id,
            target_kind,
            target_id,
            decision,
            created_at,
        } => decide_review(
            doc,
            manifest,
            state,
            decision_id,
            *target_kind,
            target_id,
            *decision,
            created_at,
        )?,
        _ => {
            return Err(collaboration_error(
                "office.collaboration.mutation_invalid",
                "The supplied mutation is not a PDF review mutation.",
            ))
        }
    }
    read_pdf_review_state(doc, manifest)?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn propose_redaction(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    state: PdfReviewState,
    proposal_id: &str,
    page_index: u32,
    rects: &[NativeOfficeCollaborationPdfRect],
    proposed_at: &str,
    reason: Option<&str>,
    text: Option<&str>,
) -> UseResult<()> {
    if page_index >= state.page_count {
        return Err(collaboration_error(
            "office.collaboration.mutation_range_invalid",
            format!(
                "PDF redaction page index {page_index} is outside the source page range 0..{}.",
                state.page_count
            ),
        )
        .with_detail("pageIndex", page_index as u64)
        .with_detail("pageCount", state.page_count as u64));
    }
    let record = redaction_record(
        proposal_id,
        page_index,
        rects,
        &manifest.actor_id,
        proposed_at,
        reason,
        text,
    )?;
    let fingerprint = canonical_json(&record)?;
    if let Some(existing) = state.redactions.by_id.get(proposal_id) {
        if canonical_json(existing)? == fingerprint {
            return Ok(());
        }
        return Err(review_match_conflict(format!(
            "PDF redaction proposal ID '{proposal_id}' already belongs to a different record."
        )));
    }
    let claim = state.claims.claim_for("redaction", proposal_id);
    if claim.is_some_and(|existing| existing != fingerprint) {
        return Err(review_match_conflict(format!(
            "PDF redaction proposal ID '{proposal_id}' was already claimed by a different record."
        )));
    }
    append_record(
        doc,
        manifest,
        "redaction-proposals",
        &record,
        "redaction",
        claim.is_none(),
    )
}

enum NativePageOperation<'a> {
    Rotate {
        page_indices: &'a [u32],
        degrees: u16,
    },
    Delete {
        page_indices: &'a [u32],
    },
    Reorder {
        page_order: &'a [u32],
    },
}

fn propose_page_operation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    state: PdfReviewState,
    page_operation_id: &str,
    operation: NativePageOperation<'_>,
    proposed_at: &str,
) -> UseResult<()> {
    validate_page_operation_range(&operation, state.page_count)?;
    let record = page_operation_record(
        page_operation_id,
        operation,
        &manifest.actor_id,
        proposed_at,
    );
    let fingerprint = canonical_json(&record)?;
    if let Some(existing) = state.page_operations.by_id.get(page_operation_id) {
        if canonical_json(existing)? == fingerprint {
            return Ok(());
        }
        return Err(review_match_conflict(format!(
            "PDF page operation ID '{page_operation_id}' already belongs to a different record."
        )));
    }
    let claim = state.claims.claim_for("page-operation", page_operation_id);
    if claim.is_some_and(|existing| existing != fingerprint) {
        return Err(review_match_conflict(format!(
            "PDF page operation ID '{page_operation_id}' was already claimed by a different record."
        )));
    }
    append_record(
        doc,
        manifest,
        "page-operations",
        &record,
        "page-operation",
        claim.is_none(),
    )
}

#[allow(clippy::too_many_arguments)]
fn decide_review(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    state: PdfReviewState,
    decision_id: &str,
    target_kind: NativeOfficeCollaborationPdfReviewTargetKind,
    target_id: &str,
    decision: NativeOfficeCollaborationPdfReviewDecision,
    created_at: &str,
) -> UseResult<()> {
    let record = decision_record(
        decision_id,
        target_kind,
        target_id,
        decision,
        &manifest.actor_id,
        created_at,
    );
    let fingerprint = canonical_json(&record)?;
    if let Some(existing) = state.decisions.by_id.get(decision_id) {
        if canonical_json(existing)? == fingerprint {
            return Ok(());
        }
        return Err(review_match_conflict(format!(
            "PDF review decision ID '{decision_id}' already belongs to a different record."
        )));
    }

    let targets = match target_kind {
        NativeOfficeCollaborationPdfReviewTargetKind::Redaction => &state.redactions.by_id,
        NativeOfficeCollaborationPdfReviewTargetKind::PageOperation => &state.page_operations.by_id,
    };
    if !targets.contains_key(target_id) {
        return Err(review_match_conflict(format!(
            "PDF review target '{}:{target_id}' does not exist.",
            target_kind.as_str()
        )));
    }
    if state
        .decided_targets
        .contains(&(target_kind.as_str().to_owned(), target_id.to_owned()))
    {
        return Err(review_match_conflict(format!(
            "PDF review target '{}:{target_id}' already has a final decision.",
            target_kind.as_str()
        )));
    }

    let claim = state.claims.claim_for("review-decision", decision_id);
    if claim.is_some_and(|existing| existing != fingerprint) {
        return Err(review_match_conflict(format!(
            "PDF review decision ID '{decision_id}' was already claimed by a different record."
        )));
    }
    append_record(
        doc,
        manifest,
        "review-decisions",
        &record,
        "review-decision",
        claim.is_none(),
    )
}

struct PdfReviewState {
    page_count: u32,
    redactions: PdfRecords,
    page_operations: PdfRecords,
    decisions: PdfRecords,
    decided_targets: HashSet<(String, String)>,
    claims: PdfClaims,
}

fn read_pdf_review_state(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<PdfReviewState> {
    let page_count = read_pdf_source_page_count(doc, manifest)?;
    let redactions = read_pdf_records(
        doc,
        &PdfRecordCollectionRoots::new(doc, manifest, "redaction-proposals"),
        "redaction proposal",
    )?;
    let page_operations = read_pdf_records(
        doc,
        &PdfRecordCollectionRoots::new(doc, manifest, "page-operations"),
        "page operation",
    )?;
    let decisions = read_pdf_records(
        doc,
        &PdfRecordCollectionRoots::new(doc, manifest, "review-decisions"),
        "review decision",
    )?;

    for id in &redactions.order {
        validate_redaction_record(&redactions.by_id[id], id, page_count)?;
    }
    for id in &page_operations.order {
        validate_page_operation_record(&page_operations.by_id[id], id, page_count)?;
    }
    let mut decided_targets = HashSet::new();
    for id in &decisions.order {
        let (target_kind, target_id) = validate_decision_record(&decisions.by_id[id], id)?;
        let targets = if target_kind == "redaction" {
            &redactions.by_id
        } else {
            &page_operations.by_id
        };
        if !targets.contains_key(&target_id) {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF review decision '{id}' references a missing {target_kind}."
            )));
        }
        if !decided_targets.insert((target_kind.to_owned(), target_id.clone())) {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF target '{target_kind}:{target_id}' has more than one final decision."
            )));
        }
    }

    let claims = read_pdf_claims(doc, manifest)?;
    claims.assert_exact_records("redaction", &redactions)?;
    claims.assert_exact_records("page-operation", &page_operations)?;
    claims.assert_exact_records("review-decision", &decisions)?;
    Ok(PdfReviewState {
        page_count,
        redactions,
        page_operations,
        decisions,
        decided_targets,
        claims,
    })
}

fn exact_u64(value: f64, minimum: u64, maximum: u64) -> Option<u64> {
    if value.fract() != 0.0 || value < minimum as f64 || value > maximum as f64 {
        return None;
    }
    Some(value as u64)
}

fn validate_redaction_record(
    record: &JsonValue,
    expected_id: &str,
    page_count: u32,
) -> UseResult<()> {
    let object = required_json_object(record, "redaction proposal")?;
    assert_allowed_keys(
        object,
        &[
            "id",
            "pageIndex",
            "proposedAt",
            "proposedBy",
            "reason",
            "rects",
            "text",
        ],
        &["id", "pageIndex", "proposedAt", "proposedBy", "rects"],
        "redaction proposal",
    )?;
    validate_record_id(object, expected_id, "redaction proposal")?;
    required_page_index(object.get("pageIndex"), page_count, "redaction proposal")?;
    let rects = object
        .get("rects")
        .and_then(JsonValue::as_array)
        .ok_or_else(|| invalid_shared_pdf("The shared PDF redaction rectangles are invalid."))?;
    if !(1..=MAX_PDF_REDACTION_RECTS).contains(&rects.len()) {
        return Err(invalid_shared_pdf(
            "The shared PDF redaction rectangle count is invalid.",
        ));
    }
    for rect in rects {
        let rect = required_json_object(rect, "redaction rectangle")?;
        assert_exact_keys(
            rect,
            &["bottom", "left", "right", "top"],
            "redaction rectangle",
        )?;
        let left = required_finite_number(rect.get("left"), "redaction left")?;
        let top = required_finite_number(rect.get("top"), "redaction top")?;
        let right = required_finite_number(rect.get("right"), "redaction right")?;
        let bottom = required_finite_number(rect.get("bottom"), "redaction bottom")?;
        if right <= left || bottom <= top {
            return Err(invalid_shared_pdf(
                "The shared PDF contains a non-positive redaction rectangle.",
            ));
        }
    }
    validate_shared_identifier_field(object, "proposedBy", "redaction proposer")?;
    validate_shared_timestamp_field(object, "proposedAt", "redaction proposal")?;
    for field in ["reason", "text"] {
        if object.contains_key(field) {
            required_json_string(object.get(field), field)?;
        }
    }
    Ok(())
}

fn validate_page_operation_record(
    record: &JsonValue,
    expected_id: &str,
    page_count: u32,
) -> UseResult<()> {
    let object = required_json_object(record, "page operation")?;
    validate_record_id(object, expected_id, "page operation")?;
    validate_shared_identifier_field(object, "proposedBy", "page operation proposer")?;
    validate_shared_timestamp_field(object, "proposedAt", "page operation proposal")?;
    let kind = required_json_string(object.get("kind"), "page operation kind")?;
    match kind {
        "rotate" => {
            assert_exact_keys(
                object,
                &[
                    "degrees",
                    "id",
                    "kind",
                    "pageIndices",
                    "proposedAt",
                    "proposedBy",
                ],
                "rotate page operation",
            )?;
            let degrees = required_exact_integer(object.get("degrees"), 90, 270, "rotation")?;
            if !matches!(degrees, 90 | 180 | 270) {
                return Err(invalid_shared_pdf(
                    "The shared PDF page rotation is unsupported.",
                ));
            }
            validate_page_indices(object.get("pageIndices"), page_count, false)?;
        }
        "delete" => {
            assert_exact_keys(
                object,
                &["id", "kind", "pageIndices", "proposedAt", "proposedBy"],
                "delete page operation",
            )?;
            let pages = validate_page_indices(object.get("pageIndices"), page_count, false)?;
            if pages.len() >= page_count as usize {
                return Err(invalid_shared_pdf(
                    "The shared PDF page deletion removes every page.",
                ));
            }
        }
        "reorder" => {
            assert_exact_keys(
                object,
                &["id", "kind", "pageOrder", "proposedAt", "proposedBy"],
                "reorder page operation",
            )?;
            validate_page_indices(object.get("pageOrder"), page_count, true)?;
        }
        _ => {
            return Err(invalid_shared_pdf(
                "The shared PDF page operation kind is unsupported.",
            ))
        }
    }
    Ok(())
}

fn validate_decision_record(record: &JsonValue, expected_id: &str) -> UseResult<(String, String)> {
    let object = required_json_object(record, "review decision")?;
    assert_exact_keys(
        object,
        &[
            "actorId",
            "createdAt",
            "decision",
            "id",
            "targetId",
            "targetKind",
        ],
        "review decision",
    )?;
    validate_record_id(object, expected_id, "review decision")?;
    let target_kind = required_json_string(object.get("targetKind"), "review target kind")?;
    if !matches!(target_kind, "redaction" | "page-operation") {
        return Err(invalid_shared_pdf(
            "The shared PDF review target kind is unsupported.",
        ));
    }
    let target_id = required_json_string(object.get("targetId"), "review target ID")?;
    validate_shared_pdf_identifier(target_id, "review target")?;
    let decision = required_json_string(object.get("decision"), "review decision")?;
    if !matches!(decision, "approve" | "reject") {
        return Err(invalid_shared_pdf(
            "The shared PDF review decision value is unsupported.",
        ));
    }
    validate_shared_identifier_field(object, "actorId", "review actor")?;
    validate_shared_timestamp_field(object, "createdAt", "review decision")?;
    Ok((target_kind.to_owned(), target_id.to_owned()))
}

fn validate_page_indices(
    value: Option<&JsonValue>,
    page_count: u32,
    complete: bool,
) -> UseResult<Vec<u32>> {
    let values = value
        .and_then(JsonValue::as_array)
        .ok_or_else(|| invalid_shared_pdf("The shared PDF page operation indices are invalid."))?;
    if (complete && values.len() != page_count as usize) || (!complete && values.is_empty()) {
        return Err(invalid_shared_pdf(
            "The shared PDF page operation index count is invalid.",
        ));
    }
    let mut seen = HashSet::new();
    let mut pages = Vec::with_capacity(values.len());
    for value in values {
        let page =
            required_exact_integer(Some(value), 0, page_count as u64 - 1, "page index")? as u32;
        if !seen.insert(page) {
            return Err(invalid_shared_pdf(
                "The shared PDF page operation contains duplicate indices.",
            ));
        }
        pages.push(page);
    }
    Ok(pages)
}

fn validate_record_id(
    object: &JsonMap<String, JsonValue>,
    expected_id: &str,
    label: &str,
) -> UseResult<()> {
    let id = required_json_string(object.get("id"), label)?;
    validate_shared_pdf_identifier(id, label)?;
    if id == expected_id {
        return Ok(());
    }
    Err(invalid_shared_pdf(format!(
        "The shared PDF {label} identity does not match its collection key."
    )))
}

fn validate_shared_identifier_field(
    object: &JsonMap<String, JsonValue>,
    field: &str,
    label: &str,
) -> UseResult<()> {
    let value = required_json_string(object.get(field), label)?;
    validate_shared_pdf_identifier(value, label)
}

fn validate_shared_timestamp_field(
    object: &JsonMap<String, JsonValue>,
    field: &str,
    label: &str,
) -> UseResult<()> {
    let value = required_json_string(object.get(field), label)?;
    if valid_iso_timestamp(value) {
        return Ok(());
    }
    Err(invalid_shared_pdf(format!(
        "The shared PDF {label} timestamp is invalid."
    )))
}

fn required_page_index(value: Option<&JsonValue>, page_count: u32, label: &str) -> UseResult<u32> {
    required_exact_integer(value, 0, page_count as u64 - 1, label).map(|value| value as u32)
}

fn required_exact_integer(
    value: Option<&JsonValue>,
    minimum: u64,
    maximum: u64,
    label: &str,
) -> UseResult<u64> {
    let value = value
        .and_then(JsonValue::as_f64)
        .ok_or_else(|| invalid_shared_pdf(format!("The shared PDF {label} is not numeric.")))?;
    exact_u64(value, minimum, maximum).ok_or_else(|| {
        invalid_shared_pdf(format!(
            "The shared PDF {label} is outside its valid range."
        ))
    })
}

fn required_finite_number(value: Option<&JsonValue>, label: &str) -> UseResult<f64> {
    value
        .and_then(JsonValue::as_f64)
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid_shared_pdf(format!("The shared PDF {label} is not finite.")))
}

fn page_operation_record(
    page_operation_id: &str,
    operation: NativePageOperation<'_>,
    proposed_by: &str,
    proposed_at: &str,
) -> JsonValue {
    match operation {
        NativePageOperation::Rotate {
            page_indices,
            degrees,
        } => serde_json::json!({
            "id": page_operation_id,
            "kind": "rotate",
            "pageIndices": page_indices,
            "degrees": degrees,
            "proposedBy": proposed_by,
            "proposedAt": proposed_at,
        }),
        NativePageOperation::Delete { page_indices } => serde_json::json!({
            "id": page_operation_id,
            "kind": "delete",
            "pageIndices": page_indices,
            "proposedBy": proposed_by,
            "proposedAt": proposed_at,
        }),
        NativePageOperation::Reorder { page_order } => serde_json::json!({
            "id": page_operation_id,
            "kind": "reorder",
            "pageOrder": page_order,
            "proposedBy": proposed_by,
            "proposedAt": proposed_at,
        }),
    }
}

fn redaction_record(
    proposal_id: &str,
    page_index: u32,
    rects: &[NativeOfficeCollaborationPdfRect],
    proposed_by: &str,
    proposed_at: &str,
    reason: Option<&str>,
    text: Option<&str>,
) -> UseResult<JsonValue> {
    let rects = rects
        .iter()
        .map(|rect| {
            Ok(serde_json::json!({
                "left": finite_json_number(rect.left)?,
                "top": finite_json_number(rect.top)?,
                "right": finite_json_number(rect.right)?,
                "bottom": finite_json_number(rect.bottom)?,
            }))
        })
        .collect::<UseResult<Vec<_>>>()?;
    let mut record = JsonMap::new();
    record.insert("id".to_owned(), JsonValue::String(proposal_id.to_owned()));
    record.insert("pageIndex".to_owned(), JsonValue::from(page_index));
    record.insert("rects".to_owned(), JsonValue::Array(rects));
    record.insert(
        "proposedBy".to_owned(),
        JsonValue::String(proposed_by.to_owned()),
    );
    record.insert(
        "proposedAt".to_owned(),
        JsonValue::String(proposed_at.to_owned()),
    );
    if let Some(reason) = reason {
        record.insert("reason".to_owned(), JsonValue::String(reason.to_owned()));
    }
    if let Some(text) = text {
        record.insert("text".to_owned(), JsonValue::String(text.to_owned()));
    }
    Ok(JsonValue::Object(record))
}

fn decision_record(
    decision_id: &str,
    target_kind: NativeOfficeCollaborationPdfReviewTargetKind,
    target_id: &str,
    decision: NativeOfficeCollaborationPdfReviewDecision,
    actor_id: &str,
    created_at: &str,
) -> JsonValue {
    serde_json::json!({
        "id": decision_id,
        "targetKind": target_kind.as_str(),
        "targetId": target_id,
        "decision": decision.as_str(),
        "actorId": actor_id,
        "createdAt": created_at,
    })
}

fn finite_json_number(value: f64) -> UseResult<JsonValue> {
    JsonNumber::from_f64(value)
        .map(JsonValue::Number)
        .ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                "A PDF redaction coordinate is not a finite JSON number.",
            )
        })
}

fn valid_pdf_rect(rect: &NativeOfficeCollaborationPdfRect) -> bool {
    [rect.left, rect.top, rect.right, rect.bottom]
        .into_iter()
        .all(f64::is_finite)
        && rect.right > rect.left
        && rect.bottom > rect.top
}

fn validate_page_operation_input(
    page_operation_id: &str,
    pages: &[u32],
    proposed_at: &str,
) -> UseResult<()> {
    validate_pdf_identifier(page_operation_id, "pageOperationId", "PDF page operation")?;
    if pages.is_empty() || pages.len() > MAX_PDF_PAGES as usize {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A PDF page operation must contain between 1 and 1,000,000 page indices.",
        )
        .with_detail("pageIndexCount", pages.len() as u64));
    }
    let mut unique = HashSet::with_capacity(pages.len());
    if pages.iter().any(|page| !unique.insert(*page)) {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A PDF page operation cannot contain duplicate page indices.",
        ));
    }
    validate_native_timestamp(proposed_at, "proposedAt", "page operation")
}

fn validate_page_operation_range(
    operation: &NativePageOperation<'_>,
    page_count: u32,
) -> UseResult<()> {
    let (pages, expected_complete, retains_page) = match operation {
        NativePageOperation::Rotate { page_indices, .. } => (*page_indices, false, true),
        NativePageOperation::Delete { page_indices } => (*page_indices, false, false),
        NativePageOperation::Reorder { page_order } => (*page_order, true, true),
    };
    if pages.iter().any(|page| *page >= page_count)
        || (expected_complete && pages.len() != page_count as usize)
        || (!retains_page && pages.len() >= page_count as usize)
    {
        return Err(collaboration_error(
            "office.collaboration.mutation_range_invalid",
            "The PDF page operation does not describe a valid subset or complete permutation of the immutable source pages.",
        )
        .with_detail("pageCount", page_count as u64)
        .with_detail("pageIndexCount", pages.len() as u64));
    }
    Ok(())
}

fn validate_native_timestamp(value: &str, field: &str, label: &str) -> UseResult<()> {
    if valid_canonical_utc_timestamp(value) {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "A native PDF {label} timestamp must use canonical UTC form YYYY-MM-DDTHH:mm:ss.sssZ."
        ),
    )
    .with_detail(field, value.to_owned()))
}

fn valid_canonical_utc_timestamp(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 24
        && bytes[10] == b'T'
        && bytes[13] == b':'
        && bytes[16] == b':'
        && bytes[19] == b'.'
        && bytes[23] == b'Z'
        && valid_date(&bytes[..10])
        && valid_time(&bytes[11..19])
        && bytes[20..23].iter().all(u8::is_ascii_digit)
}

fn valid_iso_timestamp(value: &str) -> bool {
    let bytes = value.as_bytes();
    if bytes.len() == 10 {
        return valid_date(bytes);
    }
    if bytes.len() < 20 || bytes.get(10) != Some(&b'T') || !valid_date(&bytes[..10]) {
        return false;
    }
    let zone_start = if bytes.last() == Some(&b'Z') {
        bytes.len() - 1
    } else if bytes.len() >= 6
        && matches!(bytes[bytes.len() - 6], b'+' | b'-')
        && bytes[bytes.len() - 3] == b':'
        && parse_two(&bytes[bytes.len() - 5..bytes.len() - 3]).is_some_and(|hour| hour <= 23)
        && parse_two(&bytes[bytes.len() - 2..]).is_some_and(|minute| minute <= 59)
    {
        bytes.len() - 6
    } else {
        return false;
    };
    let time = &bytes[11..zone_start];
    if time.len() < 8 || !valid_time(&time[..8]) {
        return false;
    }
    time.len() == 8
        || (time.get(8) == Some(&b'.')
            && time.len() > 9
            && time[9..].iter().all(u8::is_ascii_digit))
}

fn valid_date(value: &[u8]) -> bool {
    if value.len() != 10 || value[4] != b'-' || value[7] != b'-' {
        return false;
    }
    let Some(year) = parse_four(&value[..4]) else {
        return false;
    };
    let Some(month) = parse_two(&value[5..7]) else {
        return false;
    };
    let Some(day) = parse_two(&value[8..]) else {
        return false;
    };
    (1..=12).contains(&month) && (1..=days_in_month(year, month)).contains(&day)
}

fn valid_time(value: &[u8]) -> bool {
    value.len() == 8
        && value[2] == b':'
        && value[5] == b':'
        && parse_two(&value[..2]).is_some_and(|hour| hour <= 23)
        && parse_two(&value[3..5]).is_some_and(|minute| minute <= 59)
        && parse_two(&value[6..]).is_some_and(|second| second <= 59)
}

fn parse_two(value: &[u8]) -> Option<u32> {
    if value.len() != 2 || !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    Some(u32::from(value[0] - b'0') * 10 + u32::from(value[1] - b'0'))
}

fn parse_four(value: &[u8]) -> Option<u32> {
    if value.len() != 4 || !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    Some(
        value
            .iter()
            .fold(0, |result, digit| result * 10 + u32::from(*digit - b'0')),
    )
}

fn days_in_month(year: u32, month: u32) -> u32 {
    match month {
        2 if year % 400 == 0 || (year % 4 == 0 && year % 100 != 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

fn review_match_conflict(message: String) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message).with_suggestion(
        "Read the latest PDF review records and choose a fresh stable ID or target before retrying.",
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn canonical_json_matches_ecmascript_numbers_and_utf16_key_order() {
        let value = serde_json::json!({
            "\u{e000}": 90.0,
            "\u{10000}": 0.0000001,
            "nested": { "z": -0.0, "a": 1e21 },
        });
        assert_eq!(
            canonical_json(&value).unwrap(),
            "{\"nested\":{\"a\":1e+21,\"z\":0},\"𐀀\":1e-7,\"\":90}"
        );
    }

    #[test]
    fn native_timestamps_are_canonical_and_calendar_valid() {
        assert!(valid_canonical_utc_timestamp("2024-02-29T23:59:59.999Z"));
        assert!(!valid_canonical_utc_timestamp("2023-02-29T23:59:59.999Z"));
        assert!(!valid_canonical_utc_timestamp("2024-02-29T23:59:59Z"));
        assert!(!valid_canonical_utc_timestamp("2024-02-29T24:00:00.000Z"));
    }
}
