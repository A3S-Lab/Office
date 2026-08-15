use std::collections::HashSet;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::updates::decoder::Decode;
use yrs::{Any, Array, Doc, Map, Out, Transact, Update};

use super::*;

mod convergence;
mod support;

use support::*;

const YJS_PDF_UPDATE_BASE64: &str = "AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3C2ZpeHR1cmUtcGRmKAETYTNzLm9mZmljZS5tZXRhZGF0YQRraW5kAXcDcGRmIQETYTNzLm9mZmljZS5tZXRhZGF0YQtpbml0aWFsaXplZAEIASFhM3Mub2ZmaWNlLmJvb3RzdHJhcC5pbml0aWFsaXplcnMBdxA0MjQyNDI6YW5vbnltb3VzKAEVYTNzLm9mZmljZS5wZGYuc291cmNlBnNoYTI1NgF3QDAxMjM0NTY3ODlhYmNkZWYwMTIzNDU2Nzg5YWJjZGVmMDEyMzQ1Njc4OWFiY2RlZjAxMjM0NTY3ODlhYmNkZWYoARVhM3Mub2ZmaWNlLnBkZi5zb3VyY2UKYnl0ZUxlbmd0aAF9gEAoARVhM3Mub2ZmaWNlLnBkZi5zb3VyY2UJcGFnZUNvdW50AX0DCAEgYTNzLm9mZmljZS5wZGYuc291cmNlLWlkZW50aXRpZXMBd217ImJ5dGVMZW5ndGgiOjQwOTYsInBhZ2VDb3VudCI6Mywic2hhMjU2IjoiMDEyMzQ1Njc4OWFiY2RlZjAxMjM0NTY3ODlhYmNkZWYwMTIzNDU2Nzg5YWJjZGVmMDEyMzQ1Njc4OWFiY2RlZiJ9KAEjYTNzLm9mZmljZS5wZGYuZm9ybS12YWx1ZXMucHJlc2VuY2UOQXBwbGljYW50Lk5hbWUBeCgBIWEzcy5vZmZpY2UucGRmLmZvcm0tdmFsdWVzLmZpZWxkcydbIkFwcGxpY2FudC5OYW1lIiwiW1widmFsdWVcIixcImlkXCJdIl0Bdw5BcHBsaWNhbnQuTmFtZSgBIWEzcy5vZmZpY2UucGRmLmZvcm0tdmFsdWVzLmZpZWxkcypbIkFwcGxpY2FudC5OYW1lIiwiW1widmFsdWVcIixcInZhbHVlXCJdIl0BdwNBZGEIASBhM3Mub2ZmaWNlLnBkZi5mb3JtLXZhbHVlcy5vcmRlcgF3DkFwcGxpY2FudC5OYW1lqLLyGQQBeAGy8hkBBAE=";

#[test]
fn typed_pdf_form_value_mutations_are_browser_compatible_durable_and_idempotent() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    assert_eq!(
        pdf_form_value(&store, "Applicant.Name"),
        Some("Ada".to_owned())
    );

    let replace_request = pdf_mutation_request(
        "pdf-set-name-1",
        NativeOfficeCollaborationMutation::PdfSetFormValue {
            field_id: "Applicant.Name".to_owned(),
            value: "Grace".to_owned(),
        },
    );
    let replaced = store.mutate(replace_request.clone()).unwrap();
    assert!(replaced.state_changed);
    assert_eq!(replaced.sequence, Some(2));
    assert_eq!(
        pdf_form_value(&store, "Applicant.Name"),
        Some("Grace".to_owned())
    );

    let created = store
        .mutate(pdf_mutation_request(
            "pdf-set-email-1",
            NativeOfficeCollaborationMutation::PdfSetFormValue {
                field_id: "Applicant.Email".to_owned(),
                value: "grace@example.test".to_owned(),
            },
        ))
        .unwrap();
    assert!(created.state_changed);
    assert_eq!(created.sequence, Some(3));
    assert_eq!(
        pdf_form_value(&store, "Applicant.Email"),
        Some("grace@example.test".to_owned())
    );
    assert_eq!(
        pdf_form_order(&store),
        vec!["Applicant.Name".to_owned(), "Applicant.Email".to_owned()]
    );

    let replay = store.mutate(replace_request).unwrap();
    assert!(replay.duplicate);
    assert_eq!(replay.sequence, Some(2));

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(
        pdf_form_value(&reopened, "Applicant.Name"),
        Some("Grace".to_owned())
    );
    assert_eq!(
        pdf_form_value(&reopened, "Applicant.Email"),
        Some("grace@example.test".to_owned())
    );
}

#[test]
fn typed_pdf_form_value_validation_is_atomic_and_kind_bound() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    for (operation_id, field_id) in [
        ("pdf-empty-field", "".to_owned()),
        ("pdf-trimmed-field", " Applicant.Name".to_owned()),
        ("pdf-long-field", "😀".repeat(257)),
    ] {
        let error = store
            .mutate(pdf_mutation_request(
                operation_id,
                NativeOfficeCollaborationMutation::PdfSetFormValue {
                    field_id,
                    value: "Rejected".to_owned(),
                },
            ))
            .unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_invalid");
    }
    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);

    let markdown_root = temp.path().join("markdown-replica");
    let markdown = NativeOfficeCollaborationStore::create(create_request(&markdown_root)).unwrap();
    let error = markdown
        .mutate(mutation_request(
            "pdf-on-markdown",
            NativeOfficeCollaborationMutation::PdfSetFormValue {
                field_id: "Applicant.Name".to_owned(),
                value: "Rejected".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_kind_mismatch");
}

#[test]
fn typed_pdf_annotation_mutations_merge_leaves_tombstone_and_survive_restart() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-annotation-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let original = portable_annotation("annotation-native-1", "#ffd400", "Initial note");
    let create = NativeOfficeCollaborationMutation::PdfCreateAnnotation {
        annotation_id: "annotation-native-1".to_owned(),
        page_index: 1,
        annotation: original.clone(),
    };
    let created = store
        .mutate(pdf_mutation_request(
            "pdf-create-annotation-1",
            create.clone(),
        ))
        .unwrap();
    assert!(created.state_changed);
    assert_eq!(created.sequence, Some(2));
    let record = pdf_record(&store, "annotations", "annotation-native-1");
    assert_eq!(record["source"], "created");
    assert_eq!(record["annotation"]["color"], "#ffd400");
    assert_eq!(record["annotation"]["strokeColor"], "#ffd400");
    assert_eq!(record["annotation"]["segmentRects"][0]["origin"]["x"], 68.0);

    let mut color_update = original.clone();
    color_update["color"] = serde_json::json!("#ff0000");
    color_update["strokeColor"] = serde_json::json!("#ff0000");
    store
        .mutate(pdf_mutation_request(
            "pdf-update-annotation-color",
            NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id: "annotation-native-1".to_owned(),
                expected_annotation: original.clone(),
                next_annotation: color_update.clone(),
            },
        ))
        .unwrap();

    let mut note_update = original.clone();
    note_update["contents"] = serde_json::json!("Independent native note");
    let merged = store
        .mutate(pdf_mutation_request(
            "pdf-update-annotation-note",
            NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id: "annotation-native-1".to_owned(),
                expected_annotation: original.clone(),
                next_annotation: note_update,
            },
        ))
        .unwrap();
    assert!(merged.state_changed);
    let record = pdf_record(&store, "annotations", "annotation-native-1");
    assert_eq!(record["annotation"]["color"], "#ff0000");
    assert_eq!(record["annotation"]["strokeColor"], "#ff0000");
    assert_eq!(record["annotation"]["contents"], "Independent native note");

    let unchanged = store
        .mutate(pdf_mutation_request(
            "pdf-create-annotation-identical-after-edit",
            create,
        ))
        .unwrap();
    assert!(!unchanged.state_changed);
    assert_eq!(unchanged.sequence, None);
    assert_eq!(
        pdf_record(&store, "annotations", "annotation-native-1")["annotation"]["color"],
        "#ff0000"
    );

    let deleted = store
        .mutate(pdf_mutation_request(
            "pdf-delete-annotation-1",
            NativeOfficeCollaborationMutation::PdfDeleteAnnotation {
                annotation_id: "annotation-native-1".to_owned(),
                expected_source: NativeOfficeCollaborationPdfAnnotationSource::Created,
                expected_page_index: 1,
                expected_type: 9,
            },
        ))
        .unwrap();
    assert!(deleted.state_changed);
    assert_eq!(
        pdf_record(&store, "annotations", "annotation-native-1")["deleted"],
        true
    );

    let expected_tombstoned =
        portable_annotation("annotation-native-1", "#ff0000", "Independent native note");
    let mut tombstoned_note = expected_tombstoned.clone();
    tombstoned_note["contents"] = serde_json::json!("Retained after deletion");
    let tombstone_merge = store
        .mutate(pdf_mutation_request(
            "pdf-update-tombstoned-annotation",
            NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id: "annotation-native-1".to_owned(),
                expected_annotation: expected_tombstoned,
                next_annotation: tombstoned_note,
            },
        ))
        .unwrap();
    assert!(tombstone_merge.state_changed);
    let record = pdf_record(&store, "annotations", "annotation-native-1");
    assert_eq!(record["deleted"], true);
    assert_eq!(record["annotation"]["contents"], "Retained after deletion");

    let delete_retry = store
        .mutate(pdf_mutation_request(
            "pdf-delete-annotation-identical",
            NativeOfficeCollaborationMutation::PdfDeleteAnnotation {
                annotation_id: "annotation-native-1".to_owned(),
                expected_source: NativeOfficeCollaborationPdfAnnotationSource::Created,
                expected_page_index: 1,
                expected_type: 9,
            },
        ))
        .unwrap();
    assert!(!delete_retry.state_changed);

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    let record = pdf_record(&reopened, "annotations", "annotation-native-1");
    assert_eq!(record["deleted"], true);
    assert_eq!(record["annotation"]["contents"], "Retained after deletion");
}

#[test]
fn typed_pdf_annotation_conflicts_and_identity_guards_are_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-annotation-conflicts");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let original = portable_annotation("annotation-conflict", "#ffd400", "Initial note");
    store
        .mutate(pdf_mutation_request(
            "pdf-create-annotation-conflict",
            NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id: "annotation-conflict".to_owned(),
                page_index: 1,
                annotation: original.clone(),
            },
        ))
        .unwrap();
    let mut red = original.clone();
    red["color"] = serde_json::json!("#ff0000");
    red["strokeColor"] = serde_json::json!("#ff0000");
    store
        .mutate(pdf_mutation_request(
            "pdf-update-annotation-red",
            NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id: "annotation-conflict".to_owned(),
                expected_annotation: original.clone(),
                next_annotation: red,
            },
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    let mut blue = original.clone();
    blue["color"] = serde_json::json!("#0000ff");
    blue["strokeColor"] = serde_json::json!("#0000ff");
    let error = store
        .mutate(pdf_mutation_request(
            "pdf-update-annotation-blue-stale",
            NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id: "annotation-conflict".to_owned(),
                expected_annotation: original.clone(),
                next_annotation: blue,
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");

    let collision = store
        .mutate(pdf_mutation_request(
            "pdf-create-annotation-collision",
            NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id: "annotation-conflict".to_owned(),
                page_index: 1,
                annotation: portable_annotation(
                    "annotation-conflict",
                    "#0000ff",
                    "Different creation",
                ),
            },
        ))
        .unwrap_err();
    assert_eq!(
        collision.code,
        "office.collaboration.mutation_match_conflict"
    );

    let identity_error = store
        .mutate(pdf_mutation_request(
            "pdf-delete-annotation-wrong-identity",
            NativeOfficeCollaborationMutation::PdfDeleteAnnotation {
                annotation_id: "annotation-conflict".to_owned(),
                expected_source: NativeOfficeCollaborationPdfAnnotationSource::Base,
                expected_page_index: 1,
                expected_type: 9,
            },
        ))
        .unwrap_err();
    assert_eq!(
        identity_error.code,
        "office.collaboration.mutation_match_conflict"
    );

    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.state_vector, before.state_vector);
    assert_eq!(
        pdf_record(&store, "annotations", "annotation-conflict")["annotation"]["color"],
        "#ff0000"
    );
}

#[test]
fn typed_pdf_annotation_input_validation_rejects_unsupported_or_mismatched_values() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-annotation-validation");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    let mut unsupported = portable_annotation("annotation-invalid", "#ffd400", "Invalid");
    unsupported["type"] = serde_json::json!(2);
    for (operation_id, mutation) in [
        (
            "pdf-create-annotation-unsupported",
            NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id: "annotation-invalid".to_owned(),
                page_index: 1,
                annotation: unsupported,
            },
        ),
        (
            "pdf-create-annotation-id-mismatch",
            NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id: "annotation-invalid".to_owned(),
                page_index: 1,
                annotation: portable_annotation("another-id", "#ffd400", "Invalid"),
            },
        ),
    ] {
        let error = store
            .mutate(pdf_mutation_request(operation_id, mutation))
            .unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_invalid");
    }
    let range_error = store
        .mutate(pdf_mutation_request(
            "pdf-create-annotation-page-range",
            NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id: "annotation-page-range".to_owned(),
                page_index: 3,
                annotation: portable_annotation("annotation-page-range", "#ffd400", "Invalid")
                    .as_object()
                    .map(|annotation| {
                        let mut annotation = annotation.clone();
                        annotation.insert("pageIndex".to_owned(), serde_json::json!(3));
                        serde_json::Value::Object(annotation)
                    })
                    .unwrap(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        range_error.code,
        "office.collaboration.mutation_range_invalid"
    );

    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);
}

#[test]
fn typed_pdf_review_mutations_are_attributable_append_only_and_durable() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-review-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let proposal = NativeOfficeCollaborationMutation::PdfProposeRedaction {
        proposal_id: "redaction-native-1".to_owned(),
        page_index: 1,
        rects: vec![
            NativeOfficeCollaborationPdfRect {
                left: 10.5,
                top: 20.25,
                right: 80.75,
                bottom: 40.5,
            },
            NativeOfficeCollaborationPdfRect {
                left: 90.0,
                top: 20.25,
                right: 120.0,
                bottom: 40.5,
            },
        ],
        proposed_at: "2026-08-15T03:00:00.000Z".to_owned(),
        reason: Some("Personal data".to_owned()),
        text: Some("Account 1234".to_owned()),
    };
    let proposed = store
        .mutate(pdf_mutation_request(
            "pdf-propose-redaction-1",
            proposal.clone(),
        ))
        .unwrap();
    assert!(proposed.state_changed);
    assert_eq!(proposed.sequence, Some(2));
    assert_eq!(
        pdf_record(&store, "redaction-proposals", "redaction-native-1"),
        serde_json::json!({
            "id": "redaction-native-1",
            "pageIndex": 1,
            "rects": [
                { "left": 10.5, "top": 20.25, "right": 80.75, "bottom": 40.5 },
                { "left": 90, "top": 20.25, "right": 120, "bottom": 40.5 }
            ],
            "proposedBy": "agent-alpha",
            "proposedAt": "2026-08-15T03:00:00.000Z",
            "reason": "Personal data",
            "text": "Account 1234"
        })
    );

    let identical_retry = store
        .mutate(pdf_mutation_request(
            "pdf-propose-redaction-identical",
            proposal,
        ))
        .unwrap();
    assert!(!identical_retry.state_changed);
    assert_eq!(identical_retry.sequence, None);

    let decided = store
        .mutate(pdf_mutation_request(
            "pdf-approve-redaction-1",
            NativeOfficeCollaborationMutation::PdfDecideReview {
                decision_id: "decision-native-1".to_owned(),
                target_kind: NativeOfficeCollaborationPdfReviewTargetKind::Redaction,
                target_id: "redaction-native-1".to_owned(),
                decision: NativeOfficeCollaborationPdfReviewDecision::Approve,
                created_at: "2026-08-15T03:05:00.000Z".to_owned(),
            },
        ))
        .unwrap();
    assert!(decided.state_changed);
    assert_eq!(decided.sequence, Some(3));
    assert_eq!(
        pdf_record(&store, "review-decisions", "decision-native-1"),
        serde_json::json!({
            "id": "decision-native-1",
            "targetKind": "redaction",
            "targetId": "redaction-native-1",
            "decision": "approve",
            "actorId": "agent-alpha",
            "createdAt": "2026-08-15T03:05:00.000Z"
        })
    );

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(
        pdf_record(&reopened, "review-decisions", "decision-native-1")["actorId"],
        "agent-alpha"
    );
}

#[test]
fn typed_pdf_review_conflicts_and_invalid_geometry_are_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-review-conflict-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    for (operation_id, mutation) in [
        (
            "pdf-redaction-invalid-rect",
            NativeOfficeCollaborationMutation::PdfProposeRedaction {
                proposal_id: "redaction-invalid-rect".to_owned(),
                page_index: 0,
                rects: vec![NativeOfficeCollaborationPdfRect {
                    left: 10.0,
                    top: 10.0,
                    right: 10.0,
                    bottom: 20.0,
                }],
                proposed_at: "2026-08-15T04:00:00.000Z".to_owned(),
                reason: None,
                text: None,
            },
        ),
        (
            "pdf-redaction-invalid-time",
            NativeOfficeCollaborationMutation::PdfProposeRedaction {
                proposal_id: "redaction-invalid-time".to_owned(),
                page_index: 0,
                rects: vec![valid_pdf_rect()],
                proposed_at: "tomorrow".to_owned(),
                reason: None,
                text: None,
            },
        ),
    ] {
        let error = store
            .mutate(pdf_mutation_request(operation_id, mutation))
            .unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_invalid");
    }
    let out_of_range = store
        .mutate(pdf_mutation_request(
            "pdf-redaction-page-out-of-range",
            NativeOfficeCollaborationMutation::PdfProposeRedaction {
                proposal_id: "redaction-page-out-of-range".to_owned(),
                page_index: 3,
                rects: vec![valid_pdf_rect()],
                proposed_at: "2026-08-15T04:00:00.000Z".to_owned(),
                reason: None,
                text: None,
            },
        ))
        .unwrap_err();
    assert_eq!(
        out_of_range.code,
        "office.collaboration.mutation_range_invalid"
    );
    let after_invalid = store.inspect().unwrap();
    assert_eq!(after_invalid.current_sequence, before.current_sequence);
    assert_eq!(
        after_invalid.document_state_sha256,
        before.document_state_sha256
    );

    store
        .mutate(pdf_mutation_request(
            "pdf-propose-conflict-target",
            NativeOfficeCollaborationMutation::PdfProposeRedaction {
                proposal_id: "redaction-conflict".to_owned(),
                page_index: 0,
                rects: vec![valid_pdf_rect()],
                proposed_at: "2026-08-15T04:00:00.000Z".to_owned(),
                reason: Some("Original".to_owned()),
                text: None,
            },
        ))
        .unwrap();
    let conflict = store
        .mutate(pdf_mutation_request(
            "pdf-propose-conflicting-reuse",
            NativeOfficeCollaborationMutation::PdfProposeRedaction {
                proposal_id: "redaction-conflict".to_owned(),
                page_index: 0,
                rects: vec![valid_pdf_rect()],
                proposed_at: "2026-08-15T04:00:00.000Z".to_owned(),
                reason: Some("Changed".to_owned()),
                text: None,
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    let missing_target = store
        .mutate(pdf_mutation_request(
            "pdf-decide-missing-target",
            NativeOfficeCollaborationMutation::PdfDecideReview {
                decision_id: "decision-missing".to_owned(),
                target_kind: NativeOfficeCollaborationPdfReviewTargetKind::Redaction,
                target_id: "redaction-missing".to_owned(),
                decision: NativeOfficeCollaborationPdfReviewDecision::Reject,
                created_at: "2026-08-15T04:05:00.000Z".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        missing_target.code,
        "office.collaboration.mutation_match_conflict"
    );

    store
        .mutate(pdf_mutation_request(
            "pdf-decide-conflict-target",
            NativeOfficeCollaborationMutation::PdfDecideReview {
                decision_id: "decision-first".to_owned(),
                target_kind: NativeOfficeCollaborationPdfReviewTargetKind::Redaction,
                target_id: "redaction-conflict".to_owned(),
                decision: NativeOfficeCollaborationPdfReviewDecision::Approve,
                created_at: "2026-08-15T04:05:00.000Z".to_owned(),
            },
        ))
        .unwrap();
    let second_decision = store
        .mutate(pdf_mutation_request(
            "pdf-decide-target-twice",
            NativeOfficeCollaborationMutation::PdfDecideReview {
                decision_id: "decision-second".to_owned(),
                target_kind: NativeOfficeCollaborationPdfReviewTargetKind::Redaction,
                target_id: "redaction-conflict".to_owned(),
                decision: NativeOfficeCollaborationPdfReviewDecision::Reject,
                created_at: "2026-08-15T04:06:00.000Z".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        second_decision.code,
        "office.collaboration.mutation_match_conflict"
    );
    let final_state = store.inspect().unwrap();
    assert_eq!(final_state.current_sequence, 3);
}

#[test]
fn typed_pdf_page_operations_are_append_only_reviewable_and_durable() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-page-operation-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let rotation = NativeOfficeCollaborationMutation::PdfProposePageRotation {
        page_operation_id: "page-operation-rotate-1".to_owned(),
        page_indices: vec![0, 2],
        degrees: 90,
        proposed_at: "2026-08-15T07:00:00.000Z".to_owned(),
    };
    let rotated = store
        .mutate(pdf_mutation_request(
            "pdf-propose-page-rotation",
            rotation.clone(),
        ))
        .unwrap();
    assert!(rotated.state_changed);
    assert_eq!(rotated.sequence, Some(2));
    assert_eq!(
        pdf_record(&store, "page-operations", "page-operation-rotate-1"),
        serde_json::json!({
            "id": "page-operation-rotate-1",
            "kind": "rotate",
            "pageIndices": [0, 2],
            "degrees": 90,
            "proposedBy": "agent-alpha",
            "proposedAt": "2026-08-15T07:00:00.000Z"
        })
    );
    let retry = store
        .mutate(pdf_mutation_request(
            "pdf-propose-page-rotation-retry",
            rotation,
        ))
        .unwrap();
    assert!(!retry.state_changed);

    store
        .mutate(pdf_mutation_request(
            "pdf-propose-page-deletion",
            NativeOfficeCollaborationMutation::PdfProposePageDeletion {
                page_operation_id: "page-operation-delete-1".to_owned(),
                page_indices: vec![2],
                proposed_at: "2026-08-15T07:01:00.000Z".to_owned(),
            },
        ))
        .unwrap();
    store
        .mutate(pdf_mutation_request(
            "pdf-propose-page-reorder",
            NativeOfficeCollaborationMutation::PdfProposePageReorder {
                page_operation_id: "page-operation-reorder-1".to_owned(),
                page_order: vec![2, 0, 1],
                proposed_at: "2026-08-15T07:02:00.000Z".to_owned(),
            },
        ))
        .unwrap();
    let decision = store
        .mutate(pdf_mutation_request(
            "pdf-approve-page-deletion",
            NativeOfficeCollaborationMutation::PdfDecideReview {
                decision_id: "decision-page-operation-1".to_owned(),
                target_kind: NativeOfficeCollaborationPdfReviewTargetKind::PageOperation,
                target_id: "page-operation-delete-1".to_owned(),
                decision: NativeOfficeCollaborationPdfReviewDecision::Approve,
                created_at: "2026-08-15T07:05:00.000Z".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(decision.sequence, Some(5));
    assert_eq!(
        pdf_record(&store, "page-operations", "page-operation-delete-1"),
        serde_json::json!({
            "id": "page-operation-delete-1",
            "kind": "delete",
            "pageIndices": [2],
            "proposedBy": "agent-alpha",
            "proposedAt": "2026-08-15T07:01:00.000Z"
        })
    );
    assert_eq!(
        pdf_record(&store, "page-operations", "page-operation-reorder-1"),
        serde_json::json!({
            "id": "page-operation-reorder-1",
            "kind": "reorder",
            "pageOrder": [2, 0, 1],
            "proposedBy": "agent-alpha",
            "proposedAt": "2026-08-15T07:02:00.000Z"
        })
    );

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(
        pdf_record(&reopened, "review-decisions", "decision-page-operation-1")["targetKind"],
        "page-operation"
    );
}

#[test]
fn typed_pdf_page_operation_validation_and_identity_conflicts_are_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("pdf-page-operation-conflict-replica");
    let store = NativeOfficeCollaborationStore::create(pdf_create_request(&root)).unwrap();
    store
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    for (operation_id, mutation) in [
        (
            "pdf-page-operation-empty-pages",
            NativeOfficeCollaborationMutation::PdfProposePageRotation {
                page_operation_id: "page-operation-empty".to_owned(),
                page_indices: vec![],
                degrees: 90,
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ),
        (
            "pdf-page-operation-duplicate-pages",
            NativeOfficeCollaborationMutation::PdfProposePageDeletion {
                page_operation_id: "page-operation-duplicate".to_owned(),
                page_indices: vec![1, 1],
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ),
        (
            "pdf-page-operation-invalid-degrees",
            NativeOfficeCollaborationMutation::PdfProposePageRotation {
                page_operation_id: "page-operation-degrees".to_owned(),
                page_indices: vec![0],
                degrees: 45,
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ),
    ] {
        let error = store
            .mutate(pdf_mutation_request(operation_id, mutation))
            .unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_invalid");
    }

    for (operation_id, mutation) in [
        (
            "pdf-page-operation-out-of-range",
            NativeOfficeCollaborationMutation::PdfProposePageRotation {
                page_operation_id: "page-operation-out-of-range".to_owned(),
                page_indices: vec![3],
                degrees: 180,
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ),
        (
            "pdf-page-operation-delete-all",
            NativeOfficeCollaborationMutation::PdfProposePageDeletion {
                page_operation_id: "page-operation-delete-all".to_owned(),
                page_indices: vec![0, 1, 2],
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ),
        (
            "pdf-page-operation-incomplete-reorder",
            NativeOfficeCollaborationMutation::PdfProposePageReorder {
                page_operation_id: "page-operation-incomplete".to_owned(),
                page_order: vec![0, 1],
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ),
    ] {
        let error = store
            .mutate(pdf_mutation_request(operation_id, mutation))
            .unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_range_invalid");
    }
    let after_invalid = store.inspect().unwrap();
    assert_eq!(after_invalid.current_sequence, before.current_sequence);
    assert_eq!(
        after_invalid.document_state_sha256,
        before.document_state_sha256
    );

    store
        .mutate(pdf_mutation_request(
            "pdf-page-operation-conflict-original",
            NativeOfficeCollaborationMutation::PdfProposePageRotation {
                page_operation_id: "page-operation-conflict".to_owned(),
                page_indices: vec![0],
                degrees: 90,
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ))
        .unwrap();
    let conflict = store
        .mutate(pdf_mutation_request(
            "pdf-page-operation-conflict-reuse",
            NativeOfficeCollaborationMutation::PdfProposePageDeletion {
                page_operation_id: "page-operation-conflict".to_owned(),
                page_indices: vec![0],
                proposed_at: "2026-08-15T08:00:00.000Z".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(store.inspect().unwrap().current_sequence, 2);
}
