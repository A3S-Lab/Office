use super::*;

#[test]
fn typed_pdf_annotation_updates_converge_across_delivery_order_and_restart() {
    let temp = tempfile::tempdir().unwrap();
    let source_root = temp.path().join("pdf-annotation-order-source");
    let source = NativeOfficeCollaborationStore::create(pdf_create_request(&source_root)).unwrap();
    source
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let original = portable_annotation("annotation-order", "#ffd400", "Initial");
    source
        .mutate(pdf_mutation_request(
            "source-create-annotation",
            NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                annotation_id: "annotation-order".to_owned(),
                page_index: 1,
                annotation: original.clone(),
            },
        ))
        .unwrap();
    let mut updated = original.clone();
    updated["color"] = serde_json::json!("#ff0000");
    updated["strokeColor"] = serde_json::json!("#ff0000");
    source
        .mutate(pdf_mutation_request(
            "source-update-annotation",
            NativeOfficeCollaborationMutation::PdfUpdateAnnotation {
                annotation_id: "annotation-order".to_owned(),
                expected_annotation: original.clone(),
                next_annotation: updated,
            },
        ))
        .unwrap();
    source
        .mutate(pdf_mutation_request(
            "source-delete-annotation",
            NativeOfficeCollaborationMutation::PdfDeleteAnnotation {
                annotation_id: "annotation-order".to_owned(),
                expected_source: NativeOfficeCollaborationPdfAnnotationSource::Created,
                expected_page_index: 1,
                expected_type: 9,
            },
        ))
        .unwrap();
    let updates = source
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 8,
        })
        .unwrap()
        .updates
        .into_iter()
        .map(|event| event.update)
        .collect::<Vec<_>>();
    assert_eq!(updates.len(), 3);
    let expected = source.inspect().unwrap();

    for (permutation_index, permutation) in three_item_permutations().into_iter().enumerate() {
        let root = temp
            .path()
            .join(format!("pdf-annotation-order-{permutation_index:02}"));
        let mut create = pdf_create_request(&root);
        create.client_id = Some(920_000 + permutation_index as u64);
        let target = NativeOfficeCollaborationStore::create(create).unwrap();
        target
            .apply(pdf_apply_request(
                "bootstrap-browser-pdf",
                STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
            ))
            .unwrap();
        for (delivery_index, update_index) in permutation.into_iter().enumerate() {
            target
                .apply(pdf_apply_request(
                    &format!("deliver-{delivery_index}-{update_index}"),
                    updates[update_index].clone(),
                ))
                .unwrap_or_else(|error| panic!("permutation {permutation:?}: {error:?}"));
        }
        target
            .apply(pdf_apply_request(
                "deliver-duplicate",
                updates[permutation_index % updates.len()].clone(),
            ))
            .unwrap();
        let delivered = target.inspect().unwrap();
        assert!(!delivered.pending_updates, "permutation {permutation:?}");
        assert_eq!(
            delivered.state_vector, expected.state_vector,
            "permutation {permutation:?}"
        );
        assert_eq!(
            delivered.document_state_sha256, expected.document_state_sha256,
            "permutation {permutation:?}"
        );
        drop(target);

        let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
        assert_eq!(
            reopened.inspect().unwrap().document_state_sha256,
            expected.document_state_sha256,
            "permutation {permutation:?}"
        );
        let record = pdf_record(&reopened, "annotations", "annotation-order");
        assert_eq!(record["deleted"], true);
        assert_eq!(record["annotation"]["color"], "#ff0000");
        assert_eq!(record["annotation"]["strokeColor"], "#ff0000");
        assert_eq!(
            record["annotation"]["segmentRects"][0]["size"]["width"],
            300.0
        );
        let no_op = reopened
            .mutate(pdf_mutation_request(
                "validate-converged-annotation-state",
                NativeOfficeCollaborationMutation::PdfCreateAnnotation {
                    annotation_id: "annotation-order".to_owned(),
                    page_index: 1,
                    annotation: original.clone(),
                },
            ))
            .unwrap();
        assert!(!no_op.state_changed, "permutation {permutation:?}");
    }
}

#[test]
fn typed_pdf_review_updates_converge_across_every_delivery_order() {
    let temp = tempfile::tempdir().unwrap();
    let source_root = temp.path().join("pdf-review-order-source");
    let source = NativeOfficeCollaborationStore::create(pdf_create_request(&source_root)).unwrap();
    source
        .apply(pdf_apply_request(
            "bootstrap-browser-pdf",
            STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    for (operation_id, mutation) in [
        (
            "source-page-rotation",
            NativeOfficeCollaborationMutation::PdfProposePageRotation {
                page_operation_id: "page-operation-order-rotate".to_owned(),
                page_indices: vec![0, 2],
                degrees: 90,
                proposed_at: "2026-08-15T09:00:00.000Z".to_owned(),
            },
        ),
        (
            "source-page-deletion",
            NativeOfficeCollaborationMutation::PdfProposePageDeletion {
                page_operation_id: "page-operation-order-delete".to_owned(),
                page_indices: vec![2],
                proposed_at: "2026-08-15T09:01:00.000Z".to_owned(),
            },
        ),
        (
            "source-page-reorder",
            NativeOfficeCollaborationMutation::PdfProposePageReorder {
                page_operation_id: "page-operation-order-reorder".to_owned(),
                page_order: vec![2, 0, 1],
                proposed_at: "2026-08-15T09:02:00.000Z".to_owned(),
            },
        ),
        (
            "source-page-decision",
            NativeOfficeCollaborationMutation::PdfDecideReview {
                decision_id: "decision-page-operation-order".to_owned(),
                target_kind: NativeOfficeCollaborationPdfReviewTargetKind::PageOperation,
                target_id: "page-operation-order-delete".to_owned(),
                decision: NativeOfficeCollaborationPdfReviewDecision::Approve,
                created_at: "2026-08-15T09:05:00.000Z".to_owned(),
            },
        ),
    ] {
        source
            .mutate(pdf_mutation_request(operation_id, mutation))
            .unwrap();
    }
    let updates = source
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 8,
        })
        .unwrap()
        .updates
        .into_iter()
        .map(|event| event.update)
        .collect::<Vec<_>>();
    assert_eq!(updates.len(), 4);
    let expected = source.inspect().unwrap();
    assert!(!expected.pending_updates);

    for (permutation_index, permutation) in four_item_permutations().into_iter().enumerate() {
        let root = temp
            .path()
            .join(format!("pdf-review-order-{permutation_index:02}"));
        let mut create = pdf_create_request(&root);
        create.client_id = Some(910_000 + permutation_index as u64);
        let target = NativeOfficeCollaborationStore::create(create).unwrap();
        target
            .apply(pdf_apply_request(
                "bootstrap-browser-pdf",
                STANDARD.decode(YJS_PDF_UPDATE_BASE64).unwrap(),
            ))
            .unwrap_or_else(|error| panic!("permutation {permutation:?}: {error:?}"));
        for (delivery_index, update_index) in permutation.into_iter().enumerate() {
            let applied = target
                .apply(pdf_apply_request(
                    &format!("deliver-{delivery_index}-{update_index}"),
                    updates[update_index].clone(),
                ))
                .unwrap();
            if permutation == [0, 1, 3, 2] && delivery_index == 2 {
                assert!(target.inspect().unwrap().pending_updates);
            }
            if delivery_index == updates.len() - 1 {
                assert_eq!(
                    applied.state_vector, expected.state_vector,
                    "{permutation:?}"
                );
                assert_eq!(
                    applied.state_vector_sha256, expected.state_vector_sha256,
                    "{permutation:?}"
                );
            }
        }
        target
            .apply(pdf_apply_request(
                "deliver-duplicate",
                updates[permutation_index % updates.len()].clone(),
            ))
            .unwrap();
        let delivered = target.inspect().unwrap();
        assert!(!delivered.pending_updates, "permutation {permutation:?}");
        assert_eq!(
            delivered.state_vector, expected.state_vector,
            "permutation {permutation:?}"
        );
        assert_eq!(
            delivered.document_state_sha256, expected.document_state_sha256,
            "permutation {permutation:?}"
        );
        drop(target);

        let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
        let reopened_inspection = reopened.inspect().unwrap();
        assert!(
            !reopened_inspection.pending_updates,
            "permutation {permutation:?}"
        );
        assert_eq!(
            reopened_inspection.document_state_sha256, expected.document_state_sha256,
            "permutation {permutation:?}"
        );
        let no_op = reopened
            .mutate(pdf_mutation_request(
                "validate-converged-review-state",
                NativeOfficeCollaborationMutation::PdfProposePageRotation {
                    page_operation_id: "page-operation-order-rotate".to_owned(),
                    page_indices: vec![0, 2],
                    degrees: 90,
                    proposed_at: "2026-08-15T09:00:00.000Z".to_owned(),
                },
            ))
            .unwrap_or_else(|error| panic!("permutation {permutation:?}: {error:?}"));
        assert!(!no_op.state_changed, "permutation {permutation:?}");
        assert_eq!(
            pdf_record(
                &reopened,
                "review-decisions",
                "decision-page-operation-order"
            )["targetId"],
            "page-operation-order-delete"
        );
    }
}

fn four_item_permutations() -> Vec<[usize; 4]> {
    let mut result = Vec::with_capacity(24);
    for first in 0..4 {
        for second in 0..4 {
            for third in 0..4 {
                for fourth in 0..4 {
                    let permutation = [first, second, third, fourth];
                    if permutation.iter().copied().collect::<HashSet<_>>().len() == 4 {
                        result.push(permutation);
                    }
                }
            }
        }
    }
    assert_eq!(result.len(), 24);
    result
}

fn three_item_permutations() -> Vec<[usize; 3]> {
    let mut result = Vec::with_capacity(6);
    for first in 0..3 {
        for second in 0..3 {
            for third in 0..3 {
                let permutation = [first, second, third];
                if permutation.iter().copied().collect::<HashSet<_>>().len() == 3 {
                    result.push(permutation);
                }
            }
        }
    }
    assert_eq!(result.len(), 6);
    result
}
