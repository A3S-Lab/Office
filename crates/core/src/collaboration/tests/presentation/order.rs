use super::*;

#[test]
fn typed_presentation_element_moves_are_stable_idempotent_and_restart_safe() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replica");
    let store = initialized_presentation_store(&root, 930_001);
    for (id, after) in [
        ("element-first", Some("element-title")),
        ("element-second", Some("element-first")),
    ] {
        store
            .mutate(presentation_mutation_request(
                &format!("create-{id}"),
                NativeOfficeCollaborationMutation::PresentationCreateElement {
                    container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                    container_id: "slide-1".to_owned(),
                    element: scene_element(id, id, "shape"),
                    after_element_id: after.map(str::to_owned),
                },
            ))
            .unwrap();
    }
    assert_eq!(
        presentation_element_order(&store, "slides", "slide-1"),
        vec!["element-title", "element-first", "element-second"]
    );

    let moved = store
        .mutate(presentation_mutation_request(
            "move-second-to-front",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: Some("element-first".to_owned()),
                after_element_id: None,
            },
        ))
        .unwrap();
    assert!(moved.state_changed);
    assert_eq!(
        presentation_element_order(&store, "slides", "slide-1"),
        vec!["element-second", "element-title", "element-first"]
    );

    let idempotent = store
        .mutate(presentation_mutation_request(
            "move-second-to-front-idempotent",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: Some("element-first".to_owned()),
                after_element_id: None,
            },
        ))
        .unwrap();
    assert!(!idempotent.state_changed);

    let restored = store
        .mutate(presentation_mutation_request(
            "move-second-after-first",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: None,
                after_element_id: Some("element-first".to_owned()),
            },
        ))
        .unwrap();
    assert!(restored.state_changed);
    assert_eq!(
        presentation_element_order(&store, "slides", "slide-1"),
        vec!["element-title", "element-first", "element-second"]
    );

    for (kind, collection, container_id) in [
        (
            NativeOfficeCollaborationPresentationContainerKind::Master,
            "masters",
            "master-1",
        ),
        (
            NativeOfficeCollaborationPresentationContainerKind::Layout,
            "layouts",
            "layout-1",
        ),
    ] {
        for (id, after) in [
            (format!("{container_id}-first"), None),
            (
                format!("{container_id}-second"),
                Some(format!("{container_id}-first")),
            ),
        ] {
            store
                .mutate(presentation_mutation_request(
                    &format!("create-{id}"),
                    NativeOfficeCollaborationMutation::PresentationCreateElement {
                        container_kind: kind,
                        container_id: container_id.to_owned(),
                        element: scene_element(&id, &id, "text"),
                        after_element_id: after,
                    },
                ))
                .unwrap();
        }
        store
            .mutate(presentation_mutation_request(
                &format!("move-{container_id}-second"),
                NativeOfficeCollaborationMutation::PresentationMoveElement {
                    container_kind: kind,
                    container_id: container_id.to_owned(),
                    element_id: format!("{container_id}-second"),
                    expected_after_element_id: Some(format!("{container_id}-first")),
                    after_element_id: None,
                },
            ))
            .unwrap();
        assert_eq!(
            presentation_element_order(&store, collection, container_id),
            vec![
                format!("{container_id}-second"),
                format!("{container_id}-first"),
            ]
        );
    }

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(
        presentation_element_order(&reopened, "slides", "slide-1"),
        vec!["element-title", "element-first", "element-second"]
    );
    assert_eq!(
        presentation_element_order(&reopened, "masters", "master-1"),
        vec!["master-1-second", "master-1-first"]
    );
    assert_eq!(
        presentation_element_order(&reopened, "layouts", "layout-1"),
        vec!["layout-1-second", "layout-1-first"]
    );
}

#[test]
fn typed_presentation_element_move_conflicts_are_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_presentation_store(&temp.path().join("presentation"), 930_002);
    for (id, after) in [
        ("element-first", Some("element-title")),
        ("element-second", Some("element-first")),
    ] {
        store
            .mutate(presentation_mutation_request(
                &format!("create-{id}"),
                NativeOfficeCollaborationMutation::PresentationCreateElement {
                    container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                    container_id: "slide-1".to_owned(),
                    element: scene_element(id, id, "shape"),
                    after_element_id: after.map(str::to_owned),
                },
            ))
            .unwrap();
    }

    let before_stale = store.inspect().unwrap();
    let stale = store
        .mutate(presentation_mutation_request(
            "move-with-stale-source-position",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: Some("element-title".to_owned()),
                after_element_id: None,
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );

    let missing_anchor = store
        .mutate(presentation_mutation_request(
            "move-after-missing-anchor",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: Some("element-first".to_owned()),
                after_element_id: Some("missing-element".to_owned()),
            },
        ))
        .unwrap_err();
    assert_eq!(
        missing_anchor.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );

    let self_anchor = store
        .mutate(presentation_mutation_request(
            "move-after-self",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: Some("element-first".to_owned()),
                after_element_id: Some("element-second".to_owned()),
            },
        ))
        .unwrap_err();
    assert_eq!(self_anchor.code, "office.collaboration.mutation_invalid");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );

    store
        .mutate(presentation_mutation_request(
            "delete-first-anchor",
            NativeOfficeCollaborationMutation::PresentationDeleteElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                expected_element: scene_element("element-first", "element-first", "shape"),
            },
        ))
        .unwrap();
    let before_deleted_anchor = store.inspect().unwrap();
    let deleted_anchor = store
        .mutate(presentation_mutation_request(
            "move-after-deleted-anchor",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-second".to_owned(),
                expected_after_element_id: Some("element-title".to_owned()),
                after_element_id: Some("element-first".to_owned()),
            },
        ))
        .unwrap_err();
    assert_eq!(
        deleted_anchor.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_deleted_anchor.document_state_sha256
    );
}

#[test]
fn typed_presentation_element_move_survives_reordered_causal_delivery() {
    let temp = tempfile::tempdir().unwrap();
    let source = initialized_presentation_store(&temp.path().join("source"), 930_003);
    source
        .mutate(presentation_mutation_request(
            "create-reordered-element",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: scene_element("element-reordered", "Reordered", "shape"),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    let create_update = mutation_update(&source, 1);
    source
        .mutate(presentation_mutation_request(
            "move-reordered-element",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-reordered".to_owned(),
                expected_after_element_id: Some("element-title".to_owned()),
                after_element_id: None,
            },
        ))
        .unwrap();
    let move_update = mutation_update(&source, 2);

    let target = initialized_presentation_store(&temp.path().join("target"), 930_004);
    target
        .apply(presentation_apply_request(
            "deliver-move-before-create",
            move_update.clone(),
        ))
        .unwrap();
    target
        .apply(presentation_apply_request(
            "deliver-create-after-move",
            create_update.clone(),
        ))
        .unwrap();
    target
        .apply(presentation_apply_request(
            "deliver-create-duplicate",
            create_update,
        ))
        .unwrap();
    target
        .apply(presentation_apply_request(
            "deliver-move-duplicate",
            move_update,
        ))
        .unwrap();
    assert_eq!(
        presentation_element_order(&target, "slides", "slide-1"),
        vec!["element-reordered", "element-title"]
    );

    drop(target);
    let reopened = NativeOfficeCollaborationStore::open(temp.path().join("target")).unwrap();
    assert_eq!(
        presentation_element_order(&reopened, "slides", "slide-1"),
        vec!["element-reordered", "element-title"]
    );
}

#[test]
fn typed_presentation_element_moves_converge_with_unrelated_field_edits() {
    let temp = tempfile::tempdir().unwrap();
    let bootstrap = initialized_presentation_store(&temp.path().join("bootstrap"), 930_005);
    bootstrap
        .mutate(presentation_mutation_request(
            "create-shared-order-element",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: scene_element("element-shared", "Shared order", "shape"),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    let bootstrap_update = bootstrap.synchronize(None).unwrap().update;

    let field_editor = presentation_store_from_update(
        &temp.path().join("field-editor"),
        930_006,
        &bootstrap_update,
    );
    let order_editor = presentation_store_from_update(
        &temp.path().join("order-editor"),
        930_007,
        &bootstrap_update,
    );
    let mut edited_title = title_element();
    edited_title["text"] = json!("Concurrent title and order");
    field_editor
        .mutate(presentation_mutation_request(
            "edit-title-during-order-move",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title_element(),
                next_element: edited_title,
            },
        ))
        .unwrap();
    order_editor
        .mutate(presentation_mutation_request(
            "move-element-during-title-edit",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-shared".to_owned(),
                expected_after_element_id: Some("element-title".to_owned()),
                after_element_id: None,
            },
        ))
        .unwrap();
    let field_update = mutation_update(&field_editor, 1);
    let order_update = mutation_update(&order_editor, 1);

    let ordered =
        presentation_store_from_update(&temp.path().join("ordered"), 930_008, &bootstrap_update);
    let reordered =
        presentation_store_from_update(&temp.path().join("reordered"), 930_009, &bootstrap_update);
    for (index, update) in [field_update.clone(), order_update.clone()]
        .into_iter()
        .enumerate()
    {
        ordered
            .apply(presentation_apply_request(
                &format!("ordered-{index}"),
                update,
            ))
            .unwrap();
    }
    for (index, update) in [order_update, field_update].into_iter().enumerate() {
        reordered
            .apply(presentation_apply_request(
                &format!("reordered-{index}"),
                update,
            ))
            .unwrap();
    }

    for store in [&ordered, &reordered] {
        assert_eq!(
            presentation_element_order(store, "slides", "slide-1"),
            vec!["element-shared", "element-title"]
        );
        assert_eq!(
            presentation_element_string(store, "slides", "slide-1", "element-title", "text"),
            Some("Concurrent title and order".to_owned())
        );
    }
    assert_eq!(
        ordered.inspect().unwrap().document_state_sha256,
        reordered.inspect().unwrap().document_state_sha256
    );
}

#[test]
fn typed_presentation_element_move_collapses_concurrent_idempotent_order_entries() {
    let temp = tempfile::tempdir().unwrap();
    let bootstrap = initialized_presentation_store(&temp.path().join("bootstrap"), 930_010);
    let bootstrap_update = bootstrap.synchronize(None).unwrap().update;
    let first = presentation_store_from_update(
        &temp.path().join("first-creator"),
        930_011,
        &bootstrap_update,
    );
    let second = presentation_store_from_update(
        &temp.path().join("second-creator"),
        930_012,
        &bootstrap_update,
    );
    let element = scene_element("element-concurrent", "Concurrent", "shape");
    for (store, operation_id) in [
        (&first, "create-concurrent-first"),
        (&second, "create-concurrent-second"),
    ] {
        store
            .mutate(presentation_mutation_request(
                operation_id,
                NativeOfficeCollaborationMutation::PresentationCreateElement {
                    container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                    container_id: "slide-1".to_owned(),
                    element: element.clone(),
                    after_element_id: Some("element-title".to_owned()),
                },
            ))
            .unwrap();
    }

    let merged =
        presentation_store_from_update(&temp.path().join("merged"), 930_013, &bootstrap_update);
    for (index, update) in [mutation_update(&first, 1), mutation_update(&second, 1)]
        .into_iter()
        .enumerate()
    {
        merged
            .apply(presentation_apply_request(
                &format!("merge-concurrent-create-{index}"),
                update,
            ))
            .unwrap();
    }
    assert_eq!(
        presentation_element_order(&merged, "slides", "slide-1"),
        vec!["element-title", "element-concurrent", "element-concurrent"]
    );

    merged
        .mutate(presentation_mutation_request(
            "move-concurrent-element-to-first",
            NativeOfficeCollaborationMutation::PresentationMoveElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-concurrent".to_owned(),
                expected_after_element_id: Some("element-title".to_owned()),
                after_element_id: None,
            },
        ))
        .unwrap();
    assert_eq!(
        presentation_element_order(&merged, "slides", "slide-1"),
        vec!["element-concurrent", "element-title"]
    );
}

fn presentation_store_from_update(
    root: &Path,
    client_id: u64,
    update: &[u8],
) -> NativeOfficeCollaborationStore {
    let store =
        NativeOfficeCollaborationStore::create(presentation_create_request(root, client_id))
            .unwrap();
    store
        .apply(presentation_apply_request(
            "bootstrap-presentation-order",
            update.to_vec(),
        ))
        .unwrap();
    store
}
