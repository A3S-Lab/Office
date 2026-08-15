use std::path::Path;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{json, Value as JsonValue};
use yrs::updates::decoder::Decode;
use yrs::{Any, Array, Doc, Map, Out, Transact, Update};

use super::*;

mod order;

const YJS_PRESENTATION_UPDATE_BASE64: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../tests/fixtures/browser-presentation-collaboration-update.base64"
));

#[test]
fn typed_presentation_elements_merge_tombstone_and_survive_restart() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replica");
    let store = initialized_presentation_store(&root, 900_004);
    let original = title_element();

    let mut text_edit = original.clone();
    text_edit["text"] = json!("Native shared title");
    text_edit["x"] = json!(14);
    store
        .mutate(presentation_mutation_request(
            "presentation-update-text",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: original.clone(),
                next_element: text_edit.clone(),
            },
        ))
        .unwrap();

    let mut stale_style_edit = original.clone();
    stale_style_edit["fill"] = json!("#DBEAFE");
    store
        .mutate(presentation_mutation_request(
            "presentation-update-style",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: original.clone(),
                next_element: stale_style_edit,
            },
        ))
        .unwrap();
    assert_eq!(
        presentation_element_string(&store, "slides", "slide-1", "element-title", "text"),
        Some("Native shared title".to_owned())
    );
    assert_eq!(
        presentation_element_number(&store, "slides", "slide-1", "element-title", "x"),
        Some(14.0)
    );
    assert_eq!(
        presentation_element_string(&store, "slides", "slide-1", "element-title", "fill"),
        Some("#DBEAFE".to_owned())
    );

    let before_conflict = store.inspect().unwrap();
    let mut stale_conflict = original.clone();
    stale_conflict["text"] = json!("Conflicting title");
    let conflict = store
        .mutate(presentation_mutation_request(
            "presentation-update-conflict",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: original,
                next_element: stale_conflict,
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_conflict.document_state_sha256
    );

    let created_element = scene_element("element-native", "Native object", "shape");
    let created = store
        .mutate(presentation_mutation_request(
            "presentation-create-element",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: created_element.clone(),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    assert!(created.state_changed);
    assert_eq!(
        presentation_element_order(&store, "slides", "slide-1"),
        vec!["element-title".to_owned(), "element-native".to_owned()]
    );
    let retry = store
        .mutate(presentation_mutation_request(
            "presentation-create-element-retry",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: created_element.clone(),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    assert!(!retry.state_changed);

    let deleted = store
        .mutate(presentation_mutation_request(
            "presentation-delete-element",
            NativeOfficeCollaborationMutation::PresentationDeleteElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                expected_element: created_element.clone(),
            },
        ))
        .unwrap();
    assert!(deleted.state_changed);
    assert!(presentation_element_tombstoned(
        &store,
        "slides",
        "slide-1",
        "element-native"
    ));
    assert_eq!(
        presentation_element_order(&store, "slides", "slide-1"),
        vec!["element-title".to_owned()]
    );
    let delete_retry = store
        .mutate(presentation_mutation_request(
            "presentation-delete-element-retry",
            NativeOfficeCollaborationMutation::PresentationDeleteElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                expected_element: created_element.clone(),
            },
        ))
        .unwrap();
    assert!(!delete_retry.state_changed);
    let reuse = store
        .mutate(presentation_mutation_request(
            "presentation-reuse-deleted-element",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: created_element,
                after_element_id: None,
            },
        ))
        .unwrap_err();
    assert_eq!(reuse.code, "office.collaboration.mutation_match_conflict");

    for (kind, collection, container_id, element_id) in [
        (
            NativeOfficeCollaborationPresentationContainerKind::Master,
            "masters",
            "master-1",
            "element-master-native",
        ),
        (
            NativeOfficeCollaborationPresentationContainerKind::Layout,
            "layouts",
            "layout-1",
            "element-layout-native",
        ),
    ] {
        store
            .mutate(presentation_mutation_request(
                &format!("presentation-create-{element_id}"),
                NativeOfficeCollaborationMutation::PresentationCreateElement {
                    container_kind: kind,
                    container_id: container_id.to_owned(),
                    element: scene_element(element_id, "Container object", "text"),
                    after_element_id: None,
                },
            ))
            .unwrap();
        assert_eq!(
            presentation_element_order(&store, collection, container_id),
            vec![element_id.to_owned()]
        );
    }

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert!(presentation_element_tombstoned(
        &reopened,
        "slides",
        "slide-1",
        "element-native"
    ));
    assert_eq!(presentation_claim_count(&reopened), 5);
}

#[test]
fn typed_presentation_validation_is_kind_bound_and_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replica");
    let store = initialized_presentation_store(&root, 900_005);
    let before = store.inspect().unwrap();

    let mut reserved = scene_element("element-reserved", "Reserved", "text");
    reserved["tombstone"] = json!(true);
    let invalid = store
        .mutate(presentation_mutation_request(
            "presentation-invalid-reserved-field",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: reserved,
                after_element_id: None,
            },
        ))
        .unwrap_err();
    assert_eq!(invalid.code, "office.collaboration.mutation_invalid");

    let mut next = title_element();
    next["type"] = json!("shape");
    let immutable = store
        .mutate(presentation_mutation_request(
            "presentation-invalid-type-change",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title_element(),
                next_element: next,
            },
        ))
        .unwrap_err();
    assert_eq!(immutable.code, "office.collaboration.mutation_invalid");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );

    let markdown_root = temp.path().join("markdown-replica");
    let markdown = NativeOfficeCollaborationStore::create(create_request(&markdown_root)).unwrap();
    let kind_error = markdown
        .mutate(mutation_request(
            "presentation-on-markdown",
            NativeOfficeCollaborationMutation::PresentationDeleteElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                expected_element: title_element(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        kind_error.code,
        "office.collaboration.mutation_kind_mismatch"
    );
}

#[test]
fn typed_presentation_updates_converge_across_reordered_native_delivery() {
    let temp = tempfile::tempdir().unwrap();
    let first = initialized_presentation_store(&temp.path().join("first"), 910_001);
    let second = initialized_presentation_store(&temp.path().join("second"), 910_002);
    let mut text_edit = title_element();
    text_edit["text"] = json!("Concurrent native title");
    first
        .mutate(presentation_mutation_request(
            "concurrent-text",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title_element(),
                next_element: text_edit,
            },
        ))
        .unwrap();
    let mut fill_edit = title_element();
    fill_edit["fill"] = json!("#E0E7FF");
    second
        .mutate(presentation_mutation_request(
            "concurrent-fill",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title_element(),
                next_element: fill_edit,
            },
        ))
        .unwrap();
    let first_update = mutation_update(&first, 1);
    let second_update = mutation_update(&second, 1);

    let ordered = initialized_presentation_store(&temp.path().join("ordered"), 910_003);
    let reordered = initialized_presentation_store(&temp.path().join("reordered"), 910_004);
    for (index, update) in [first_update.clone(), second_update.clone()]
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
    for (index, update) in [second_update, first_update.clone(), first_update]
        .into_iter()
        .enumerate()
    {
        reordered
            .apply(presentation_apply_request(
                &format!("reordered-{index}"),
                update,
            ))
            .unwrap();
    }

    for store in [&ordered, &reordered] {
        assert_eq!(
            presentation_element_string(store, "slides", "slide-1", "element-title", "text"),
            Some("Concurrent native title".to_owned())
        );
        assert_eq!(
            presentation_element_string(store, "slides", "slide-1", "element-title", "fill"),
            Some("#E0E7FF".to_owned())
        );
    }
    assert_eq!(
        ordered.inspect().unwrap().document_state_sha256,
        reordered.inspect().unwrap().document_state_sha256
    );
}

#[test]
fn typed_presentation_claims_fail_closed_after_concurrent_identity_collision() {
    let temp = tempfile::tempdir().unwrap();
    let first = initialized_presentation_store(&temp.path().join("first"), 920_001);
    let second = initialized_presentation_store(&temp.path().join("second"), 920_002);
    for (store, operation_id, text) in [
        (&first, "collision-first", "Ada object"),
        (&second, "collision-second", "Grace object"),
    ] {
        store
            .mutate(presentation_mutation_request(
                operation_id,
                NativeOfficeCollaborationMutation::PresentationCreateElement {
                    container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                    container_id: "slide-1".to_owned(),
                    element: scene_element("element-collision", text, "shape"),
                    after_element_id: None,
                },
            ))
            .unwrap();
    }
    let target = initialized_presentation_store(&temp.path().join("target"), 920_003);
    target
        .apply(presentation_apply_request(
            "deliver-collision-first",
            mutation_update(&first, 1),
        ))
        .unwrap();
    target
        .apply(presentation_apply_request(
            "deliver-collision-second",
            mutation_update(&second, 1),
        ))
        .unwrap();

    let error = target
        .mutate(presentation_mutation_request(
            "mutation-after-collision",
            NativeOfficeCollaborationMutation::PresentationDeleteElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                expected_element: title_element(),
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.content_invalid");
    assert!(error.message.contains("concurrently assigned"));
}

fn initialized_presentation_store(root: &Path, client_id: u64) -> NativeOfficeCollaborationStore {
    let store =
        NativeOfficeCollaborationStore::create(presentation_create_request(root, client_id))
            .unwrap();
    store
        .apply(presentation_apply_request(
            "bootstrap-browser-presentation",
            STANDARD
                .decode(YJS_PRESENTATION_UPDATE_BASE64.trim())
                .unwrap(),
        ))
        .unwrap();
    store
}

fn presentation_create_request(
    root: &Path,
    client_id: u64,
) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-presentation".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Presentation,
        actor_id: "agent-presentation".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-presentation-1".to_owned(),
        namespace: None,
        client_id: Some(client_id),
        initial_update: None,
    }
}

fn presentation_apply_request(
    operation_id: &str,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-presentation".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-presentation".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Presentation,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn presentation_mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-presentation".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-presentation".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Presentation,
        mutation,
        if_state_vector: None,
    }
}

fn title_element() -> JsonValue {
    scene_element_with_geometry(
        "element-title",
        "Shared presentation",
        "text",
        10,
        10,
        80,
        20,
        32,
        "transparent",
        true,
        "center",
    )
}

fn scene_element(id: &str, text: &str, element_type: &str) -> JsonValue {
    scene_element_with_geometry(
        id,
        text,
        element_type,
        20,
        20,
        40,
        20,
        18,
        "#F8FAFC",
        false,
        "left",
    )
}

#[allow(clippy::too_many_arguments)]
fn scene_element_with_geometry(
    id: &str,
    text: &str,
    element_type: &str,
    x: u32,
    y: u32,
    width: u32,
    height: u32,
    font_size: u32,
    fill: &str,
    bold: bool,
    align: &str,
) -> JsonValue {
    json!({
        "id": id,
        "type": element_type,
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "text": text,
        "fontSize": font_size,
        "color": "#172033",
        "fill": fill,
        "bold": bold,
        "align": align,
    })
}

fn presentation_peer(store: &NativeOfficeCollaborationStore) -> Doc {
    let exported = store.synchronize(None).unwrap();
    let peer = Doc::with_client_id(818_184);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    peer
}

fn presentation_element_field(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
    element_id: &str,
    field: &str,
) -> Option<Any> {
    let peer = presentation_peer(store);
    let containers = peer.get_or_insert_map(format!("a3s.office.presentation.{collection}"));
    let transaction = peer.transact();
    let container = match containers.get(&transaction, container_id) {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation container: {value:?}"),
    };
    let elements = match container.get(&transaction, "elements") {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation elements root: {value:?}"),
    };
    let element = match elements.get(&transaction, element_id) {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation element: {value:?}"),
    };
    match element.get(&transaction, field) {
        Some(Out::Any(value)) => Some(value),
        None => None,
        value => panic!("unexpected Presentation element field: {value:?}"),
    }
}

fn presentation_element_string(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
    element_id: &str,
    field: &str,
) -> Option<String> {
    match presentation_element_field(store, collection, container_id, element_id, field) {
        Some(Any::String(value)) => Some(value.to_string()),
        None => None,
        value => panic!("unexpected Presentation string field: {value:?}"),
    }
}

fn presentation_element_number(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
    element_id: &str,
    field: &str,
) -> Option<f64> {
    match presentation_element_field(store, collection, container_id, element_id, field) {
        Some(Any::Number(value)) => Some(value),
        None => None,
        value => panic!("unexpected Presentation number field: {value:?}"),
    }
}

fn presentation_element_tombstoned(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
    element_id: &str,
) -> bool {
    matches!(
        presentation_element_field(store, collection, container_id, element_id, "tombstone"),
        Some(Any::Bool(true))
    )
}

fn presentation_element_order(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
) -> Vec<String> {
    let peer = presentation_peer(store);
    let containers = peer.get_or_insert_map(format!("a3s.office.presentation.{collection}"));
    let transaction = peer.transact();
    let container = match containers.get(&transaction, container_id) {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation container: {value:?}"),
    };
    let order = match container.get(&transaction, "elementOrder") {
        Some(Out::YArray(value)) => value,
        value => panic!("unexpected Presentation element order: {value:?}"),
    };
    (0..order.len(&transaction))
        .map(|index| match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            value => panic!("unexpected Presentation order entry: {value:?}"),
        })
        .collect()
}

fn presentation_claim_count(store: &NativeOfficeCollaborationStore) -> u32 {
    let peer = presentation_peer(store);
    let claims = peer.get_or_insert_array("a3s.office.presentation.record-claims");
    let count = claims.len(&peer.transact());
    count
}

fn mutation_update(store: &NativeOfficeCollaborationStore, after_sequence: u64) -> Vec<u8> {
    let batch = store
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(after_sequence),
            limit: 10,
        })
        .unwrap();
    assert_eq!(batch.updates.len(), 1);
    batch.updates[0].update.clone()
}
