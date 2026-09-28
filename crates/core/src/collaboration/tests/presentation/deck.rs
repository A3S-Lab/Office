use super::*;

const TITLE: &str = "Shared presentation";
const BODY: &str = "Body";

#[test]
fn presentation_slide_move_rewrites_only_the_order_array() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replica");
    let store = initialized_presentation_store(&root, 940_001);
    assert_eq!(
        presentation_slide_order(&store),
        vec!["slide-1".to_owned(), "slide-2".to_owned()]
    );
    assert_shape_text(&store);

    let before_same = store.inspect().unwrap();
    let same = store
        .mutate(presentation_mutation_request(
            "slide-already-there",
            slide_move("slide-2", Some("slide-1"), Some("slide-1")),
        ))
        .unwrap();
    assert!(!same.state_changed);
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_same.document_state_sha256
    );
    assert_shape_text(&store);

    let before_stale = store.inspect().unwrap();
    let stale = store
        .mutate(presentation_mutation_request(
            "slide-stale-predecessor",
            slide_move("slide-2", None, None),
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );
    assert_eq!(
        presentation_slide_order(&store),
        vec!["slide-1".to_owned(), "slide-2".to_owned()]
    );
    assert_shape_text(&store);

    let moved = store
        .mutate(presentation_mutation_request(
            "slide-to-front",
            slide_move("slide-2", Some("slide-1"), None),
        ))
        .unwrap();
    assert!(moved.state_changed);
    assert_eq!(
        presentation_slide_order(&store),
        vec!["slide-2".to_owned(), "slide-1".to_owned()]
    );
    assert_shape_text(&store);
    assert_eq!(
        presentation_background(&store, "slides", "slide-1").as_deref(),
        Some("#FFFFFF")
    );
    assert_eq!(
        presentation_background(&store, "slides", "slide-2").as_deref(),
        Some("#FFFFFF")
    );
}

#[test]
fn presentation_group_paths_commit_together_or_not_at_all() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replica");
    let store = initialized_presentation_store(&root, 940_002);
    for (id, after) in [
        ("element-a", Some("element-title")),
        ("element-b", Some("element-a")),
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
    let order_before = presentation_element_order(&store, "slides", "slide-1");
    assert_shape_text(&store);
    assert!(presentation_group_ids(&store, "slides", "slide-1", "element-a").is_empty());
    assert!(presentation_group_ids(&store, "slides", "slide-1", "element-title").is_empty());

    let before_stale = store.inspect().unwrap();
    let stale = store
        .mutate(presentation_mutation_request(
            "group-one-member-stale",
            set_group(vec![
                group_member("element-a", &["nope"], &["group-1"]),
                group_member("element-b", &[], &["group-1"]),
            ]),
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );
    assert!(presentation_group_ids(&store, "slides", "slide-1", "element-a").is_empty());
    assert!(presentation_group_ids(&store, "slides", "slide-1", "element-b").is_empty());
    assert_shape_text(&store);
    assert_eq!(
        presentation_element_order(&store, "slides", "slide-1"),
        order_before
    );

    let grouped = store
        .mutate(presentation_mutation_request(
            "group-both",
            set_group(vec![
                group_member("element-a", &[], &["group-1"]),
                group_member("element-b", &[], &["group-1"]),
            ]),
        ))
        .unwrap();
    assert!(grouped.state_changed);
    assert_eq!(
        presentation_group_ids(&store, "slides", "slide-1", "element-a"),
        vec!["group-1".to_owned()]
    );
    assert_eq!(
        presentation_group_ids(&store, "slides", "slide-1", "element-b"),
        vec!["group-1".to_owned()]
    );
    assert!(presentation_group_ids(&store, "slides", "slide-1", "element-title").is_empty());
    assert_shape_text(&store);
    assert_eq!(
        presentation_element_string(&store, "slides", "slide-1", "element-a", "text").as_deref(),
        Some("element-a")
    );

    let before_partial = store.inspect().unwrap();
    let partial = store
        .mutate(presentation_mutation_request(
            "group-second-member-stale",
            set_group(vec![
                group_member("element-a", &["group-1"], &["group-2"]),
                group_member("element-b", &[], &["group-2"]),
            ]),
        ))
        .unwrap_err();
    assert_eq!(partial.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_partial.document_state_sha256
    );
    assert_eq!(
        presentation_group_ids(&store, "slides", "slide-1", "element-a"),
        vec!["group-1".to_owned()]
    );
    assert_eq!(
        presentation_group_ids(&store, "slides", "slide-1", "element-b"),
        vec!["group-1".to_owned()]
    );
    assert_shape_text(&store);
}

#[test]
fn presentation_background_is_one_field_on_slide_master_and_layout() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replica");
    let store = initialized_presentation_store(&root, 940_003);
    assert_shape_text(&store);
    let order_before = presentation_slide_order(&store);

    write_background(
        &store,
        "slide-background-same",
        SlideKind::Slide,
        "slide-1",
        true,
    );
    write_background(
        &store,
        "slide-background",
        SlideKind::Slide,
        "slide-1",
        false,
    );
    write_background(
        &store,
        "master-background",
        SlideKind::Master,
        "master-1",
        false,
    );
    write_background(
        &store,
        "layout-background",
        SlideKind::Layout,
        "layout-1",
        false,
    );

    assert_eq!(
        presentation_background(&store, "slides", "slide-1").as_deref(),
        Some("#010203")
    );
    assert_eq!(
        presentation_background(&store, "slides", "slide-2").as_deref(),
        Some("#FFFFFF")
    );
    assert_eq!(
        presentation_background(&store, "masters", "master-1").as_deref(),
        Some("#020304")
    );
    assert_eq!(
        presentation_background(&store, "layouts", "layout-1").as_deref(),
        Some("#030405")
    );
    assert_eq!(presentation_slide_order(&store), order_before);
    assert_shape_text(&store);
    assert_eq!(
        presentation_element_string(&store, "slides", "slide-1", "element-title", "text")
            .as_deref(),
        Some(TITLE)
    );
}

enum SlideKind {
    Slide,
    Master,
    Layout,
}

fn write_background(
    store: &NativeOfficeCollaborationStore,
    operation_id: &str,
    kind: SlideKind,
    container_id: &str,
    idempotent: bool,
) {
    let (container_kind, collection, next) = match kind {
        SlideKind::Slide => (
            NativeOfficeCollaborationPresentationContainerKind::Slide,
            "slides",
            "#010203",
        ),
        SlideKind::Master => (
            NativeOfficeCollaborationPresentationContainerKind::Master,
            "masters",
            "#020304",
        ),
        SlideKind::Layout => (
            NativeOfficeCollaborationPresentationContainerKind::Layout,
            "layouts",
            "#030405",
        ),
    };
    let current = presentation_background(store, collection, container_id);
    let before_stale = store.inspect().unwrap();
    let stale_expected = match &current {
        Some(_) => None,
        None => Some("#000000".to_owned()),
    };
    let stale = store
        .mutate(presentation_mutation_request(
            &format!("{operation_id}-stale"),
            NativeOfficeCollaborationMutation::PresentationSetBackground {
                container_kind,
                container_id: container_id.to_owned(),
                expected_background: stale_expected,
                next_background: Some(next.to_owned()),
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );
    assert_eq!(
        presentation_background(store, collection, container_id),
        current
    );
    assert_shape_text(store);

    if idempotent {
        let before_same = store.inspect().unwrap();
        let same = store
            .mutate(presentation_mutation_request(
                &format!("{operation_id}-same"),
                NativeOfficeCollaborationMutation::PresentationSetBackground {
                    container_kind,
                    container_id: container_id.to_owned(),
                    expected_background: current.clone(),
                    next_background: current.clone(),
                },
            ))
            .unwrap();
        assert!(!same.state_changed);
        assert_eq!(
            store.inspect().unwrap().document_state_sha256,
            before_same.document_state_sha256
        );
        return;
    }

    let wrote = store
        .mutate(presentation_mutation_request(
            operation_id,
            NativeOfficeCollaborationMutation::PresentationSetBackground {
                container_kind,
                container_id: container_id.to_owned(),
                expected_background: current,
                next_background: Some(next.to_owned()),
            },
        ))
        .unwrap();
    assert!(wrote.state_changed);
    assert_eq!(
        presentation_background(store, collection, container_id).as_deref(),
        Some(next)
    );
    assert_shape_text(store);
}

fn slide_move(
    slide_id: &str,
    expected_after_slide_id: Option<&str>,
    after_slide_id: Option<&str>,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::PresentationMoveSlide {
        slide_id: slide_id.to_owned(),
        expected_after_slide_id: expected_after_slide_id.map(str::to_owned),
        after_slide_id: after_slide_id.map(str::to_owned),
    }
}

fn set_group(
    members: Vec<NativeOfficeCollaborationPresentationGroupMember>,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::PresentationSetGroup {
        container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
        container_id: "slide-1".to_owned(),
        members,
    }
}

fn group_member(
    element_id: &str,
    expected_group_ids: &[&str],
    next_group_ids: &[&str],
) -> NativeOfficeCollaborationPresentationGroupMember {
    NativeOfficeCollaborationPresentationGroupMember {
        element_id: element_id.to_owned(),
        expected_group_ids: expected_group_ids
            .iter()
            .map(|value| (*value).to_owned())
            .collect(),
        next_group_ids: next_group_ids
            .iter()
            .map(|value| (*value).to_owned())
            .collect(),
    }
}

fn assert_shape_text(store: &NativeOfficeCollaborationStore) {
    assert_eq!(
        presentation_element_string(store, "slides", "slide-1", "element-title", "text").as_deref(),
        Some(TITLE)
    );
    assert_eq!(
        presentation_element_string(store, "slides", "slide-2", "element-body", "text").as_deref(),
        Some(BODY)
    );
}

fn presentation_slide_order(store: &NativeOfficeCollaborationStore) -> Vec<String> {
    let peer = presentation_peer(store);
    let order = peer.get_or_insert_array("a3s.office.presentation.slide-order");
    let transaction = peer.transact();
    (0..order.len(&transaction))
        .map(|index| match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            value => panic!("unexpected slide order entry: {value:?}"),
        })
        .collect()
}

fn presentation_background(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
) -> Option<String> {
    let peer = presentation_peer(store);
    let containers = peer.get_or_insert_map(format!("a3s.office.presentation.{collection}"));
    let transaction = peer.transact();
    let container = match containers.get(&transaction, container_id) {
        Some(Out::YMap(value)) => value,
        value => panic!("unexpected Presentation container: {value:?}"),
    };
    match container.get(&transaction, "background") {
        None => None,
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        value => panic!("unexpected Presentation background: {value:?}"),
    }
}

fn presentation_group_ids(
    store: &NativeOfficeCollaborationStore,
    collection: &str,
    container_id: &str,
    element_id: &str,
) -> Vec<String> {
    match presentation_element_field(store, collection, container_id, element_id, "groupIds") {
        None => Vec::new(),
        Some(Any::Array(values)) => values
            .iter()
            .map(|value| match value {
                Any::String(value) => value.to_string(),
                value => panic!("unexpected group id: {value:?}"),
            })
            .collect(),
        value => panic!("unexpected group path: {value:?}"),
    }
}
