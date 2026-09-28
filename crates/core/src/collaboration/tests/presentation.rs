use std::collections::BTreeMap;
use std::path::Path;

use base64::engine::general_purpose::STANDARD;
use serde_json::{json, Value as JsonValue};
use yrs::updates::decoder::Decode;
use yrs::{Any, Array, Doc, Map, Out, Transact, Update};

use super::*;

mod deck;
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

#[test]
fn find_lists_presentation_element_text_by_identity() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-find");
    let store = initialized_presentation_store(&root, 900_014);
    let (kind, found) = store.find_text("Shared presentation", 10).unwrap();
    assert_eq!(kind, NativeOfficeCollaborationArtifactKind::Presentation);
    assert_eq!(found.match_count, 1);
    assert!(!found.truncated);
    assert_eq!(found.matches[0].occurrence, 1);
    assert_eq!(found.matches[0].container_kind.as_deref(), Some("slide"));
    assert_eq!(found.matches[0].container_id.as_deref(), Some("slide-1"));
    assert_eq!(
        found.matches[0].element_id.as_deref(),
        Some("element-title")
    );
    assert_eq!(found.matches[0].index_utf16, 0);
}

#[test]
fn replace_presentation_text_changes_only_the_matched_span() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("presentation-replace-text");
    let store = initialized_presentation_store(&root, 900_015);
    let (kind, found) = store.find_text("presentation", 10).unwrap();
    assert_eq!(kind, NativeOfficeCollaborationArtifactKind::Presentation);
    store
        .mutate(presentation_mutation_request(
            "replace-live-title",
            NativeOfficeCollaborationMutation::PresentationReplaceText {
                search: "presentation".to_owned(),
                replacement: "deck".to_owned(),
                expected_matches: found.match_count as u32,
                occurrence: Some(1),
                container_kind: None,
                container_id: None,
                element_id: None,
                index_utf16: None,
            },
        ))
        .unwrap();
    assert_eq!(
        presentation_element_string(&store, "slides", "slide-1", "element-title", "text")
            .as_deref(),
        Some("Shared deck")
    );
    assert_eq!(
        presentation_element_number(&store, "slides", "slide-1", "element-title", "x"),
        Some(10.0)
    );

    let conflict = store
        .mutate(presentation_mutation_request(
            "replace-stale-count",
            NativeOfficeCollaborationMutation::PresentationReplaceText {
                search: "presentation".to_owned(),
                replacement: "slides".to_owned(),
                expected_matches: 1,
                occurrence: None,
                container_kind: None,
                container_id: None,
                element_id: None,
                index_utf16: None,
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
}

#[test]
fn presentation_splice_is_element_local_and_survives_an_unrelated_shape_edit() {
    let temp = tempfile::tempdir().unwrap();
    let first = initialized_presentation_store(&temp.path().join("first"), 900_041);
    let second = initialized_presentation_store(&temp.path().join("second"), 900_042);
    let title = projected_element_text(&first, "element-title");
    assert_eq!(title, "Shared presentation");
    let end = u32::try_from(title.encode_utf16().count()).unwrap();
    let inserted = first
        .mutate(presentation_mutation_request(
            "presentation-splice-cjk",
            NativeOfficeCollaborationMutation::PresentationSplice {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                index_utf16: end,
                delete_utf16: 0,
                expected_slice: String::new(),
                insert: "中".to_owned(),
            },
        ))
        .unwrap();
    let NativeOfficeCollaborationFrameCaret::Presentation {
        element_id,
        index_utf16,
        ..
    } = inserted.caret.unwrap()
    else {
        panic!("presentation splice must return an element caret");
    };
    assert_eq!(element_id, "element-title");
    assert_eq!(index_utf16, Some(end + 1));

    second
        .apply(presentation_apply_request(
            "apply-first-splice",
            first.synchronize(None).unwrap().update,
        ))
        .unwrap();
    let followed = second
        .mutate(presentation_mutation_request(
            "presentation-splice-follow",
            NativeOfficeCollaborationMutation::PresentationSplice {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                index_utf16: end + 1,
                delete_utf16: 0,
                expected_slice: String::new(),
                insert: "文".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(
        followed.caret.unwrap(),
        NativeOfficeCollaborationFrameCaret::Presentation {
            container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
            container_id: "slide-1".to_owned(),
            element_id: "element-title".to_owned(),
            index_utf16: Some(end + 2),
        }
    );
    first
        .apply(presentation_apply_request(
            "apply-second-splice",
            second
                .synchronize(Some(&first.synchronize(None).unwrap().state_vector))
                .unwrap()
                .update,
        ))
        .unwrap();

    assert_eq!(
        projected_element_text(&first, "element-title"),
        "Shared presentation中文"
    );
    assert_eq!(
        projected_element_text(&second, "element-title"),
        "Shared presentation中文"
    );

    let mut styled = title_element();
    styled["fill"] = json!("#DBEAFE");
    first
        .mutate(presentation_mutation_request(
            "presentation-fill-does-not-move-text",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title_element(),
                next_element: styled,
            },
        ))
        .unwrap();
    assert_eq!(
        projected_element_text(&first, "element-title"),
        "Shared presentation中文"
    );

    let drifted = first
        .mutate(presentation_mutation_request(
            "presentation-splice-drift",
            NativeOfficeCollaborationMutation::PresentationSplice {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                index_utf16: 0,
                delete_utf16: 1,
                expected_slice: "X".to_owned(),
                insert: "Y".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(drifted.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        projected_element_text(&first, "element-title"),
        "Shared presentation中文"
    );
}

fn projected_element_text(store: &NativeOfficeCollaborationStore, element_id: &str) -> String {
    let projection = store.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Presentation { containers } = projection.content
    else {
        panic!("expected a presentation projection");
    };
    containers
        .iter()
        .flat_map(|container| container.elements.iter())
        .find(|element| element.element_id == element_id)
        .unwrap_or_else(|| panic!("missing element {element_id}"))
        .element["text"]
        .as_str()
        .unwrap()
        .to_owned()
}

#[test]
fn presentation_replica_export_writes_every_shape_text_or_returns_no_package() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_presentation_store(&temp.path().join("export-presentation"), 900_014);
    let slide = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"element-title\" name=\"Title\"/><p:txBody><a:p><a:r><a:t>Old</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:cNvPr id=\"kept\" name=\"Kept\"/><p:txBody><a:p><a:r><a:t>KEEP</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
    let details = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"element-body\" name=\"Body\"/><p:txBody><a:p><a:r><a:t>Old</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
    let mut parts = BTreeMap::new();
    parts.insert(
        "ppt/slides/slide1.xml".to_owned(),
        slide.as_bytes().to_vec(),
    );
    parts.insert(
        "ppt/slides/slide2.xml".to_owned(),
        details.as_bytes().to_vec(),
    );
    parts.insert("ppt/presentation.xml".to_owned(), b"<kept/>".to_vec());
    let containers = [
        ("slide-1", "ppt/slides/slide1.xml"),
        ("slide-2", "ppt/slides/slide2.xml"),
    ];

    let written = export_presentation_replica(&store, parts.clone(), &containers).unwrap();
    let cover = String::from_utf8(written["ppt/slides/slide1.xml"].clone()).unwrap();
    let body = String::from_utf8(written["ppt/slides/slide2.xml"].clone()).unwrap();
    assert!(cover.contains("<a:t>Shared presentation</a:t>"));
    assert!(cover.contains("<a:t>KEEP</a:t>"));
    assert!(body.contains("<a:t>Body</a:t>"));
    assert_eq!(
        written["ppt/presentation.xml"],
        parts["ppt/presentation.xml"]
    );

    let two_runs = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"element-body\" name=\"Body\"/><p:txBody><a:p><a:r><a:t>A</a:t></a:r><a:r><a:t>B</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
    let mut split = parts.clone();
    split.insert(
        "ppt/slides/slide2.xml".to_owned(),
        two_runs.as_bytes().to_vec(),
    );
    assert!(export_presentation_replica(&store, split, &containers).is_err());

    assert!(export_presentation_replica(
        &store,
        parts.clone(),
        &[("slide-1", "ppt/slides/slide1.xml")],
    )
    .is_err());

    let missing_shape = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"other\" name=\"Other\"/><p:txBody><a:p><a:r><a:t>Old</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
    let mut unnamed = parts;
    unnamed.insert(
        "ppt/slides/slide2.xml".to_owned(),
        missing_shape.as_bytes().to_vec(),
    );
    assert!(export_presentation_replica(&store, unnamed, &containers).is_err());
}

#[test]
fn reordered_presentation_updates_converge_and_reopen_the_merged_export() {
    let temp = tempfile::tempdir().unwrap();
    let title = title_element();
    let body = body_element();
    let mut title_text = title.clone();
    title_text["text"] = json!("Concurrent native title");
    let mut title_fill = title.clone();
    title_fill["fill"] = json!("#E0E7FF");
    let mut body_fill = body.clone();
    body_fill["fill"] = json!("#FDE68A");
    let updates = [
        offline_presentation_update(
            &temp,
            930_111,
            "offline-title-text",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title,
                next_element: title_text,
            },
        ),
        offline_presentation_update(
            &temp,
            930_112,
            "offline-title-fill",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element_id: "element-title".to_owned(),
                expected_element: title_element(),
                next_element: title_fill,
            },
        ),
        offline_presentation_update(
            &temp,
            930_113,
            "offline-body-fill",
            NativeOfficeCollaborationMutation::PresentationUpdateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-2".to_owned(),
                element_id: "element-body".to_owned(),
                expected_element: body,
                next_element: body_fill,
            },
        ),
    ];

    let mut hashes = Vec::new();
    for (permutation_index, permutation) in
        three_presentation_permutations().into_iter().enumerate()
    {
        let ordered = [
            updates[permutation[0]].clone(),
            updates[permutation[1]].clone(),
            updates[permutation[2]].clone(),
        ];
        let store = deliver_presentation_updates(
            &temp.path().join(format!("order-{permutation_index}")),
            930_200 + permutation_index as u64,
            &ordered,
        );
        assert_eq!(
            presentation_element_string(&store, "slides", "slide-1", "element-title", "text"),
            Some("Concurrent native title".to_owned())
        );
        assert_eq!(
            presentation_element_string(&store, "slides", "slide-1", "element-title", "fill"),
            Some("#E0E7FF".to_owned())
        );
        assert_eq!(
            presentation_element_string(&store, "slides", "slide-2", "element-body", "text"),
            Some("Body".to_owned())
        );
        assert_eq!(
            presentation_element_string(&store, "slides", "slide-2", "element-body", "fill"),
            Some("#FDE68A".to_owned())
        );
        hashes.push(store.inspect().unwrap().document_state_sha256);
    }
    assert!(hashes.iter().all(|hash| hash == &hashes[0]));

    let repeated = [
        updates[0].clone(),
        updates[0].clone(),
        updates[1].clone(),
        updates[2].clone(),
    ];
    let replayed = deliver_presentation_updates(&temp.path().join("repeated"), 930_220, &repeated);
    assert_eq!(replayed.inspect().unwrap().document_state_sha256, hashes[0]);

    let slide = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"element-title\" name=\"Title\"/><p:txBody><a:p><a:r><a:t>Old</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:cNvPr id=\"kept\" name=\"Kept\"/><p:txBody><a:p><a:r><a:t>KEEP</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
    let details = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:sp><p:cNvPr id=\"element-body\" name=\"Body\"/><p:txBody><a:p><a:r><a:t>Old</a:t></a:r></a:p></p:txBody></p:sp></p:sld>";
    let mut parts = BTreeMap::new();
    parts.insert(
        "ppt/slides/slide1.xml".to_owned(),
        slide.as_bytes().to_vec(),
    );
    parts.insert(
        "ppt/slides/slide2.xml".to_owned(),
        details.as_bytes().to_vec(),
    );
    parts.insert("ppt/presentation.xml".to_owned(), b"<kept/>".to_vec());
    let written = export_presentation_replica(
        &replayed,
        parts.clone(),
        &[
            ("slide-1", "ppt/slides/slide1.xml"),
            ("slide-2", "ppt/slides/slide2.xml"),
        ],
    )
    .unwrap();
    assert_eq!(
        written["ppt/presentation.xml"],
        parts["ppt/presentation.xml"]
    );
    let cover = String::from_utf8(written["ppt/slides/slide1.xml"].clone()).unwrap();
    let body_xml = String::from_utf8(written["ppt/slides/slide2.xml"].clone()).unwrap();
    assert_eq!(
        reopen_exported_shape_text(&cover, "element-title").unwrap(),
        "Concurrent native title"
    );
    assert_eq!(reopen_exported_shape_text(&cover, "kept").unwrap(), "KEEP");
    assert_eq!(
        reopen_exported_shape_text(&body_xml, "element-body").unwrap(),
        "Body"
    );
}

fn body_element() -> JsonValue {
    json!({
        "id": "element-body",
        "type": "shape",
        "x": 20,
        "y": 25,
        "width": 60,
        "height": 40,
        "text": "Body",
        "fontSize": 18,
        "color": "#172033",
        "fill": "#DCE6FB",
        "bold": false,
        "align": "left",
    })
}

fn offline_presentation_update(
    temp: &tempfile::TempDir,
    client_id: u64,
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> Vec<u8> {
    let store = initialized_presentation_store(&temp.path().join(operation_id), client_id);
    let base = store.synchronize(None).unwrap().state_vector;
    store
        .mutate(presentation_mutation_request(operation_id, mutation))
        .unwrap();
    store.synchronize(Some(&base)).unwrap().update
}

fn deliver_presentation_updates(
    root: &Path,
    client_id: u64,
    updates: &[Vec<u8>],
) -> NativeOfficeCollaborationStore {
    let store = initialized_presentation_store(root, client_id);
    for (index, update) in updates.iter().enumerate() {
        store
            .apply(presentation_apply_request(
                &format!("deliver-{client_id}-{index}"),
                update.clone(),
            ))
            .unwrap();
    }
    store
}

fn three_presentation_permutations() -> Vec<[usize; 3]> {
    let mut result = Vec::with_capacity(6);
    for first in 0..3 {
        for second in 0..3 {
            for third in 0..3 {
                let permutation = [first, second, third];
                if permutation
                    .iter()
                    .copied()
                    .collect::<std::collections::HashSet<_>>()
                    .len()
                    == 3
                {
                    result.push(permutation);
                }
            }
        }
    }
    assert_eq!(result.len(), 6);
    result
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
