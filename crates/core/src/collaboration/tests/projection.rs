use super::*;

#[test]
fn projects_canonical_markdown_after_reordered_delivery_and_restart() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("projected-markdown-replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    store
        .apply(apply_request(
            "projection-reordered-second",
            STANDARD.decode(YJS_REORDERED_SECOND_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let pending = store.project().unwrap_err();
    assert!(matches!(
        pending.code.as_str(),
        "office.collaboration.not_initialized" | "office.collaboration.projection_incomplete"
    ));

    store
        .apply(apply_request(
            "projection-reordered-first",
            STANDARD.decode(YJS_REORDERED_FIRST_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let projection = store.project().unwrap();
    assert_eq!(projection.sequence, 2);
    assert_eq!(projection.artifact_id, "fixture-markdown");
    assert_eq!(
        projection.content,
        NativeOfficeCollaborationProjectedContent::Markdown {
            source: "AB".to_owned()
        }
    );

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(reopened.project().unwrap(), projection);
}

#[test]
fn document_projection_exposes_stable_identity_for_guarded_replacement() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("projected-document-replica");
    let store = NativeOfficeCollaborationStore::create(document_create_request(&root)).unwrap();
    store
        .apply(document_apply_request(
            "projection-bootstrap-document",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();

    let before = store.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        comments,
        page_color,
        track_changes,
        ..
    } = before.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Hello 😀 world");
    assert_eq!(page_color.as_deref(), Some("#F8FAFC"));
    assert_eq!(track_changes, Some(true));
    assert!(comments.is_empty());
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].paragraph_id.as_deref(), Some("00000001"));
    assert_eq!(paragraphs[0].text_id.as_deref(), Some("00000002"));
    assert_eq!(paragraphs[0].container_path, vec!["documentSection"]);
    assert!(paragraphs[0].replaceable);

    store
        .mutate(document_mutation_request(
            "projection-replace-paragraph",
            NativeOfficeCollaborationMutation::DocumentReplaceParagraph {
                paragraph_id: "00000001".to_owned(),
                expected_text_id: "00000002".to_owned(),
                expected_text: "Hello 😀 world".to_owned(),
                replacement: "Agent and human now share this paragraph.".to_owned(),
            },
        ))
        .unwrap();
    let after = store.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        ..
    } = &after.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Agent and human now share this paragraph.");
    assert_eq!(paragraphs[0].text_id.as_deref(), Some("00000003"));

    drop(store);
    assert_eq!(
        NativeOfficeCollaborationStore::open(&root)
            .unwrap()
            .project()
            .unwrap(),
        after
    );
}

#[test]
fn browser_edit_after_projection_makes_stale_agent_paragraph_guard_fail_closed() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("browser-edit-projection-replica");
    let store = NativeOfficeCollaborationStore::create(document_create_request(&root)).unwrap();
    store
        .apply(document_apply_request(
            "projection-browser-bootstrap",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let stale_projection = store.project().unwrap();

    let browser = Doc::with_client_id(424_299);
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&store.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    let before_browser_edit = browser.transact().state_vector();
    let fragment = browser.get_or_insert_xml_fragment("a3s.office.document.content");
    let (paragraph, text) = {
        let transaction = browser.transact();
        let paragraph = fragment
            .successors(&transaction)
            .find_map(|node| match node {
                XmlOut::Element(element) if element.tag().as_ref() == "paragraph" => Some(element),
                _ => None,
            })
            .unwrap();
        let text = paragraph
            .children(&transaction)
            .find_map(|node| match node {
                XmlOut::Text(text) => Some(text),
                _ => None,
            })
            .unwrap();
        (paragraph, text)
    };
    let mut transaction = browser.transact_mut();
    let end = text.len(&transaction);
    text.insert(&mut transaction, end, " edited by user");
    paragraph.insert_attribute(&mut transaction, "textId", "00000003");
    drop(transaction);
    let browser_update = browser
        .transact()
        .encode_state_as_update_v1(&before_browser_edit);
    store
        .apply(document_apply_request(
            "projection-browser-user-edit",
            browser_update,
        ))
        .unwrap();

    let conflict = store
        .mutate(document_mutation_request(
            "projection-stale-agent-replace",
            NativeOfficeCollaborationMutation::DocumentReplaceParagraph {
                paragraph_id: "00000001".to_owned(),
                expected_text_id: "00000002".to_owned(),
                expected_text: "Hello 😀 world".to_owned(),
                replacement: "This stale write must not win.".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    let current = store.project().unwrap();
    assert_ne!(current.state_vector, stale_projection.state_vector);
    let NativeOfficeCollaborationProjectedContent::Document {
        plain_text,
        paragraphs,
        ..
    } = current.content
    else {
        panic!("expected Document projection");
    };
    assert_eq!(plain_text, "Hello 😀 world edited by user");
    assert_eq!(paragraphs[0].text_id.as_deref(), Some("00000003"));
}
