use super::*;
use yrs::{XmlElementPrelim, XmlTextPrelim};

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
            source: "AB".to_owned(),
            slices: vec![NativeOfficeCollaborationMarkdownSlice {
                index: 0,
                start_utf16: 0,
                end_utf16: 2,
                text: "AB".to_owned(),
            }],
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

#[test]
fn document_replica_export_writes_every_paragraph_or_returns_no_package() {
    let temp = tempfile::tempdir().unwrap();
    let store = NativeOfficeCollaborationStore::create(document_create_request(
        &temp.path().join("export-document"),
    ))
    .unwrap();
    store
        .apply(document_apply_request(
            "export-bootstrap-document",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
        .mutate(document_mutation_request(
            "export-insert-paragraph",
            NativeOfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id: "00000001".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000020".to_owned(),
                text_id: "00000021".to_owned(),
                text: "BETA".to_owned(),
            },
        ))
        .unwrap();

    let document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Hello</w:t></w:r></w:p><w:p w14:paraId="00000020"><w:r><w:t>OLD</w:t></w:r></w:p><w:p w14:paraId="00000099"><w:r><w:t>KEEP</w:t></w:r></w:p></w:body></w:document>"#;
    let mut parts = std::collections::BTreeMap::new();
    parts.insert("word/document.xml".to_owned(), document.to_vec());
    parts.insert("word/styles.xml".to_owned(), b"<styles/>".to_vec());

    let written = export_document_replica(&store, parts.clone()).unwrap();
    let xml = String::from_utf8(written["word/document.xml"].clone()).unwrap();
    assert!(xml.contains("Hello 😀 world"));
    assert!(xml.contains("BETA"));
    assert!(xml.contains("KEEP"));
    assert_eq!(written["word/styles.xml"], parts["word/styles.xml"]);

    let two_runs = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Hello</w:t></w:r></w:p><w:p w14:paraId="00000020"><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>"#;
    let mut split = parts.clone();
    split.insert("word/document.xml".to_owned(), two_runs.to_vec());
    assert!(export_document_replica(&store, split).is_err());

    let missing = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"#;
    let mut missing_parts = parts.clone();
    missing_parts.insert("word/document.xml".to_owned(), missing.to_vec());
    assert!(export_document_replica(&store, missing_parts).is_err());

    let browser = Doc::with_client_id(424_301);
    browser
        .transact_mut()
        .apply_update(Update::decode_v1(&store.synchronize(None).unwrap().update).unwrap())
        .unwrap();
    let before = browser.transact().state_vector();
    let fragment = browser.get_or_insert_xml_fragment("a3s.office.document.content");
    {
        let mut transaction = browser.transact_mut();
        let section = fragment
            .successors(&transaction)
            .find_map(|node| match node {
                XmlOut::Element(element) if element.tag().as_ref() == "documentSection" => {
                    Some(element)
                }
                _ => None,
            })
            .unwrap();
        let element = section.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
        element.push_back(&mut transaction, XmlTextPrelim::new("orphan"));
    }
    store
        .apply(document_apply_request(
            "export-paragraph-without-id",
            browser.transact().encode_state_as_update_v1(&before),
        ))
        .unwrap();
    let unnamed = export_document_replica(&store, parts).unwrap_err();
    assert_eq!(unnamed.code, "office.collaboration.snapshot_invalid");
    assert!(unnamed.message.contains("paraId"));
}

#[test]
fn reordered_document_updates_converge_and_reopen_the_merged_export() {
    let temp = tempfile::tempdir().unwrap();
    let updates = [
        offline_document_update(
            &temp,
            940_111,
            "offline-splice",
            NativeOfficeCollaborationMutation::DocumentSplice {
                paragraph_id: "00000001".to_owned(),
                text_id: "00000002".to_owned(),
                index_utf16: 6,
                delete_utf16: 2,
                expected_slice: "😀".to_owned(),
                insert: "🦀".to_owned(),
            },
        ),
        offline_document_update(
            &temp,
            940_112,
            "offline-page-color",
            NativeOfficeCollaborationMutation::DocumentSetPageColor {
                page_color: "#101828".to_owned(),
            },
        ),
        offline_document_update(
            &temp,
            940_113,
            "offline-track-changes",
            NativeOfficeCollaborationMutation::DocumentSetTrackChanges {
                track_changes: false,
            },
        ),
    ];

    let mut hashes = Vec::new();
    for (permutation_index, permutation) in three_document_permutations().into_iter().enumerate() {
        let ordered = [
            updates[permutation[0]].clone(),
            updates[permutation[1]].clone(),
            updates[permutation[2]].clone(),
        ];
        let store = deliver_document_updates(
            &temp.path().join(format!("order-{permutation_index}")),
            940_200 + permutation_index as u64,
            &ordered,
        );
        let state = document_state(&store);
        assert_eq!(
            state.paragraphs,
            vec![document_paragraph("00000001", "00000002", "Hello 🦀 world")]
        );
        assert_eq!(state.page_color.as_deref(), Some("#101828"));
        assert_eq!(state.track_changes, Some(false));
        hashes.push(store.inspect().unwrap().document_state_sha256);
    }
    assert!(hashes.iter().all(|hash| hash == &hashes[0]));

    let repeated = [
        updates[0].clone(),
        updates[0].clone(),
        updates[1].clone(),
        updates[2].clone(),
    ];
    let replayed = deliver_document_updates(&temp.path().join("repeated"), 940_220, &repeated);
    assert_eq!(replayed.inspect().unwrap().document_state_sha256, hashes[0]);

    let document = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:body><w:p w14:paraId=\"00000001\"><w:r><w:t>Hello 😀 world</w:t></w:r></w:p><w:p w14:paraId=\"00000099\"><w:r><w:t>KEEP</w:t></w:r></w:p></w:body></w:document>";
    let mut parts = std::collections::BTreeMap::new();
    parts.insert("word/document.xml".to_owned(), document.as_bytes().to_vec());
    parts.insert("word/styles.xml".to_owned(), b"<styles/>".to_vec());
    let written = export_document_replica(&replayed, parts.clone()).unwrap();
    assert_eq!(written["word/styles.xml"], parts["word/styles.xml"]);
    let xml = String::from_utf8(written["word/document.xml"].clone()).unwrap();
    assert!(xml.contains("KEEP"));

    let package = temp.path().join("merged.docx");
    std::fs::write(&package, b"snapshot").unwrap();
    let imported = import_document_snapshot(
        &package,
        &written,
        &temp.path().join("reopened"),
        "merged-letter",
    )
    .unwrap();
    let reopened = document_state(&imported.store);
    assert_eq!(
        reopened
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.text.as_str())
            .collect::<Vec<_>>(),
        vec!["Hello 🦀 world", "KEEP"]
    );
}

fn offline_document_update(
    temp: &tempfile::TempDir,
    client_id: u64,
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> Vec<u8> {
    let store = document_store_with_client(&temp.path().join(operation_id), client_id);
    let base = store.synchronize(None).unwrap().state_vector;
    store
        .mutate(document_mutation_request(operation_id, mutation))
        .unwrap();
    store.synchronize(Some(&base)).unwrap().update
}

fn deliver_document_updates(
    root: &std::path::Path,
    client_id: u64,
    updates: &[Vec<u8>],
) -> NativeOfficeCollaborationStore {
    let store = document_store_with_client(root, client_id);
    for (index, update) in updates.iter().enumerate() {
        store
            .apply(document_apply_request(
                &format!("deliver-{client_id}-{index}"),
                update.clone(),
            ))
            .unwrap();
    }
    store
}

fn document_store_with_client(
    root: &std::path::Path,
    client_id: u64,
) -> NativeOfficeCollaborationStore {
    let mut request = document_create_request(root);
    request.client_id = Some(client_id);
    request.operation_id = format!("create-document-{client_id}");
    let store = NativeOfficeCollaborationStore::create(request).unwrap();
    store
        .apply(document_apply_request(
            "bootstrap-browser-document",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
}

fn three_document_permutations() -> Vec<[usize; 3]> {
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
