use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, GetString, Out, ReadTxn, Transact, Update, Xml, XmlFragment, XmlOut};

use super::super::*;

const YJS_COMPLEX_DOCUMENT_UPDATE_BASE64: &str = "AS2z8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3GGZpeHR1cmUtY29tcGxleC1kb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGEEa2luZAF3CGRvY3VtZW50KAETYTNzLm9mZmljZS5tZXRhZGF0YQtpbml0aWFsaXplZAF4CAEhYTNzLm9mZmljZS5ib290c3RyYXAuaW5pdGlhbGl6ZXJzAXceNDI0MjQzOmJyb3dzZXItY29tcGxleC1maXh0dXJlBwEbYTNzLm9mZmljZS5kb2N1bWVudC5jb250ZW50Aw9kb2N1bWVudFNlY3Rpb24HALPyGQYDCmJ1bGxldExpc3QHALPyGQcDCGxpc3RJdGVtBwCz8hkIAwlwYXJhZ3JhcGgHALPyGQkGBACz8hkKC0xpc3QgYW5jaG9yKACz8hkJC3BhcmFncmFwaElkAXcIMDAwMDAxMDEoALPyGQkGdGV4dElkAXcIMDAwMDAxMDKHs/IZCQMJcGFyYWdyYXBoBwCz8hkYBgQAs/IZGQlMaXN0IHRhaWwoALPyGRgLcGFyYWdyYXBoSWQBdwgwMDAwMDExMCgAs/IZGAZ0ZXh0SWQBdwgwMDAwMDExMYez8hkHAwV0YWJsZQcAs/IZJQMIdGFibGVSb3cHALPyGSYDCXRhYmxlQ2VsbAcAs/IZJwMJcGFyYWdyYXBoBwCz8hkoBgQAs/IZKQpPdXRlciBjZWxsKACz8hkoC3BhcmFncmFwaElkAXcIMDAwMDAyMTEoALPyGSgGdGV4dElkAXcIMDAwMDAyMTKHs/IZKAMFdGFibGUHALPyGTYDCHRhYmxlUm93BwCz8hk3Awl0YWJsZUNlbGwHALPyGTgDCXBhcmFncmFwaAcAs/IZOQYEALPyGToNTmVzdGVkIHRhcmdldCgAs/IZOQtwYXJhZ3JhcGhJZAF3CDAwMDAwMzExKACz8hk5BnRleHRJZAF3CDAwMDAwMzEyh7PyGTkDCXBhcmFncmFwaAcAs/IZSgYEALPyGUsLTmVzdGVkIHRhaWwoALPyGUoLcGFyYWdyYXBoSWQBdwgwMDAwMDMyMCgAs/IZSgZ0ZXh0SWQBdwgwMDAwMDMyMSgAs/IZNwVyb3dJZAF3CDAwMDAwMzAxKACz8hk3CXJvd1RleHRJZAF3CDAwMDAwMzAyKACz8hkmBXJvd0lkAXcIMDAwMDAyMDEoALPyGSYJcm93VGV4dElkAXcIMDAwMDAyMDIoALPyGQYCaWQBdxhkb2N1bWVudC1zZWN0aW9uLWNvbXBsZXgA";

#[derive(Debug, PartialEq, Eq)]
struct ComplexDocumentParagraph {
    paragraph_id: String,
    text_id: String,
    parent_tag: String,
    text: String,
}

#[derive(Debug, PartialEq, Eq)]
struct ComplexDocumentRow {
    row_id: String,
    row_text_id: String,
}

#[derive(Debug, PartialEq, Eq)]
struct ComplexDocumentState {
    paragraphs: Vec<ComplexDocumentParagraph>,
    rows: Vec<ComplexDocumentRow>,
}

#[test]
fn typed_nested_document_mutations_converge_and_rotate_all_word_identities() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("complex-document-replica");
    let store = NativeOfficeCollaborationStore::create(create_request(&root)).unwrap();
    store
        .apply(apply_request(
            "bootstrap-complex-browser-document",
            STANDARD.decode(YJS_COMPLEX_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    assert_eq!(document_state(&store), expected_initial_state(),);

    let replace_before = store.inspect().unwrap().state_vector;
    let replaced = store
        .mutate(mutation_request(
            "replace-nested-text",
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: "Nested target".to_owned(),
                replacement: "Native nested".to_owned(),
                expected_matches: 1,
            },
        ))
        .unwrap();
    assert_eq!(replaced.sequence, Some(2));
    let replace_update = store.synchronize(Some(&replace_before)).unwrap().update;

    let list_before = store.inspect().unwrap().state_vector;
    let list_inserted = store
        .mutate(mutation_request(
            "insert-list-paragraph",
            NativeOfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id: "00000101".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000105".to_owned(),
                text_id: "00000106".to_owned(),
                text: "Native list".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(list_inserted.sequence, Some(3));
    let list_update = store.synchronize(Some(&list_before)).unwrap().update;

    let cell_before = store.inspect().unwrap().state_vector;
    let cell_inserted = store
        .mutate(mutation_request(
            "insert-nested-cell-paragraph",
            NativeOfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id: "00000311".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000315".to_owned(),
                text_id: "00000316".to_owned(),
                text: "Native cell".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(cell_inserted.sequence, Some(4));
    let cell_update = store.synchronize(Some(&cell_before)).unwrap().update;

    let delete_before = store.inspect().unwrap().state_vector;
    let deleted = store
        .mutate(mutation_request(
            "delete-nested-cell-tail",
            NativeOfficeCollaborationMutation::DocumentDeleteParagraph {
                paragraph_id: "00000320".to_owned(),
                expected_text_id: "00000321".to_owned(),
                expected_text: "Nested tail".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(deleted.sequence, Some(5));
    let delete_update = store.synchronize(Some(&delete_before)).unwrap().update;

    assert!(!replace_update.is_empty());
    assert!(!list_update.is_empty());
    assert!(!cell_update.is_empty());
    assert!(!delete_update.is_empty());

    let expected = expected_final_state();
    assert_eq!(document_state(&store), expected);
    let before_conflict = store.inspect().unwrap();
    let conflict = store
        .mutate(mutation_request(
            "delete-nested-cell-stale",
            NativeOfficeCollaborationMutation::DocumentDeleteParagraph {
                paragraph_id: "00000315".to_owned(),
                expected_text_id: "00000317".to_owned(),
                expected_text: "Native cell".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    let after_conflict = store.inspect().unwrap();
    assert_eq!(
        after_conflict.document_state_sha256,
        before_conflict.document_state_sha256
    );
    assert_eq!(
        after_conflict.current_sequence,
        before_conflict.current_sequence
    );

    drop(store);
    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(document_state(&reopened), expected);
    let events = reopened
        .events(NativeOfficeCollaborationEventsRequest {
            after_sequence: Some(1),
            limit: 4,
        })
        .unwrap();
    assert_eq!(events.updates.len(), 4);
    assert!(events
        .updates
        .iter()
        .all(|event| event.operation_kind == NativeOfficeCollaborationOperationKind::Mutate));
}

fn create_request(root: &std::path::Path) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-complex-document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "agent-complex".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-complex-document".to_owned(),
        namespace: None,
        client_id: Some(900_003),
        initial_update: None,
    }
}

fn apply_request(operation_id: &str, update: Vec<u8>) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-complex".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-complex-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-complex".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-complex-document".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Document,
        mutation,
        if_state_vector: None,
    }
}

fn document_state(store: &NativeOfficeCollaborationStore) -> ComplexDocumentState {
    let update = store.synchronize(None).unwrap().update;
    let peer = Doc::with_client_id(818_183);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&update).unwrap())
        .unwrap();
    let fragment = peer.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = peer.transact();
    let mut paragraphs = Vec::new();
    let mut rows = Vec::new();
    for node in fragment.successors(&transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        match element.tag().as_ref() {
            "paragraph" => {
                let parent_tag = match element.parent() {
                    Some(XmlOut::Element(parent)) => parent.tag().to_string(),
                    parent => panic!("unexpected paragraph parent: {parent:?}"),
                };
                paragraphs.push(ComplexDocumentParagraph {
                    paragraph_id: string_attribute(&element, &transaction, "paragraphId"),
                    text_id: string_attribute(&element, &transaction, "textId"),
                    parent_tag,
                    text: element
                        .children(&transaction)
                        .filter_map(|child| match child {
                            XmlOut::Text(text) => Some(text.get_string(&transaction)),
                            _ => None,
                        })
                        .collect(),
                });
            }
            "tableRow" => rows.push(ComplexDocumentRow {
                row_id: string_attribute(&element, &transaction, "rowId"),
                row_text_id: string_attribute(&element, &transaction, "rowTextId"),
            }),
            _ => {}
        }
    }
    ComplexDocumentState { paragraphs, rows }
}

fn string_attribute<T: ReadTxn>(element: &impl Xml, transaction: &T, name: &str) -> String {
    match element.get_attribute(transaction, name) {
        Some(Out::Any(Any::String(value))) => value.to_string(),
        value => panic!("unexpected XML attribute '{name}': {value:?}"),
    }
}

fn expected_initial_state() -> ComplexDocumentState {
    ComplexDocumentState {
        paragraphs: vec![
            paragraph("00000101", "00000102", "listItem", "List anchor"),
            paragraph("00000110", "00000111", "listItem", "List tail"),
            paragraph("00000211", "00000212", "tableCell", "Outer cell"),
            paragraph("00000311", "00000312", "tableCell", "Nested target"),
            paragraph("00000320", "00000321", "tableCell", "Nested tail"),
        ],
        rows: vec![row("00000201", "00000202"), row("00000301", "00000302")],
    }
}

fn expected_final_state() -> ComplexDocumentState {
    ComplexDocumentState {
        paragraphs: vec![
            paragraph("00000101", "00000102", "listItem", "List anchor"),
            paragraph("00000105", "00000106", "listItem", "Native list"),
            paragraph("00000110", "00000111", "listItem", "List tail"),
            paragraph("00000211", "00000212", "tableCell", "Outer cell"),
            paragraph("00000311", "00000313", "tableCell", "Native nested"),
            paragraph("00000315", "00000316", "tableCell", "Native cell"),
        ],
        rows: vec![row("00000201", "00000205"), row("00000301", "00000305")],
    }
}

fn paragraph(
    paragraph_id: &str,
    text_id: &str,
    parent_tag: &str,
    text: &str,
) -> ComplexDocumentParagraph {
    ComplexDocumentParagraph {
        paragraph_id: paragraph_id.to_owned(),
        text_id: text_id.to_owned(),
        parent_tag: parent_tag.to_owned(),
        text: text.to_owned(),
    }
}

fn row(row_id: &str, row_text_id: &str) -> ComplexDocumentRow {
    ComplexDocumentRow {
        row_id: row_id.to_owned(),
        row_text_id: row_text_id.to_owned(),
    }
}
