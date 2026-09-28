use std::collections::HashMap;
use std::sync::Arc;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::types::xml::XmlIn;
use yrs::updates::decoder::Decode;
use yrs::{
    Any, Doc, GetString, Out, ReadTxn, Transact, Update, Xml, XmlElementPrelim, XmlFragment,
    XmlOut, XmlTextPrelim,
};

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
                occurrence: None,
                paragraph_id: None,
                text_id: None,
                index_utf16: None,
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

#[test]
fn document_table_row_frame_rewrites_ancestor_identity_or_writes_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let store =
        NativeOfficeCollaborationStore::create(create_request(&temp.path().join("table-row")))
            .unwrap();
    store
        .apply(apply_request(
            "bootstrap-complex-browser-document",
            STANDARD.decode(YJS_COMPLEX_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();
    let stale = store
        .mutate(mutation_request(
            "insert-row-stale",
            NativeOfficeCollaborationMutation::DocumentInsertTableRow {
                anchor_row_id: "00000301".to_owned(),
                expected_row_text_id: "00000399".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                row_id: "00000330".to_owned(),
                row_text_id: "00000331".to_owned(),
                paragraph_id: "00000332".to_owned(),
                text_id: "00000333".to_owned(),
                text: "Added row".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );
    assert_eq!(document_state(&store), expected_initial_state());

    store
        .mutate(mutation_request(
            "insert-row",
            NativeOfficeCollaborationMutation::DocumentInsertTableRow {
                anchor_row_id: "00000301".to_owned(),
                expected_row_text_id: "00000302".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                row_id: "00000330".to_owned(),
                row_text_id: "00000331".to_owned(),
                paragraph_id: "00000332".to_owned(),
                text_id: "00000333".to_owned(),
                text: "Added row".to_owned(),
            },
        ))
        .unwrap();
    let inserted = document_state(&store);
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000101" && paragraph.text == "List anchor"
    }));
    assert!(inserted
        .rows
        .iter()
        .any(|row| row.row_id == "00000301" && row.row_text_id == "00000302"));
    assert!(inserted
        .rows
        .iter()
        .any(|row| row.row_id == "00000201" && row.row_text_id == "00000203"));
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000332"
            && paragraph.text_id == "00000333"
            && paragraph.parent_tag == "tableCell"
            && paragraph.text == "Added row"
    }));

    let wide_cells =
        NativeOfficeCollaborationStore::create(create_request(&temp.path().join("wide-cells")))
            .unwrap();
    wide_cells
        .apply(apply_request(
            "bootstrap-wide-cells",
            store.synchronize(None).unwrap().update,
        ))
        .unwrap();
    let peer = Doc::with_client_id(424_251);
    let exported = wide_cells.synchronize(None).unwrap().update;
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported).unwrap())
        .unwrap();
    let peer_before = peer.transact().state_vector();
    {
        let fragment = peer.get_or_insert_xml_fragment("a3s.office.document.content");
        let mut transaction = peer.transact_mut();
        let row = fragment
            .successors(&transaction)
            .find_map(|node| match node {
                XmlOut::Element(element)
                    if element.tag().as_ref() == "tableRow"
                        && string_attribute(&element, &transaction, "rowId") == "00000330" =>
                {
                    Some(element)
                }
                _ => None,
            })
            .unwrap();
        row.push_back(&mut transaction, XmlElementPrelim::empty("tableCell"));
    }
    wide_cells
        .apply(apply_request(
            "add-second-cell",
            peer.transact().encode_state_as_update_v1(&peer_before),
        ))
        .unwrap();
    let before_cells = wide_cells.inspect().unwrap();
    let extra_cell = wide_cells
        .mutate(mutation_request(
            "delete-two-cell-row",
            NativeOfficeCollaborationMutation::DocumentDeleteTableRow {
                row_id: "00000330".to_owned(),
                expected_row_text_id: "00000331".to_owned(),
                paragraph_id: "00000332".to_owned(),
                expected_text_id: "00000333".to_owned(),
                expected_text: "Added row".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        extra_cell.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        wide_cells.inspect().unwrap().document_state_sha256,
        before_cells.document_state_sha256
    );

    let before_wide = store.inspect().unwrap();
    let wide = store
        .mutate(mutation_request(
            "delete-wide-row",
            NativeOfficeCollaborationMutation::DocumentDeleteTableRow {
                row_id: "00000301".to_owned(),
                expected_row_text_id: "00000302".to_owned(),
                paragraph_id: "00000311".to_owned(),
                expected_text_id: "00000312".to_owned(),
                expected_text: "Nested target".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        wide.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_wide.document_state_sha256
    );

    let before_last = store.inspect().unwrap();
    let last = store
        .mutate(mutation_request(
            "delete-last-outer-row",
            NativeOfficeCollaborationMutation::DocumentDeleteTableRow {
                row_id: "00000201".to_owned(),
                expected_row_text_id: "00000203".to_owned(),
                paragraph_id: "00000211".to_owned(),
                expected_text_id: "00000212".to_owned(),
                expected_text: "Outer cell".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        last.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_last.document_state_sha256
    );

    store
        .mutate(mutation_request(
            "delete-added-row",
            NativeOfficeCollaborationMutation::DocumentDeleteTableRow {
                row_id: "00000330".to_owned(),
                expected_row_text_id: "00000331".to_owned(),
                paragraph_id: "00000332".to_owned(),
                expected_text_id: "00000333".to_owned(),
                expected_text: "Added row".to_owned(),
            },
        ))
        .unwrap();
    let deleted = document_state(&store);
    assert!(deleted
        .paragraphs
        .iter()
        .all(|paragraph| paragraph.paragraph_id != "00000332"));
    assert!(deleted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000101" && paragraph.text == "List anchor"
    }));
    assert!(deleted
        .rows
        .iter()
        .any(|row| row.row_id == "00000301" && row.row_text_id == "00000302"));
    assert!(deleted
        .rows
        .iter()
        .any(|row| row.row_id == "00000201" && row.row_text_id == "00000204"));
    assert_eq!(deleted.rows.len(), 2);
}

#[test]
fn document_list_item_frame_rewrites_ancestor_identity_or_writes_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let store =
        NativeOfficeCollaborationStore::create(create_request(&temp.path().join("list-items")))
            .unwrap();
    store
        .apply(apply_request(
            "bootstrap-list-items",
            STANDARD.decode(YJS_COMPLEX_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    let stale = store
        .mutate(mutation_request(
            "insert-stale-list-item",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000199".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000120".to_owned(),
                text_id: "00000121".to_owned(),
                text: "Added item".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );

    let not_leading = store
        .mutate(mutation_request(
            "insert-from-list-tail",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000110".to_owned(),
                expected_text_id: "00000111".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000120".to_owned(),
                text_id: "00000121".to_owned(),
                text: "Added item".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        not_leading.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );

    let cell = store
        .mutate(mutation_request(
            "insert-from-table-cell",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000211".to_owned(),
                expected_text_id: "00000212".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000120".to_owned(),
                text_id: "00000121".to_owned(),
                text: "Added item".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        cell.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );

    store
        .mutate(mutation_request(
            "insert-list-item",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000120".to_owned(),
                text_id: "00000121".to_owned(),
                text: "Added item".to_owned(),
            },
        ))
        .unwrap();
    let inserted = document_state(&store);
    assert_eq!(
        list_item_paragraphs(&store),
        vec![
            vec!["00000101".to_owned(), "00000110".to_owned()],
            vec!["00000120".to_owned()],
        ]
    );
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000120"
            && paragraph.text_id == "00000121"
            && paragraph.parent_tag == "listItem"
            && paragraph.text == "Added item"
    }));
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000211" && paragraph.text == "Outer cell"
    }));
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000311" && paragraph.text == "Nested target"
    }));
    assert!(inserted
        .rows
        .iter()
        .any(|row| row.row_id == "00000201" && row.row_text_id == "00000202"));

    let before_wide = store.inspect().unwrap();
    let wide = store
        .mutate(mutation_request(
            "delete-two-paragraph-item",
            NativeOfficeCollaborationMutation::DocumentDeleteListItem {
                paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                expected_text: "List anchor".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        wide.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_wide.document_state_sha256
    );

    store
        .mutate(mutation_request(
            "delete-added-item",
            NativeOfficeCollaborationMutation::DocumentDeleteListItem {
                paragraph_id: "00000120".to_owned(),
                expected_text_id: "00000121".to_owned(),
                expected_text: "Added item".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(
        list_item_paragraphs(&store),
        vec![vec!["00000101".to_owned(), "00000110".to_owned()]]
    );

    let before_last = store.inspect().unwrap();
    let last = store
        .mutate(mutation_request(
            "delete-last-item",
            NativeOfficeCollaborationMutation::DocumentDeleteListItem {
                paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                expected_text: "List anchor".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        last.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_last.document_state_sha256
    );
    assert!(document_state(&store).paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000320" && paragraph.text == "Nested tail"
    }));

    store
        .mutate(mutation_request(
            "insert-list-item-before",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::Before,
                paragraph_id: "00000130".to_owned(),
                text_id: "00000131".to_owned(),
                text: "Leading item".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(
        list_item_paragraphs(&store),
        vec![
            vec!["00000130".to_owned()],
            vec!["00000101".to_owned(), "00000110".to_owned()],
        ]
    );

    let nested =
        NativeOfficeCollaborationStore::create(create_request(&temp.path().join("list-in-cell")))
            .unwrap();
    nested
        .apply(apply_request(
            "bootstrap-list-in-cell",
            STANDARD.decode(YJS_COMPLEX_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    seed_list_in_outer_cell(&nested);
    let before_nested = nested.inspect().unwrap();
    let drifted = nested
        .mutate(mutation_request(
            "insert-drifted-cell-list",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000401".to_owned(),
                expected_text_id: "00000499".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000410".to_owned(),
                text_id: "00000411".to_owned(),
                text: "Cell item".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(drifted.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        nested.inspect().unwrap().document_state_sha256,
        before_nested.document_state_sha256
    );
    nested
        .mutate(mutation_request(
            "insert-cell-list-item",
            NativeOfficeCollaborationMutation::DocumentInsertListItem {
                anchor_paragraph_id: "00000401".to_owned(),
                expected_text_id: "00000402".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000410".to_owned(),
                text_id: "00000411".to_owned(),
                text: "Cell item".to_owned(),
            },
        ))
        .unwrap();
    let rotated = document_state(&nested);
    assert!(rotated
        .rows
        .iter()
        .any(|row| row.row_id == "00000201" && row.row_text_id == "00000203"));
    assert!(rotated
        .rows
        .iter()
        .any(|row| row.row_id == "00000301" && row.row_text_id == "00000302"));
    assert!(rotated.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000211" && paragraph.text == "Outer cell"
    }));
    assert!(rotated.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000410" && paragraph.text == "Cell item"
    }));
}

#[test]
fn document_section_frame_keeps_the_observed_section_or_writes_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let store =
        NativeOfficeCollaborationStore::create(create_request(&temp.path().join("sections")))
            .unwrap();
    store
        .apply(apply_request(
            "bootstrap-sections",
            STANDARD.decode(YJS_COMPLEX_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    let missing = store
        .mutate(mutation_request(
            "insert-missing-section",
            NativeOfficeCollaborationMutation::DocumentInsertSection {
                anchor_section_id: "missing-section".to_owned(),
                expected_paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                section_id: "document-section-added".to_owned(),
                paragraph_id: "00000501".to_owned(),
                text_id: "00000502".to_owned(),
                text: "Added section".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(missing.code, "office.collaboration.mutation_target_missing");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );

    let stale = store
        .mutate(mutation_request(
            "insert-stale-section-paragraph",
            NativeOfficeCollaborationMutation::DocumentInsertSection {
                anchor_section_id: "document-section-complex".to_owned(),
                expected_paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000199".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                section_id: "document-section-added".to_owned(),
                paragraph_id: "00000501".to_owned(),
                text_id: "00000502".to_owned(),
                text: "Added section".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before.document_state_sha256
    );

    let before_last = store.inspect().unwrap();
    let last = store
        .mutate(mutation_request(
            "delete-only-section",
            NativeOfficeCollaborationMutation::DocumentDeleteSection {
                section_id: "document-section-complex".to_owned(),
                paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                expected_text: "List anchor".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        last.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_last.document_state_sha256
    );

    store
        .mutate(mutation_request(
            "insert-section-after",
            NativeOfficeCollaborationMutation::DocumentInsertSection {
                anchor_section_id: "document-section-complex".to_owned(),
                expected_paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                section_id: "document-section-added".to_owned(),
                paragraph_id: "00000501".to_owned(),
                text_id: "00000502".to_owned(),
                text: "Added section".to_owned(),
            },
        ))
        .unwrap();
    let inserted = document_state(&store);
    assert_eq!(
        section_ids(&store),
        vec![
            "document-section-complex".to_owned(),
            "document-section-added".to_owned(),
        ]
    );
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000501"
            && paragraph.text_id == "00000502"
            && paragraph.parent_tag == "documentSection"
            && paragraph.text == "Added section"
    }));
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000101" && paragraph.text == "List anchor"
    }));
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000110" && paragraph.text == "List tail"
    }));
    assert!(inserted.paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000211" && paragraph.text == "Outer cell"
    }));
    assert!(inserted
        .rows
        .iter()
        .any(|row| row.row_id == "00000201" && row.row_text_id == "00000202"));
    assert!(inserted
        .rows
        .iter()
        .any(|row| row.row_id == "00000301" && row.row_text_id == "00000302"));

    let before_wide = store.inspect().unwrap();
    let wide = store
        .mutate(mutation_request(
            "delete-original-section",
            NativeOfficeCollaborationMutation::DocumentDeleteSection {
                section_id: "document-section-complex".to_owned(),
                paragraph_id: "00000101".to_owned(),
                expected_text_id: "00000102".to_owned(),
                expected_text: "List anchor".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        wide.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_wide.document_state_sha256
    );

    let before_stale_delete = store.inspect().unwrap();
    let stale_delete = store
        .mutate(mutation_request(
            "delete-stale-added-section",
            NativeOfficeCollaborationMutation::DocumentDeleteSection {
                section_id: "document-section-added".to_owned(),
                paragraph_id: "00000501".to_owned(),
                expected_text_id: "00000599".to_owned(),
                expected_text: "Added section".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        stale_delete.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale_delete.document_state_sha256
    );

    store
        .mutate(mutation_request(
            "delete-added-section",
            NativeOfficeCollaborationMutation::DocumentDeleteSection {
                section_id: "document-section-added".to_owned(),
                paragraph_id: "00000501".to_owned(),
                expected_text_id: "00000502".to_owned(),
                expected_text: "Added section".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(
        section_ids(&store),
        vec!["document-section-complex".to_owned()]
    );
    assert!(document_state(&store).paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000320" && paragraph.text == "Nested tail"
    }));

    store
        .mutate(mutation_request(
            "insert-section-before",
            NativeOfficeCollaborationMutation::DocumentInsertSection {
                anchor_section_id: "document-section-complex".to_owned(),
                expected_paragraph_id: "00000211".to_owned(),
                expected_text_id: "00000212".to_owned(),
                position: NativeOfficeCollaborationParagraphPosition::Before,
                section_id: "document-section-leading".to_owned(),
                paragraph_id: "00000510".to_owned(),
                text_id: "00000511".to_owned(),
                text: "Leading section".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(
        section_ids(&store),
        vec![
            "document-section-leading".to_owned(),
            "document-section-complex".to_owned(),
        ]
    );
    assert!(document_state(&store).paragraphs.iter().any(|paragraph| {
        paragraph.paragraph_id == "00000110" && paragraph.text == "List tail"
    }));
}

fn section_ids(store: &NativeOfficeCollaborationStore) -> Vec<String> {
    let doc = Doc::with_client_id(900_003);
    let exported = store.synchronize(None).unwrap().update;
    doc.transact_mut()
        .apply_update(Update::decode_v1(&exported).unwrap())
        .unwrap();
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = doc.transact();
    fragment
        .children(&transaction)
        .map(|node| match node {
            XmlOut::Element(element) if element.tag().as_ref() == "documentSection" => {
                string_attribute(&element, &transaction, "id")
            }
            other => panic!("unexpected top-level node: {other:?}"),
        })
        .collect()
}

fn list_item_paragraphs(store: &NativeOfficeCollaborationStore) -> Vec<Vec<String>> {
    let doc = Doc::with_client_id(900_003);
    let exported = store.synchronize(None).unwrap().update;
    doc.transact_mut()
        .apply_update(Update::decode_v1(&exported).unwrap())
        .unwrap();
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = doc.transact();
    let mut items = Vec::new();
    for node in fragment.successors(&transaction) {
        let XmlOut::Element(element) = node else {
            continue;
        };
        if element.tag().as_ref() != "listItem" {
            continue;
        }
        let mut paragraphs = Vec::new();
        for child in element.successors(&transaction) {
            let XmlOut::Element(child) = child else {
                continue;
            };
            if child.tag().as_ref() != "paragraph" {
                continue;
            }
            if let Some(id) = string_attribute_option(&child, &transaction, "paragraphId") {
                paragraphs.push(id);
            }
        }
        items.push(paragraphs);
    }
    items
}

fn seed_list_in_outer_cell(store: &NativeOfficeCollaborationStore) {
    let peer = Doc::with_client_id(424_252);
    let exported = store.synchronize(None).unwrap().update;
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported).unwrap())
        .unwrap();
    let peer_before = peer.transact().state_vector();
    {
        let fragment = peer.get_or_insert_xml_fragment("a3s.office.document.content");
        let mut transaction = peer.transact_mut();
        let paragraph = fragment
            .successors(&transaction)
            .find_map(|node| match node {
                XmlOut::Element(element)
                    if element.tag().as_ref() == "paragraph"
                        && string_attribute(&element, &transaction, "paragraphId")
                            == "00000211" =>
                {
                    Some(element)
                }
                _ => None,
            })
            .unwrap();
        let XmlOut::Element(cell) = paragraph.parent().unwrap() else {
            panic!("outer paragraph parent");
        };
        let paragraph = XmlElementPrelim {
            tag: Arc::from("paragraph"),
            attributes: HashMap::from([
                (Arc::from("paragraphId"), "00000401".to_owned()),
                (Arc::from("textId"), "00000402".to_owned()),
            ]),
            children: vec![XmlIn::from(XmlTextPrelim::new("Cell list"))],
        };
        let item = XmlElementPrelim {
            tag: Arc::from("listItem"),
            attributes: HashMap::new(),
            children: vec![XmlIn::from(paragraph)],
        };
        cell.push_back(
            &mut transaction,
            XmlElementPrelim {
                tag: Arc::from("bulletList"),
                attributes: HashMap::new(),
                children: vec![XmlIn::from(item)],
            },
        );
    }
    store
        .apply(apply_request(
            "seed-cell-list",
            peer.transact().encode_state_as_update_v1(&peer_before),
        ))
        .unwrap();
}

fn string_attribute_option(
    element: &yrs::XmlElementRef,
    transaction: &impl ReadTxn,
    name: &str,
) -> Option<String> {
    match element.get_attribute(transaction, name) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        _ => None,
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
