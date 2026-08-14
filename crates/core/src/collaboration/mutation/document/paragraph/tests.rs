use yrs::{Transact, Xml, XmlFragment};

use super::*;
use crate::collaboration::document::new_replica_document;
use crate::collaboration::{
    NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationMode,
};

#[test]
fn paragraph_contract_requires_canonical_bounded_word_identities() {
    for value in ["", "00000000", "80000000", "0000000a", " 0000000A"] {
        let error = validate_insert_paragraph("00000001", value, "00000002", "value").unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_invalid");
    }
    validate_insert_paragraph("00000001", "0000000A", "7FFFFFFF", "value").unwrap();
}

#[test]
fn paragraph_insert_and_delete_preserve_section_schema_and_identity_guards() {
    let doc = document_with_paragraphs(&[
        ("00000001", "00000002", "First"),
        ("00000003", "00000004", "Last"),
    ]);
    let manifest = manifest();

    insert_paragraph(
        &doc,
        &manifest,
        "00000001",
        NativeOfficeCollaborationParagraphPosition::After,
        "0000000A",
        "0000000B",
        "Native",
    )
    .unwrap();

    assert_eq!(
        paragraphs(&doc),
        vec![
            (
                "00000001".to_owned(),
                "00000002".to_owned(),
                "First".to_owned()
            ),
            (
                "0000000A".to_owned(),
                "0000000B".to_owned(),
                "Native".to_owned()
            ),
            (
                "00000003".to_owned(),
                "00000004".to_owned(),
                "Last".to_owned()
            ),
        ]
    );
    let collision = insert_paragraph(
        &doc,
        &manifest,
        "00000001",
        NativeOfficeCollaborationParagraphPosition::Before,
        "0000000A",
        "0000000C",
        "Collision",
    )
    .unwrap_err();
    assert_eq!(
        collision.code,
        "office.collaboration.mutation_identity_conflict"
    );

    let stale = delete_paragraph(&doc, &manifest, "0000000A", "0000000C", "Native").unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    delete_paragraph(&doc, &manifest, "0000000A", "0000000B", "Native").unwrap();
    assert_eq!(paragraphs(&doc).len(), 2);
}

#[test]
fn paragraph_delete_keeps_the_required_last_section_block() {
    let doc = document_with_paragraphs(&[("00000001", "00000002", "Only")]);
    let error = delete_paragraph(&doc, &manifest(), "00000001", "00000002", "Only").unwrap_err();
    assert_eq!(
        error.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(paragraphs(&doc).len(), 1);
}

#[test]
fn paragraph_mutations_support_list_items_and_table_cells_with_row_identity_rotation() {
    let doc = document_with_list_and_table();
    let manifest = manifest();

    insert_paragraph(
        &doc,
        &manifest,
        "00000001",
        NativeOfficeCollaborationParagraphPosition::After,
        "00000005",
        "00000006",
        "List inserted",
    )
    .unwrap();
    assert_eq!(paragraph_parent_tag(&doc, "00000005"), "listItem");
    assert_eq!(table_row_text_id(&doc, "00000020"), "00000021");
    delete_paragraph(&doc, &manifest, "00000005", "00000006", "List inserted").unwrap();

    insert_paragraph(
        &doc,
        &manifest,
        "00000012",
        NativeOfficeCollaborationParagraphPosition::Before,
        "00000014",
        "00000015",
        "Cell inserted",
    )
    .unwrap();
    assert_eq!(paragraph_parent_tag(&doc, "00000014"), "tableCell");
    assert_eq!(table_row_text_id(&doc, "00000020"), "00000022");
    delete_paragraph(&doc, &manifest, "00000014", "00000015", "Cell inserted").unwrap();
    assert_eq!(table_row_text_id(&doc, "00000020"), "00000023");
}

#[test]
fn paragraph_insert_supports_table_headers_and_blockquotes() {
    let doc = document_with_header_and_blockquote();
    let manifest = manifest();

    insert_paragraph(
        &doc,
        &manifest,
        "00000001",
        NativeOfficeCollaborationParagraphPosition::After,
        "00000003",
        "00000004",
        "Header inserted",
    )
    .unwrap();
    assert_eq!(paragraph_parent_tag(&doc, "00000003"), "tableHeader");
    assert_eq!(table_row_text_id(&doc, "00000020"), "00000022");

    insert_paragraph(
        &doc,
        &manifest,
        "00000010",
        NativeOfficeCollaborationParagraphPosition::After,
        "00000012",
        "00000013",
        "Quote inserted",
    )
    .unwrap();
    assert_eq!(paragraph_parent_tag(&doc, "00000012"), "blockquote");
    assert_eq!(table_row_text_id(&doc, "00000020"), "00000022");
}

#[test]
fn paragraph_delete_preserves_required_nested_container_blocks() {
    for container_tag in ["listItem", "tableCell", "tableHeader", "blockquote"] {
        let (doc, paragraph_id, text_id, text) =
            document_with_single_nested_paragraph(container_tag);
        let error = delete_paragraph(&doc, &manifest(), paragraph_id, text_id, text).unwrap_err();
        assert_eq!(
            error.code,
            "office.collaboration.mutation_structure_conflict"
        );
        assert_eq!(paragraphs(&doc).len(), 1);
    }

    let doc = document_with_list_head_followed_only_by_nested_list();
    let error =
        delete_paragraph(&doc, &manifest(), "00000001", "00000002", "Required head").unwrap_err();
    assert_eq!(
        error.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(paragraph_parent_tag(&doc, "00000001"), "listItem");
}

#[test]
fn paragraph_insert_rejects_partial_ancestor_row_identity_before_writing() {
    let (doc, _, _, _) = document_with_single_nested_paragraph("tableCell");
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let row = {
        let transaction = doc.transact();
        fragment
            .successors(&transaction)
            .find_map(|node| match node {
                XmlOut::Element(element) if element.tag().as_ref() == "tableRow" => Some(element),
                _ => None,
            })
            .expect("table row")
    };
    row.remove_attribute(&mut doc.transact_mut(), &"rowTextId");

    let error = insert_paragraph(
        &doc,
        &manifest(),
        "00000001",
        NativeOfficeCollaborationParagraphPosition::After,
        "00000003",
        "00000004",
        "Rejected",
    )
    .unwrap_err();
    assert_eq!(error.code, "office.collaboration.content_invalid");
    assert_eq!(paragraphs(&doc).len(), 1);
}

#[test]
fn paragraph_mutations_reject_malformed_container_ancestry() {
    let doc = new_replica_document(
        7,
        "a3s.office",
        NativeOfficeCollaborationArtifactKind::Document,
    );
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let mut transaction = doc.transact_mut();
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    let malformed = section.push_back(&mut transaction, XmlElementPrelim::empty("listItem"));
    let paragraph = malformed.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
    paragraph.insert_attribute(&mut transaction, PARAGRAPH_ID_ATTRIBUTE, "00000001");
    paragraph.insert_attribute(&mut transaction, TEXT_ID_ATTRIBUTE, "00000002");
    paragraph.push_back(&mut transaction, XmlTextPrelim::new("Malformed"));
    drop(transaction);

    let error = insert_paragraph(
        &doc,
        &manifest(),
        "00000001",
        NativeOfficeCollaborationParagraphPosition::After,
        "00000003",
        "00000004",
        "Rejected",
    )
    .unwrap_err();
    assert_eq!(
        error.code,
        "office.collaboration.mutation_structure_conflict"
    );
    assert_eq!(paragraphs(&doc).len(), 1);
}

fn document_with_paragraphs(values: &[(&str, &str, &str)]) -> yrs::Doc {
    let doc = new_replica_document(
        7,
        "a3s.office",
        NativeOfficeCollaborationArtifactKind::Document,
    );
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let mut transaction = doc.transact_mut();
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    section.insert_attribute(&mut transaction, "id", "document-section-1");
    for (paragraph_id, text_id, text) in values {
        let paragraph = section.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
        paragraph.insert_attribute(&mut transaction, PARAGRAPH_ID_ATTRIBUTE, *paragraph_id);
        paragraph.insert_attribute(&mut transaction, TEXT_ID_ATTRIBUTE, *text_id);
        if !text.is_empty() {
            paragraph.push_back(&mut transaction, XmlTextPrelim::new(*text));
        }
    }
    drop(transaction);
    doc
}

fn document_with_list_and_table() -> yrs::Doc {
    let doc = new_replica_document(
        7,
        "a3s.office",
        NativeOfficeCollaborationArtifactKind::Document,
    );
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let mut transaction = doc.transact_mut();
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    let list = section.push_back(&mut transaction, XmlElementPrelim::empty("bulletList"));
    let item = list.push_back(&mut transaction, XmlElementPrelim::empty("listItem"));
    push_identity_paragraph(
        &item,
        &mut transaction,
        "00000001",
        "00000002",
        "List anchor",
    );
    push_identity_paragraph(&item, &mut transaction, "00000003", "00000004", "List tail");

    let table = section.push_back(&mut transaction, XmlElementPrelim::empty("table"));
    let row = table.push_back(&mut transaction, XmlElementPrelim::empty("tableRow"));
    row.insert_attribute(&mut transaction, "rowId", "00000020");
    row.insert_attribute(&mut transaction, "rowTextId", "00000021");
    let cell = row.push_back(&mut transaction, XmlElementPrelim::empty("tableCell"));
    push_identity_paragraph(
        &cell,
        &mut transaction,
        "00000010",
        "00000011",
        "Cell anchor",
    );
    push_identity_paragraph(&cell, &mut transaction, "00000012", "00000013", "Cell tail");
    drop(transaction);
    doc
}

fn document_with_header_and_blockquote() -> yrs::Doc {
    let doc = new_replica_document(
        7,
        "a3s.office",
        NativeOfficeCollaborationArtifactKind::Document,
    );
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let mut transaction = doc.transact_mut();
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    let table = section.push_back(&mut transaction, XmlElementPrelim::empty("table"));
    let row = table.push_back(&mut transaction, XmlElementPrelim::empty("tableRow"));
    row.insert_attribute(&mut transaction, "rowId", "00000020");
    row.insert_attribute(&mut transaction, "rowTextId", "00000021");
    let header = row.push_back(&mut transaction, XmlElementPrelim::empty("tableHeader"));
    push_identity_paragraph(&header, &mut transaction, "00000001", "00000002", "Header");
    let blockquote = section.push_back(&mut transaction, XmlElementPrelim::empty("blockquote"));
    push_identity_paragraph(
        &blockquote,
        &mut transaction,
        "00000010",
        "00000011",
        "Quote",
    );
    drop(transaction);
    doc
}

fn document_with_single_nested_paragraph(
    container_tag: &str,
) -> (yrs::Doc, &'static str, &'static str, &'static str) {
    let doc = new_replica_document(
        7,
        "a3s.office",
        NativeOfficeCollaborationArtifactKind::Document,
    );
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let mut transaction = doc.transact_mut();
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    let container = match container_tag {
        "listItem" => {
            let list = section.push_back(&mut transaction, XmlElementPrelim::empty("orderedList"));
            list.push_back(&mut transaction, XmlElementPrelim::empty("listItem"))
        }
        "tableCell" | "tableHeader" => {
            let table = section.push_back(&mut transaction, XmlElementPrelim::empty("table"));
            let row = table.push_back(&mut transaction, XmlElementPrelim::empty("tableRow"));
            row.insert_attribute(&mut transaction, "rowId", "00000020");
            row.insert_attribute(&mut transaction, "rowTextId", "00000021");
            row.push_back(&mut transaction, XmlElementPrelim::empty(container_tag))
        }
        "blockquote" => section.push_back(&mut transaction, XmlElementPrelim::empty("blockquote")),
        _ => panic!("unsupported test container: {container_tag}"),
    };
    push_identity_paragraph(&container, &mut transaction, "00000001", "00000002", "Only");
    drop(transaction);
    (doc, "00000001", "00000002", "Only")
}

fn document_with_list_head_followed_only_by_nested_list() -> yrs::Doc {
    let doc = new_replica_document(
        7,
        "a3s.office",
        NativeOfficeCollaborationArtifactKind::Document,
    );
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let mut transaction = doc.transact_mut();
    let section = fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    let list = section.push_back(&mut transaction, XmlElementPrelim::empty("bulletList"));
    let item = list.push_back(&mut transaction, XmlElementPrelim::empty("listItem"));
    push_identity_paragraph(
        &item,
        &mut transaction,
        "00000001",
        "00000002",
        "Required head",
    );
    let nested = item.push_back(&mut transaction, XmlElementPrelim::empty("bulletList"));
    let nested_item = nested.push_back(&mut transaction, XmlElementPrelim::empty("listItem"));
    push_identity_paragraph(
        &nested_item,
        &mut transaction,
        "00000003",
        "00000004",
        "Nested",
    );
    drop(transaction);
    doc
}

fn push_identity_paragraph(
    container: &XmlElementRef,
    transaction: &mut yrs::TransactionMut<'_>,
    paragraph_id: &str,
    text_id: &str,
    text: &str,
) {
    let paragraph = container.push_back(transaction, XmlElementPrelim::empty("paragraph"));
    paragraph.insert_attribute(transaction, PARAGRAPH_ID_ATTRIBUTE, paragraph_id);
    paragraph.insert_attribute(transaction, TEXT_ID_ATTRIBUTE, text_id);
    paragraph.push_back(transaction, XmlTextPrelim::new(text));
}

fn paragraph_parent_tag(doc: &yrs::Doc, paragraph_id: &str) -> String {
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = doc.transact();
    fragment
        .successors(&transaction)
        .find_map(|node| match node {
            XmlOut::Element(element)
                if document_identity_attribute(&element, &transaction, PARAGRAPH_ID_ATTRIBUTE)
                    .ok()
                    .flatten()
                    .as_deref()
                    == Some(paragraph_id) =>
            {
                match element.parent() {
                    Some(XmlOut::Element(parent)) => Some(parent.tag().to_string()),
                    _ => None,
                }
            }
            _ => None,
        })
        .expect("paragraph parent")
}

fn table_row_text_id(doc: &yrs::Doc, row_id: &str) -> String {
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = doc.transact();
    fragment
        .successors(&transaction)
        .find_map(|node| match node {
            XmlOut::Element(element)
                if element.tag().as_ref() == "tableRow"
                    && document_identity_attribute(&element, &transaction, "rowId")
                        .ok()
                        .flatten()
                        .as_deref()
                        == Some(row_id) =>
            {
                document_identity_attribute(&element, &transaction, "rowTextId")
                    .ok()
                    .flatten()
            }
            _ => None,
        })
        .expect("table row text identity")
}

fn paragraphs(doc: &yrs::Doc) -> Vec<(String, String, String)> {
    let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = doc.transact();
    fragment
        .successors(&transaction)
        .filter_map(|node| match node {
            XmlOut::Element(element) if element.tag().as_ref() == "paragraph" => {
                let paragraph_id =
                    document_identity_attribute(&element, &transaction, PARAGRAPH_ID_ATTRIBUTE)
                        .unwrap()?;
                let text_id =
                    document_identity_attribute(&element, &transaction, TEXT_ID_ATTRIBUTE)
                        .unwrap()?;
                Some((
                    paragraph_id,
                    text_id,
                    deletable_paragraph_text(&element, &transaction).unwrap(),
                ))
            }
            _ => None,
        })
        .collect()
}

fn manifest() -> NativeOfficeCollaborationManifest {
    NativeOfficeCollaborationManifest {
        format: String::new(),
        schema_version: 1,
        protocol: String::new(),
        protocol_version: 1,
        namespace: "a3s.office".to_owned(),
        artifact_id: "document".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Document,
        actor_id: "agent".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        client_id: 7,
    }
}
