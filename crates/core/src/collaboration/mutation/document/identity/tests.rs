use yrs::{Doc, Transact, XmlElementPrelim, XmlFragment};

use super::*;

#[test]
fn table_row_ancestry_accepts_the_exact_depth_limit_and_rejects_one_more() {
    let (_at_limit_document, at_limit) =
        paragraph_with_element_ancestors(MAX_DOCUMENT_ANCESTOR_DEPTH);
    assert!(ancestor_table_rows(std::slice::from_ref(&at_limit))
        .unwrap()
        .is_empty());

    let (_beyond_limit_document, beyond_limit) =
        paragraph_with_element_ancestors(MAX_DOCUMENT_ANCESTOR_DEPTH + 1);
    let error = ancestor_table_rows(std::slice::from_ref(&beyond_limit)).unwrap_err();
    assert_eq!(error.code, "office.collaboration.content_invalid");
}

fn paragraph_with_element_ancestors(depth: usize) -> (Doc, yrs::XmlElementRef) {
    assert!(depth > 0);
    let document = Doc::new();
    let fragment = document.get_or_insert_xml_fragment("document");
    let mut transaction = document.transact_mut();
    let mut container =
        fragment.push_back(&mut transaction, XmlElementPrelim::empty("documentSection"));
    for _ in 1..depth {
        container = container.push_back(&mut transaction, XmlElementPrelim::empty("blockquote"));
    }
    let paragraph = container.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
    drop(transaction);
    (document, paragraph)
}
