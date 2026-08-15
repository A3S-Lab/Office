#![allow(dead_code)]

mod pdf;
mod presentation;
mod spreadsheet;

#[allow(unused_imports)]
pub use pdf::{pdf_collaboration_fixture, pdf_snapshot};
#[allow(unused_imports)]
pub use presentation::{
    presentation_collaboration_fixture, presentation_element, presentation_element_order,
    presentation_element_tombstoned, presentation_scene_element, presentation_slide_body_element,
    presentation_slide_title_element,
};
#[allow(unused_imports)]
pub use spreadsheet::{spreadsheet_cell, spreadsheet_collaboration_fixture};

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, GetString, Out, ReadTxn, Transact, Update, Xml, XmlFragment, XmlOut};

const COMPLEX_DOCUMENT_FIXTURE_BASE64: &str = "AS2z8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3GGZpeHR1cmUtY29tcGxleC1kb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGEEa2luZAF3CGRvY3VtZW50KAETYTNzLm9mZmljZS5tZXRhZGF0YQtpbml0aWFsaXplZAF4CAEhYTNzLm9mZmljZS5ib290c3RyYXAuaW5pdGlhbGl6ZXJzAXceNDI0MjQzOmJyb3dzZXItY29tcGxleC1maXh0dXJlBwEbYTNzLm9mZmljZS5kb2N1bWVudC5jb250ZW50Aw9kb2N1bWVudFNlY3Rpb24HALPyGQYDCmJ1bGxldExpc3QHALPyGQcDCGxpc3RJdGVtBwCz8hkIAwlwYXJhZ3JhcGgHALPyGQkGBACz8hkKC0xpc3QgYW5jaG9yKACz8hkJC3BhcmFncmFwaElkAXcIMDAwMDAxMDEoALPyGQkGdGV4dElkAXcIMDAwMDAxMDKHs/IZCQMJcGFyYWdyYXBoBwCz8hkYBgQAs/IZGQlMaXN0IHRhaWwoALPyGRgLcGFyYWdyYXBoSWQBdwgwMDAwMDExMCgAs/IZGAZ0ZXh0SWQBdwgwMDAwMDExMYez8hkHAwV0YWJsZQcAs/IZJQMIdGFibGVSb3cHALPyGSYDCXRhYmxlQ2VsbAcAs/IZJwMJcGFyYWdyYXBoBwCz8hkoBgQAs/IZKQpPdXRlciBjZWxsKACz8hkoC3BhcmFncmFwaElkAXcIMDAwMDAyMTEoALPyGSgGdGV4dElkAXcIMDAwMDAyMTKHs/IZKAMFdGFibGUHALPyGTYDCHRhYmxlUm93BwCz8hk3Awl0YWJsZUNlbGwHALPyGTgDCXBhcmFncmFwaAcAs/IZOQYEALPyGToNTmVzdGVkIHRhcmdldCgAs/IZOQtwYXJhZ3JhcGhJZAF3CDAwMDAwMzExKACz8hk5BnRleHRJZAF3CDAwMDAwMzEyh7PyGTkDCXBhcmFncmFwaAcAs/IZSgYEALPyGUsLTmVzdGVkIHRhaWwoALPyGUoLcGFyYWdyYXBoSWQBdwgwMDAwMDMyMCgAs/IZSgZ0ZXh0SWQBdwgwMDAwMDMyMSgAs/IZNwVyb3dJZAF3CDAwMDAwMzAxKACz8hk3CXJvd1RleHRJZAF3CDAwMDAwMzAyKACz8hkmBXJvd0lkAXcIMDAwMDAyMDEoALPyGSYJcm93VGV4dElkAXcIMDAwMDAyMDIoALPyGQYCaWQBdxhkb2N1bWVudC1zZWN0aW9uLWNvbXBsZXgA";

#[derive(Debug, PartialEq, Eq)]
pub struct DocumentParagraph {
    pub paragraph_id: String,
    pub text_id: String,
    pub parent_tag: String,
    pub text: String,
}

#[derive(Debug, PartialEq, Eq)]
pub struct DocumentRow {
    pub row_id: String,
    pub row_text_id: String,
}

#[derive(Debug, PartialEq, Eq)]
pub struct DocumentSnapshot {
    pub paragraphs: Vec<DocumentParagraph>,
    pub rows: Vec<DocumentRow>,
}

pub fn complex_document_fixture() -> Vec<u8> {
    STANDARD.decode(COMPLEX_DOCUMENT_FIXTURE_BASE64).unwrap()
}

pub fn document_snapshot(update: &[u8]) -> DocumentSnapshot {
    let document = Doc::with_client_id(710_011);
    document
        .transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    let fragment = document.get_or_insert_xml_fragment("a3s.office.document.content");
    let transaction = document.transact();
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
                paragraphs.push(DocumentParagraph {
                    paragraph_id: attribute(&element, &transaction, "paragraphId"),
                    text_id: attribute(&element, &transaction, "textId"),
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
            "tableRow" => rows.push(DocumentRow {
                row_id: attribute(&element, &transaction, "rowId"),
                row_text_id: attribute(&element, &transaction, "rowTextId"),
            }),
            _ => {}
        }
    }
    DocumentSnapshot { paragraphs, rows }
}

fn attribute<T: ReadTxn>(element: &impl Xml, transaction: &T, name: &str) -> String {
    match element.get_attribute(transaction, name) {
        Some(Out::Any(Any::String(value))) => value.to_string(),
        value => panic!("unexpected XML attribute '{name}': {value:?}"),
    }
}
