use std::path::Path;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{json, Value as JsonValue};

use super::pdf::{initialized_pdf_store, pdf_form_mutation};
use super::presentation::{initialized_presentation_store, presentation_mutation_request};
use super::spreadsheet::{initialized_spreadsheet_store, spreadsheet_mutation_request};
use super::{
    apply_request, create_request, document_apply_request, document_create_request,
    document_mutation_request, mutation_request, YJS_DOCUMENT_UPDATE_BASE64,
    YJS_MARKDOWN_UPDATE_BASE64,
};
use crate::collaboration::{
    NativeOfficeCollaborationMutation, NativeOfficeCollaborationParagraphPosition,
    NativeOfficeCollaborationPdfFormField, NativeOfficeCollaborationPresentationContainerKind,
    NativeOfficeCollaborationProjectedContent, NativeOfficeCollaborationProjection,
    NativeOfficeCollaborationStore,
};

const PDF_SOURCE_SHA256: &str = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";

#[test]
fn locate_then_patch_markdown_preserves_an_unrelated_slice() {
    let temp = tempfile::tempdir().unwrap();
    let store = markdown_store(&temp.path().join("patch"));
    let read = store.project().unwrap();
    let alpha = markdown_slice(&read, "# Shared");
    let beta = markdown_slice(&read, "Yjs to Yrs.");
    assert_ne!(alpha.start_utf16, beta.start_utf16);
    let outside = store
        .mutate(mutation_request(
            "markdown-outside",
            splice(
                beta.start_utf16,
                beta.end_utf16 - beta.start_utf16,
                &beta.text,
                "Yjs to Yrs!",
            ),
        ))
        .unwrap();
    assert!(outside.state_changed);
    let moved = store.project().unwrap();
    assert_eq!(markdown_slice(&moved, "Yjs to Yrs!").text, "Yjs to Yrs!");
    assert_ne!(markdown_slice(&moved, "Yjs to Yrs!").text, beta.text);

    store
        .mutate(mutation_request(
            "markdown-patch",
            splice(
                alpha.start_utf16,
                alpha.end_utf16 - alpha.start_utf16,
                &alpha.text,
                "# Shared live",
            ),
        ))
        .unwrap();
    let patched = markdown_source_projection(&store);
    assert_eq!(patched, "# Shared live\n\nYjs to Yrs!");

    let stale_store = markdown_store(&temp.path().join("stale"));
    let stale = stale_store.project().unwrap();
    let stale_alpha = markdown_slice(&stale, "# Shared");
    let stale_beta = markdown_slice(&stale, "Yjs to Yrs.");
    stale_store
        .mutate(mutation_request(
            "markdown-stale-outside",
            splice(
                stale_alpha.start_utf16,
                stale_alpha.end_utf16 - stale_alpha.start_utf16,
                &stale_alpha.text,
                "Title",
            ),
        ))
        .unwrap();
    let before = stale_store.inspect().unwrap();
    let error = stale_store
        .mutate(mutation_request(
            "markdown-stale-patch",
            splice(
                stale_alpha.start_utf16,
                stale_alpha.end_utf16 - stale_alpha.start_utf16,
                &stale_alpha.text,
                "# Shared live",
            ),
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    assert_unchanged(&stale_store, &before.document_state_sha256);
    assert_eq!(
        markdown_source_projection(&stale_store),
        "Title\n\nYjs to Yrs."
    );
    assert_eq!(
        markdown_slice(&stale_store.project().unwrap(), "Yjs to Yrs.").text,
        stale_beta.text
    );

    let whole_store = markdown_store(&temp.path().join("whole"));
    let whole = whole_store.project().unwrap();
    let whole_beta = markdown_slice(&whole, "Yjs to Yrs.");
    let NativeOfficeCollaborationProjectedContent::Markdown { source, .. } = &whole.content else {
        panic!("markdown projection");
    };
    whole_store
        .mutate(mutation_request(
            "markdown-whole-outside",
            splice(
                whole_beta.start_utf16,
                whole_beta.end_utf16 - whole_beta.start_utf16,
                &whole_beta.text,
                "Yjs to Yrs!",
            ),
        ))
        .unwrap();
    whole_store
        .mutate(mutation_request(
            "markdown-whole-replace",
            NativeOfficeCollaborationMutation::MarkdownReplace {
                expected_markdown: source.clone(),
                markdown: source.replacen("# Shared", "# Shared live", 1),
            },
        ))
        .unwrap();
    assert_eq!(
        markdown_source_projection(&whole_store),
        "# Shared live\n\nYjs to Yrs!"
    );
}

#[test]
fn locate_then_patch_document_preserves_an_unrelated_paragraph() {
    let temp = tempfile::tempdir().unwrap();
    let store = document_store(&temp.path().join("patch"));
    seed_document_beta(&store);
    let read = store.project().unwrap();
    let alpha = document_paragraph(&read, "Hello 😀 world");
    let beta = document_paragraph(&read, "BETA");
    assert!(alpha.end_utf16 > alpha.start_utf16);
    assert_eq!(alpha.start_utf16, 0);
    store
        .mutate(document_mutation_request(
            "document-outside",
            replace_paragraph(beta, "BETA2"),
        ))
        .unwrap();
    let moved = store.project().unwrap();
    assert_eq!(document_paragraph(&moved, "BETA2").text, "BETA2");
    assert_ne!(document_paragraph(&moved, "BETA2").text_id, beta.text_id);

    store
        .mutate(document_mutation_request(
            "document-patch",
            replace_paragraph(alpha, "Hello agent"),
        ))
        .unwrap();
    let patched = store.project().unwrap();
    assert_eq!(
        document_paragraph(&patched, "Hello agent").text,
        "Hello agent"
    );
    assert_eq!(document_paragraph(&patched, "BETA2").text, "BETA2");

    let stale_store = document_store(&temp.path().join("stale"));
    seed_document_beta(&stale_store);
    let stale = stale_store.project().unwrap();
    let stale_alpha = document_paragraph(&stale, "Hello 😀 world");
    let stale_beta = document_paragraph(&stale, "BETA");
    stale_store
        .mutate(document_mutation_request(
            "document-stale-outside",
            replace_paragraph(stale_alpha, "Hello human"),
        ))
        .unwrap();
    let before = stale_store.inspect().unwrap();
    let error = stale_store
        .mutate(document_mutation_request(
            "document-stale-patch",
            replace_paragraph(stale_alpha, "Hello agent"),
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    assert_unchanged(&stale_store, &before.document_state_sha256);
    let stale_after = stale_store.project().unwrap();
    assert_eq!(
        document_paragraph(&stale_after, "Hello human").text,
        "Hello human"
    );
    assert_eq!(
        document_paragraph(&stale_after, "BETA").text,
        stale_beta.text
    );

    let whole_store = document_store(&temp.path().join("whole"));
    seed_document_beta(&whole_store);
    let whole = whole_store.project().unwrap();
    let whole_beta = document_paragraph(&whole, "BETA");
    let NativeOfficeCollaborationProjectedContent::Document { plain_text, .. } = &whole.content
    else {
        panic!("document projection");
    };
    whole_store
        .mutate(document_mutation_request(
            "document-whole-outside",
            replace_paragraph(whole_beta, "BETA2"),
        ))
        .unwrap();
    let before_whole = whole_store.inspect().unwrap();
    let error = whole_store
        .mutate(document_mutation_request(
            "document-whole-replace",
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: plain_text.clone(),
                replacement: plain_text.replacen("Hello 😀 world", "Hello agent", 1),
                expected_matches: 1,
                occurrence: None,
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    assert_unchanged(&whole_store, &before_whole.document_state_sha256);
    assert_eq!(
        document_paragraph(&whole_store.project().unwrap(), "BETA2").text,
        "BETA2"
    );
}

#[test]
fn locate_then_patch_spreadsheet_preserves_an_unrelated_cell() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("patch"));
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-seed-beta",
            set_cell("sheet-data", 2, 0, None, json!({ "v": 2, "m": "2" })),
        ))
        .unwrap();
    let read = store.project().unwrap();
    assert_projection_has_addresses(&read);
    let alpha = spreadsheet_cell(&read, 1, 0);
    let beta = spreadsheet_cell(&read, 2, 0);
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-outside",
            set_cell(
                "sheet-data",
                2,
                0,
                Some(beta.clone()),
                json!({ "v": 3, "m": "3" }),
            ),
        ))
        .unwrap();
    let moved = spreadsheet_cell(&store.project().unwrap(), 2, 0);
    assert_ne!(moved, beta);

    let mut next_alpha = alpha.clone();
    next_alpha["v"] = json!(11);
    next_alpha["m"] = json!("11");
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-patch",
            set_cell("sheet-data", 1, 0, Some(alpha.clone()), next_alpha),
        ))
        .unwrap();
    let patched = store.project().unwrap();
    assert_eq!(cell_number(&patched, 1, 0), 11.0);
    assert_eq!(cell_number(&patched, 2, 0), 3.0);

    let stale_store = initialized_spreadsheet_store(&temp.path().join("stale"));
    stale_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-stale-seed",
            set_cell("sheet-data", 2, 0, None, json!({ "v": 2, "m": "2" })),
        ))
        .unwrap();
    let stale = stale_store.project().unwrap();
    let stale_alpha = spreadsheet_cell(&stale, 1, 0);
    let stale_beta = spreadsheet_cell(&stale, 2, 0);
    let mut conflicting = stale_alpha.clone();
    conflicting["v"] = json!(99);
    conflicting["m"] = json!("99");
    stale_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-stale-outside",
            set_cell("sheet-data", 1, 0, Some(stale_alpha.clone()), conflicting),
        ))
        .unwrap();
    let before = stale_store.inspect().unwrap();
    let mut agent_next = stale_alpha.clone();
    agent_next["v"] = json!(11);
    agent_next["m"] = json!("11");
    let error = stale_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-stale-patch",
            set_cell("sheet-data", 1, 0, Some(stale_alpha), agent_next),
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    assert_unchanged(&stale_store, &before.document_state_sha256);
    assert_eq!(
        spreadsheet_cell(&stale_store.project().unwrap(), 2, 0),
        stale_beta.clone()
    );
    assert_eq!(cell_number(&stale_store.project().unwrap(), 1, 0), 99.0);

    let whole_store = initialized_spreadsheet_store(&temp.path().join("whole"));
    whole_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-whole-seed",
            set_cell("sheet-data", 2, 0, None, json!({ "v": 2, "m": "2" })),
        ))
        .unwrap();
    let whole = whole_store.project().unwrap();
    let whole_alpha = spreadsheet_cell(&whole, 1, 0);
    let whole_beta = spreadsheet_cell(&whole, 2, 0);
    whole_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-whole-outside",
            set_cell(
                "sheet-data",
                2,
                0,
                Some(whole_beta.clone()),
                json!({ "v": 3, "m": "3" }),
            ),
        ))
        .unwrap();
    let mut whole_next = whole_alpha.clone();
    whole_next["v"] = json!(11);
    whole_next["m"] = json!("11");
    whole_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-whole-replace",
            NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
                sheet_id: "sheet-data".to_owned(),
                changes: vec![
                    super::spreadsheet::spreadsheet_cell_change(
                        1,
                        0,
                        Some(whole_alpha.clone()),
                        Some(whole_next),
                    ),
                    super::spreadsheet::spreadsheet_cell_change(
                        2,
                        0,
                        Some(whole_beta.clone()),
                        Some(whole_beta.clone()),
                    ),
                ],
            },
        ))
        .unwrap();
    let replaced = whole_store.project().unwrap();
    assert_eq!(cell_number(&replaced, 1, 0), 11.0);
    assert_eq!(cell_number(&replaced, 2, 0), 3.0);
}

#[test]
fn locate_then_patch_presentation_preserves_an_unrelated_element() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_presentation_store(&temp.path().join("patch"), 900_104);
    store
        .mutate(presentation_mutation_request(
            "presentation-seed-beta",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: scene_element("element-beta", "BETA"),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    let read = store.project().unwrap();
    assert_projection_has_addresses(&read);
    let alpha = presentation_element(&read, "element-title");
    let beta = presentation_element(&read, "element-beta");
    let mut beta_next = beta.clone();
    beta_next["text"] = json!("BETA2");
    store
        .mutate(presentation_mutation_request(
            "presentation-outside",
            update_element("element-beta", beta, beta_next),
        ))
        .unwrap();
    assert_eq!(
        presentation_element(&store.project().unwrap(), "element-beta")["text"],
        "BETA2"
    );

    let mut alpha_next = alpha.clone();
    alpha_next["text"] = json!("ALPHA2");
    store
        .mutate(presentation_mutation_request(
            "presentation-patch",
            update_element("element-title", alpha, alpha_next),
        ))
        .unwrap();
    let patched = store.project().unwrap();
    assert_eq!(
        presentation_element(&patched, "element-title")["text"],
        "ALPHA2"
    );
    assert_eq!(
        presentation_element(&patched, "element-beta")["text"],
        "BETA2"
    );

    let stale_store = initialized_presentation_store(&temp.path().join("stale"), 900_105);
    stale_store
        .mutate(presentation_mutation_request(
            "presentation-stale-seed",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: scene_element("element-beta", "BETA"),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    let stale = stale_store.project().unwrap();
    let stale_alpha = presentation_element(&stale, "element-title");
    let stale_beta = presentation_element(&stale, "element-beta");
    let mut conflicting = stale_alpha.clone();
    conflicting["text"] = json!("HUMAN");
    stale_store
        .mutate(presentation_mutation_request(
            "presentation-stale-outside",
            update_element("element-title", stale_alpha.clone(), conflicting),
        ))
        .unwrap();
    let before = stale_store.inspect().unwrap();
    let mut agent_next = stale_alpha.clone();
    agent_next["text"] = json!("ALPHA2");
    let error = stale_store
        .mutate(presentation_mutation_request(
            "presentation-stale-patch",
            update_element("element-title", stale_alpha, agent_next),
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    assert_unchanged(&stale_store, &before.document_state_sha256);
    assert_eq!(
        presentation_element(&stale_store.project().unwrap(), "element-beta")["text"],
        stale_beta["text"]
    );
    assert_eq!(
        presentation_element(&stale_store.project().unwrap(), "element-title")["text"],
        "HUMAN"
    );

    let whole_store = initialized_presentation_store(&temp.path().join("whole"), 900_106);
    whole_store
        .mutate(presentation_mutation_request(
            "presentation-whole-seed",
            NativeOfficeCollaborationMutation::PresentationCreateElement {
                container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
                container_id: "slide-1".to_owned(),
                element: scene_element("element-beta", "BETA"),
                after_element_id: Some("element-title".to_owned()),
            },
        ))
        .unwrap();
    let whole = whole_store.project().unwrap();
    let whole_alpha = presentation_element(&whole, "element-title");
    let whole_beta = presentation_element(&whole, "element-beta");
    let mut live_beta = whole_beta.clone();
    live_beta["text"] = json!("BETA2");
    whole_store
        .mutate(presentation_mutation_request(
            "presentation-whole-outside",
            update_element("element-beta", whole_beta.clone(), live_beta),
        ))
        .unwrap();
    let mut whole_alpha_next = whole_alpha.clone();
    whole_alpha_next["text"] = json!("ALPHA2");
    whole_store
        .mutate(presentation_mutation_request(
            "presentation-whole-alpha",
            update_element("element-title", whole_alpha.clone(), whole_alpha_next),
        ))
        .unwrap();
    whole_store
        .mutate(presentation_mutation_request(
            "presentation-whole-echo",
            update_element("element-beta", whole_beta.clone(), whole_beta),
        ))
        .unwrap();
    let replaced = whole_store.project().unwrap();
    assert_eq!(
        presentation_element(&replaced, "element-title")["text"],
        "ALPHA2"
    );
    assert_eq!(
        presentation_element(&replaced, "element-beta")["text"],
        "BETA2"
    );
}

#[test]
fn locate_then_patch_pdf_preserves_an_unrelated_form_field() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_pdf_store(&temp.path().join("patch"));
    store
        .mutate(pdf_form_mutation(
            "pdf-seed-email",
            NativeOfficeCollaborationMutation::PdfSetFormValue {
                field_id: "Applicant.Email".to_owned(),
                value: "beta@example.test".to_owned(),
                expected_value: None,
            },
        ))
        .unwrap();
    let read = store.project().unwrap();
    assert_projection_has_addresses(&read);
    let encoded = serde_json::to_string(&read).unwrap();
    assert!(!encoded.contains(PDF_SOURCE_SHA256));
    assert!(!encoded.contains("byteLength"));
    assert!(!encoded.contains("projection_unsupported"));
    let name = pdf_field(&read, "Applicant.Name");
    let email = pdf_field(&read, "Applicant.Email");
    assert_eq!(name.value, "Ada");
    store
        .mutate(pdf_form_mutation(
            "pdf-outside",
            set_pdf_field("Applicant.Email", "beta2@example.test", &email.value),
        ))
        .unwrap();
    assert_eq!(
        pdf_field(&store.project().unwrap(), "Applicant.Email").value,
        "beta2@example.test"
    );

    store
        .mutate(pdf_form_mutation(
            "pdf-patch",
            set_pdf_field("Applicant.Name", "Grace", &name.value),
        ))
        .unwrap();
    let patched = store.project().unwrap();
    assert_eq!(pdf_field(&patched, "Applicant.Name").value, "Grace");
    assert_eq!(
        pdf_field(&patched, "Applicant.Email").value,
        "beta2@example.test"
    );

    let stale_store = initialized_pdf_store(&temp.path().join("stale"));
    stale_store
        .mutate(pdf_form_mutation(
            "pdf-stale-seed",
            NativeOfficeCollaborationMutation::PdfSetFormValue {
                field_id: "Applicant.Email".to_owned(),
                value: "beta@example.test".to_owned(),
                expected_value: None,
            },
        ))
        .unwrap();
    let stale = stale_store.project().unwrap();
    let stale_name = pdf_field(&stale, "Applicant.Name");
    let stale_email = pdf_field(&stale, "Applicant.Email");
    stale_store
        .mutate(pdf_form_mutation(
            "pdf-stale-outside",
            set_pdf_field("Applicant.Name", "Other", &stale_name.value),
        ))
        .unwrap();
    let before = stale_store.inspect().unwrap();
    let error = stale_store
        .mutate(pdf_form_mutation(
            "pdf-stale-patch",
            set_pdf_field("Applicant.Name", "Grace", &stale_name.value),
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    assert_unchanged(&stale_store, &before.document_state_sha256);
    let stale_after = stale_store.project().unwrap();
    assert_eq!(pdf_field(&stale_after, "Applicant.Name").value, "Other");
    assert_eq!(
        pdf_field(&stale_after, "Applicant.Email").value,
        stale_email.value
    );

    let whole_store = initialized_pdf_store(&temp.path().join("whole"));
    whole_store
        .mutate(pdf_form_mutation(
            "pdf-whole-seed",
            NativeOfficeCollaborationMutation::PdfSetFormValue {
                field_id: "Applicant.Email".to_owned(),
                value: "beta@example.test".to_owned(),
                expected_value: None,
            },
        ))
        .unwrap();
    let whole = whole_store.project().unwrap();
    let whole_name = pdf_field(&whole, "Applicant.Name");
    let whole_email = pdf_field(&whole, "Applicant.Email");
    whole_store
        .mutate(pdf_form_mutation(
            "pdf-whole-outside",
            set_pdf_field("Applicant.Email", "beta2@example.test", &whole_email.value),
        ))
        .unwrap();
    whole_store
        .mutate(pdf_form_mutation(
            "pdf-whole-name",
            set_pdf_field("Applicant.Name", "Grace", &whole_name.value),
        ))
        .unwrap();
    whole_store
        .mutate(pdf_form_mutation(
            "pdf-whole-echo",
            set_pdf_field("Applicant.Email", &whole_email.value, &whole_email.value),
        ))
        .unwrap();
    let replaced = whole_store.project().unwrap();
    assert_eq!(pdf_field(&replaced, "Applicant.Name").value, "Grace");
    assert_eq!(
        pdf_field(&replaced, "Applicant.Email").value,
        "beta2@example.test"
    );
    let NativeOfficeCollaborationProjectedContent::Pdf {
        page_count,
        annotations,
        form_fields,
    } = &replaced.content
    else {
        panic!("pdf projection");
    };
    assert_eq!(*page_count, 3);
    assert!(
        annotations.is_empty()
            || annotations
                .iter()
                .all(|annotation| annotation.page_index < *page_count)
    );
    assert!(form_fields
        .iter()
        .any(|field| field.field_id == "Applicant.Email"));
}

fn markdown_store(root: &Path) -> NativeOfficeCollaborationStore {
    let store = NativeOfficeCollaborationStore::create(create_request(root)).unwrap();
    store
        .apply(apply_request(
            "bootstrap-markdown",
            STANDARD.decode(YJS_MARKDOWN_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
}

fn document_store(root: &Path) -> NativeOfficeCollaborationStore {
    let store = NativeOfficeCollaborationStore::create(document_create_request(root)).unwrap();
    store
        .apply(document_apply_request(
            "bootstrap-document",
            STANDARD.decode(YJS_DOCUMENT_UPDATE_BASE64).unwrap(),
        ))
        .unwrap();
    store
}

fn seed_document_beta(store: &NativeOfficeCollaborationStore) {
    let read = store.project().unwrap();
    let alpha = document_paragraph(&read, "Hello 😀 world");
    store
        .mutate(document_mutation_request(
            "document-seed-beta",
            NativeOfficeCollaborationMutation::DocumentInsertParagraph {
                anchor_paragraph_id: alpha.paragraph_id.clone().unwrap(),
                position: NativeOfficeCollaborationParagraphPosition::After,
                paragraph_id: "00000020".to_owned(),
                text_id: "00000021".to_owned(),
                text: "BETA".to_owned(),
            },
        ))
        .unwrap();
}

fn splice(
    index_utf16: u32,
    delete_utf16: u32,
    expected_text: &str,
    insert: &str,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::MarkdownSplice {
        index_utf16,
        delete_utf16,
        expected_text: expected_text.to_owned(),
        insert: insert.to_owned(),
    }
}

fn markdown_slice<'a>(
    projection: &'a NativeOfficeCollaborationProjection,
    text: &str,
) -> &'a crate::collaboration::NativeOfficeCollaborationMarkdownSlice {
    let NativeOfficeCollaborationProjectedContent::Markdown { slices, .. } = &projection.content
    else {
        panic!(
            "expected markdown projection, got {:?}",
            projection.artifact_kind
        );
    };
    slices
        .iter()
        .find(|slice| slice.text == text)
        .unwrap_or_else(|| {
            panic!("missing markdown slice {text:?}");
        })
}

fn markdown_source_projection(store: &NativeOfficeCollaborationStore) -> String {
    let projection = store.project().unwrap();
    let NativeOfficeCollaborationProjectedContent::Markdown { source, .. } = projection.content
    else {
        panic!("expected markdown projection");
    };
    source
}

fn document_paragraph<'a>(
    projection: &'a NativeOfficeCollaborationProjection,
    text: &str,
) -> &'a crate::collaboration::NativeOfficeCollaborationDocumentParagraph {
    let NativeOfficeCollaborationProjectedContent::Document { paragraphs, .. } =
        &projection.content
    else {
        panic!("expected document projection");
    };
    paragraphs
        .iter()
        .find(|paragraph| paragraph.text == text)
        .unwrap_or_else(|| {
            panic!("missing paragraph {text:?}");
        })
}

fn replace_paragraph(
    paragraph: &crate::collaboration::NativeOfficeCollaborationDocumentParagraph,
    replacement: &str,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::DocumentReplaceParagraph {
        paragraph_id: paragraph.paragraph_id.clone().expect("paragraph id"),
        expected_text_id: paragraph.text_id.clone().expect("text id"),
        expected_text: paragraph.text.clone(),
        replacement: replacement.to_owned(),
    }
}

fn spreadsheet_cell(
    projection: &NativeOfficeCollaborationProjection,
    row: u32,
    column: u32,
) -> JsonValue {
    let NativeOfficeCollaborationProjectedContent::Spreadsheet { sheets } = &projection.content
    else {
        panic!("expected spreadsheet projection");
    };
    sheets
        .iter()
        .flat_map(|sheet| sheet.cells.iter())
        .find(|cell| cell.row == row && cell.column == column)
        .unwrap_or_else(|| panic!("missing cell {row}:{column}"))
        .cell
        .clone()
}

fn cell_number(projection: &NativeOfficeCollaborationProjection, row: u32, column: u32) -> f64 {
    spreadsheet_cell(projection, row, column)["v"]
        .as_f64()
        .unwrap_or_else(|| panic!("cell {row}:{column} value is not a number"))
}

fn set_cell(
    sheet_id: &str,
    row: u32,
    column: u32,
    expected_cell: Option<JsonValue>,
    next_cell: JsonValue,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::SpreadsheetSetCell {
        sheet_id: sheet_id.to_owned(),
        row,
        column,
        expected_cell,
        next_cell,
    }
}

fn presentation_element(
    projection: &NativeOfficeCollaborationProjection,
    element_id: &str,
) -> JsonValue {
    let NativeOfficeCollaborationProjectedContent::Presentation { containers } =
        &projection.content
    else {
        panic!("expected presentation projection");
    };
    containers
        .iter()
        .flat_map(|container| container.elements.iter())
        .find(|element| element.element_id == element_id)
        .unwrap_or_else(|| panic!("missing element {element_id}"))
        .element
        .clone()
}

fn update_element(
    element_id: &str,
    expected_element: JsonValue,
    next_element: JsonValue,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::PresentationUpdateElement {
        container_kind: NativeOfficeCollaborationPresentationContainerKind::Slide,
        container_id: "slide-1".to_owned(),
        element_id: element_id.to_owned(),
        expected_element,
        next_element,
    }
}

fn scene_element(id: &str, text: &str) -> JsonValue {
    json!({
        "id": id,
        "type": "shape",
        "x": 20,
        "y": 20,
        "width": 40,
        "height": 20,
        "text": text,
        "fontSize": 18,
        "color": "#172033",
        "fill": "#F8FAFC",
        "bold": false,
        "align": "left",
    })
}

fn pdf_field<'a>(
    projection: &'a NativeOfficeCollaborationProjection,
    field_id: &str,
) -> &'a NativeOfficeCollaborationPdfFormField {
    let NativeOfficeCollaborationProjectedContent::Pdf { form_fields, .. } = &projection.content
    else {
        panic!("expected pdf projection");
    };
    form_fields
        .iter()
        .find(|field| field.field_id == field_id)
        .unwrap_or_else(|| {
            panic!("missing form field {field_id}");
        })
}

fn set_pdf_field(
    field_id: &str,
    value: &str,
    expected_value: &str,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::PdfSetFormValue {
        field_id: field_id.to_owned(),
        value: value.to_owned(),
        expected_value: Some(expected_value.to_owned()),
    }
}

fn assert_projection_has_addresses(projection: &NativeOfficeCollaborationProjection) {
    assert_eq!(projection.version, 4);
    let encoded = serde_json::to_string(projection).unwrap();
    assert!(!encoded.contains("projection_unsupported"));
    assert!(!encoded.contains("byteLength"));
}

fn assert_unchanged(store: &NativeOfficeCollaborationStore, state_sha256: &str) {
    let after = store.inspect().unwrap();
    assert_eq!(after.document_state_sha256, state_sha256);
}
