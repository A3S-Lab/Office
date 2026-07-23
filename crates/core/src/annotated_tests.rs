use crate::{
    DocumentKind, NativeOfficeAnnotatedOptions, NativeOfficeEditor, OfficeNodeType,
    SpreadsheetCellValue, DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT, MAX_NATIVE_OFFICE_ANNOTATED_LIMIT,
};

#[tokio::test]
async fn annotated_view_is_typed_bounded_and_covers_all_formats() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("annotated.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    word.set_text("/body/p[1]", "Native Word").unwrap();
    let word_view = word.snapshot().unwrap().annotated_view().unwrap();
    assert_eq!(word_view.kind, DocumentKind::Word);
    assert_eq!(word_view.limit, DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT);
    assert!(word_view.entries.iter().any(|entry| {
        entry.path == "/body/p[1]"
            && entry.node_type == OfficeNodeType::Paragraph
            && entry.text == "Native Word"
    }));
    assert!(word_view.text.contains("[/body/p[1]] [Paragraph]"));
    assert!(!word_view.text.contains(word_path.to_str().unwrap()));

    let spreadsheet_path = temp.path().join("annotated.xlsx");
    let mut spreadsheet = NativeOfficeEditor::create(&spreadsheet_path).await.unwrap();
    spreadsheet
        .set_cell_value(
            "/Sheet1/B2",
            SpreadsheetCellValue::Formula {
                expression: "SUM(A1:A2)".into(),
            },
        )
        .unwrap();
    let spreadsheet_view = spreadsheet.snapshot().unwrap().annotated_view().unwrap();
    let cell = spreadsheet_view
        .entries
        .iter()
        .find(|entry| entry.path == "/Sheet1/B2")
        .unwrap();
    assert_eq!(cell.node_type, OfficeNodeType::Cell);
    assert_eq!(cell.format["formula"], "SUM(A1:A2)");
    assert_eq!(cell.format["valueType"], "Number");
    assert!(spreadsheet_view.text.contains("formula=SUM(A1:A2)"));

    let presentation_path = temp.path().join("annotated.pptx");
    let mut presentation = NativeOfficeEditor::create(&presentation_path)
        .await
        .unwrap();
    let slide = presentation.add_slide("/", "Native Deck").unwrap();
    presentation.add_shape(&slide, "Body text").unwrap();
    let presentation_view = presentation.snapshot().unwrap().annotated_view().unwrap();
    assert!(presentation_view.entries.iter().any(|entry| {
        entry.path == "/slide[1]"
            && entry.node_type == OfficeNodeType::Slide
            && entry.text.contains("Native Deck")
    }));
    assert!(presentation_view
        .entries
        .iter()
        .any(|entry| entry.node_type == OfficeNodeType::Shape && entry.text == "Body text"));
}

#[tokio::test]
async fn annotated_view_reports_truncation_and_rejects_invalid_limits() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("bounded.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_text("/body/p[1]", "safe-unicode-界".repeat(1_000))
        .unwrap();
    editor.add_paragraph("/body", "second").unwrap();
    let document = editor.snapshot().unwrap();

    let view = document
        .annotated(NativeOfficeAnnotatedOptions { limit: 1 })
        .unwrap();
    assert_eq!(view.returned, 1);
    assert!(view.total > view.returned);
    assert!(view.truncated);
    assert!(view.entries[0].text_truncated);
    assert!(view.text.contains("showed 1 of"));
    assert!(view.entries[0]
        .text
        .is_char_boundary(view.entries[0].text.len()));

    for limit in [0, MAX_NATIVE_OFFICE_ANNOTATED_LIMIT + 1] {
        let error = document
            .annotated(NativeOfficeAnnotatedOptions { limit })
            .unwrap_err();
        assert_eq!(error.code, "use.office.annotated_limit_invalid");
        assert_eq!(error.details["limit"], limit);
    }
}

#[test]
fn annotated_contracts_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeOfficeAnnotatedOptions>();
    assert_send_sync::<crate::NativeOfficeAnnotatedEntry>();
    assert_send_sync::<crate::NativeOfficeAnnotatedView>();
}
