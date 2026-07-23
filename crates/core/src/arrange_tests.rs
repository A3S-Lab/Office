use crate::{
    NativeOfficeEditor, NativeOfficeInsertPosition, NativeOfficeMutation, NativeOfficePartType,
    SpreadsheetCellValue,
};

#[test]
fn arrangement_mutations_have_a_stable_typed_batch_schema() {
    let mutations = vec![
        NativeOfficeMutation::Move {
            path: "/body/p[2]".into(),
            target_parent: Some("/body".into()),
            position: Some(NativeOfficeInsertPosition::at_index(0)),
        },
        NativeOfficeMutation::Copy {
            path: "/Sheet1".into(),
            target_parent: Some("/".into()),
            position: Some(NativeOfficeInsertPosition::before("/Data")),
            name: Some("Copy".into()),
        },
        NativeOfficeMutation::Swap {
            path: "/slide[1]".into(),
            with: "/slide[2]".into(),
        },
    ];
    let encoded = serde_json::to_value(&mutations).unwrap();
    assert_eq!(
        encoded,
        serde_json::json!([
            {
                "operation": "move",
                "path": "/body/p[2]",
                "to": "/body",
                "position": { "kind": "index", "index": 0 }
            },
            {
                "operation": "copy",
                "path": "/Sheet1",
                "to": "/",
                "position": { "kind": "before", "path": "/Data" },
                "name": "Copy"
            },
            {
                "operation": "swap",
                "path": "/slide[1]",
                "with": "/slide[2]"
            }
        ])
    );
    let decoded: Vec<NativeOfficeMutation> = serde_json::from_value(encoded).unwrap();
    assert_eq!(decoded, mutations);
}

#[tokio::test]
async fn word_move_copy_and_swap_are_typed_atomic_and_loss_preserving() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("arrange.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/body/p[1]", "A").unwrap();
    editor.add_paragraph("/body", "B").unwrap();
    editor.add_paragraph("/body", "C").unwrap();

    assert_eq!(
        editor
            .move_node(
                "/body/p[1]",
                None,
                Some(NativeOfficeInsertPosition::after("/body/p[2]")),
            )
            .unwrap(),
        "/body/p[2]"
    );
    assert_eq!(word_texts(&editor), ["B", "A", "C"]);
    assert_eq!(
        editor.copy_node("/body/p[2]", None, None, None).unwrap(),
        "/body/p[3]"
    );
    assert_eq!(word_texts(&editor), ["B", "A", "A", "C"]);
    let swapped = editor.swap_nodes("/body/p[1]", "/body/p[4]").unwrap();
    assert_eq!(swapped.first, "/body/p[4]");
    assert_eq!(swapped.second, "/body/p[1]");
    assert_eq!(word_texts(&editor), ["C", "A", "A", "B"]);

    let before = editor.package().content_sha256();
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::Move {
                path: "/body/p[1]".into(),
                target_parent: None,
                position: Some(NativeOfficeInsertPosition::at_index(3)),
            },
            NativeOfficeMutation::Copy {
                path: "/body/p[99]".into(),
                target_parent: None,
                position: None,
                name: None,
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.node_not_found");
    assert_eq!(editor.package().content_sha256(), before);
    assert_eq!(word_texts(&editor), ["C", "A", "A", "B"]);

    editor.add_table("/body", 1, 2).unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .copy_node("/body/tbl[1]/tr[1]/tc[1]", None, None, None)
        .unwrap_err();
    assert_eq!(error.code, "use.office.word_copy_unsupported");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn spreadsheet_arrangement_covers_sheets_and_plain_rows_and_fails_closed_for_formulas() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("arrange.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for (row, value) in [(1, "A"), (2, "B"), (3, "C")] {
        editor
            .set_cell_value(
                format!("/Sheet1/A{row}"),
                SpreadsheetCellValue::Text {
                    value: value.into(),
                },
            )
            .unwrap();
    }

    assert_eq!(
        editor
            .move_node(
                "/Sheet1/row[1]",
                None,
                Some(NativeOfficeInsertPosition::after("/Sheet1/row[2]")),
            )
            .unwrap(),
        "/Sheet1/row[2]"
    );
    assert_eq!(spreadsheet_row_texts(&editor, "/Sheet1"), ["B", "A", "C"]);
    assert_eq!(
        editor
            .copy_node("/Sheet1/row[2]", None, None, None)
            .unwrap(),
        "/Sheet1/row[3]"
    );
    let swapped = editor
        .swap_nodes("/Sheet1/row[1]", "/Sheet1/row[4]")
        .unwrap();
    assert_eq!(swapped.first, "/Sheet1/row[4]");
    assert_eq!(swapped.second, "/Sheet1/row[1]");
    assert_eq!(
        spreadsheet_row_texts(&editor, "/Sheet1"),
        ["C", "A", "A", "B"]
    );

    editor.add_worksheet("Data").unwrap();
    assert_eq!(
        editor
            .copy_node(
                "/Sheet1",
                Some("/".into()),
                Some(NativeOfficeInsertPosition::at_index(1)),
                Some("Copy".into()),
            )
            .unwrap(),
        "/Copy"
    );
    editor
        .move_node("/Data", None, Some(NativeOfficeInsertPosition::at_index(0)))
        .unwrap();
    editor.swap_nodes("/Data", "/Copy").unwrap();
    assert_eq!(worksheet_paths(&editor), ["/Copy", "/Sheet1", "/Data"]);

    editor
        .set_cell_value(
            "/Sheet1/B1",
            SpreadsheetCellValue::Formula {
                expression: "A2".into(),
            },
        )
        .unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .move_node(
            "/Sheet1/row[1]",
            None,
            Some(NativeOfficeInsertPosition::after("/Sheet1/row[2]")),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_row_arrange_unsupported");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn presentation_arrangement_covers_slides_and_shapes_and_rejects_owned_graphs() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("arrange.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for title in ["A", "B", "C"] {
        editor.add_slide("/", title).unwrap();
    }
    editor
        .move_node(
            "/slide[1]",
            None,
            Some(NativeOfficeInsertPosition::after("/slide[2]")),
        )
        .unwrap();
    assert_eq!(slide_texts(&editor), ["B", "A", "C"]);
    assert_eq!(
        editor.copy_node("/slide[2]", None, None, None).unwrap(),
        "/slide[3]"
    );
    editor.swap_nodes("/slide[1]", "/slide[4]").unwrap();
    assert_eq!(slide_texts(&editor), ["C", "A", "A", "B"]);

    let shape_path = temp.path().join("shapes.pptx");
    let mut shapes = NativeOfficeEditor::create(&shape_path).await.unwrap();
    shapes.add_slide("/", "A").unwrap();
    shapes.add_shape("/slide[1]", "B").unwrap();
    shapes.add_shape("/slide[1]", "C").unwrap();
    shapes
        .move_node(
            "/slide[1]/shape[1]",
            None,
            Some(NativeOfficeInsertPosition::after("/slide[1]/shape[2]")),
        )
        .unwrap();
    shapes
        .copy_node("/slide[1]/shape[2]", None, None, None)
        .unwrap();
    shapes
        .swap_nodes("/slide[1]/shape[1]", "/slide[1]/shape[4]")
        .unwrap();
    assert_eq!(shape_texts(&shapes, "/slide[1]"), ["C", "A", "A", "B"]);
    let ids = shapes
        .snapshot()
        .unwrap()
        .get("/slide[1]", 1)
        .unwrap()
        .children
        .iter()
        .filter_map(|shape| shape.format.get("id"))
        .cloned()
        .collect::<std::collections::BTreeSet<_>>();
    assert_eq!(ids.len(), 4);

    let slide_part = shapes
        .snapshot()
        .unwrap()
        .get("/slide[1]", 0)
        .unwrap()
        .format["part"]
        .clone();
    let slide_xml = String::from_utf8(
        shapes
            .package()
            .xml_part(&slide_part)
            .unwrap()
            .raw()
            .to_vec(),
    )
    .unwrap();
    let placeholder_xml = slide_xml.replacen(
        "<p:cNvSpPr txBox=\"1\"/><p:nvPr/>",
        "<p:cNvSpPr txBox=\"1\"/><p:nvPr><p:ph type=\"body\"/></p:nvPr>",
        1,
    );
    assert_ne!(placeholder_xml, slide_xml);
    shapes
        .replace_xml_part(&slide_part, placeholder_xml)
        .unwrap();
    let before = shapes.package().content_sha256();
    let error = shapes
        .copy_node("/slide[1]/shape[1]", None, None, None)
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.presentation_object_copy_unsupported"
    );
    assert_eq!(shapes.package().content_sha256(), before);

    shapes
        .add_part("/slide[1]", NativeOfficePartType::Chart)
        .unwrap();
    let before = shapes.package().content_sha256();
    let error = shapes.copy_node("/slide[1]", None, None, None).unwrap_err();
    assert_eq!(error.code, "use.office.presentation_slide_copy_unsupported");
    assert_eq!(shapes.package().content_sha256(), before);
}

fn word_texts(editor: &NativeOfficeEditor) -> Vec<String> {
    editor
        .snapshot()
        .unwrap()
        .get("/body", 1)
        .unwrap()
        .children
        .into_iter()
        .filter(|node| node.tag == "p")
        .map(|node| node.text)
        .collect()
}

fn spreadsheet_row_texts(editor: &NativeOfficeEditor, sheet: &str) -> Vec<String> {
    editor
        .snapshot()
        .unwrap()
        .get(sheet, 1)
        .unwrap()
        .children
        .into_iter()
        .map(|node| node.text)
        .collect()
}

fn worksheet_paths(editor: &NativeOfficeEditor) -> Vec<String> {
    editor
        .snapshot()
        .unwrap()
        .get("/", 1)
        .unwrap()
        .children
        .into_iter()
        .map(|node| node.path)
        .collect()
}

fn slide_texts(editor: &NativeOfficeEditor) -> Vec<String> {
    editor
        .snapshot()
        .unwrap()
        .get("/", 1)
        .unwrap()
        .children
        .into_iter()
        .map(|node| node.text)
        .collect()
}

fn shape_texts(editor: &NativeOfficeEditor, slide: &str) -> Vec<String> {
    editor
        .snapshot()
        .unwrap()
        .get(slide, 1)
        .unwrap()
        .children
        .into_iter()
        .map(|node| node.text)
        .collect()
}
