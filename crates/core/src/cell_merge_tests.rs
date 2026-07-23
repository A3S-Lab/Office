use crate::{NativeOfficeEditor, NativeOfficeMutation, SpreadsheetCellValue};

const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEET_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

#[test]
fn merge_mutations_have_stable_typed_json_and_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeOfficeMutation>();

    let merge = NativeOfficeMutation::MergeCells {
        path: "/Sheet1/A1:B2".into(),
    };
    assert_eq!(
        serde_json::to_value(&merge).unwrap(),
        serde_json::json!({
            "operation": "merge-cells",
            "path": "/Sheet1/A1:B2"
        })
    );
    let unmerge: NativeOfficeMutation = serde_json::from_value(serde_json::json!({
        "operation": "unmerge-cells",
        "path": "/Sheet1/A1:B2"
    }))
    .unwrap();
    assert_eq!(
        unmerge,
        NativeOfficeMutation::UnmergeCells {
            path: "/Sheet1/A1:B2".into()
        }
    );
}

#[tokio::test]
async fn merges_normalize_round_trip_and_render_without_materializing_blank_cells() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("merged.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1",
            SpreadsheetCellValue::Text {
                value: "Quarter".into(),
            },
        )
        .unwrap();

    assert_eq!(
        editor.merge_cells("/Sheet1/B2:A1").unwrap(),
        "/Sheet1/A1:B2"
    );
    let worksheet = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(worksheet.contains("<mergeCells count=\"1\"><mergeCell ref=\"A1:B2\"/></mergeCells>"));

    let snapshot = editor.snapshot().unwrap();
    let anchor = snapshot.get("/Sheet1/A1", 0).unwrap();
    assert_eq!(anchor.format["merge"], "A1:B2");
    assert_eq!(anchor.format["mergeAnchor"], "true");
    let blank = snapshot.get("/Sheet1/B2", 0).unwrap();
    assert_eq!(blank.text, "");
    assert_eq!(blank.format["merge"], "A1:B2");
    assert_eq!(blank.format["mergeAnchor"], "false");
    let range = snapshot.get("/Sheet1/A1:B2", 0).unwrap();
    assert_eq!(range.format["merge"], "true");
    let merge = snapshot.query("mergeCell").unwrap();
    assert_eq!(merge.len(), 1);
    assert_eq!(merge[0].path, "/Sheet1/mergeCell[1]");
    assert_eq!(merge[0].format["ref"], "A1:B2");

    let html = snapshot.html_view().unwrap();
    assert!(html.content.contains("data-merge=\"A1:B2\""));
    assert!(html.content.contains("data-merge-anchor=\"true\""));
    let svg = snapshot.svg_view().unwrap();
    assert!(svg.content.contains("data-merge=\"A1:B2\""));
    assert!(svg.content.contains("data-merge-anchor=\"true\""));

    editor.save().await.unwrap();
    let mut reopened = NativeOfficeEditor::open(&path).await.unwrap();
    let before = reopened.package().content_sha256();
    assert_eq!(
        reopened.merge_cells("/sheet1/a1:b2").unwrap(),
        "/Sheet1/A1:B2"
    );
    assert!(!reopened.is_dirty());
    assert_eq!(reopened.package().content_sha256(), before);

    assert_eq!(
        reopened.unmerge_cells("/Sheet1/A1:B2").unwrap(),
        "/Sheet1/A1:B2"
    );
    assert!(!part_text(&reopened, "xl/worksheets/sheet1.xml").contains("mergeCells"));
    assert!(!reopened
        .snapshot()
        .unwrap()
        .get("/Sheet1/A1", 0)
        .unwrap()
        .format
        .contains_key("merge"));

    reopened.save().await.unwrap();
    let mut idempotent = NativeOfficeEditor::open(&path).await.unwrap();
    idempotent.unmerge_cells("/Sheet1/A1:B2").unwrap();
    assert!(!idempotent.is_dirty());
}

#[tokio::test]
async fn merge_rejects_overlaps_and_unmerge_requires_an_exact_existing_range_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("precision.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.merge_cells("/Sheet1/A1:B1").unwrap();
    editor.merge_cells("/Sheet1/A2:B2").unwrap();

    let before = editor.package().content_sha256();
    let error = editor.merge_cells("/Sheet1/B1:C1").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_overlap");
    assert_eq!(error.details["requested"], "B1:C1");
    assert_eq!(error.details["existing"], "A1:B1");
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor.unmerge_cells("/Sheet1/B1:C1").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_not_exact");
    assert_eq!(error.details["validRanges"], serde_json::json!(["A1:B1"]));
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor.unmerge_cells("/Sheet1/A1").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_not_exact");
    assert_eq!(error.details["validRanges"], serde_json::json!(["A1:B1"]));
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor.unmerge_cells("/Sheet1/A1:B2").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_not_exact");
    assert_eq!(
        error.details["validRanges"],
        serde_json::json!(["A1:B1", "A2:B2"])
    );
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::UnmergeCells {
                path: "/Sheet1/A1:B1".into(),
            },
            NativeOfficeMutation::MergeCells {
                path: "/Sheet1/A2:C2".into(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_overlap");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn merge_preserves_schema_order_strict_namespaces_and_unknown_container_data() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let worksheet = format!(
        "<worksheet xmlns=\"{STRICT_SPREADSHEET_NAMESPACE}\" xmlns:v=\"urn:vendor\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>anchor</t></is></c></row></sheetData><autoFilter ref=\"A1:F20\"/><mergeCells count=\"1\" v:count=\"vendor\"><mergeCell ref=\"A1:B1\"/><v:extension v:value=\"keep\"/></mergeCells><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>TRUE</formula></cfRule></conditionalFormatting><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>"
    );
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor.merge_cells("/Sheet1/C1:D1").unwrap();
    let edited = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(edited.contains(STRICT_SPREADSHEET_NAMESPACE));
    assert!(edited.contains("v:count=\"vendor\""));
    assert!(edited.contains("<v:extension v:value=\"keep\"/>"));
    assert!(edited.contains("count=\"2\""));
    assert!(edited.contains("<mergeCell ref=\"C1:D1\"/>"));
    assert!(edited.find("<autoFilter").unwrap() < edited.find("<mergeCells").unwrap());
    assert!(edited.find("<mergeCells").unwrap() < edited.find("<conditionalFormatting").unwrap());

    editor.unmerge_cells("/Sheet1/C1:D1").unwrap();
    let preserved = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(preserved.contains("v:count=\"vendor\""));
    assert!(preserved.contains("<v:extension v:value=\"keep\"/>"));
    assert!(preserved.contains("count=\"1\""));

    let before = editor.package().content_sha256();
    let error = editor.unmerge_cells("/Sheet1/A1:B1").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_unknown_content");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn merge_rejects_table_intersections_and_non_spreadsheets() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let table_part = "xl/tables/table1.xml";
    package
        .set_part(
            table_part,
            format!(
                "<table xmlns=\"{SPREADSHEET_NAMESPACE}\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"A1:B3\"><autoFilter ref=\"A1:B3\"/><tableColumns count=\"2\"><tableColumn id=\"1\" name=\"A\"/><tableColumn id=\"2\" name=\"B\"/></tableColumns></table>"
            )
            .into_bytes(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        table_part,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/worksheets/_rels/sheet1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        "../tables/table1.xml",
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let before = editor.package().content_sha256();
    let error = editor.merge_cells("/Sheet1/A2:B2").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_merge_table_overlap");
    assert_eq!(error.details["table"], "Table1");
    assert_eq!(error.details["tableRange"], "A1:B3");
    assert_eq!(editor.package().content_sha256(), before);
    editor.merge_cells("/Sheet1/C1:D1").unwrap();

    let word_path = temp.path().join("word.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    let error = word.merge_cells("/Sheet1/A1:B1").unwrap_err();
    assert_eq!(error.code, "use.office.mutation_type_unsupported");
}

fn part_text(editor: &NativeOfficeEditor, part: &str) -> String {
    String::from_utf8(editor.package().part(part).unwrap().to_vec()).unwrap()
}
