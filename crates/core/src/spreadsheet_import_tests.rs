use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficeReplayArtifact,
    NativeSpreadsheetDelimitedFormat, NativeSpreadsheetDelimitedImport,
    NativeSpreadsheetFrozenPane, OfficeNodeType, SpreadsheetCellValue,
};

fn text(value: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Text {
        value: value.to_string(),
    }
}

fn cell(editor: &NativeOfficeEditor, path: &str) -> crate::DocumentNode {
    editor.snapshot().unwrap().get(path, 0).unwrap()
}

#[test]
fn delimited_import_and_frozen_pane_have_closed_typed_batch_contracts() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetDelimitedImport>();
    assert_send_sync::<NativeSpreadsheetFrozenPane>();

    let mutation = NativeOfficeMutation::ImportSpreadsheetDelimited {
        sheet: "/Sheet1".into(),
        import: NativeSpreadsheetDelimitedImport::new(
            "Name\tValue\nAlpha\t42",
            NativeSpreadsheetDelimitedFormat::Tsv,
        )
        .with_header(true)
        .with_start_cell("B2"),
    };
    assert_eq!(
        serde_json::to_value(mutation).unwrap(),
        serde_json::json!({
            "operation": "import-spreadsheet-delimited",
            "sheet": "/Sheet1",
            "import": {
                "content": "Name\tValue\nAlpha\t42",
                "format": "tsv",
                "header": true,
                "startCell": "B2"
            }
        })
    );

    let pane = NativeOfficeMutation::SetSpreadsheetFrozenPane {
        sheet: "/Sheet1".into(),
        pane: NativeSpreadsheetFrozenPane::new(2, 0, "B3"),
    };
    assert_eq!(
        serde_json::to_value(pane).unwrap(),
        serde_json::json!({
            "operation": "set-spreadsheet-frozen-pane",
            "sheet": "/Sheet1",
            "pane": {
                "frozenRows": 2,
                "frozenColumns": 0,
                "topLeftCell": "B3"
            }
        })
    );
}

#[tokio::test]
async fn import_handles_bom_quotes_blank_rows_types_header_filter_and_freeze() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("import.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let content = "\u{feff}Name,Value,Note\r\n\"Alpha, Inc\",001,\"line 1\nline 2\"\r\n\r\nBeta,TRUE,2026-07-17";

    let receipt = editor
        .import_spreadsheet_delimited(
            "/Sheet1",
            NativeSpreadsheetDelimitedImport::new(content, NativeSpreadsheetDelimitedFormat::Csv)
                .with_header(true)
                .with_start_cell("B2"),
        )
        .unwrap();

    assert_eq!(receipt.path, "/Sheet1/B2:D5");
    assert_eq!(receipt.range.as_deref(), Some("B2:D5"));
    assert_eq!(receipt.row_count, 4);
    assert_eq!(receipt.column_count, 3);
    assert_eq!(receipt.filter_path.as_deref(), Some("/Sheet1/autofilter"));
    assert_eq!(receipt.freeze_path.as_deref(), Some("/Sheet1/freeze"));
    assert!(receipt.changed);

    assert_eq!(cell(&editor, "/Sheet1/B3").text, "Alpha, Inc");
    assert_eq!(cell(&editor, "/Sheet1/C3").text, "001");
    assert_eq!(cell(&editor, "/Sheet1/C3").format["valueType"], "Number");
    assert_eq!(cell(&editor, "/Sheet1/D3").text, "line 1\nline 2");
    assert!(editor.snapshot().unwrap().get("/Sheet1/B4", 0).is_err());
    assert_eq!(cell(&editor, "/Sheet1/C5").text, "true");
    assert_eq!(cell(&editor, "/Sheet1/C5").format["valueType"], "Boolean");
    assert_eq!(
        cell(&editor, "/Sheet1/D5").format["numberFormat"],
        "yyyy-mm-dd"
    );

    let snapshot = editor.snapshot().unwrap();
    let filter = snapshot.get("/Sheet1/autofilter", 0).unwrap();
    assert_eq!(filter.format["ref"], "B2:D5");
    let freeze = snapshot.get("/Sheet1/freeze", 0).unwrap();
    assert_eq!(freeze.node_type, OfficeNodeType::FrozenPane);
    assert_eq!(freeze.format["frozenRows"], "2");
    assert_eq!(freeze.format["frozenColumns"], "0");
    assert_eq!(freeze.format["topLeftCell"], "B3");
    assert_eq!(freeze.format["nativeMutable"], "true");
}

#[tokio::test]
async fn import_upserts_existing_cells_preserves_ragged_columns_and_clears_explicit_empties() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("upsert.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for (path, value) in [
        ("/Sheet1/A2", "left"),
        ("/Sheet1/B2", "old-b2"),
        ("/Sheet1/C2", "old-c2"),
        ("/Sheet1/D2", "right-2"),
        ("/Sheet1/B3", "old-b3"),
        ("/Sheet1/C3", "old-c3"),
        ("/Sheet1/D3", "right-3"),
    ] {
        editor.set_cell_value(path, text(value)).unwrap();
    }

    let result = editor
        .import_spreadsheet_delimited(
            "/Sheet1",
            NativeSpreadsheetDelimitedImport::new(
                "new-b2,\n,new-c3",
                NativeSpreadsheetDelimitedFormat::Csv,
            )
            .with_start_cell("B2"),
        )
        .unwrap();
    assert_eq!(result.range.as_deref(), Some("B2:C3"));
    assert_eq!(cell(&editor, "/Sheet1/A2").text, "left");
    assert_eq!(cell(&editor, "/Sheet1/B2").text, "new-b2");
    assert_eq!(cell(&editor, "/Sheet1/C2").text, "");
    assert_eq!(cell(&editor, "/Sheet1/B3").text, "");
    assert_eq!(cell(&editor, "/Sheet1/C3").text, "new-c3");
    assert_eq!(cell(&editor, "/Sheet1/D2").text, "right-2");
    assert_eq!(cell(&editor, "/Sheet1/D3").text, "right-3");

    let quoted_empty = editor
        .import_spreadsheet_delimited(
            "/Sheet1",
            NativeSpreadsheetDelimitedImport::new("\"\"", NativeSpreadsheetDelimitedFormat::Csv)
                .with_start_cell("A2"),
        )
        .unwrap();
    assert_eq!(quoted_empty.row_count, 1);
    assert_eq!(quoted_empty.column_count, 1);
    assert_eq!(cell(&editor, "/Sheet1/A2").text, "");
}

#[tokio::test]
async fn import_validates_before_commit_and_failed_batches_roll_back() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("rollback.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetCellValue {
                path: "/Sheet1/A1".into(),
                value: text("must roll back"),
            },
            NativeOfficeMutation::ImportSpreadsheetDelimited {
                sheet: "/Sheet1".into(),
                import: NativeSpreadsheetDelimitedImport::new(
                    "one,two",
                    NativeSpreadsheetDelimitedFormat::Csv,
                )
                .with_start_cell("XFD1"),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_import_column_limit");
    assert_eq!(editor.package().content_sha256(), before);

    let long = "x".repeat(32_768);
    let error = editor
        .import_spreadsheet_delimited(
            "/Sheet1",
            NativeSpreadsheetDelimitedImport::new(long, NativeSpreadsheetDelimitedFormat::Csv),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_import_cell_limit");
    assert_eq!(editor.package().content_sha256(), before);

    for malformed in ["\"unclosed", "\"closed\"suffix", "unquoted\"quote"] {
        let error = editor
            .import_spreadsheet_delimited(
                "/Sheet1",
                NativeSpreadsheetDelimitedImport::new(
                    malformed,
                    NativeSpreadsheetDelimitedFormat::Csv,
                ),
            )
            .unwrap_err();
        assert_eq!(
            error.code,
            "use.office.spreadsheet_import_delimited_invalid"
        );
        assert_eq!(editor.package().content_sha256(), before);
    }
}

#[tokio::test]
async fn frozen_pane_has_a_typed_semantic_lifecycle_and_atomic_validation() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("freeze.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let path = editor
        .set_spreadsheet_frozen_pane("/Sheet1", NativeSpreadsheetFrozenPane::new(1, 2, "C2"))
        .unwrap();
    assert_eq!(path, "/Sheet1/freeze");
    let snapshot = editor.snapshot().unwrap();
    let pane = snapshot.get("/Sheet1/freeze", 0).unwrap();
    assert_eq!(pane.format["frozenRows"], "1");
    assert_eq!(pane.format["frozenColumns"], "2");
    assert_eq!(pane.format["topLeftCell"], "C2");
    assert_eq!(pane.format["activePane"], "bottomRight");
    assert_eq!(snapshot.query("frozen-pane").unwrap().len(), 1);

    let before = editor.package().content_sha256();
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetCellValue {
                path: "/Sheet1/A1".into(),
                value: text("must roll back"),
            },
            NativeOfficeMutation::SetSpreadsheetFrozenPane {
                sheet: "/Sheet1".into(),
                pane: NativeSpreadsheetFrozenPane::new(2, 0, "A2"),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_freeze_geometry_invalid");
    assert_eq!(editor.package().content_sha256(), before);
    assert!(editor.snapshot().unwrap().get("/Sheet1/A1", 0).is_err());

    editor
        .set_spreadsheet_frozen_pane("/Sheet1", NativeSpreadsheetFrozenPane::new(2, 0, "A3"))
        .unwrap();
    let pane = editor.snapshot().unwrap().get("/Sheet1/freeze", 0).unwrap();
    assert_eq!(pane.format["activePane"], "bottomLeft");
    editor.remove("/Sheet1/freeze").unwrap();
    assert!(editor.snapshot().unwrap().get("/Sheet1/freeze", 0).is_err());
}

#[tokio::test]
async fn import_preserves_strict_spreadsheetml_and_freeze_fails_closed_on_unknown_content() {
    const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-import.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let worksheet = std::str::from_utf8(package.part("xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .replace(TRANSITIONAL, STRICT);
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor
        .import_spreadsheet_delimited(
            "/Sheet1",
            NativeSpreadsheetDelimitedImport::new(
                "Date,Amount\n2026-07-17,42",
                NativeSpreadsheetDelimitedFormat::Csv,
            )
            .with_header(true),
        )
        .unwrap();
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains(STRICT));
    assert!(!worksheet.contains(TRANSITIONAL));
    assert!(worksheet.contains("<autoFilter ref=\"A1:B2\""));
    assert!(worksheet.contains("<pane ySplit=\"1\""));

    let worksheet = worksheet.replacen(
        " state=\"frozen\"",
        " xmlns:v=\"urn:vendor\" v:keep=\"true\" state=\"frozen\"",
        1,
    );
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/freeze", 0)
            .unwrap()
            .format["nativeMutable"],
        "false"
    );
    let before = editor.package().content_sha256();
    let error = editor
        .set_spreadsheet_frozen_pane("/Sheet1", NativeSpreadsheetFrozenPane::new(1, 0, "A2"))
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_freeze_unknown_content");
    assert_eq!(editor.package().content_sha256(), before);
    let error = editor.remove("/Sheet1/freeze").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_freeze_unknown_content");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn header_import_freeze_state_has_a_lifecycle_and_exact_replay() {
    let temp = tempfile::tempdir().unwrap();
    let source_path = temp.path().join("source.xlsx");
    let target_path = temp.path().join("target.xlsx");
    let mut source = NativeOfficeEditor::create(&source_path).await.unwrap();
    source
        .import_spreadsheet_delimited(
            "/Sheet1",
            NativeSpreadsheetDelimitedImport::new(
                "Date,Amount,Formula\n2026-07-17,42,=B2*2",
                NativeSpreadsheetDelimitedFormat::Csv,
            )
            .with_header(true),
        )
        .unwrap();

    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        NativeOfficeMutation::SetSpreadsheetFrozenPane { sheet, pane }
            if sheet == "/Sheet1" && pane.frozen_rows == 1
    )));
    let mut target = NativeOfficeEditor::create(&target_path).await.unwrap();
    target.apply_replay(&artifact).unwrap();
    assert_eq!(
        target.package().content_sha256(),
        source.package().content_sha256()
    );

    source.remove("/Sheet1/freeze").unwrap();
    assert!(source.snapshot().unwrap().get("/Sheet1/freeze", 0).is_err());
    assert_eq!(
        cell(&source, "/Sheet1/A2").format["numberFormat"],
        "yyyy-mm-dd"
    );
}
