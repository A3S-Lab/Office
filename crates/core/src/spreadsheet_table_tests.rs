use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficeReplayArtifact,
    NativeSpreadsheetNamedRange, NativeSpreadsheetTable, NativeSpreadsheetTableStyle,
    OfficeNodeType, SpreadsheetCellValue,
};

#[test]
fn spreadsheet_tables_have_a_closed_typed_batch_contract() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetTable>();
    assert_send_sync::<NativeOfficeMutation>();

    let mutation = NativeOfficeMutation::AddSpreadsheetTable {
        sheet: "/Sheet1".into(),
        table: NativeSpreadsheetTable::new("Sales", "A1:C4", ["Name", "Qty", "Price"])
            .with_totals_row(true)
            .with_style(NativeSpreadsheetTableStyle::Dark { number: 1 })
            .with_first_column(true),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "add-spreadsheet-table",
            "sheet": "/Sheet1",
            "table": {
                "name": "Sales",
                "range": "A1:C4",
                "columns": [
                    {"name": "Name"},
                    {"name": "Qty"},
                    {"name": "Price"}
                ],
                "headerRow": true,
                "totalsRow": true,
                "style": {"family": "dark", "number": 1},
                "showFirstColumn": true,
                "showLastColumn": false,
                "showRowStripes": true,
                "showColumnStripes": false
            }
        })
    );
}

#[tokio::test]
async fn spreadsheet_tables_have_a_native_typed_lifecycle() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("tables.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A2",
            SpreadsheetCellValue::Text {
                value: "Keyboard".into(),
            },
        )
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/B2",
            SpreadsheetCellValue::Number { value: "2".into() },
        )
        .unwrap();

    let table_path = editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C4", ["Name", "Qty", "Price"]),
        )
        .unwrap();
    assert_eq!(table_path, "/Sheet1/table[1]");
    let snapshot = editor.snapshot().unwrap();
    let table = snapshot.get(&table_path, 1).unwrap();
    assert_eq!(table.node_type, OfficeNodeType::Table);
    assert_eq!(table.format["name"], "Sales");
    assert_eq!(table.format["displayName"], "Sales");
    assert_eq!(table.format["ref"], "A1:C4");
    assert_eq!(table.format["headerRow"], "true");
    assert_eq!(table.format["totalsRow"], "false");
    assert_eq!(table.format["styleFamily"], "medium");
    assert_eq!(table.format["styleNumber"], "2");
    assert_eq!(table.format["nativeMutable"], "true");
    assert_eq!(table.children.len(), 4);
    assert_eq!(table.children[2].text, "Price");
    assert_eq!(table.children[3].node_type, OfficeNodeType::AutoFilter);
    assert_eq!(snapshot.get("/Sheet1/A1", 0).unwrap().text, "Name");
    assert_eq!(snapshot.query("table[name=Sales]").unwrap().len(), 1);
    assert!(editor.package().contains_part("xl/tables/table1.xml"));

    let updated_path = editor
        .set_spreadsheet_table(
            &table_path,
            NativeSpreadsheetTable::new("Inventory", "B2:D6", ["Item", "Units", "Cost"])
                .with_display_name("InventoryView")
                .with_totals_row(true)
                .with_style(NativeSpreadsheetTableStyle::Dark { number: 3 })
                .with_first_column(true)
                .with_row_stripes(false)
                .with_column_stripes(true),
        )
        .unwrap();
    assert_eq!(updated_path, table_path);
    let snapshot = editor.snapshot().unwrap();
    let table = snapshot.get(&table_path, 1).unwrap();
    assert_eq!(table.format["name"], "Inventory");
    assert_eq!(table.format["displayName"], "InventoryView");
    assert_eq!(table.format["ref"], "B2:D6");
    assert_eq!(table.format["totalsRow"], "true");
    assert_eq!(table.format["styleFamily"], "dark");
    assert_eq!(table.format["styleNumber"], "3");
    assert_eq!(table.format["showFirstColumn"], "true");
    assert_eq!(table.format["showRowStripes"], "false");
    assert_eq!(table.format["showColumnStripes"], "true");
    assert_eq!(snapshot.get("/Sheet1/B2", 0).unwrap().text, "Item");
    let table_xml =
        std::str::from_utf8(editor.package().part("xl/tables/table1.xml").unwrap()).unwrap();
    assert!(table_xml.contains("<autoFilter ref=\"B2:D5\"/>"));
    assert!(table_xml.contains("name=\"TableStyleDark3\""));

    editor.save().await.unwrap();
    let mut reopened = NativeOfficeEditor::open(&path).await.unwrap();
    assert_eq!(
        reopened
            .snapshot()
            .unwrap()
            .get(&table_path, 0)
            .unwrap()
            .format["displayName"],
        "InventoryView"
    );
    reopened.remove(&table_path).unwrap();
    assert!(!reopened.package().contains_part("xl/tables/table1.xml"));
    assert!(reopened
        .snapshot()
        .unwrap()
        .query("table")
        .unwrap()
        .is_empty());
    let worksheet =
        std::str::from_utf8(reopened.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(!worksheet.contains("tableParts"));
}

#[tokio::test]
async fn spreadsheet_table_paths_are_numbered_within_each_worksheet() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table-paths.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_worksheet("Data").unwrap();

    let first = editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("FirstTable", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let second_sheet_first = editor
        .add_spreadsheet_table(
            "/Data",
            NativeSpreadsheetTable::new("DataTable", "A1:B3", ["One", "Two"]),
        )
        .unwrap();

    assert_eq!(first, "/Sheet1/table[1]");
    assert_eq!(second_sheet_first, "/Data/table[1]");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get(&second_sheet_first, 0)
            .unwrap()
            .format["name"],
        "DataTable"
    );
}

#[tokio::test]
async fn spreadsheet_table_ids_include_nonstandard_part_locations() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table-ids.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    package
        .set_part(
            "xl/imported/list-object.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="41" name="Imported" displayName="Imported" ref="A1:A2" headerRowCount="1" totalsRowCount="0" totalsRowShown="0"><autoFilter ref="A1:A2"/><tableColumns count="1"><tableColumn id="1" name="Imported"/></tableColumns></table>"#
                .to_vec(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        "xl/imported/list-object.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Created", "A1:B3", ["One", "Two"]),
        )
        .unwrap();

    let table_xml =
        std::str::from_utf8(editor.package().part("xl/tables/table1.xml").unwrap()).unwrap();
    assert!(table_xml.contains("id=\"42\""));
}

#[tokio::test]
async fn spreadsheet_table_validation_is_atomic_and_cross_feature_aware() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table-validation.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_named_range(NativeSpreadsheetNamedRange::new(
            "ReservedTable",
            "'Sheet1'!$A$1",
        ))
        .unwrap();
    let before = editor.package().content_sha256();
    let collision = editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("ReservedTable", "A1:B3", ["One", "Two"]),
        )
        .unwrap_err();
    assert_eq!(
        collision.code,
        "use.office.spreadsheet_table_name_collision"
    );
    assert_eq!(editor.package().content_sha256(), before);

    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("FirstTable", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let after_first = editor.package().content_sha256();
    let overlap = editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("SecondTable", "B3:C5", ["Two", "Three"]),
        )
        .unwrap_err();
    assert_eq!(overlap.code, "use.office.spreadsheet_table_overlap");
    assert_eq!(editor.package().content_sha256(), after_first);

    for (table, code) in [
        (
            NativeSpreadsheetTable::new("A1", "D1:E3", ["D", "E"]),
            "use.office.spreadsheet_table_name_invalid",
        ),
        (
            NativeSpreadsheetTable::new("BadWidth", "D1:F3", ["D", "E"]),
            "use.office.spreadsheet_table_column_count_invalid",
        ),
        (
            NativeSpreadsheetTable::new("DuplicateColumns", "D1:E3", ["Name", "name"]),
            "use.office.spreadsheet_table_column_duplicate",
        ),
        (
            NativeSpreadsheetTable::new("NoData", "D1:E1", ["D", "E"]),
            "use.office.spreadsheet_table_data_rows_invalid",
        ),
    ] {
        let error = editor.add_spreadsheet_table("/Sheet1", table).unwrap_err();
        assert_eq!(error.code, code);
        assert_eq!(editor.package().content_sha256(), after_first);
    }
}

#[tokio::test]
async fn spreadsheet_tables_are_exactly_replayable() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table-replay.xlsx");
    let restored_path = temp.path().join("table-replay-restored.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A2",
            SpreadsheetCellValue::Text {
                value: "Mouse".into(),
            },
        )
        .unwrap();
    let table_path = editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Products", "A1:B4", ["Product", "Quantity"])
                .with_style(NativeSpreadsheetTableStyle::Light { number: 9 })
                .with_last_column(true),
        )
        .unwrap();
    editor
        .set_spreadsheet_table(
            &table_path,
            NativeSpreadsheetTable::new("Inventory", "B2:C5", ["Item", "Units"])
                .with_display_name("InventoryView")
                .with_totals_row(true)
                .with_style(NativeSpreadsheetTableStyle::Dark { number: 2 })
                .with_row_stripes(false)
                .with_column_stripes(true),
        )
        .unwrap();

    let document = editor.snapshot().unwrap();
    let artifact = NativeOfficeReplayArtifact::dump(&document, "/").unwrap();
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        NativeOfficeMutation::AddSpreadsheetTable { table, .. }
            if table.name == "Inventory"
                && table.display_name.as_deref() == Some("InventoryView")
                && table.range == "B2:C5"
    )));

    let expected_sha256 = editor.package().content_sha256();
    let mut restored = NativeOfficeEditor::create(&restored_path).await.unwrap();
    let result = restored.apply_replay(&artifact).unwrap();
    assert_eq!(result.applied, artifact.mutations.len());
    assert_eq!(restored.package().content_sha256(), expected_sha256);
}

#[tokio::test]
async fn spreadsheet_tables_preserve_strict_ooxml_and_unknown_owned_boundaries() {
    const TRANSITIONAL_SPREADSHEET: &str =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_SPREADSHEET: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const TRANSITIONAL_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-tables.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    for part_name in [
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
    ] {
        let bytes = package.part(part_name).unwrap();
        let xml = std::str::from_utf8(bytes)
            .unwrap()
            .replace(TRANSITIONAL_SPREADSHEET, STRICT_SPREADSHEET)
            .replace(TRANSITIONAL_RELATIONSHIPS, STRICT_RELATIONSHIPS);
        package.set_part(part_name, xml.into_bytes()).unwrap();
    }
    editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("StrictTable", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let table_xml =
        std::str::from_utf8(editor.package().part("xl/tables/table1.xml").unwrap()).unwrap();
    assert!(table_xml.contains(STRICT_SPREADSHEET));
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains(STRICT_RELATIONSHIPS));
    let relationships = std::str::from_utf8(
        editor
            .package()
            .part("xl/worksheets/_rels/sheet1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(relationships.contains(&format!("{STRICT_RELATIONSHIPS}/table")));

    let vendor_table = table_xml
        .replace("<table ", "<table xmlns:v=\"urn:vendor\" v:owner=\"keep\" ")
        .replace("<tableStyleInfo ", "<tableStyleInfo v:style=\"keep\" ")
        .replace("</table>", "<extLst><v:payload/></extLst></table>");
    let mut package = editor.package().clone();
    package
        .set_part("xl/tables/table1.xml", vendor_table.into_bytes())
        .unwrap();
    editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("StrictRenamed", "A1:B4", ["First", "Second"]),
        )
        .unwrap();
    let preserved =
        std::str::from_utf8(editor.package().part("xl/tables/table1.xml").unwrap()).unwrap();
    assert!(preserved.contains("v:owner=\"keep\""));
    assert!(preserved.contains("v:style=\"keep\""));
    assert!(preserved.contains("<extLst><v:payload/></extLst>"));
    assert!(preserved.contains(STRICT_SPREADSHEET));

    let unsupported = preserved.replacen("<tableColumn ", "<tableColumn v:column=\"keep\" ", 1);
    let mut package = editor.package().clone();
    package
        .set_part("xl/tables/table1.xml", unsupported.into_bytes())
        .unwrap();
    editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Blocked", "A1:B4", ["First", "Second"]),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_table_unknown_content");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn spreadsheet_tables_fail_closed_for_unsupported_imported_state() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("unsupported-table-state.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Imported", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let base = editor.package().clone();
    let table_xml = std::str::from_utf8(base.part("xl/tables/table1.xml").unwrap()).unwrap();
    let variants = [
        table_xml.replace(
            "<autoFilter ref=\"A1:B3\"/>",
            "<autoFilter ref=\"A1:B3\"><filterColumn colId=\"0\"><filters><dateGroupItem year=\"2026\" month=\"7\" dateTimeGrouping=\"month\"/></filters></filterColumn></autoFilter>",
        ),
        table_xml.replace(
            "<tableColumns",
            "<sortState ref=\"A2:B3\"><sortCondition ref=\"A2:A3\"/></sortState><tableColumns",
        ),
        table_xml.replace("<table ", "<table tableType=\"queryTable\" "),
        table_xml.replace("<table ", "<table connectionId=\"7\" "),
        table_xml.replace("<autoFilter ref=\"A1:B3\"/>", ""),
        table_xml.replace("<autoFilter ref=\"A1:B3\"/>", "<autoFilter ref=\"A1:A3\"/>"),
        table_xml.replace(
            "totalsRowCount=\"0\" totalsRowShown=\"0\"",
            "totalsRowCount=\"1\" totalsRowShown=\"0\"",
        ),
    ];

    for table_xml in variants {
        let mut package = base.clone();
        package
            .set_part("xl/tables/table1.xml", table_xml.into_bytes())
            .unwrap();
        let mut imported = NativeOfficeEditor::from_package(package).unwrap();
        assert_eq!(
            imported
                .snapshot()
                .unwrap()
                .get("/Sheet1/table[1]", 0)
                .unwrap()
                .format["nativeMutable"],
            "false"
        );
        let before = imported.package().content_sha256();
        let error = imported
            .set_spreadsheet_table(
                "/Sheet1/table[1]",
                NativeSpreadsheetTable::new("Blocked", "C1:D3", ["Three", "Four"]),
            )
            .unwrap_err();
        assert_eq!(error.code, "use.office.spreadsheet_table_unknown_content");
        assert_eq!(imported.package().content_sha256(), before);
    }
}

#[tokio::test]
async fn spreadsheet_tables_reject_invalid_collection_counts() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("invalid-table-count.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Counted", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let worksheet = std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .replace("<tableParts count=\"1\">", "<tableParts count=\"invalid\">");
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();

    let error = NativeOfficeEditor::from_package(package).unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_table_invalid");
}

#[tokio::test]
async fn spreadsheet_table_set_and_remove_reject_owned_relationships() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("related-table.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Related", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let mut package = editor.package().clone();
    crate::opc_edit::add_external_relationship(
        &mut package,
        "xl/tables/_rels/table1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
        "https://example.com/query",
    )
    .unwrap();
    let mut imported = NativeOfficeEditor::from_package(package).unwrap();

    for mutation in ["set", "remove"] {
        let before = imported.package().content_sha256();
        let error = if mutation == "set" {
            imported
                .set_spreadsheet_table(
                    "/Sheet1/table[1]",
                    NativeSpreadsheetTable::new("Blocked", "C1:D3", ["Three", "Four"]),
                )
                .unwrap_err()
        } else {
            imported.remove("/Sheet1/table[1]").unwrap_err()
        };
        assert_eq!(
            error.code,
            "use.office.spreadsheet_table_relationships_unsupported"
        );
        assert_eq!(imported.package().content_sha256(), before);
    }

    let mut package = editor.package().clone();
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/_rels/workbook.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        "tables/table1.xml",
    )
    .unwrap();
    let mut imported = NativeOfficeEditor::from_package(package).unwrap();
    let before = imported.package().content_sha256();
    let error = imported
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Blocked", "C1:D3", ["Three", "Four"]),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_table_relationships_unsupported"
    );
    assert_eq!(imported.package().content_sha256(), before);
}

#[tokio::test]
async fn spreadsheet_table_final_removal_fails_closed_for_unknown_collection_data() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("unknown-table-parts.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("OnlyTable", "A1:B3", ["One", "Two"]),
        )
        .unwrap();
    let worksheet = std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .replace(
            "<tableParts count=\"1\">",
            "<tableParts xmlns:v=\"urn:vendor\" v:owner=\"keep\" count=\"1\">",
        );
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();
    let error = editor.remove("/Sheet1/table[1]").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_table_unknown_content");
    assert_eq!(editor.package().content_sha256(), before);
}
