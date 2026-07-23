use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficeReplayArtifact,
    NativeSpreadsheetAutoFilter, NativeSpreadsheetDynamicFilter, NativeSpreadsheetFilterColumn,
    NativeSpreadsheetFilterCriteria, NativeSpreadsheetTable, OfficeNodeType, SpreadsheetCellValue,
};

#[test]
fn spreadsheet_filters_have_a_closed_typed_contract() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetAutoFilter>();
    assert_send_sync::<NativeSpreadsheetFilterCriteria>();

    let mutation = NativeOfficeMutation::AddSpreadsheetAutoFilter {
        sheet: "/Sheet1".into(),
        filter: NativeSpreadsheetAutoFilter::new("A1:C20")
            .with_filter(
                0,
                NativeSpreadsheetFilterCriteria::Values {
                    values: vec!["Open".into(), "Closed".into()],
                    include_blanks: true,
                },
            )
            .with_filter(
                2,
                NativeSpreadsheetFilterCriteria::GreaterThan {
                    value: "100".into(),
                },
            ),
    };
    assert_eq!(
        serde_json::to_value(mutation).unwrap(),
        serde_json::json!({
            "operation": "add-spreadsheet-auto-filter",
            "sheet": "/Sheet1",
            "filter": {
                "range": "A1:C20",
                "columns": [
                    {
                        "column": 0,
                        "criteria": {
                            "type": "values",
                            "values": ["Open", "Closed"],
                            "includeBlanks": true
                        }
                    },
                    {
                        "column": 2,
                        "criteria": {"type": "greater-than", "value": "100"}
                    }
                ]
            }
        })
    );
}

#[tokio::test]
async fn every_typed_filter_criterion_round_trips_semantically() {
    let criteria = vec![
        NativeSpreadsheetFilterCriteria::Values {
            values: vec!["Open".into(), "Closed".into()],
            include_blanks: true,
        },
        NativeSpreadsheetFilterCriteria::Equals {
            value: "exact*?~".into(),
        },
        NativeSpreadsheetFilterCriteria::NotEquals {
            value: "excluded".into(),
        },
        NativeSpreadsheetFilterCriteria::Contains {
            value: "middle".into(),
        },
        NativeSpreadsheetFilterCriteria::DoesNotContain {
            value: "blocked".into(),
        },
        NativeSpreadsheetFilterCriteria::BeginsWith {
            value: "prefix".into(),
        },
        NativeSpreadsheetFilterCriteria::EndsWith {
            value: "suffix".into(),
        },
        NativeSpreadsheetFilterCriteria::GreaterThan { value: "10".into() },
        NativeSpreadsheetFilterCriteria::GreaterThanOrEqual { value: "20".into() },
        NativeSpreadsheetFilterCriteria::LessThan { value: "90".into() },
        NativeSpreadsheetFilterCriteria::LessThanOrEqual { value: "80".into() },
        NativeSpreadsheetFilterCriteria::Between {
            lower: "30".into(),
            upper: "70".into(),
        },
        NativeSpreadsheetFilterCriteria::NotBetween {
            lower: "40".into(),
            upper: "60".into(),
        },
        NativeSpreadsheetFilterCriteria::Blanks,
        NativeSpreadsheetFilterCriteria::NonBlanks,
        NativeSpreadsheetFilterCriteria::Top { count: 10 },
        NativeSpreadsheetFilterCriteria::TopPercent { percent: 20 },
        NativeSpreadsheetFilterCriteria::Bottom { count: 5 },
        NativeSpreadsheetFilterCriteria::BottomPercent { percent: 15 },
        NativeSpreadsheetFilterCriteria::Dynamic {
            kind: NativeSpreadsheetDynamicFilter::AboveAverage,
        },
    ];
    let expected = NativeSpreadsheetAutoFilter {
        range: "A1:T10".into(),
        columns: criteria
            .into_iter()
            .enumerate()
            .map(|(column, criteria)| {
                NativeSpreadsheetFilterColumn::new(u32::try_from(column).unwrap(), criteria)
            })
            .collect(),
    };
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("all-filter-criteria.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let filter_path = editor
        .add_spreadsheet_auto_filter("/Sheet1", expected.clone())
        .unwrap();
    let node = editor.snapshot().unwrap().get(&filter_path, 2).unwrap();
    assert_eq!(
        NativeSpreadsheetAutoFilter::from_semantic_node(&node).unwrap(),
        expected
    );
}

#[tokio::test]
async fn worksheet_auto_filters_have_a_native_typed_lifecycle() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("worksheet-filter.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1",
            SpreadsheetCellValue::Text {
                value: "Status".into(),
            },
        )
        .unwrap();

    let filter_path = editor
        .add_spreadsheet_auto_filter(
            "/Sheet1",
            NativeSpreadsheetAutoFilter::new("A1:C20")
                .with_filter(
                    0,
                    NativeSpreadsheetFilterCriteria::Values {
                        values: vec!["Open".into(), "Closed".into()],
                        include_blanks: true,
                    },
                )
                .with_filter(
                    1,
                    NativeSpreadsheetFilterCriteria::Contains {
                        value: "east*?~".into(),
                    },
                )
                .with_filter(
                    2,
                    NativeSpreadsheetFilterCriteria::Between {
                        lower: "10".into(),
                        upper: "100".into(),
                    },
                ),
        )
        .unwrap();
    assert_eq!(filter_path, "/Sheet1/autofilter");
    let snapshot = editor.snapshot().unwrap();
    let filter = snapshot.get(&filter_path, 2).unwrap();
    assert_eq!(filter.node_type, OfficeNodeType::AutoFilter);
    assert_eq!(filter.format["ref"], "A1:C20");
    assert_eq!(filter.format["nativeMutable"], "true");
    assert_eq!(filter.children[0].format["criteriaType"], "values");
    assert_eq!(filter.children[0].children[1].text, "Closed");
    assert_eq!(filter.children[1].format["criteriaType"], "contains");
    assert_eq!(snapshot.query("autofilter").unwrap().len(), 1);
    assert_eq!(
        snapshot
            .query("filtercolumn[criteriaType=between]")
            .unwrap()
            .len(),
        1
    );
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet
        .contains("<filters blank=\"1\"><filter val=\"Open\"/><filter val=\"Closed\"/></filters>"));
    assert!(worksheet.contains("val=\"*east~*~?~~*\""));
    assert!(worksheet.contains("operator=\"greaterThanOrEqual\" val=\"10\""));

    editor
        .set_spreadsheet_auto_filter(
            &filter_path,
            NativeSpreadsheetAutoFilter::new("B2:D30")
                .with_filter(0, NativeSpreadsheetFilterCriteria::Blanks)
                .with_filter(
                    1,
                    NativeSpreadsheetFilterCriteria::TopPercent { percent: 10 },
                )
                .with_filter(
                    2,
                    NativeSpreadsheetFilterCriteria::Dynamic {
                        kind: NativeSpreadsheetDynamicFilter::ThisMonth,
                    },
                ),
        )
        .unwrap();
    let updated = editor.snapshot().unwrap().get(&filter_path, 1).unwrap();
    assert_eq!(updated.format["ref"], "B2:D30");
    assert_eq!(updated.children[2].format["dynamicKind"], "thisMonth");

    editor.save().await.unwrap();
    let mut reopened = NativeOfficeEditor::open(&path).await.unwrap();
    assert_eq!(
        reopened
            .snapshot()
            .unwrap()
            .get(&filter_path, 0)
            .unwrap()
            .format["ref"],
        "B2:D30"
    );
    reopened.remove(&filter_path).unwrap();
    assert!(reopened
        .snapshot()
        .unwrap()
        .query("autofilter")
        .unwrap()
        .is_empty());
}

#[tokio::test]
async fn table_auto_filters_share_the_typed_filter_value() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table-filter.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let table_path = editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C20", ["Region", "Status", "Amount"])
                .with_filter(
                    0,
                    NativeSpreadsheetFilterCriteria::EndsWith {
                        value: "West".into(),
                    },
                )
                .with_filter(2, NativeSpreadsheetFilterCriteria::Bottom { count: 5 }),
        )
        .unwrap();
    let snapshot = editor.snapshot().unwrap();
    let table = snapshot.get(&table_path, 2).unwrap();
    assert_eq!(table.format["nativeMutable"], "true");
    let filter = table
        .children
        .iter()
        .find(|child| child.node_type == OfficeNodeType::AutoFilter)
        .unwrap();
    assert_eq!(filter.path, "/Sheet1/table[1]/autofilter");
    assert_eq!(filter.children[0].format["criteriaType"], "ends-with");
    let reconstructed = NativeSpreadsheetTable::from_semantic_node(&table).unwrap();
    assert_eq!(reconstructed.filters.len(), 2);
    assert_eq!(reconstructed.filters[1].column, 2);
    let table_xml =
        std::str::from_utf8(editor.package().part("xl/tables/table1.xml").unwrap()).unwrap();
    assert!(table_xml.contains("<autoFilter ref=\"A1:C20\"><filterColumn colId=\"0\">"));
    assert!(table_xml.contains("<top10 percent=\"0\" top=\"0\" val=\"5\"/>"));
}

#[tokio::test]
async fn spreadsheet_filter_validation_and_geometry_are_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("filter-validation.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Reserved", "A1:B4", ["One", "Two"]),
        )
        .unwrap();
    let before = editor.package().content_sha256();
    let overlap = editor
        .add_spreadsheet_auto_filter("/Sheet1", NativeSpreadsheetAutoFilter::new("B1:C10"))
        .unwrap_err();
    assert_eq!(overlap.code, "use.office.spreadsheet_filter_table_overlap");
    assert_eq!(editor.package().content_sha256(), before);

    let invalid = [
        NativeSpreadsheetAutoFilter::new("D1:E10").with_filter(
            2,
            NativeSpreadsheetFilterCriteria::Equals { value: "x".into() },
        ),
        NativeSpreadsheetAutoFilter {
            range: "D1:E10".into(),
            columns: vec![
                NativeSpreadsheetFilterColumn::new(0, NativeSpreadsheetFilterCriteria::Blanks),
                NativeSpreadsheetFilterColumn::new(0, NativeSpreadsheetFilterCriteria::NonBlanks),
            ],
        },
        NativeSpreadsheetAutoFilter::new("D1:E10").with_filter(
            0,
            NativeSpreadsheetFilterCriteria::TopPercent { percent: 0 },
        ),
    ];
    let codes = [
        "use.office.spreadsheet_filter_column_invalid",
        "use.office.spreadsheet_filter_column_duplicate",
        "use.office.spreadsheet_filter_percent_invalid",
    ];
    for (filter, code) in invalid.into_iter().zip(codes) {
        let error = editor
            .add_spreadsheet_auto_filter("/Sheet1", filter)
            .unwrap_err();
        assert_eq!(error.code, code);
        assert_eq!(editor.package().content_sha256(), before);
    }

    let error = editor
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Reserved", "A1:B4", ["One", "Two"])
                .with_header_row(false)
                .with_filter(0, NativeSpreadsheetFilterCriteria::Blanks),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_table_filter_header_required"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn imported_filter_extensions_and_sort_state_fail_closed() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("unsupported-filter.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_auto_filter(
            "/Sheet1",
            NativeSpreadsheetAutoFilter::new("A1:B10")
                .with_filter(0, NativeSpreadsheetFilterCriteria::Blanks),
        )
        .unwrap();
    let worksheet = std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .replace(
            "<filters blank=\"1\"/>",
            "<filters><dateGroupItem year=\"2026\" month=\"7\" dateTimeGrouping=\"month\"/></filters>",
        )
        .replace(
            "</autoFilter>",
            "<sortState ref=\"A2:B10\"><sortCondition ref=\"A2:A10\"/></sortState></autoFilter>",
        );
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut imported = NativeOfficeEditor::from_package(package).unwrap();
    let node = imported
        .snapshot()
        .unwrap()
        .get("/Sheet1/autofilter", 2)
        .unwrap();
    assert_eq!(node.format["nativeMutable"], "false");
    assert_eq!(node.children[0].format["criteriaType"], "unsupported");
    let before = imported.package().content_sha256();
    let error = imported
        .set_spreadsheet_auto_filter(
            "/Sheet1/autofilter",
            NativeSpreadsheetAutoFilter::new("C1:D10"),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_filter_unknown_content");
    assert_eq!(imported.package().content_sha256(), before);
    let error = imported.remove("/Sheet1/autofilter").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_filter_unknown_content");
    assert_eq!(imported.package().content_sha256(), before);
}

#[tokio::test]
async fn worksheet_auto_filters_preserve_strict_spreadsheetml() {
    const TRANSITIONAL_SPREADSHEET: &str =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_SPREADSHEET: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const TRANSITIONAL_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-filter.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    for part_name in [
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
    ] {
        let xml = std::str::from_utf8(package.part(part_name).unwrap())
            .unwrap()
            .replace(TRANSITIONAL_SPREADSHEET, STRICT_SPREADSHEET)
            .replace(TRANSITIONAL_RELATIONSHIPS, STRICT_RELATIONSHIPS);
        package.set_part(part_name, xml.into_bytes()).unwrap();
    }
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .add_spreadsheet_auto_filter(
            "/Sheet1",
            NativeSpreadsheetAutoFilter::new("A1:B20").with_filter(
                1,
                NativeSpreadsheetFilterCriteria::DoesNotContain {
                    value: "draft".into(),
                },
            ),
        )
        .unwrap();
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains(STRICT_SPREADSHEET));
    assert!(worksheet.contains("<autoFilter ref=\"A1:B20\">"));
    assert!(worksheet.find("<autoFilter").unwrap() < worksheet.find("<pageMargins").unwrap());

    editor
        .set_spreadsheet_auto_filter(
            "/Sheet1/autofilter",
            NativeSpreadsheetAutoFilter::new("C1:D30").with_filter(
                0,
                NativeSpreadsheetFilterCriteria::Dynamic {
                    kind: NativeSpreadsheetDynamicFilter::AboveAverage,
                },
            ),
        )
        .unwrap();
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains(STRICT_SPREADSHEET));
    assert!(worksheet.contains("<dynamicFilter type=\"aboveAverage\"/>"));
    editor.remove("/Sheet1/autofilter").unwrap();
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains(STRICT_SPREADSHEET));
    assert!(!worksheet.contains("<autoFilter"));
}

#[tokio::test]
async fn table_date_color_icon_and_sort_filters_fail_closed() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("unsupported-table-filters.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Imported", "A1:B10", ["One", "Two"])
                .with_filter(0, NativeSpreadsheetFilterCriteria::Blanks),
        )
        .unwrap();
    let original = editor.package().clone();
    let table_xml = std::str::from_utf8(original.part("xl/tables/table1.xml").unwrap()).unwrap();
    let unsupported = [
        (
            "date-group",
            table_xml.replace(
                "<filters blank=\"1\"/>",
                "<filters><dateGroupItem year=\"2026\" month=\"7\" dateTimeGrouping=\"month\"/></filters>",
            ),
        ),
        (
            "color",
            table_xml.replace("<filters blank=\"1\"/>", "<colorFilter dxfId=\"0\"/>"),
        ),
        (
            "icon",
            table_xml.replace(
                "<filters blank=\"1\"/>",
                "<iconFilter iconSet=\"3Arrows\" iconId=\"0\"/>",
            ),
        ),
        (
            "sort-state",
            table_xml.replace(
                "</autoFilter>",
                "<sortState ref=\"A2:B10\"><sortCondition ref=\"A2:A10\"/></sortState></autoFilter>",
            ),
        ),
    ];

    for (case, xml) in unsupported {
        let mut package = original.clone();
        package
            .set_part("xl/tables/table1.xml", xml.into_bytes())
            .unwrap();
        let mut imported = NativeOfficeEditor::from_package(package).unwrap();
        let node = imported
            .snapshot()
            .unwrap()
            .get("/Sheet1/table[1]", 2)
            .unwrap();
        assert_eq!(node.format["nativeMutable"], "false", "{case}");
        let filter = node
            .children
            .iter()
            .find(|child| child.node_type == OfficeNodeType::AutoFilter)
            .unwrap();
        assert_eq!(filter.format["nativeMutable"], "false", "{case}");
        let before = imported.package().content_sha256();
        let error = imported
            .set_spreadsheet_table(
                "/Sheet1/table[1]",
                NativeSpreadsheetTable::new("Blocked", "A1:B10", ["One", "Two"]),
            )
            .unwrap_err();
        assert_eq!(
            error.code, "use.office.spreadsheet_table_unknown_content",
            "{case}"
        );
        assert_eq!(imported.package().content_sha256(), before, "{case}");
    }
}

#[tokio::test]
async fn unknown_filter_attributes_and_comments_roll_back() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("unknown-filter-content.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_auto_filter(
            "/Sheet1",
            NativeSpreadsheetAutoFilter::new("A1:B10")
                .with_filter(0, NativeSpreadsheetFilterCriteria::Blanks),
        )
        .unwrap();
    let original = editor.package().clone();
    let worksheet =
        std::str::from_utf8(original.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let unsupported = [
        worksheet.replace(
            "<autoFilter ",
            "<autoFilter xmlns:v=\"urn:vendor\" v:owner=\"keep\" ",
        ),
        worksheet.replace(
            "<filters blank=\"1\"/>",
            "<!--vendor-comment--><filters blank=\"1\"/>",
        ),
    ];

    for xml in unsupported {
        let mut package = original.clone();
        package
            .set_part("xl/worksheets/sheet1.xml", xml.into_bytes())
            .unwrap();
        let mut imported = NativeOfficeEditor::from_package(package).unwrap();
        let before = imported.package().content_sha256();
        let error = imported
            .set_spreadsheet_auto_filter(
                "/Sheet1/autofilter",
                NativeSpreadsheetAutoFilter::new("C1:D20"),
            )
            .unwrap_err();
        assert_eq!(error.code, "use.office.spreadsheet_filter_unknown_content");
        assert_eq!(imported.package().content_sha256(), before);
        let error = imported.remove("/Sheet1/autofilter").unwrap_err();
        assert_eq!(error.code, "use.office.spreadsheet_filter_unknown_content");
        assert_eq!(imported.package().content_sha256(), before);
    }
}

#[tokio::test]
async fn spreadsheet_auto_filters_are_exactly_replayable() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("filter-replay.xlsx");
    let restored_path = temp.path().join("filter-replay-restored.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_worksheet("Data").unwrap();
    editor
        .add_spreadsheet_auto_filter(
            "/Sheet1",
            NativeSpreadsheetAutoFilter::new("A1:B20")
                .with_filter(0, NativeSpreadsheetFilterCriteria::NonBlanks),
        )
        .unwrap();
    editor
        .add_spreadsheet_table(
            "/Data",
            NativeSpreadsheetTable::new("Products", "A1:B10", ["Name", "Price"]).with_filter(
                1,
                NativeSpreadsheetFilterCriteria::TopPercent { percent: 20 },
            ),
        )
        .unwrap();
    let document = editor.snapshot().unwrap();
    let artifact = NativeOfficeReplayArtifact::dump(&document, "/").unwrap();
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        NativeOfficeMutation::AddSpreadsheetAutoFilter { filter, .. }
            if filter.columns.len() == 1
    )));
    let expected = editor.package().content_sha256();
    let mut restored = NativeOfficeEditor::create(&restored_path).await.unwrap();
    restored.apply_replay(&artifact).unwrap();
    assert_eq!(restored.package().content_sha256(), expected);
}
