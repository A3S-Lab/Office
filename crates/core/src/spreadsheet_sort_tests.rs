use crate::{
    NativeOfficeComment, NativeOfficeEditor, NativeOfficeHyperlink, NativeOfficeImage,
    NativeOfficeMutation, NativeOfficePartType, NativeOfficeRgbColor, NativeSpreadsheetAutoFilter,
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetDataValidation,
    NativeSpreadsheetDataValidationType, NativeSpreadsheetDifferentialFormat,
    NativeSpreadsheetSort, NativeSpreadsheetSortDirection, NativeSpreadsheetSortKey,
    NativeSpreadsheetTable, OfficeNodeType, SpreadsheetCellValue,
};

fn text(value: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Text {
        value: value.to_string(),
    }
}

fn number(value: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Number {
        value: value.to_string(),
    }
}

fn cell(editor: &NativeOfficeEditor, path: &str) -> String {
    editor.snapshot().unwrap().get(path, 0).unwrap().text
}

#[test]
fn spreadsheet_sort_has_a_closed_typed_batch_contract() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetSort>();
    assert_send_sync::<NativeSpreadsheetSortKey>();

    let mutation = NativeOfficeMutation::SortSpreadsheetRange {
        path: "/Sheet1/A1:D100".into(),
        sort: NativeSpreadsheetSort::new(vec![
            NativeSpreadsheetSortKey::descending("B"),
            NativeSpreadsheetSortKey::ascending("C"),
        ])
        .with_header(true)
        .with_case_sensitive(true),
    };
    assert_eq!(
        serde_json::to_value(mutation).unwrap(),
        serde_json::json!({
            "operation": "sort-spreadsheet-range",
            "path": "/Sheet1/A1:D100",
            "sort": {
                "keys": [
                    {"column": "B", "direction": "descending"},
                    {"column": "C", "direction": "ascending"}
                ],
                "header": true,
                "caseSensitive": true
            }
        })
    );
}

#[tokio::test]
async fn sort_is_stable_multi_key_partial_range_and_metadata_has_a_lifecycle() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("stable-sort.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for (path, value) in [
        ("/Sheet1/A1", "ID"),
        ("/Sheet1/B1", "Score"),
        ("/Sheet1/C1", "Name"),
        ("/Sheet1/D1", "Fixed"),
        ("/Sheet1/A2", "first"),
        ("/Sheet1/C2", "Alpha"),
        ("/Sheet1/D2", "slot-2"),
        ("/Sheet1/A3", "second"),
        ("/Sheet1/C3", "alpha"),
        ("/Sheet1/D3", "slot-3"),
        ("/Sheet1/A4", "third"),
        ("/Sheet1/C4", "Beta"),
        ("/Sheet1/D4", "slot-4"),
        ("/Sheet1/A5", "fourth"),
        ("/Sheet1/C5", "Zulu"),
        ("/Sheet1/D5", "slot-5"),
    ] {
        editor.set_cell_value(path, text(value)).unwrap();
    }
    for (path, value) in [
        ("/Sheet1/B2", "20"),
        ("/Sheet1/B3", "20"),
        ("/Sheet1/B4", "20"),
        ("/Sheet1/B5", "10"),
    ] {
        editor.set_cell_value(path, number(value)).unwrap();
    }

    let request = NativeSpreadsheetSort::new(vec![
        NativeSpreadsheetSortKey::descending("B"),
        NativeSpreadsheetSortKey::ascending("C"),
    ])
    .with_header(true);
    let sort_path = editor
        .sort_spreadsheet_range("/Sheet1/A1:C5", request.clone())
        .unwrap();
    assert_eq!(sort_path, "/Sheet1/sort");
    assert_eq!(cell(&editor, "/Sheet1/A1"), "ID");
    assert_eq!(cell(&editor, "/Sheet1/A2"), "first");
    assert_eq!(cell(&editor, "/Sheet1/A3"), "second");
    assert_eq!(cell(&editor, "/Sheet1/A4"), "third");
    assert_eq!(cell(&editor, "/Sheet1/A5"), "fourth");
    assert_eq!(cell(&editor, "/Sheet1/D2"), "slot-2");
    assert_eq!(cell(&editor, "/Sheet1/D3"), "slot-3");
    assert_eq!(cell(&editor, "/Sheet1/D4"), "slot-4");
    assert_eq!(cell(&editor, "/Sheet1/D5"), "slot-5");

    let state = editor.snapshot().unwrap().get(&sort_path, 1).unwrap();
    assert_eq!(state.node_type, OfficeNodeType::SortState);
    assert_eq!(state.format["ref"], "A1:C5");
    assert_eq!(state.format["header"], "true");
    assert_eq!(state.format["caseSensitive"], "false");
    assert_eq!(state.format["nativeMutable"], "true");
    assert_eq!(state.children.len(), 2);
    assert_eq!(state.children[0].path, "/Sheet1/sort/key[1]");
    assert_eq!(state.children[0].format["ref"], "B2:B5");
    assert_eq!(state.children[0].format["direction"], "descending");
    assert_eq!(state.children[1].format["ref"], "C2:C5");
    assert_eq!(
        NativeSpreadsheetSort::from_semantic_node(&state).unwrap(),
        request
    );

    editor.save().await.unwrap();
    let mut reopened = NativeOfficeEditor::open(&path).await.unwrap();
    assert_eq!(
        reopened
            .snapshot()
            .unwrap()
            .get(&sort_path, 1)
            .unwrap()
            .format["ref"],
        "A1:C5"
    );
    reopened.remove(&sort_path).unwrap();
    assert!(reopened.snapshot().unwrap().get(&sort_path, 0).is_err());
    assert_eq!(cell(&reopened, "/Sheet1/A2"), "first");
}

#[tokio::test]
async fn sort_orders_numbers_before_text_and_keeps_blanks_last_in_both_directions() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-types.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for (row, id) in [
        (1, "ten"),
        (2, "text-two"),
        (3, "blank-one"),
        (4, "minus"),
        (5, "apple"),
        (6, "blank-two"),
    ] {
        editor
            .set_cell_value(format!("/Sheet1/A{row}"), text(id))
            .unwrap();
    }
    editor.set_cell_value("/Sheet1/B1", number("10")).unwrap();
    editor.set_cell_value("/Sheet1/B2", text("2")).unwrap();
    editor.set_cell_value("/Sheet1/B4", number("-1")).unwrap();
    editor.set_cell_value("/Sheet1/B5", text("apple")).unwrap();

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:B6",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")]),
        )
        .unwrap();
    assert_eq!(
        (1..=6)
            .map(|row| cell(&editor, &format!("/Sheet1/A{row}")))
            .collect::<Vec<_>>(),
        [
            "minus",
            "ten",
            "text-two",
            "apple",
            "blank-one",
            "blank-two"
        ]
    );

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:B6",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::descending("B")]),
        )
        .unwrap();
    assert_eq!(
        (1..=6)
            .map(|row| cell(&editor, &format!("/Sheet1/A{row}")))
            .collect::<Vec<_>>(),
        [
            "ten",
            "minus",
            "apple",
            "text-two",
            "blank-one",
            "blank-two"
        ]
    );
}

#[tokio::test]
async fn sort_materializes_sparse_destination_rows_without_moving_other_columns() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sparse-sort.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A2", text("two")).unwrap();
    editor.set_cell_value("/Sheet1/B2", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/D2", text("fixed-two"))
        .unwrap();
    editor.set_cell_value("/Sheet1/A4", text("four")).unwrap();
    editor.set_cell_value("/Sheet1/B4", number("1")).unwrap();
    editor
        .set_cell_value("/Sheet1/D4", text("fixed-four"))
        .unwrap();

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A2:B4",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")]),
        )
        .unwrap();
    assert_eq!(cell(&editor, "/Sheet1/A2"), "four");
    assert_eq!(cell(&editor, "/Sheet1/A3"), "two");
    assert!(editor.snapshot().unwrap().get("/Sheet1/A4", 0).is_err());
    assert_eq!(cell(&editor, "/Sheet1/D2"), "fixed-two");
    assert_eq!(cell(&editor, "/Sheet1/D4"), "fixed-four");
    editor.save().await.unwrap();
    NativeOfficeEditor::open(&path).await.unwrap();
}

#[tokio::test]
async fn sort_updates_the_used_dimension_after_sparse_rows_move() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-dimension.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A3", number("1")).unwrap();

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:A3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
        )
        .unwrap();

    assert_eq!(cell(&editor, "/Sheet1/A1"), "1");
    assert!(editor.snapshot().unwrap().get("/Sheet1/A3", 0).is_err());
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains("<dimension ref=\"A1\"/>"));
}

#[tokio::test]
async fn sort_rejects_formulas_and_merged_cells_and_batch_failure_rolls_back() {
    let temp = tempfile::tempdir().unwrap();
    let formula_path = temp.path().join("formula-sort.xlsx");
    let mut formula = NativeOfficeEditor::create(&formula_path).await.unwrap();
    formula.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    formula
        .set_cell_value(
            "/Sheet1/B1",
            SpreadsheetCellValue::Formula {
                expression: "A1*2".into(),
            },
        )
        .unwrap();
    let error = formula
        .sort_spreadsheet_range(
            "/Sheet1/A1:B1",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_sort_formula_unsupported"
    );

    let merged_path = temp.path().join("merged-sort.xlsx");
    let mut merged = NativeOfficeEditor::create(&merged_path).await.unwrap();
    merged.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    merged.set_cell_value("/Sheet1/A2", number("1")).unwrap();
    merged.merge_cells("/Sheet1/A1:B1").unwrap();
    let error = merged
        .sort_spreadsheet_range(
            "/Sheet1/A1:B2",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_sort_merge_overlap");

    let rollback_path = temp.path().join("rollback-sort.xlsx");
    let mut rollback = NativeOfficeEditor::create(&rollback_path).await.unwrap();
    rollback.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    rollback.set_cell_value("/Sheet1/A2", number("1")).unwrap();
    let before = rollback.package().content_sha256();
    let error = rollback
        .apply_batch(&[
            NativeOfficeMutation::SortSpreadsheetRange {
                path: "/Sheet1/A1:A2".into(),
                sort: NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
            },
            NativeOfficeMutation::SetText {
                path: "/Missing/A1".into(),
                text: "fail".into(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.node_not_found");
    assert_eq!(rollback.package().content_sha256(), before);
    assert_eq!(cell(&rollback, "/Sheet1/A1"), "2");
}

#[tokio::test]
async fn sort_moves_hyperlinks_comments_validations_and_conditional_formats_with_records() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-sidecars.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for (path, value) in [
        ("/Sheet1/A1", "Key"),
        ("/Sheet1/B1", "Record"),
        ("/Sheet1/C1", "Rule"),
        ("/Sheet1/B2", "two"),
        ("/Sheet1/B3", "one"),
    ] {
        editor.set_cell_value(path, text(value)).unwrap();
    }
    editor.set_cell_value("/Sheet1/A2", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/A3", number("1")).unwrap();
    editor
        .set_hyperlink(
            "/Sheet1/B2",
            NativeOfficeHyperlink::external("https://example.com/two")
                .unwrap()
                .with_display("two"),
        )
        .unwrap();
    editor
        .add_comment(
            "/Sheet1/B2",
            NativeOfficeComment::new("Alice", "Moves with two").unwrap(),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "C2",
                "Yes,No",
            ),
        )
        .unwrap();
    editor
        .add_conditional_format(
            "/Sheet1",
            NativeSpreadsheetConditionalFormat::new(
                "C2",
                NativeSpreadsheetConditionalFormatRule::CellIs {
                    operator: NativeSpreadsheetConditionalFormatOperator::GreaterThan,
                    formula1: "0".into(),
                    formula2: None,
                    format: NativeSpreadsheetDifferentialFormat::default()
                        .with_fill(NativeOfficeRgbColor::new(198, 239, 206)),
                },
            ),
        )
        .unwrap();

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:C3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")])
                .with_header(true),
        )
        .unwrap();
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get("/Sheet1/B3/comment", 0).unwrap().text,
        "Moves with two"
    );
    assert!(snapshot.get("/Sheet1/B2/comment", 0).is_err());
    assert_eq!(
        snapshot.get("/Sheet1/B3/hyperlink", 0).unwrap().format["target"],
        "https://example.com/two"
    );
    assert!(snapshot.get("/Sheet1/B2/hyperlink", 0).is_err());
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains("<dataValidation"));
    assert!(worksheet.contains("sqref=\"C3\""));
    assert!(worksheet.contains("<conditionalFormatting sqref=\"C3\""));
    let comments_part = editor
        .package()
        .part_names()
        .find(|part| part.starts_with("xl/comments"))
        .unwrap()
        .to_string();
    let comments = std::str::from_utf8(editor.package().part(&comments_part).unwrap()).unwrap();
    assert!(comments.contains("ref=\"B3\""));
    let vml_part = editor
        .package()
        .part_names()
        .find(|part| part.starts_with("xl/drawings/vmlDrawing"))
        .unwrap()
        .to_string();
    let vml = std::str::from_utf8(editor.package().part(&vml_part).unwrap()).unwrap();
    assert!(vml.contains("<x:Row>2</x:Row>"));

    editor.save().await.unwrap();
    let reopened = NativeOfficeEditor::open(&path).await.unwrap();
    assert_eq!(
        reopened
            .snapshot()
            .unwrap()
            .get("/Sheet1/B3/comment", 0)
            .unwrap()
            .text,
        "Moves with two"
    );
}

#[tokio::test]
async fn sort_supports_exact_table_and_worksheet_autofilter_ranges() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-filter-table.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:B4", ["Name", "Rank"]),
        )
        .unwrap();
    for (path, value) in [
        ("/Sheet1/A2", "Three"),
        ("/Sheet1/A3", "One"),
        ("/Sheet1/A4", "Two"),
    ] {
        editor.set_cell_value(path, text(value)).unwrap();
    }
    for (path, value) in [
        ("/Sheet1/B2", "3"),
        ("/Sheet1/B3", "1"),
        ("/Sheet1/B4", "2"),
    ] {
        editor.set_cell_value(path, number(value)).unwrap();
    }
    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:B4",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")])
                .with_header(true),
        )
        .unwrap();
    assert_eq!(cell(&editor, "/Sheet1/A2"), "One");
    assert_eq!(cell(&editor, "/Sheet1/A3"), "Two");
    assert_eq!(cell(&editor, "/Sheet1/A4"), "Three");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/table[1]", 0)
            .unwrap()
            .format["ref"],
        "A1:B4"
    );

    editor
        .add_spreadsheet_auto_filter("/Sheet1", NativeSpreadsheetAutoFilter::new("D1:E4"))
        .unwrap();
    for (path, value) in [
        ("/Sheet1/D1", "Name"),
        ("/Sheet1/E1", "Rank"),
        ("/Sheet1/D2", "Three"),
        ("/Sheet1/D3", "One"),
        ("/Sheet1/D4", "Two"),
    ] {
        editor.set_cell_value(path, text(value)).unwrap();
    }
    for (path, value) in [
        ("/Sheet1/E2", "3"),
        ("/Sheet1/E3", "1"),
        ("/Sheet1/E4", "2"),
    ] {
        editor.set_cell_value(path, number(value)).unwrap();
    }
    editor
        .sort_spreadsheet_range(
            "/Sheet1/D2:E4",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("E")]),
        )
        .unwrap();
    assert_eq!(cell(&editor, "/Sheet1/D2"), "One");
    assert_eq!(cell(&editor, "/Sheet1/D3"), "Two");
    assert_eq!(cell(&editor, "/Sheet1/D4"), "Three");
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get("/Sheet1/autofilter", 0).unwrap().format["ref"],
        "D1:E4"
    );
    assert_eq!(
        snapshot.get("/Sheet1/sort", 0).unwrap().format["ref"],
        "D2:E4"
    );

    let before = editor.package().content_sha256();
    let error = editor
        .sort_spreadsheet_range(
            "/Sheet1/D1:E3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("E")])
                .with_header(true),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_sort_filter_partial");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn sort_fails_closed_for_table_totals_rows() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-table-totals.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Totals", "A1:B4", ["Name", "Rank"]).with_totals_row(true),
        )
        .unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:B4",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")])
                .with_header(true),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_sort_table_totals_unsupported"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn sort_preserves_strict_spreadsheetml_and_rejects_unknown_existing_state() {
    const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-sort.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A1", text("Key")).unwrap();
    editor.set_cell_value("/Sheet1/A2", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/A3", number("1")).unwrap();
    let worksheet = std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        .replace(TRANSITIONAL, STRICT);
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:A3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")])
                .with_header(true),
        )
        .unwrap();
    assert_eq!(cell(&editor, "/Sheet1/A2"), "1");
    let worksheet =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(worksheet.contains(STRICT));
    assert!(!worksheet.contains(TRANSITIONAL));
    assert!(worksheet.contains("<sortState ref=\"A1:A3\">"));

    let unknown = worksheet.replacen(
        "<sortState ",
        "<sortState xmlns:v=\"urn:vendor\" v:method=\"custom\" ",
        1,
    );
    editor
        .replace_xml_part("/xl/worksheets/sheet1.xml", unknown)
        .unwrap();
    let state = editor.snapshot().unwrap().get("/Sheet1/sort", 1).unwrap();
    assert_eq!(state.format["nativeMutable"], "false");
    let before = editor.package().content_sha256();
    let error = editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:A3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::descending("A")])
                .with_header(true),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_sort_unknown_content");
    assert_eq!(editor.package().content_sha256(), before);
    assert_eq!(
        editor.remove("/Sheet1/sort").unwrap_err().code,
        "use.office.spreadsheet_sort_unknown_content"
    );
}

#[tokio::test]
async fn sort_validates_paths_ranges_and_ordered_keys_before_mutating() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-validation.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/B2", number("1")).unwrap();
    let before = editor.package().content_sha256();

    let cases = [
        (
            "/Sheet1/A1:B2",
            NativeSpreadsheetSort::new(Vec::new()),
            "use.office.spreadsheet_sort_key_limit",
        ),
        (
            "/Sheet1/A1:B2",
            NativeSpreadsheetSort::new(vec![
                NativeSpreadsheetSortKey::ascending("A"),
                NativeSpreadsheetSortKey::descending("a"),
            ]),
            "use.office.spreadsheet_sort_column_duplicate",
        ),
        (
            "/Sheet1/A1:B2",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("C")]),
            "use.office.spreadsheet_sort_column_outside_range",
        ),
        (
            "/Sheet1/A1:B1",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")])
                .with_header(true),
            "use.office.spreadsheet_sort_range_empty",
        ),
        (
            "/Sheet1/A1:A2",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A"); 65]),
            "use.office.spreadsheet_sort_key_limit",
        ),
        (
            "/Sheet1/A1:XFD7",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
            "use.office.spreadsheet_sort_range_limit",
        ),
    ];
    for (path, request, code) in cases {
        assert_eq!(
            editor
                .sort_spreadsheet_range(path, request)
                .unwrap_err()
                .code,
            code
        );
        assert_eq!(editor.package().content_sha256(), before);
    }
    assert_eq!(
        editor
            .sort_spreadsheet_range(
                "/Sheet1/not-a-range",
                NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
            )
            .unwrap_err()
            .code,
        "use.office.mutation_path_unsupported"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn sort_auto_detects_the_used_range_and_honors_case_sensitivity() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-used-range.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for (path, value) in [
        ("/Sheet1/A1", "ID"),
        ("/Sheet1/B1", "Name"),
        ("/Sheet1/A2", "lower"),
        ("/Sheet1/B2", "a"),
        ("/Sheet1/A3", "upper"),
        ("/Sheet1/B3", "B"),
    ] {
        editor.set_cell_value(path, text(value)).unwrap();
    }
    editor
        .sort_spreadsheet_range(
            "/Sheet1",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")])
                .with_header(true)
                .with_case_sensitive(true),
        )
        .unwrap();
    assert_eq!(cell(&editor, "/Sheet1/A2"), "upper");
    assert_eq!(cell(&editor, "/Sheet1/A3"), "lower");
    let state = editor.snapshot().unwrap().get("/Sheet1/sort", 1).unwrap();
    assert_eq!(state.format["ref"], "A1:B3");
    assert_eq!(state.format["caseSensitive"], "true");
}

#[tokio::test]
async fn sort_moves_supported_spreadsheet_picture_anchors_with_records() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-picture.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A1", text("Key")).unwrap();
    editor.set_cell_value("/Sheet1/A2", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/A3", number("1")).unwrap();
    let picture = editor
        .add_image(
            "/Sheet1/B2",
            NativeOfficeImage::from_bytes(crate::image_tests::PNG_1X1)
                .unwrap()
                .with_name("Movable")
                .with_width_px(20)
                .with_height_px(20),
        )
        .unwrap();
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get(&picture.path, 0)
            .unwrap()
            .format["anchorCell"],
        "B2"
    );

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:B3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")])
                .with_header(true),
        )
        .unwrap();
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get(&picture.path, 0)
            .unwrap()
            .format["anchorCell"],
        "B3"
    );
}

#[tokio::test]
async fn sort_rejects_drawing_anchors_that_cannot_follow_one_record_losslessly() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-crossing-picture.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A1", text("Key")).unwrap();
    editor.set_cell_value("/Sheet1/A2", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/A3", number("1")).unwrap();
    editor
        .add_image(
            "/Sheet1/A2",
            NativeOfficeImage::from_bytes(crate::image_tests::PNG_1X1)
                .unwrap()
                .with_width_px(20)
                .with_height_px(20),
        )
        .unwrap();
    let drawing_part = editor
        .package()
        .part_names()
        .find(|part| part.starts_with("xl/drawings/drawing"))
        .unwrap()
        .to_string();
    let drawing = std::str::from_utf8(editor.package().part(&drawing_part).unwrap())
        .unwrap()
        .replace("oneCellAnchor", "twoCellAnchor")
        .replacen(
            "<xdr:ext",
            "<xdr:to><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:ext",
            1,
        );
    let mut package = editor.package().clone();
    package
        .set_part(&drawing_part, drawing.into_bytes())
        .unwrap();
    editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();

    let error = editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:A3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")])
                .with_header(true),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_sort_drawing_unsupported"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn sort_invalidates_chart_caches_after_physical_row_changes() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-chart-cache.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A1", text("Key")).unwrap();
    editor.set_cell_value("/Sheet1/A2", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/A3", number("1")).unwrap();
    let chart = editor
        .add_part("/Sheet1", NativeOfficePartType::Chart)
        .unwrap();
    editor
        .replace_xml_part(
            &chart.part,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Sheet1!$A$2:$A$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )
        .unwrap();
    assert!(std::str::from_utf8(
        editor
            .package()
            .part(chart.part.trim_start_matches('/'))
            .unwrap()
    )
    .unwrap()
    .contains("numCache"));

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:A3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")])
                .with_header(true),
        )
        .unwrap();
    let chart_xml = std::str::from_utf8(
        editor
            .package()
            .part(chart.part.trim_start_matches('/'))
            .unwrap(),
    )
    .unwrap();
    assert!(!chart_xml.contains("numCache"));
    assert!(chart_xml.contains("<c:f>Sheet1!$A$2:$A$3</c:f>"));
}

#[tokio::test]
async fn sort_rejects_worksheets_that_own_pivot_tables_before_commit() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-pivot.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/A2", number("1")).unwrap();
    let pivot_part = "xl/pivotTables/pivotTable1.xml";
    let mut package = editor.package().clone();
    package
        .set_part(
            pivot_part,
            br#"<?xml version="1.0" encoding="UTF-8"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Pivot1" cacheId="1"/>"#.to_vec(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        pivot_part,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/worksheets/_rels/sheet1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
        "../pivotTables/pivotTable1.xml",
    )
    .unwrap();
    editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();

    let error = editor
        .sort_spreadsheet_range(
            "/Sheet1/A1:A2",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("A")]),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_sort_pivot_unsupported");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn sort_preserves_prefixed_xml_sparse_gaps_and_destination_row_properties() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sort-prefixed.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let worksheet = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><s:dimension ref="A2:D4"/><s:sheetData><s:row r="2" ht="20" customHeight="1"><s:c r="A2" t="inlineStr"><s:is><s:t>two</s:t></s:is></s:c><s:c r="B2"><s:v>2</s:v></s:c><s:c r="D2" t="inlineStr"><s:is><s:t>slot-two</s:t></s:is></s:c></s:row><s:row r="4" ht="40" customHeight="1"><s:c r="A4" t="inlineStr"><s:is><s:t>one</s:t></s:is></s:c><s:c r="B4"><s:v>1</s:v></s:c><s:c r="D4" t="inlineStr"><s:is><s:t>slot-four</s:t></s:is></s:c></s:row></s:sheetData></s:worksheet>"#;
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.to_vec())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor
        .sort_spreadsheet_range(
            "/Sheet1/A2:B4",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")]),
        )
        .unwrap();
    assert_eq!(cell(&editor, "/Sheet1/A2"), "one");
    assert_eq!(cell(&editor, "/Sheet1/A3"), "two");
    assert_eq!(cell(&editor, "/Sheet1/D2"), "slot-two");
    assert_eq!(cell(&editor, "/Sheet1/D4"), "slot-four");
    let xml =
        std::str::from_utf8(editor.package().part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(xml.contains("<s:row customHeight=\"1\" ht=\"20\" r=\"2\">"));
    assert!(xml.contains("<s:row r=\"3\">"));
    assert!(xml.contains("<s:row customHeight=\"1\" ht=\"40\" r=\"4\">"));
    assert!(xml.contains("<s:sortState ref=\"A2:B4\">"));
    assert!(!xml.contains("<row"));
    assert!(!xml.contains("<c "));
}

#[test]
fn direction_defaults_to_ascending_when_deserialized() {
    let sort: NativeSpreadsheetSort = serde_json::from_value(serde_json::json!({
        "keys": [{"column": "A"}]
    }))
    .unwrap();
    assert_eq!(
        sort.keys[0].direction,
        NativeSpreadsheetSortDirection::Ascending
    );
    assert!(!sort.header);
    assert!(!sort.case_sensitive);
}
