use super::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficeRgbColor, NativeSpreadsheetBorder,
    NativeSpreadsheetCellFormat, NativeSpreadsheetFill, NativeSpreadsheetReadingOrder,
    NativeSpreadsheetVerticalAlignment, SpreadsheetCellValue,
};

fn rich_cell_format() -> NativeSpreadsheetCellFormat {
    NativeSpreadsheetCellFormat {
        number_format: Some("#,##0.00;[Red]-#,##0.00".into()),
        fill: Some(NativeSpreadsheetFill::Solid {
            color: NativeOfficeRgbColor::new(0xAA, 0xBB, 0xCC),
        }),
        border: None,
        vertical_alignment: Some(NativeSpreadsheetVerticalAlignment::Distributed),
        wrap_text: Some(true),
        text_rotation: Some(45),
        indent: Some(2),
        shrink_to_fit: Some(true),
        reading_order: Some(NativeSpreadsheetReadingOrder::RightToLeft),
    }
}

fn assert_package_unchanged(
    actual: &super::NativeOfficePackage,
    expected: &super::NativeOfficePackage,
) {
    let actual_names = actual.part_names().collect::<Vec<_>>();
    let expected_names = expected.part_names().collect::<Vec<_>>();
    assert_eq!(actual_names, expected_names);
    for name in expected_names {
        assert_eq!(actual.part(name).unwrap(), expected.part(name).unwrap());
    }
}

#[test]
fn cell_format_mutation_has_a_typed_stable_json_contract() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetCellFormat>();
    assert_send_sync::<NativeSpreadsheetFill>();
    assert_send_sync::<NativeSpreadsheetReadingOrder>();
    assert_send_sync::<NativeSpreadsheetVerticalAlignment>();

    let mutation = NativeOfficeMutation::SetCellFormat {
        path: "/Sheet1/A1:C2".into(),
        format: rich_cell_format(),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "set-cell-format",
            "path": "/Sheet1/A1:C2",
            "format": {
                "numberFormat": "#,##0.00;[Red]-#,##0.00",
                "fill": {
                    "kind": "solid",
                    "color": { "red": 170, "green": 187, "blue": 204 }
                },
                "verticalAlignment": "distributed",
                "wrapText": true,
                "textRotation": 45,
                "indent": 2,
                "shrinkToFit": true,
                "readingOrder": "right-to-left"
            }
        })
    );
    assert!(
        serde_json::from_value::<NativeOfficeMutation>(serde_json::json!({
            "operation": "set-cell-format",
            "path": "/Sheet1/A1",
            "format": { "gradient": "red-blue" }
        }))
        .is_err()
    );
}

#[tokio::test]
async fn native_spreadsheet_writes_and_deduplicates_cell_format() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("cell-format.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1:B2",
            SpreadsheetCellValue::Number {
                value: "1234.5".into(),
            },
        )
        .unwrap();
    editor
        .set_cell_format("/Sheet1/A1:B2", rich_cell_format())
        .unwrap();

    for path in ["/Sheet1/A1", "/Sheet1/B2"] {
        let cell = editor.snapshot().unwrap().get(path, 0).unwrap();
        assert_eq!(cell.format["numberFormat"], "#,##0.00;[Red]-#,##0.00");
        assert_eq!(cell.format["fill"], "AABBCC");
        assert_eq!(cell.format["verticalAlignment"], "distributed");
        assert_eq!(cell.format["wrapText"], "true");
        assert_eq!(cell.format["textRotation"], "45");
        assert_eq!(cell.format["indent"], "2");
        assert_eq!(cell.format["shrinkToFit"], "true");
        assert_eq!(cell.format["readingOrder"], "rtl");
    }

    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    assert!(styles.contains("formatCode=\"#,##0.00;[Red]-#,##0.00\""));
    assert!(styles.contains("patternType=\"solid\""));
    assert!(styles.contains("rgb=\"FFAABBCC\""));
    assert!(styles.contains("applyNumberFormat=\"1\""));
    assert!(styles.contains("applyFill=\"1\""));
    assert!(styles.contains("applyAlignment=\"1\""));

    let document = editor.snapshot().unwrap();
    for rendered in [document.html_view().unwrap(), document.svg_view().unwrap()] {
        for expected in [
            "data-fill=\"AABBCC\"",
            "data-vertical-alignment=\"distributed\"",
            "data-wrap-text=\"true\"",
            "data-text-rotation=\"45\"",
            "data-indent=\"2\"",
            "data-shrink-to-fit=\"true\"",
            "data-reading-order=\"rtl\"",
        ] {
            assert!(rendered.content.contains(expected), "missing {expected}");
        }
    }

    let first_styles = editor.package().part("xl/styles.xml").unwrap().to_vec();
    editor
        .set_cell_format("/Sheet1/A1:B2", rich_cell_format())
        .unwrap();
    assert_eq!(
        editor.package().part("xl/styles.xml").unwrap(),
        first_styles
    );
}

#[tokio::test]
async fn native_spreadsheet_cell_format_clears_fill_and_direction_explicitly() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("clear-cell-format.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_format("/Sheet1/A1", rich_cell_format())
        .unwrap();
    editor
        .set_cell_format(
            "/Sheet1/A1",
            NativeSpreadsheetCellFormat {
                number_format: Some("general".into()),
                fill: Some(NativeSpreadsheetFill::None),
                wrap_text: Some(false),
                text_rotation: Some(0),
                indent: Some(0),
                shrink_to_fit: Some(false),
                reading_order: Some(NativeSpreadsheetReadingOrder::Context),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();

    let cell = editor.snapshot().unwrap().get("/Sheet1/A1", 0).unwrap();
    assert_eq!(cell.format["numberFormat"], "General");
    assert!(!cell.format.contains_key("fill"));
    assert_eq!(cell.format["wrapText"], "false");
    assert_eq!(cell.format["textRotation"], "0");
    assert_eq!(cell.format["indent"], "0");
    assert_eq!(cell.format["shrinkToFit"], "false");
    assert!(!cell.format.contains_key("readingOrder"));
}

#[tokio::test]
async fn number_format_aliases_round_trip_through_semantic_reads() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("number-format-aliases.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let formats = [
        ("A1", "general", "General"),
        ("B1", "number", "#,##0.00"),
        ("C1", "currency", "\"$\"#,##0.00"),
        (
            "D1",
            "accounting",
            "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        ),
        ("E1", "percent", "0.00%"),
        ("F1", "scientific", "0.00E+00"),
        ("G1", "text", "@"),
        ("H1", "date", "yyyy-mm-dd"),
        ("I1", "time", "h:mm:ss"),
        ("J1", "datetime", "yyyy-mm-dd h:mm:ss"),
    ];
    for (cell, alias, expected) in formats {
        editor
            .set_cell_format(
                format!("/Sheet1/{cell}"),
                NativeSpreadsheetCellFormat {
                    number_format: Some(alias.into()),
                    ..NativeSpreadsheetCellFormat::default()
                },
            )
            .unwrap();
        let node = editor
            .snapshot()
            .unwrap()
            .get(&format!("/Sheet1/{cell}"), 0)
            .unwrap();
        assert_eq!(node.format["numberFormat"], expected);
    }

    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    assert_eq!(styles.matches("<numFmt ").count(), 4);
}

#[tokio::test]
async fn custom_number_format_ids_do_not_collide_with_existing_xf_references() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("number-format-id.xlsx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.set_text_format(
        "/Sheet1/A1",
        super::NativeOfficeTextFormat {
            bold: Some(true),
            ..super::NativeOfficeTextFormat::default()
        },
    )
    .unwrap();

    let mut package = seed.package().clone();
    let mut styles = String::from_utf8(package.part("xl/styles.xml").unwrap().to_vec()).unwrap();
    let value = styles.rfind("numFmtId=\"0\"").unwrap() + "numFmtId=\"".len();
    styles.replace_range(value..=value, "164");
    package
        .set_part("xl/styles.xml", styles.into_bytes())
        .unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_cell_format(
            "/Sheet1/A1",
            NativeSpreadsheetCellFormat {
                number_format: Some("yyyy-mm-dd h:mm:ss.000".into()),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();
    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    assert!(styles.contains("numFmtId=\"165\" formatCode=\"yyyy-mm-dd h:mm:ss.000\""));
}

#[tokio::test]
async fn cell_format_preserves_unrelated_xf_and_alignment_data() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("preserve-cell-format.xlsx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.set_text_format(
        "/Sheet1/A1",
        super::NativeOfficeTextFormat {
            bold: Some(true),
            alignment: Some(super::NativeOfficeHorizontalAlignment::Center),
            ..super::NativeOfficeTextFormat::default()
        },
    )
    .unwrap();

    let mut package = seed.package().clone();
    let styles = String::from_utf8(package.part("xl/styles.xml").unwrap().to_vec()).unwrap();
    let cell_xfs_end = styles.find("</cellXfs>").unwrap();
    let xf_start = styles[..cell_xfs_end].rfind("<xf ").unwrap();
    let xf_end = xf_start + styles[xf_start..].find('>').unwrap();
    let xf_attribute_position = if styles.as_bytes()[xf_end - 1] == b'/' {
        xf_end - 1
    } else {
        xf_end
    };
    let mut styles = styles;
    styles.insert_str(xf_attribute_position, " quotePrefix=\"1\" dataKeep=\"xf\"");
    styles = styles.replacen(
        "<alignment horizontal=\"center\"/>",
        "<alignment horizontal=\"center\" relativeIndent=\"3\" dataKeep=\"alignment\"/>",
        1,
    );
    package
        .set_part("xl/styles.xml", styles.into_bytes())
        .unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_cell_format(
            "/Sheet1/A1",
            NativeSpreadsheetCellFormat {
                wrap_text: Some(true),
                vertical_alignment: Some(NativeSpreadsheetVerticalAlignment::Top),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();
    let cell = editor.snapshot().unwrap().get("/Sheet1/A1", 0).unwrap();
    assert_eq!(cell.format["bold"], "true");
    assert_eq!(cell.format["alignment"], "center");
    assert_eq!(cell.format["wrapText"], "true");
    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    assert_eq!(styles.matches("dataKeep=\"xf\"").count(), 2);
    assert_eq!(styles.matches("quotePrefix=\"1\"").count(), 2);
    assert_eq!(styles.matches("dataKeep=\"alignment\"").count(), 2);
    assert_eq!(styles.matches("relativeIndent=\"3\"").count(), 2);
}

#[tokio::test]
async fn invalid_cell_format_rolls_back_the_entire_batch() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("invalid-cell-format.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let original = editor.package().clone();
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/Sheet1/A1".into(),
                text: "Changed".into(),
            },
            NativeOfficeMutation::SetCellFormat {
                path: "/Sheet1/A1".into(),
                format: NativeSpreadsheetCellFormat {
                    number_format: Some("0.00;0.00;0.00;0.00;0.00".into()),
                    ..NativeSpreadsheetCellFormat::default()
                },
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.number_format_invalid");
    assert_package_unchanged(editor.package(), &original);

    let error = editor
        .set_cell_format(
            "/Sheet1/A1",
            NativeSpreadsheetCellFormat {
                text_rotation: Some(181),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.text_rotation_invalid");
    assert_package_unchanged(editor.package(), &original);

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/Sheet1/A1".into(),
                text: "Changed again".into(),
            },
            NativeOfficeMutation::SetCellFormat {
                path: "/Sheet1/A1".into(),
                format: NativeSpreadsheetCellFormat {
                    border: Some(NativeSpreadsheetBorder::default()),
                    ..NativeSpreadsheetCellFormat::default()
                },
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.cell_border_empty");
    assert_package_unchanged(editor.package(), &original);
}
