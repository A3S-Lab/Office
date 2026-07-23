use super::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficeRgbColor, NativeSpreadsheetBorder,
    NativeSpreadsheetBorderLine, NativeSpreadsheetBorderStyle, NativeSpreadsheetCellFormat,
};

fn rich_border() -> NativeSpreadsheetBorder {
    NativeSpreadsheetBorder {
        left: Some(NativeSpreadsheetBorderLine::Line {
            style: NativeSpreadsheetBorderStyle::Thin,
            color: Some(NativeOfficeRgbColor::new(0x11, 0x22, 0x33)),
        }),
        right: Some(NativeSpreadsheetBorderLine::Line {
            style: NativeSpreadsheetBorderStyle::MediumDashed,
            color: None,
        }),
        top: Some(NativeSpreadsheetBorderLine::Line {
            style: NativeSpreadsheetBorderStyle::Double,
            color: Some(NativeOfficeRgbColor::new(0x44, 0x55, 0x66)),
        }),
        bottom: Some(NativeSpreadsheetBorderLine::None),
        diagonal: Some(NativeSpreadsheetBorderLine::Line {
            style: NativeSpreadsheetBorderStyle::SlantDashDot,
            color: Some(NativeOfficeRgbColor::new(0x77, 0x88, 0x99)),
        }),
        diagonal_up: Some(true),
        diagonal_down: Some(false),
    }
}

#[test]
fn cell_border_has_a_typed_stable_json_contract() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetBorder>();
    assert_send_sync::<NativeSpreadsheetBorderLine>();
    assert_send_sync::<NativeSpreadsheetBorderStyle>();

    let mutation = NativeOfficeMutation::SetCellFormat {
        path: "/Sheet1/D4".into(),
        format: NativeSpreadsheetCellFormat {
            border: Some(rich_border()),
            ..NativeSpreadsheetCellFormat::default()
        },
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "set-cell-format",
            "path": "/Sheet1/D4",
            "format": {
                "border": {
                    "left": {
                        "kind": "line",
                        "style": "thin",
                        "color": { "red": 17, "green": 34, "blue": 51 }
                    },
                    "right": { "kind": "line", "style": "mediumDashed" },
                    "top": {
                        "kind": "line",
                        "style": "double",
                        "color": { "red": 68, "green": 85, "blue": 102 }
                    },
                    "bottom": { "kind": "none" },
                    "diagonal": {
                        "kind": "line",
                        "style": "slantDashDot",
                        "color": { "red": 119, "green": 136, "blue": 153 }
                    },
                    "diagonalUp": true,
                    "diagonalDown": false
                }
            }
        })
    );
}

#[tokio::test]
async fn native_spreadsheet_writes_reads_clears_and_deduplicates_cell_borders() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("cell-borders.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_format(
            "/Sheet1/A1:B2",
            NativeSpreadsheetCellFormat {
                border: Some(rich_border()),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();

    for path in ["/Sheet1/A1", "/Sheet1/B2"] {
        let cell = editor.snapshot().unwrap().get(path, 0).unwrap();
        assert_eq!(cell.format["borderLeft"], "thin");
        assert_eq!(cell.format["borderLeftColor"], "112233");
        assert_eq!(cell.format["borderRight"], "mediumDashed");
        assert_eq!(cell.format["borderTop"], "double");
        assert_eq!(cell.format["borderTopColor"], "445566");
        assert!(!cell.format.contains_key("borderBottom"));
        assert_eq!(cell.format["borderDiagonal"], "slantDashDot");
        assert_eq!(cell.format["borderDiagonalColor"], "778899");
        assert_eq!(cell.format["borderDiagonalUp"], "true");
        assert_eq!(cell.format["borderDiagonalDown"], "false");
    }

    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    assert_eq!(styles.matches("<border").count(), 3);
    assert!(styles.contains("applyBorder=\"1\""));
    assert!(styles.contains("diagonalUp=\"1\""));
    assert!(styles.contains("diagonalDown=\"0\""));
    assert!(styles.contains("style=\"mediumDashed\""));
    assert!(styles.contains("style=\"slantDashDot\""));

    let document = editor.snapshot().unwrap();
    for rendered in [document.html_view().unwrap(), document.svg_view().unwrap()] {
        for expected in [
            "data-border-left=\"thin\"",
            "data-border-left-color=\"112233\"",
            "data-border-right=\"mediumDashed\"",
            "data-border-top=\"double\"",
            "data-border-top-color=\"445566\"",
            "data-border-diagonal=\"slantDashDot\"",
            "data-border-diagonal-color=\"778899\"",
            "data-border-diagonal-up=\"true\"",
            "data-border-diagonal-down=\"false\"",
        ] {
            assert!(rendered.content.contains(expected), "missing {expected}");
        }
    }

    let first_styles = editor.package().part("xl/styles.xml").unwrap().to_vec();
    editor
        .set_cell_format(
            "/Sheet1/A1:B2",
            NativeSpreadsheetCellFormat {
                border: Some(rich_border()),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();
    assert_eq!(
        editor.package().part("xl/styles.xml").unwrap(),
        first_styles
    );
}

#[tokio::test]
async fn native_spreadsheet_supports_every_standard_cell_border_style() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("all-cell-border-styles.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let styles = [
        (NativeSpreadsheetBorderStyle::Thin, "thin"),
        (NativeSpreadsheetBorderStyle::Medium, "medium"),
        (NativeSpreadsheetBorderStyle::Thick, "thick"),
        (NativeSpreadsheetBorderStyle::Double, "double"),
        (NativeSpreadsheetBorderStyle::Dashed, "dashed"),
        (NativeSpreadsheetBorderStyle::Dotted, "dotted"),
        (NativeSpreadsheetBorderStyle::DashDot, "dashDot"),
        (NativeSpreadsheetBorderStyle::DashDotDot, "dashDotDot"),
        (NativeSpreadsheetBorderStyle::Hair, "hair"),
        (NativeSpreadsheetBorderStyle::MediumDashed, "mediumDashed"),
        (NativeSpreadsheetBorderStyle::MediumDashDot, "mediumDashDot"),
        (
            NativeSpreadsheetBorderStyle::MediumDashDotDot,
            "mediumDashDotDot",
        ),
        (NativeSpreadsheetBorderStyle::SlantDashDot, "slantDashDot"),
    ];
    for (index, (style, expected)) in styles.into_iter().enumerate() {
        let reference = super::spreadsheet_reference::column_name((index + 1) as u32);
        let path = format!("/Sheet1/{reference}1");
        editor
            .set_cell_format(
                &path,
                NativeSpreadsheetCellFormat {
                    border: Some(NativeSpreadsheetBorder {
                        top: Some(NativeSpreadsheetBorderLine::Line { style, color: None }),
                        ..NativeSpreadsheetBorder::default()
                    }),
                    ..NativeSpreadsheetCellFormat::default()
                },
            )
            .unwrap();
        assert_eq!(
            editor.snapshot().unwrap().get(&path, 0).unwrap().format["borderTop"],
            expected
        );
    }
}

#[tokio::test]
async fn cell_border_updates_preserve_unrelated_border_xml() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("preserve-cell-border.xlsx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.set_cell_format(
        "/Sheet1/A1",
        NativeSpreadsheetCellFormat {
            border: Some(rich_border()),
            ..NativeSpreadsheetCellFormat::default()
        },
    )
    .unwrap();

    let mut package = seed.package().clone();
    let mut styles = String::from_utf8(package.part("xl/styles.xml").unwrap().to_vec()).unwrap();
    let borders_end = styles.find("</borders>").unwrap();
    let border_start = styles[..borders_end].rfind("<border ").unwrap();
    styles.insert_str(border_start + "<border".len(), " dataKeep=\"border\"");
    styles = styles.replacen(
        "<left style=\"thin\"><color rgb=\"FF112233\"/></left>",
        "<left style=\"thin\" dataKeep=\"left\"><color rgb=\"FF112233\" dataKeep=\"color\"/></left>",
        1,
    );
    let borders_end = styles.find("</borders>").unwrap();
    let border_end = styles[..borders_end].rfind("</border>").unwrap();
    styles.insert_str(
        border_end,
        "<vertical style=\"hair\" dataKeep=\"vertical\"/>",
    );
    package
        .set_part("xl/styles.xml", styles.into_bytes())
        .unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_cell_format(
            "/Sheet1/A1",
            NativeSpreadsheetCellFormat {
                border: Some(NativeSpreadsheetBorder {
                    top: Some(NativeSpreadsheetBorderLine::Line {
                        style: NativeSpreadsheetBorderStyle::Thick,
                        color: Some(NativeOfficeRgbColor::new(0xDE, 0xAD, 0xBE)),
                    }),
                    ..NativeSpreadsheetBorder::default()
                }),
                ..NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();

    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    for preserved in [
        "dataKeep=\"border\"",
        "dataKeep=\"left\"",
        "dataKeep=\"color\"",
        "dataKeep=\"vertical\"",
    ] {
        assert_eq!(styles.matches(preserved).count(), 2, "missing {preserved}");
    }
    assert!(styles.contains("<top style=\"thick\"><color rgb=\"FFDEADBE\"/></top>"));
    let cell = editor.snapshot().unwrap().get("/Sheet1/A1", 0).unwrap();
    assert_eq!(cell.format["borderLeft"], "thin");
    assert_eq!(cell.format["borderTop"], "thick");
}
