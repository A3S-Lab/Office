use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficeRgbColor,
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatIconSet,
    NativeSpreadsheetConditionalFormatOperator, NativeSpreadsheetConditionalFormatRule,
    NativeSpreadsheetConditionalFormatThreshold, NativeSpreadsheetConditionalFormatThresholdKind,
    NativeSpreadsheetConditionalFormatTimePeriod, NativeSpreadsheetDifferentialFormat,
    OfficeNodeType,
};

const STRICT_SPREADSHEET_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

fn rgb(value: u32) -> NativeOfficeRgbColor {
    NativeOfficeRgbColor::new(
        ((value >> 16) & 0xff) as u8,
        ((value >> 8) & 0xff) as u8,
        (value & 0xff) as u8,
    )
}

fn highlight(value: u32) -> NativeSpreadsheetDifferentialFormat {
    NativeSpreadsheetDifferentialFormat::default()
        .with_fill(rgb(value))
        .with_font_color(rgb(0x112233))
        .with_bold(true)
}

fn comparison(range: &str) -> NativeSpreadsheetConditionalFormat {
    NativeSpreadsheetConditionalFormat::new(
        range,
        NativeSpreadsheetConditionalFormatRule::CellIs {
            operator: NativeSpreadsheetConditionalFormatOperator::GreaterThan,
            formula1: "80".into(),
            formula2: None,
            format: highlight(0xC6EFCE),
        },
    )
}

#[test]
fn conditional_formatting_has_closed_typed_json_and_is_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetConditionalFormat>();
    assert_send_sync::<NativeSpreadsheetConditionalFormatRule>();
    assert_send_sync::<NativeSpreadsheetConditionalFormatThreshold>();

    let mutation = NativeOfficeMutation::AddConditionalFormat {
        sheet: "/Sheet1".into(),
        conditional_format: NativeSpreadsheetConditionalFormat::new(
            "A2:A20",
            NativeSpreadsheetConditionalFormatRule::DataBar {
                color: rgb(0x638EC6),
                min: NativeSpreadsheetConditionalFormatThreshold::min(),
                max: NativeSpreadsheetConditionalFormatThreshold::number("100"),
                show_value: false,
                min_length: Some(5),
                max_length: Some(95),
            },
        ),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "add-conditional-format",
            "sheet": "/Sheet1",
            "conditionalFormat": {
                "ranges": ["A2:A20"],
                "rule": {
                    "type": "dataBar",
                    "color": {"red": 99, "green": 142, "blue": 198},
                    "min": {"kind": "min"},
                    "max": {"kind": "number", "value": "100"},
                    "showValue": false,
                    "minLength": 5,
                    "maxLength": 95
                }
            }
        })
    );

    assert!(
        serde_json::from_value::<NativeOfficeMutation>(serde_json::json!({
            "operation": "add-conditional-format",
            "sheet": "/Sheet1",
            "conditionalFormat": {
                "ranges": ["A1"],
                "rule": {"type": "script", "formula": "TRUE"}
            }
        }))
        .is_err()
    );
}

#[tokio::test]
async fn conditional_formatting_lifecycle_and_visual_families_are_native() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("conditional-formatting.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    assert_eq!(
        editor
            .add_conditional_format("/sheet1", comparison("A2:A11"))
            .unwrap(),
        "/Sheet1/cf[1]"
    );
    editor
        .add_conditional_format(
            "/Sheet1",
            NativeSpreadsheetConditionalFormat::new(
                "B2:B11",
                NativeSpreadsheetConditionalFormatRule::DataBar {
                    color: rgb(0x638EC6),
                    min: NativeSpreadsheetConditionalFormatThreshold::min(),
                    max: NativeSpreadsheetConditionalFormatThreshold::max(),
                    show_value: true,
                    min_length: None,
                    max_length: None,
                },
            ),
        )
        .unwrap();
    editor
        .add_conditional_format(
            "/Sheet1",
            NativeSpreadsheetConditionalFormat::new(
                "C2:C11",
                NativeSpreadsheetConditionalFormatRule::ColorScale {
                    min: NativeSpreadsheetConditionalFormatThreshold::min(),
                    min_color: rgb(0xF8696B),
                    mid: Some(NativeSpreadsheetConditionalFormatThreshold::percentile(
                        "50",
                    )),
                    mid_color: Some(rgb(0xFFEB84)),
                    max: NativeSpreadsheetConditionalFormatThreshold::max(),
                    max_color: rgb(0x63BE7B),
                },
            ),
        )
        .unwrap();
    editor
        .add_conditional_format(
            "/Sheet1",
            NativeSpreadsheetConditionalFormat::new(
                "D2:D11",
                NativeSpreadsheetConditionalFormatRule::IconSet {
                    icon_set: NativeSpreadsheetConditionalFormatIconSet::ThreeTrafficLights1,
                    thresholds: Vec::new(),
                    reverse: true,
                    show_value: false,
                },
            ),
        )
        .unwrap();

    let worksheet = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(worksheet.contains("type=\"cellIs\""));
    assert!(worksheet.contains("dxfId=\"0\""));
    assert!(worksheet.contains("<dataBar>"));
    assert!(worksheet.contains("<colorScale>"));
    assert!(worksheet.contains("iconSet=\"3TrafficLights1\""));
    let styles = part_text(&editor, "xl/styles.xml");
    assert!(styles.contains("<dxfs count=\"1\">"));
    assert!(styles.contains("rgb=\"FFC6EFCE\""));
    assert!(styles.contains("rgb=\"FF112233\""));

    let snapshot = editor.snapshot().unwrap();
    let cell_is = snapshot.get("/Sheet1/cf[1]", 0).unwrap();
    assert_eq!(cell_is.node_type, OfficeNodeType::ConditionalFormatting);
    assert_eq!(cell_is.format["type"], "cellIs");
    assert_eq!(cell_is.format["operator"], "greaterThan");
    assert_eq!(cell_is.format["formula1"], "80");
    assert_eq!(cell_is.format["fill"], "C6EFCE");
    assert_eq!(cell_is.format["fontColor"], "112233");
    assert_eq!(cell_is.format["fontBold"], "true");
    assert_eq!(cell_is.format["nativeMutable"], "true");
    assert_eq!(snapshot.query("conditionalFormatting").unwrap().len(), 4);
    assert_eq!(
        snapshot
            .query("conditionalFormatting[type=iconSet]")
            .unwrap()
            .len(),
        1
    );
    let icon = snapshot.get("/Sheet1/cf[4]", 0).unwrap();
    assert_eq!(icon.format["thresholds"], "percent:0;percent:33;percent:66");

    let replacement = NativeSpreadsheetConditionalFormat::new(
        "E2:E20",
        NativeSpreadsheetConditionalFormatRule::Formula {
            formula: "MOD(E2,2)=0".into(),
            format: highlight(0xBDD7EE),
        },
    )
    .with_stop_if_true(true);
    assert_eq!(
        editor
            .set_conditional_format("/sheet1/conditional-formatting[1]", replacement)
            .unwrap(),
        "/Sheet1/cf[1]"
    );
    let updated = editor.snapshot().unwrap().get("/Sheet1/cf[1]", 0).unwrap();
    assert_eq!(updated.format["type"], "expression");
    assert_eq!(updated.format["formula"], "MOD(E2,2)=0");
    assert_eq!(updated.format["ref"], "E2:E20");
    assert_eq!(updated.format["priority"], "1");
    assert_eq!(updated.format["stopIfTrue"], "true");

    editor.save().await.unwrap();
    let mut reopened = NativeOfficeEditor::open(&path).await.unwrap();
    reopened.remove("/Sheet1/cf[2]").unwrap();
    assert_eq!(
        reopened
            .snapshot()
            .unwrap()
            .query("conditionalFormatting")
            .unwrap()
            .len(),
        3
    );
}

#[tokio::test]
async fn classic_conditional_format_families_round_trip_semantically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("classic-cf.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let format = highlight(0xFFEB9C);
    let rules = vec![
        NativeSpreadsheetConditionalFormatRule::ContainsText {
            text: "error".into(),
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::NotContainsText {
            text: "ok".into(),
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::BeginsWith {
            text: "A".into(),
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::EndsWith {
            text: "Z".into(),
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::Top {
            rank: 10,
            percent: true,
            bottom: true,
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::AboveAverage {
            above: false,
            equal: true,
            standard_deviations: Some(2),
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::DuplicateValues {
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::UniqueValues {
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::ContainsBlanks {
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::NotContainsBlanks {
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::ContainsErrors {
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::NotContainsErrors {
            format: format.clone(),
        },
        NativeSpreadsheetConditionalFormatRule::TimePeriod {
            period: NativeSpreadsheetConditionalFormatTimePeriod::ThisMonth,
            format,
        },
    ];
    for (index, rule) in rules.into_iter().enumerate() {
        editor
            .add_conditional_format(
                "/Sheet1",
                NativeSpreadsheetConditionalFormat::new(
                    format!("A{}:A{}", index + 1, index + 2),
                    rule,
                ),
            )
            .unwrap();
    }
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.query("conditionalFormatting").unwrap().len(), 13);
    assert_eq!(
        snapshot.get("/Sheet1/cf[1]", 0).unwrap().format["text"],
        "error"
    );
    assert_eq!(
        snapshot.get("/Sheet1/cf[5]", 0).unwrap().format["percent"],
        "true"
    );
    assert_eq!(
        snapshot.get("/Sheet1/cf[6]", 0).unwrap().format["above"],
        "false"
    );
    assert_eq!(
        snapshot.get("/Sheet1/cf[13]", 0).unwrap().format["period"],
        "thisMonth"
    );
    assert!(snapshot
        .query("conditionalFormatting")
        .unwrap()
        .iter()
        .all(|node| node.format["nativeMutable"] == "true"));
}

#[tokio::test]
async fn invalid_conditional_formats_roll_back_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("invalid-cf.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_conditional_format("/Sheet1", comparison("A1:A10"))
        .unwrap();
    let before = editor.package().content_sha256();

    let error = editor
        .add_conditional_format(
            "/Sheet1",
            NativeSpreadsheetConditionalFormat::new(
                "B1:B10",
                NativeSpreadsheetConditionalFormatRule::CellIs {
                    operator: NativeSpreadsheetConditionalFormatOperator::Between,
                    formula1: "1".into(),
                    formula2: None,
                    format: highlight(0xFF0000),
                },
            ),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_conditional_format_formula2_required"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .add_conditional_format("/Sheet1", comparison("C1:C10").with_range("C10:D20"))
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_conditional_format_range_overlap"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::AddConditionalFormat {
                sheet: "/Sheet1".into(),
                conditional_format: comparison("E1:E10"),
            },
            NativeOfficeMutation::AddConditionalFormat {
                sheet: "/Sheet1".into(),
                conditional_format: NativeSpreadsheetConditionalFormat::new(
                    "F1:F10",
                    NativeSpreadsheetConditionalFormatRule::ColorScale {
                        min: NativeSpreadsheetConditionalFormatThreshold::min(),
                        min_color: rgb(0xFFFFFF),
                        mid: Some(NativeSpreadsheetConditionalFormatThreshold {
                            kind: NativeSpreadsheetConditionalFormatThresholdKind::Percentile,
                            value: Some("101".into()),
                        }),
                        mid_color: Some(rgb(0xFFFF00)),
                        max: NativeSpreadsheetConditionalFormatThreshold::max(),
                        max_color: rgb(0x00FF00),
                    },
                ),
            },
        ])
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_conditional_format_threshold_invalid"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn strict_and_unknown_conditional_format_content_is_preserved_or_rejected() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-cf.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let worksheet = format!(
        "<worksheet xmlns=\"{STRICT_SPREADSHEET_NAMESPACE}\" xmlns:v=\"urn:vendor\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><sheetData/><conditionalFormatting sqref=\"A1:A5\" v:owner=\"keep\"><cfRule type=\"expression\" priority=\"1\" v:id=\"keep\"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>"
    );
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor
        .set_conditional_format(
            "/Sheet1/cf[1]",
            NativeSpreadsheetConditionalFormat::new(
                "A1:A5",
                NativeSpreadsheetConditionalFormatRule::Formula {
                    formula: "A1<10".into(),
                    format: NativeSpreadsheetDifferentialFormat::default(),
                },
            ),
        )
        .unwrap();
    editor
        .add_conditional_format("/Sheet1", comparison("B1:B5"))
        .unwrap();
    let edited = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(edited.contains(STRICT_SPREADSHEET_NAMESPACE));
    assert!(edited.contains("v:owner=\"keep\""));
    assert!(edited.contains("v:id=\"keep\""));
    assert!(edited.find("conditionalFormatting").unwrap() < edited.find("pageMargins").unwrap());

    editor.remove("/Sheet1/cf[2]").unwrap();
    let before = editor.package().content_sha256();
    let error = editor.remove("/Sheet1/cf[1]").unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_conditional_format_unknown_content"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let worksheet = part_text(&editor, "xl/worksheets/sheet1.xml").replacen(
        "</formula>",
        "</formula><v:payload/>",
        1,
    );
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .set_conditional_format("/Sheet1/cf[1]", comparison("A1:A5"))
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_conditional_format_unknown_content"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn shared_sqref_rules_cannot_change_ranges_implicitly() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("shared-cf.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let worksheet = part_text(&editor, "xl/worksheets/sheet1.xml").replacen(
        "<pageMargins",
        "<conditionalFormatting sqref=\"A1:A5\"><cfRule type=\"expression\" priority=\"1\"><formula>A1&gt;0</formula></cfRule><cfRule type=\"expression\" priority=\"2\"><formula>A1&lt;10</formula></cfRule></conditionalFormatting><pageMargins",
        1,
    );
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .set_conditional_format(
            "/Sheet1/cf[1]",
            NativeSpreadsheetConditionalFormat::new(
                "B1:B5",
                NativeSpreadsheetConditionalFormatRule::Formula {
                    formula: "B1>0".into(),
                    format: NativeSpreadsheetDifferentialFormat::default(),
                },
            ),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_conditional_format_shared_range"
    );
    assert_eq!(editor.package().content_sha256(), before);

    editor
        .set_conditional_format(
            "/Sheet1/cf[1]",
            NativeSpreadsheetConditionalFormat::new(
                "A1:A5",
                NativeSpreadsheetConditionalFormatRule::Formula {
                    formula: "A1>=1".into(),
                    format: NativeSpreadsheetDifferentialFormat::default(),
                },
            ),
        )
        .unwrap();
    editor.remove("/Sheet1/cf[1]").unwrap();
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .query("conditionalFormatting")
            .unwrap()
            .len(),
        1
    );
}

fn part_text(editor: &NativeOfficeEditor, name: &str) -> String {
    String::from_utf8(editor.package().part(name).unwrap().to_vec()).unwrap()
}
