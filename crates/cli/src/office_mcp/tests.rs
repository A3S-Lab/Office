use super::*;
use a3s_office::{
    NativeOfficeComment, NativeOfficeCommentPosition, NativeOfficeHighlightColor,
    NativeOfficeHyperlinkTarget, NativeOfficeMutation, NativeOfficeRgbColor, NativeOfficeTextCase,
    NativeOfficeTextFormat, NativeOfficeTextMatchMode, NativeOfficeTextScript,
    NativeOfficeUnderline, NativeSpreadsheetBorder, NativeSpreadsheetBorderLine,
    NativeSpreadsheetBorderStyle, NativeSpreadsheetCellFormat,
    NativeSpreadsheetConditionalFormatOperator, NativeSpreadsheetConditionalFormatRule,
    NativeSpreadsheetFill, NativeSpreadsheetNamedRangeScope, NativeSpreadsheetReadingOrder,
    NativeSpreadsheetSortDirection, NativeSpreadsheetVerticalAlignment,
};

#[test]
fn native_office_server_exposes_bounded_tools_and_confirmed_compat_install() {
    let server = NativeOfficeMcpServer::new();
    let tools = server.tool_router.list_all();
    let mut names: Vec<&str> = tools
        .iter()
        .map(|tool| tool.name.as_ref())
        .collect::<Vec<_>>();
    names.sort_unstable();
    assert_eq!(
        names,
        [
            "office_apply_batch",
            "office_close",
            "office_create",
            "office_get",
            "office_install_compat",
            "office_list",
            "office_merge_template",
            "office_open",
            "office_query",
            "office_raw_xml",
            "office_save",
            "office_validate",
            "office_view",
        ]
    );

    let annotations = |name: &str| {
        tools
            .iter()
            .find(|tool| tool.name == name)
            .and_then(|tool| tool.annotations.as_ref())
            .unwrap_or_else(|| panic!("{name} must declare MCP annotations"))
    };
    for name in [
        "office_validate",
        "office_list",
        "office_get",
        "office_query",
        "office_raw_xml",
    ] {
        let annotation = annotations(name);
        assert_eq!(annotation.read_only_hint, Some(true), "{name}");
        assert_eq!(annotation.open_world_hint, Some(false), "{name}");
    }
    assert_eq!(
        annotations("office_apply_batch").read_only_hint,
        Some(false)
    );
    assert_eq!(annotations("office_save").destructive_hint, Some(true));
}

#[test]
fn office_view_schema_exposes_typed_view_options() {
    let schema = schemars::schema_for!(OfficeViewInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    assert!(encoded.contains("annotated"));
    assert!(encoded.contains("screenshot"));
    assert!(encoded.contains("issues"));
    assert!(encoded.contains("issueType"));
    assert!(encoded.contains("missing_alt_text"));
    assert!(encoded.contains("limit"));
    assert!(encoded.contains("output"));
    assert!(encoded.contains("timeoutMs"));

    let input: OfficeViewInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "view": "screenshot",
        "output": "report.png",
        "timeoutMs": 30_000
    }))
    .unwrap();
    assert_eq!(input.view, OfficeView::Screenshot);
    assert_eq!(input.output.as_deref(), Some("report.png"));
    assert_eq!(input.timeout_ms, Some(30_000));

    let issues: OfficeViewInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "view": "issues",
        "issueType": "missing_alt_text",
        "limit": 10
    }))
    .unwrap();
    assert_eq!(issues.view, OfficeView::Issues);
    assert_eq!(
        issues.issue_type,
        Some(input::OfficeIssueFilter::MissingAltText)
    );
    assert_eq!(issues.limit, Some(10));

    let annotated: OfficeViewInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "view": "annotated",
        "limit": 25
    }))
    .unwrap();
    assert_eq!(annotated.view, OfficeView::Annotated);
    assert_eq!(annotated.limit, Some(25));
}

#[test]
fn office_batch_schema_exposes_typed_text_formatting() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "set-text-format",
        "fontFamily",
        "fontSizeCentipoints",
        "textColor",
        "alignment",
        "underline",
        "script",
        "strikethrough",
        "doubleStrikethrough",
        "textCase",
        "highlight",
        "language",
        "small-caps",
        "superscript",
        "justify",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "mutations": [{
            "operation": "set-text-format",
            "path": "/body/p[1]/r[1]",
            "format": {
                "bold": true,
                "underline": "double",
                "script": "superscript",
                "strikethrough": false,
                "doubleStrikethrough": true,
                "textCase": "small-caps",
                "highlight": "yellow",
                "language": "en-US",
                "fontSizeCentipoints": 1200,
                "textColor": { "red": 18, "green": 52, "blue": 86 }
            }
        }]
    }))
    .unwrap();
    let mutation = input.mutations.into_iter().next().unwrap();
    let native = mutation.into_native().unwrap();
    assert!(matches!(
        native,
        NativeOfficeMutation::SetTextFormat {
            format: NativeOfficeTextFormat {
                bold: Some(true),
                underline: Some(NativeOfficeUnderline::Double),
                script: Some(NativeOfficeTextScript::Superscript),
                strikethrough: Some(false),
                double_strikethrough: Some(true),
                text_case: Some(NativeOfficeTextCase::SmallCaps),
                highlight: Some(NativeOfficeHighlightColor::Yellow),
                language: Some(ref language),
                font_size_centipoints: Some(1200),
                text_color: Some(NativeOfficeRgbColor {
                    red: 18,
                    green: 52,
                    blue: 86
                }),
                ..
            },
            ..
        } if language == "en-US"
    ));
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_cell_formatting() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "set-cell-format",
        "numberFormat",
        "fill",
        "solid",
        "border",
        "diagonalUp",
        "mediumDashDotDot",
        "slantDashDot",
        "verticalAlignment",
        "distributed",
        "wrapText",
        "textRotation",
        "indent",
        "shrinkToFit",
        "readingOrder",
        "right-to-left",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "mutations": [{
            "operation": "set-cell-format",
            "path": "/Sheet1/A1:C2",
            "format": {
                "numberFormat": "currency",
                "fill": {
                    "kind": "solid",
                    "color": { "red": 18, "green": 52, "blue": 86 }
                },
                "border": {
                    "left": {
                        "kind": "line",
                        "style": "dashDot",
                        "color": { "red": 1, "green": 2, "blue": 3 }
                    },
                    "bottom": { "kind": "none" },
                    "diagonalUp": true
                },
                "verticalAlignment": "distributed",
                "wrapText": true,
                "textRotation": 45,
                "indent": 2,
                "shrinkToFit": false,
                "readingOrder": "right-to-left"
            }
        }]
    }))
    .unwrap();
    let mutation = input.mutations.into_iter().next().unwrap();
    assert!(matches!(
        mutation.into_native().unwrap(),
        NativeOfficeMutation::SetCellFormat {
            format: NativeSpreadsheetCellFormat {
                number_format: Some(ref number_format),
                fill: Some(NativeSpreadsheetFill::Solid {
                    color: NativeOfficeRgbColor {
                        red: 18,
                        green: 52,
                        blue: 86,
                    },
                }),
                border: Some(NativeSpreadsheetBorder {
                    left: Some(NativeSpreadsheetBorderLine::Line {
                        style: NativeSpreadsheetBorderStyle::DashDot,
                        color: Some(NativeOfficeRgbColor {
                            red: 1,
                            green: 2,
                            blue: 3,
                        }),
                    }),
                    bottom: Some(NativeSpreadsheetBorderLine::None),
                    diagonal_up: Some(true),
                    ..
                }),
                vertical_alignment: Some(NativeSpreadsheetVerticalAlignment::Distributed),
                wrap_text: Some(true),
                text_rotation: Some(45),
                indent: Some(2),
                shrink_to_fit: Some(false),
                reading_order: Some(NativeSpreadsheetReadingOrder::RightToLeft),
            },
            ..
        } if number_format == "currency"
    ));

    let unknown = serde_json::from_value::<OfficeBatchInput>(serde_json::json!({
        "session": "report",
        "mutations": [{
            "operation": "set-cell-format",
            "path": "/Sheet1/A1",
            "format": { "gradient": "red-blue" }
        }]
    }));
    assert!(unknown.is_err());
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_merges() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in ["merge-cells", "unmerge-cells"] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [
            { "operation": "merge-cells", "path": "/Sheet1/A1:B2" },
            { "operation": "unmerge-cells", "path": "/Sheet1/C1:D2" }
        ]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::MergeCells { ref path } if path == "/Sheet1/A1:B2"
    ));
    assert!(matches!(
        input.mutations[1].clone().into_native().unwrap(),
        NativeOfficeMutation::UnmergeCells { ref path } if path == "/Sheet1/C1:D2"
    ));
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_sorting() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "sort-spreadsheet-range",
        "caseSensitive",
        "ascending",
        "descending",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "sort-spreadsheet-range",
            "path": "/Sheet1/A1:D100",
            "sort": {
                "keys": [
                    {"column": "B", "direction": "descending"},
                    {"column": "C"}
                ],
                "header": true,
                "caseSensitive": false
            }
        }]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::SortSpreadsheetRange { ref path, ref sort }
            if path == "/Sheet1/A1:D100"
                && sort.header
                && !sort.case_sensitive
                && sort.keys[0].direction == NativeSpreadsheetSortDirection::Descending
                && sort.keys[1].direction == NativeSpreadsheetSortDirection::Ascending
    ));

    let unknown = serde_json::from_value::<OfficeBatchInput>(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "sort-spreadsheet-range",
            "path": "/Sheet1/A1:D100",
            "sort": {
                "keys": [{"column": "B", "direction": "random"}],
                "locale": "en-US"
            }
        }]
    }));
    assert!(unknown.is_err());
}

#[test]
fn office_batch_schema_exposes_native_spreadsheet_recalculation() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    assert!(
        encoded.contains("recalculate-spreadsheet-formulas"),
        "{encoded}"
    );

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "recalculate-spreadsheet-formulas"
        }]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::RecalculateSpreadsheetFormulas
    ));
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_import_and_frozen_panes() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "import-spreadsheet-delimited",
        "set-spreadsheet-frozen-pane",
        "startCell",
        "frozenRows",
        "frozenColumns",
        "topLeftCell",
        "csv",
        "tsv",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [
            {
                "operation": "import-spreadsheet-delimited",
                "sheet": "/Sheet1",
                "import": {
                    "content": "Name,Value\nAlpha,42",
                    "format": "csv",
                    "header": true,
                    "startCell": "B2"
                }
            },
            {
                "operation": "set-spreadsheet-frozen-pane",
                "sheet": "/Sheet1",
                "pane": {
                    "frozenRows": 2,
                    "frozenColumns": 0,
                    "topLeftCell": "B3"
                }
            }
        ]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::ImportSpreadsheetDelimited { ref sheet, ref import }
            if sheet == "/Sheet1"
                && import.header
                && import.start_cell == "B2"
                && import.format == a3s_office::NativeSpreadsheetDelimitedFormat::Csv
    ));
    assert!(matches!(
        input.mutations[1].clone().into_native().unwrap(),
        NativeOfficeMutation::SetSpreadsheetFrozenPane { ref sheet, ref pane }
            if sheet == "/Sheet1"
                && pane.frozen_rows == 2
                && pane.frozen_columns == 0
                && pane.top_left_cell == "B3"
    ));

    let unknown = serde_json::from_value::<OfficeBatchInput>(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "import-spreadsheet-delimited",
            "sheet": "/Sheet1",
            "import": {"content": "a", "format": "json"}
        }]
    }));
    assert!(unknown.is_err());
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_data_validation() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "add-data-validation",
        "set-data-validation",
        "textLength",
        "notBetween",
        "inCellDropdown",
        "information",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "add-data-validation",
            "sheet": "/Sheet1",
            "validation": {
                "type": "whole",
                "ranges": ["A2:A20"],
                "operator": "between",
                "formula1": "1",
                "formula2": "100",
                "allowBlank": false,
                "errorStyle": "warning"
            }
        }]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::AddDataValidation {
            ref sheet,
            validation: a3s_office::NativeSpreadsheetDataValidation {
                validation_type: a3s_office::NativeSpreadsheetDataValidationType::Whole,
                operator: Some(
                    a3s_office::NativeSpreadsheetDataValidationOperator::Between
                ),
                allow_blank: false,
                error_style:
                    a3s_office::NativeSpreadsheetDataValidationErrorStyle::Warning,
                ..
            }
        } if sheet == "/Sheet1"
    ));

    let unknown = serde_json::from_value::<OfficeBatchInput>(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "add-data-validation",
            "sheet": "/Sheet1",
            "validation": {
                "type": "list",
                "ranges": ["A1"],
                "formula1": "A,B",
                "script": "alert(1)"
            }
        }]
    }));
    assert!(unknown.is_err());
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_conditional_formatting() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "add-conditional-format",
        "set-conditional-format",
        "conditionalFormat",
        "cellIs",
        "containsText",
        "last7Days",
        "dataBar",
        "colorScale",
        "iconSet",
        "3TrafficLights1",
        "fontColor",
        "percentile",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "add-conditional-format",
            "sheet": "/Sheet1",
            "conditionalFormat": {
                "ranges": ["A2:A20"],
                "stopIfTrue": true,
                "rule": {
                    "type": "cellIs",
                    "operator": "greaterThan",
                    "formula1": "80",
                    "format": {
                        "fill": {"red": 198, "green": 239, "blue": 206},
                        "fontColor": {"red": 0, "green": 97, "blue": 0},
                        "bold": true
                    }
                }
            }
        }]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::AddConditionalFormat {
            ref sheet,
            conditional_format: a3s_office::NativeSpreadsheetConditionalFormat {
                stop_if_true: true,
                rule: NativeSpreadsheetConditionalFormatRule::CellIs {
                    operator: NativeSpreadsheetConditionalFormatOperator::GreaterThan,
                    ref formula1,
                    ref format,
                    ..
                },
                ..
            }
        } if sheet == "/Sheet1"
            && formula1 == "80"
            && format.fill == Some(NativeOfficeRgbColor::new(198, 239, 206))
            && format.bold == Some(true)
    ));

    let unknown = serde_json::from_value::<OfficeBatchInput>(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "add-conditional-format",
            "sheet": "/Sheet1",
            "conditionalFormat": {
                "ranges": ["A1"],
                "rule": {"type": "script", "formula": "TRUE"}
            }
        }]
    }));
    assert!(unknown.is_err());
}

#[test]
fn office_batch_schema_exposes_typed_spreadsheet_named_ranges() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "add-named-range",
        "set-named-range",
        "namedRange",
        "scope",
        "comment",
        "volatile",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "add-named-range",
            "namedRange": {
                "name": "Status",
                "ref": "A2:A20",
                "scope": "Sheet1",
                "comment": "Workflow status",
                "volatile": true
            }
        }]
    }))
    .unwrap();
    assert!(matches!(
        input.mutations[0].clone().into_native().unwrap(),
        NativeOfficeMutation::AddNamedRange {
            named_range: a3s_office::NativeSpreadsheetNamedRange {
                ref name,
                ref reference,
                scope: NativeSpreadsheetNamedRangeScope::Worksheet(ref sheet),
                volatile: true,
                ..
            }
        } if name == "Status" && reference == "A2:A20" && sheet == "Sheet1"
    ));

    let unknown = serde_json::from_value::<OfficeBatchInput>(serde_json::json!({
        "session": "workbook",
        "mutations": [{
            "operation": "add-named-range",
            "namedRange": {
                "name": "Status",
                "ref": "'Sheet1'!$A$1",
                "script": "alert(1)"
            }
        }]
    }));
    assert!(unknown.is_err());
}

#[test]
fn office_batch_schema_exposes_typed_text_replacement() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in ["replace-text", "find", "replace", "literal", "regex"] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "mutations": [{
            "operation": "replace-text",
            "path": "/body",
            "replacement": {
                "find": "Q([1-4]) 2025",
                "replace": "Q$1 2026",
                "mode": "regex"
            }
        }]
    }))
    .unwrap();
    let mutation = input.mutations.into_iter().next().unwrap();
    assert!(matches!(
        mutation.into_native().unwrap(),
        NativeOfficeMutation::ReplaceText {
            replacement: a3s_office::NativeOfficeTextReplacement {
                mode: NativeOfficeTextMatchMode::Regex,
                ref find,
                ref replace,
            },
            ..
        } if find == "Q([1-4]) 2025" && replace == "Q$1 2026"
    ));
}

#[test]
fn office_batch_schema_exposes_typed_hyperlinks() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "set-hyperlink",
        "external",
        "internal",
        "uri",
        "location",
        "display",
        "tooltip",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "report",
        "mutations": [{
            "operation": "set-hyperlink",
            "path": "/body/p[1]",
            "hyperlink": {
                "target": {
                    "kind": "external",
                    "uri": "https://example.com/report"
                },
                "display": "Report",
                "tooltip": "Open report"
            }
        }]
    }))
    .unwrap();
    let mutation = input.mutations.into_iter().next().unwrap();
    assert!(matches!(
        mutation.into_native().unwrap(),
        NativeOfficeMutation::SetHyperlink {
            hyperlink: a3s_office::NativeOfficeHyperlink {
                target: NativeOfficeHyperlinkTarget::External { ref uri },
                display: Some(ref display),
                tooltip: Some(ref tooltip),
            },
            ..
        } if uri == "https://example.com/report"
            && display == "Report"
            && tooltip == "Open report"
    ));
}

#[test]
fn office_batch_schema_exposes_typed_legacy_comments() {
    let schema = schemars::schema_for!(OfficeBatchInput);
    let encoded = serde_json::to_string(&schema).unwrap();
    for expected in [
        "add-comment",
        "set-comment",
        "author",
        "initials",
        "position",
        "xEmu",
        "yEmu",
    ] {
        assert!(encoded.contains(expected), "missing {expected}");
    }

    let input: OfficeBatchInput = serde_json::from_value(serde_json::json!({
        "session": "deck",
        "mutations": [{
            "operation": "add-comment",
            "parent": "/slide[1]",
            "comment": {
                "author": "Alice",
                "text": "Review this",
                "initials": "AL",
                "position": { "xEmu": 914400, "yEmu": 457200 }
            }
        }]
    }))
    .unwrap();
    let mutation = input.mutations.into_iter().next().unwrap();
    assert!(matches!(
        mutation.into_native().unwrap(),
        NativeOfficeMutation::AddComment {
            comment: NativeOfficeComment {
                ref author,
                ref text,
                initials: Some(ref initials),
                position: Some(NativeOfficeCommentPosition {
                    x_emu: 914_400,
                    y_emu: 457_200,
                }),
            },
            ..
        } if author == "Alice" && text == "Review this" && initials == "AL"
    ));
}
