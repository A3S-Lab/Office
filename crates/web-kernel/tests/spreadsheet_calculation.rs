use a3s_office_web_kernel::{
    calculate_spreadsheet, SpreadsheetCalculatedCell, SpreadsheetCalculationRequest,
    SpreadsheetCoordinate, SpreadsheetInputCell, SpreadsheetInputSheet, SpreadsheetValue,
    OFFICE_KERNEL_PROTOCOL_VERSION,
};
use serde::Deserialize;

#[test]
fn calculates_sparse_dependencies_with_the_shared_formula_parser() {
    let result = calculate_spreadsheet(&request()).unwrap();

    assert_eq!(result.engine, "wasm");
    assert_eq!(result.cells.len(), 2);
    assert_eq!(
        result.cells[0].value,
        SpreadsheetValue::Number { value: 5.0 }
    );
    assert_eq!(
        result.cells[1].value,
        SpreadsheetValue::Number { value: 10.0 }
    );
    assert_eq!(
        result.calculation_order,
        vec![
            SpreadsheetCoordinate {
                sheet_id: "sheet-1".into(),
                row: 0,
                column: 1,
            },
            SpreadsheetCoordinate {
                sheet_id: "sheet-1".into(),
                row: 1,
                column: 1,
            },
        ]
    );
    assert!(result.issues.is_empty());
}

#[test]
fn reports_cycles_and_unsupported_functions_without_losing_valid_cells() {
    let mut request = request();
    request.sheets[0].cells.extend([
        SpreadsheetInputCell {
            row: 2,
            column: 1,
            formula: Some("=A3S_UNKNOWN(A1)".into()),
            value: SpreadsheetValue::Number { value: 41.0 },
        },
        SpreadsheetInputCell {
            row: 3,
            column: 1,
            formula: Some("=B5".into()),
            value: SpreadsheetValue::Number { value: 7.0 },
        },
        SpreadsheetInputCell {
            row: 4,
            column: 1,
            formula: Some("=B4".into()),
            value: SpreadsheetValue::Number { value: 8.0 },
        },
    ]);

    let result = calculate_spreadsheet(&request).unwrap();

    assert_eq!(result.cells.len(), 2);
    assert!(result
        .issues
        .iter()
        .any(|issue| issue.code == "office.kernel.spreadsheet.formula_unsupported"));
    assert!(result
        .issues
        .iter()
        .any(|issue| issue.code == "office.kernel.spreadsheet.circular_reference"));
}

#[test]
fn matches_the_javascript_scalar_contract_fixtures() {
    let fixtures: Vec<ParityFixture> = serde_json::from_str(include_str!(
        "../../../tests/fixtures/spreadsheet-kernel-parity.json"
    ))
    .unwrap();

    for fixture in fixtures {
        let request = SpreadsheetCalculationRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetCalculation".into(),
            request_id: 11,
            revision: 4,
            document_revision: 9,
            sheets: fixture.sheets,
            targets: fixture.targets,
        };
        let result = calculate_spreadsheet(&request).unwrap();
        assert_eq!(
            result.cells, fixture.expected_cells,
            "{} cells",
            fixture.name
        );
        assert_eq!(
            result.calculation_order, fixture.expected_order,
            "{} order",
            fixture.name
        );
        assert_eq!(
            result
                .issues
                .into_iter()
                .map(|issue| ExpectedIssue {
                    cell: issue.cell,
                    code: issue.code,
                })
                .collect::<Vec<_>>(),
            fixture.expected_issues,
            "{} issues",
            fixture.name
        );
    }
}

#[test]
fn rejects_malformed_targets_and_oversized_formulas() {
    let mut malformed = request();
    malformed.targets = vec![SpreadsheetCoordinate {
        sheet_id: "missing".into(),
        row: 0,
        column: 0,
    }];
    assert_eq!(
        calculate_spreadsheet(&malformed).unwrap_err().code,
        "office.kernel.spreadsheet.target_invalid"
    );

    let mut oversized = request();
    oversized.sheets[0].cells[2].formula = Some(format!("={}", "1".repeat(8_192)));
    assert_eq!(
        calculate_spreadsheet(&oversized).unwrap_err().code,
        "office.kernel.spreadsheet.formula_invalid"
    );

    let mut oversized_targets = request();
    oversized_targets.targets = vec![
        SpreadsheetCoordinate {
            sheet_id: "sheet-1".into(),
            row: 0,
            column: 0,
        };
        100_001
    ];
    assert_eq!(
        calculate_spreadsheet(&oversized_targets).unwrap_err().code,
        "office.kernel.spreadsheet.target_limit_exceeded"
    );
}

#[test]
fn bounds_deep_dependency_chains_before_the_wasm_stack() {
    let chain_length = 300_u32;
    let mut request = request();
    request.sheets[0].cells = (0..chain_length)
        .map(|row| SpreadsheetInputCell {
            row,
            column: 0,
            formula: Some(if row == 0 {
                "=1".into()
            } else {
                format!("=A{row}+1")
            }),
            value: SpreadsheetValue::Number {
                value: f64::from(row + 1),
            },
        })
        .collect();
    request.targets = vec![SpreadsheetCoordinate {
        sheet_id: "sheet-1".into(),
        row: chain_length - 1,
        column: 0,
    }];

    let result = calculate_spreadsheet(&request).unwrap();

    assert!(result.cells.is_empty());
    assert!(result
        .issues
        .iter()
        .any(|issue| { issue.code == "office.kernel.spreadsheet.dependency_depth_exceeded" }));
}

#[test]
fn bounds_text_produced_by_formula_calculation() {
    let mut request = request();
    request.sheets[0].cells = vec![
        SpreadsheetInputCell {
            row: 0,
            column: 0,
            formula: None,
            value: SpreadsheetValue::Text {
                value: "a".repeat(20_000),
            },
        },
        SpreadsheetInputCell {
            row: 0,
            column: 1,
            formula: None,
            value: SpreadsheetValue::Text {
                value: "b".repeat(20_000),
            },
        },
        SpreadsheetInputCell {
            row: 0,
            column: 2,
            formula: Some("=CONCAT(A1,B1)".into()),
            value: SpreadsheetValue::Blank,
        },
    ];

    let result = calculate_spreadsheet(&request).unwrap();

    assert_eq!(
        result.cells,
        vec![SpreadsheetCalculatedCell {
            sheet_id: "sheet-1".into(),
            row: 0,
            column: 2,
            value: SpreadsheetValue::Error {
                value: "#VALUE!".into(),
            },
        }]
    );
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ParityFixture {
    name: String,
    sheets: Vec<SpreadsheetInputSheet>,
    targets: Vec<SpreadsheetCoordinate>,
    expected_cells: Vec<SpreadsheetCalculatedCell>,
    expected_order: Vec<SpreadsheetCoordinate>,
    expected_issues: Vec<ExpectedIssue>,
}

#[derive(Debug, PartialEq, Eq, Deserialize)]
struct ExpectedIssue {
    cell: SpreadsheetCoordinate,
    code: String,
}

fn request() -> SpreadsheetCalculationRequest {
    SpreadsheetCalculationRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "spreadsheetCalculation".into(),
        request_id: 7,
        revision: 3,
        document_revision: 2,
        sheets: vec![SpreadsheetInputSheet {
            id: "sheet-1".into(),
            name: "Sheet 1".into(),
            cells: vec![
                SpreadsheetInputCell {
                    row: 0,
                    column: 0,
                    formula: None,
                    value: SpreadsheetValue::Number { value: 2.0 },
                },
                SpreadsheetInputCell {
                    row: 1,
                    column: 0,
                    formula: None,
                    value: SpreadsheetValue::Number { value: 3.0 },
                },
                SpreadsheetInputCell {
                    row: 0,
                    column: 1,
                    formula: Some("=SUM(A1:A2)".into()),
                    value: SpreadsheetValue::Blank,
                },
                SpreadsheetInputCell {
                    row: 1,
                    column: 1,
                    formula: Some("=B1*2".into()),
                    value: SpreadsheetValue::Blank,
                },
            ],
        }],
        targets: Vec::new(),
    }
}
