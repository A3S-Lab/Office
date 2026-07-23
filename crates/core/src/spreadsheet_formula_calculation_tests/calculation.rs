use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeOfficePartType, NativeOfficeReplayArtifact,
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatRule,
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationType,
    NativeSpreadsheetDifferentialFormat, NativeSpreadsheetNamedRange, NativeSpreadsheetTable,
    SpreadsheetCellValue, SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaFunctionRegistry,
    SpreadsheetFormulaValue, MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES,
    MAX_SPREADSHEET_FORMULA_TEXT_BYTES,
};

fn number(value: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Number {
        value: value.to_string(),
    }
}

fn text(value: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Text {
        value: value.to_string(),
    }
}

fn boolean(value: bool) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Boolean { value }
}

fn formula(expression: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Formula {
        expression: expression.to_string(),
    }
}

#[tokio::test]
async fn calculation_evaluates_dependency_order_operators_and_typed_functions() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("calculate.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/A2", text("ignored"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/B1", formula("A1*3"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/C1", formula("SUM(A1:B1)+ROUND(2.55,1)"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/D1", formula("IF(C1=10.6,\"yes\",\"no\")"))
        .unwrap();
    editor.set_cell_value("/Sheet1/E1", formula("1/0")).unwrap();

    let calculation = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap();
    assert_eq!(calculation.formula_count, 4);
    assert_eq!(calculation.spill_cell_count, 0);
    assert_eq!(
        calculation
            .cells
            .iter()
            .map(|cell| (cell.cell.path(), cell.value.clone()))
            .collect::<Vec<_>>(),
        [
            (
                "/Sheet1/B1".to_string(),
                SpreadsheetFormulaValue::Number { value: "6".into() }
            ),
            (
                "/Sheet1/C1".to_string(),
                SpreadsheetFormulaValue::Number {
                    value: "10.6".into()
                }
            ),
            (
                "/Sheet1/D1".to_string(),
                SpreadsheetFormulaValue::Text {
                    value: "yes".into()
                }
            ),
            (
                "/Sheet1/E1".to_string(),
                SpreadsheetFormulaValue::error(SpreadsheetFormulaErrorLiteral::DivisionByZero)
            ),
        ]
    );
}

#[tokio::test]
async fn calculation_covers_the_closed_scalar_function_registry() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("functions.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/A2", text("ignored"))
        .unwrap();
    editor.set_cell_value("/Sheet1/A3", boolean(true)).unwrap();
    for (path, expression) in [
        ("/Sheet1/B1", "SUM(A1:A3)"),
        ("/Sheet1/B2", "AVERAGE(A1:A3)"),
        ("/Sheet1/B3", "MIN(A1:A3)"),
        ("/Sheet1/B4", "MAX(A1:A3)"),
        ("/Sheet1/B5", "COUNT(\"3\",TRUE,A1:A3)"),
        ("/Sheet1/B6", "COUNTA(A1:A3)"),
        ("/Sheet1/B7", "ABS(-2)"),
        ("/Sheet1/B8", "SQRT(9)"),
        ("/Sheet1/B9", "POWER(2,3)"),
        ("/Sheet1/B10", "MOD(-3,2)"),
        ("/Sheet1/B11", "ROUND(-2.55,1)"),
        ("/Sheet1/B12", "IFERROR(1/0,7)"),
        ("/Sheet1/B13", "AND(A1:A3)"),
        ("/Sheet1/B14", "OR(0,A2:A2)"),
        ("/Sheet1/B15", "NOT(\"TRUE\")"),
        ("/Sheet1/B16", "CONCAT(\"v=\",A1)"),
        ("/Sheet1/B17", "ROW(A3)"),
        ("/Sheet1/B18", "COLUMN(B1)"),
        ("/Sheet1/B19", "IF(TRUE,,9)"),
        ("/Sheet1/B20", "IF(FALSE,9)"),
        ("/Sheet1/B21", "NA()"),
        ("/Sheet1/B22", "PI()"),
    ] {
        editor.set_cell_value(path, formula(expression)).unwrap();
    }

    let calculation = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap();
    let values = calculation
        .cells
        .iter()
        .map(|cell| (cell.cell.path(), cell.value.clone()))
        .collect::<std::collections::BTreeMap<_, _>>();
    for (path, expected) in [
        ("/Sheet1/B1", "2"),
        ("/Sheet1/B2", "2"),
        ("/Sheet1/B3", "2"),
        ("/Sheet1/B4", "2"),
        ("/Sheet1/B5", "3"),
        ("/Sheet1/B6", "3"),
        ("/Sheet1/B7", "2"),
        ("/Sheet1/B8", "3"),
        ("/Sheet1/B9", "8"),
        ("/Sheet1/B10", "1"),
        ("/Sheet1/B11", "-2.6"),
        ("/Sheet1/B12", "7"),
        ("/Sheet1/B17", "3"),
        ("/Sheet1/B18", "2"),
    ] {
        assert_eq!(
            values[path],
            SpreadsheetFormulaValue::Number {
                value: expected.into()
            },
            "{path}"
        );
    }
    assert_eq!(
        values["/Sheet1/B13"],
        SpreadsheetFormulaValue::Boolean { value: true }
    );
    assert_eq!(
        values["/Sheet1/B14"],
        SpreadsheetFormulaValue::Boolean { value: false }
    );
    assert_eq!(
        values["/Sheet1/B15"],
        SpreadsheetFormulaValue::Boolean { value: false }
    );
    assert_eq!(
        values["/Sheet1/B16"],
        SpreadsheetFormulaValue::Text {
            value: "v=2".into()
        }
    );
    assert_eq!(values["/Sheet1/B19"], SpreadsheetFormulaValue::Blank);
    assert_eq!(
        values["/Sheet1/B20"],
        SpreadsheetFormulaValue::Boolean { value: false }
    );
    assert_eq!(
        values["/Sheet1/B21"],
        SpreadsheetFormulaValue::error(SpreadsheetFormulaErrorLiteral::NotAvailable)
    );
    let SpreadsheetFormulaValue::Number { value } = &values["/Sheet1/B22"] else {
        panic!("PI did not return a number");
    };
    assert!((value.parse::<f64>().unwrap() - std::f64::consts::PI).abs() < f64::EPSILON);
}

#[tokio::test]
async fn calculation_plans_dynamic_array_spills_and_detects_obstructions() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("spill.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SEQUENCE(2,3,1,1)"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/E1", formula("TRANSPOSE(A1#)"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/F2", text("blocked"))
        .unwrap();

    let calculation = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap();
    assert_eq!(calculation.spill_cell_count, 5);
    assert_eq!(calculation.cells[0].spill_range.as_deref(), Some("A1:C2"));
    assert_eq!(
        calculation.cells[0].value,
        SpreadsheetFormulaValue::Array {
            rows: vec![
                vec![
                    SpreadsheetFormulaValue::Number { value: "1".into() },
                    SpreadsheetFormulaValue::Number { value: "2".into() },
                    SpreadsheetFormulaValue::Number { value: "3".into() },
                ],
                vec![
                    SpreadsheetFormulaValue::Number { value: "4".into() },
                    SpreadsheetFormulaValue::Number { value: "5".into() },
                    SpreadsheetFormulaValue::Number { value: "6".into() },
                ],
            ]
        }
    );
    assert_eq!(
        calculation.cells[1].value,
        SpreadsheetFormulaValue::error(SpreadsheetFormulaErrorLiteral::Spill)
    );
    assert_eq!(calculation.cells[1].spill_range, None);
}

#[tokio::test]
async fn calculation_resolves_table_items_current_rows_and_local_references() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("table-references.xlsx"))
        .await
        .unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C5", ["Item", "Qty", "Unit Price"])
                .with_display_name("SalesView")
                .with_totals_row(true),
        )
        .unwrap();
    for (path, value) in [
        ("/Sheet1/B2", number("2")),
        ("/Sheet1/C2", number("10")),
        ("/Sheet1/B3", number("3")),
        ("/Sheet1/C3", number("20")),
        ("/Sheet1/C4", number("30")),
        ("/Sheet1/B5", number("999")),
        ("/Sheet1/C5", number("999")),
    ] {
        editor.set_cell_value(path, value).unwrap();
    }
    editor
        .set_cell_value("/Sheet1/B4", formula("B2+B3"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A2", formula("Sales[@Qty]"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A3", formula("SUM(Sales[[#This Row],[Qty]])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A4", formula("[@Qty]"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/E1", formula("SUM(SalesView[Qty])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/F1", formula("SUM(Sales[[Qty]:[Unit Price]])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/G1", formula("SUM(Sales[#All])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/H1", formula("SUM(Sales[#Data])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/I1", formula("COUNTA(Sales[[#Headers],[Qty]])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/J1", formula("SUM(Sales[[#Totals],[Qty]])"))
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/K1",
            formula("SUM(Sales[[#Headers],[#Totals],[Qty]])"),
        )
        .unwrap();

    let graph = editor
        .snapshot()
        .unwrap()
        .formula_dependency_graph()
        .unwrap();
    for path in [
        "/Sheet1/A2",
        "/Sheet1/A3",
        "/Sheet1/A4",
        "/Sheet1/E1",
        "/Sheet1/F1",
        "/Sheet1/G1",
        "/Sheet1/H1",
        "/Sheet1/I1",
        "/Sheet1/J1",
        "/Sheet1/K1",
    ] {
        let node = graph
            .nodes
            .iter()
            .find(|node| node.cell.path() == path)
            .unwrap();
        assert!(node.unresolved_references.is_empty(), "{path}");
    }
    for path in ["/Sheet1/A4", "/Sheet1/E1", "/Sheet1/F1"] {
        let node = graph
            .nodes
            .iter()
            .find(|node| node.cell.path() == path)
            .unwrap();
        assert_eq!(
            node.dependencies
                .iter()
                .map(|cell| cell.path())
                .collect::<Vec<_>>(),
            ["/Sheet1/B4"]
        );
    }
    for path in ["/Sheet1/G1", "/Sheet1/H1"] {
        let node = graph
            .nodes
            .iter()
            .find(|node| node.cell.path() == path)
            .unwrap();
        assert_eq!(
            node.dependencies
                .iter()
                .map(|cell| cell.path())
                .collect::<Vec<_>>(),
            ["/Sheet1/A2", "/Sheet1/A3", "/Sheet1/A4", "/Sheet1/B4"]
        );
    }

    let calculation = editor.recalculate_spreadsheet_formulas().unwrap();
    let values = calculation
        .cells
        .iter()
        .map(|cell| (cell.cell.path(), cell.value.clone()))
        .collect::<std::collections::BTreeMap<_, _>>();
    for (path, expected) in [
        ("/Sheet1/A2", "2"),
        ("/Sheet1/A3", "3"),
        ("/Sheet1/A4", "5"),
        ("/Sheet1/E1", "10"),
        ("/Sheet1/F1", "70"),
        ("/Sheet1/G1", "2078"),
        ("/Sheet1/H1", "80"),
        ("/Sheet1/I1", "1"),
        ("/Sheet1/J1", "999"),
        ("/Sheet1/K1", "999"),
    ] {
        assert_eq!(
            values[path],
            SpreadsheetFormulaValue::Number {
                value: expected.into()
            },
            "{path}"
        );
    }
    let artifact = NativeOfficeReplayArtifact::dump(&editor.snapshot().unwrap(), "/").unwrap();
    let mut restored =
        NativeOfficeEditor::create(temp.path().join("table-references-restored.xlsx"))
            .await
            .unwrap();
    restored.apply_replay(&artifact).unwrap();
    assert_eq!(
        restored.package().content_sha256(),
        editor.package().content_sha256()
    );
}

#[tokio::test]
async fn calculation_rejects_table_local_and_missing_item_rows_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let mut outside = NativeOfficeEditor::create(temp.path().join("table-local-outside.xlsx"))
        .await
        .unwrap();
    outside
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C4", ["Item", "Qty", "Price"]),
        )
        .unwrap();
    outside
        .set_cell_value("/Sheet1/E2", formula("SUM([@Qty])"))
        .unwrap();
    let graph = outside
        .snapshot()
        .unwrap()
        .formula_dependency_graph()
        .unwrap();
    let node = graph
        .nodes
        .iter()
        .find(|node| node.cell.path() == "/Sheet1/E2")
        .unwrap();
    assert_eq!(
        node.unresolved_references[0].kind,
        crate::SpreadsheetFormulaUnresolvedReferenceKind::StructuredReference
    );
    let before = outside.package().content_sha256();
    let error = outside.recalculate_spreadsheet_formulas().unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_formula_structured_reference_unsupported"
    );
    assert_eq!(outside.package().content_sha256(), before);

    for (file_name, table, expression) in [
        (
            "table-missing-header.xlsx",
            NativeSpreadsheetTable::new("Data", "A1:B3", ["Name", "Value"]).with_header_row(false),
            "SUM(Data[#Headers])",
        ),
        (
            "table-missing-totals.xlsx",
            NativeSpreadsheetTable::new("Data", "A1:B3", ["Name", "Value"]),
            "SUM(Data[#Totals])",
        ),
    ] {
        let mut editor = NativeOfficeEditor::create(temp.path().join(file_name))
            .await
            .unwrap();
        editor.add_spreadsheet_table("/Sheet1", table).unwrap();
        editor
            .set_cell_value("/Sheet1/D1", formula(expression))
            .unwrap();
        let before = editor.package().content_sha256();
        let error = editor.recalculate_spreadsheet_formulas().unwrap_err();
        assert_eq!(
            error.code, "use.office.spreadsheet_formula_structured_reference_unsupported",
            "{expression}"
        );
        assert_eq!(editor.package().content_sha256(), before, "{expression}");
    }
}

#[tokio::test]
async fn table_mutation_rewrites_structured_formula_identities_and_columns() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("table-formula-rewrite.xlsx"))
        .await
        .unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C4", ["Item", "Qty", "Price"])
                .with_display_name("SalesView"),
        )
        .unwrap();
    editor.set_cell_value("/Sheet1/B2", number("2")).unwrap();
    editor.set_cell_value("/Sheet1/B3", number("3")).unwrap();
    editor
        .set_cell_value("/Sheet1/A2", formula("[@Qty]"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/E1", formula("SUM(SalesView[Qty])"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/F1", formula("SUM(Sales[Qty])"))
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/G1",
            formula("CONCAT(\"Sales[Qty]=\",SUM(Sales[Qty]))"),
        )
        .unwrap();
    editor
        .add_named_range(NativeSpreadsheetNamedRange::new(
            "TableQuantity",
            "SUM(Sales[Qty])",
        ))
        .unwrap();
    editor
        .add_conditional_format(
            "/Sheet1",
            NativeSpreadsheetConditionalFormat::new(
                "B2:B4",
                NativeSpreadsheetConditionalFormatRule::Formula {
                    formula: "SalesView[Qty]>0".into(),
                    format: NativeSpreadsheetDifferentialFormat::default(),
                },
            ),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Custom,
                "C2:C4",
                "SUM(Sales[Qty])>0",
            ),
        )
        .unwrap();
    editor.recalculate_spreadsheet_formulas().unwrap();

    editor
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Orders", "A1:C4", ["Product", "Units", "Cost"])
                .with_display_name("OrdersView"),
        )
        .unwrap();

    let snapshot = editor.snapshot().unwrap();
    for (path, expected) in [
        ("/Sheet1/A2", "[@Units]"),
        ("/Sheet1/E1", "SUM(OrdersView[Units])"),
        ("/Sheet1/F1", "SUM(Orders[Units])"),
        ("/Sheet1/G1", "CONCAT(\"Sales[Qty]=\",SUM(Orders[Units]))"),
    ] {
        assert_eq!(snapshot.get(path, 0).unwrap().format["formula"], expected);
    }
    assert_eq!(
        snapshot
            .get("/namedrange[@name=TableQuantity][@scope=workbook]", 0,)
            .unwrap()
            .format["ref"],
        "SUM(Orders[Units])"
    );
    assert_eq!(
        snapshot.get("/Sheet1/cf[1]", 0).unwrap().format["formula"],
        "OrdersView[Units]>0"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[1]", 0).unwrap().format["formula1"],
        "SUM(Orders[Units])>0"
    );

    let calculation = editor.recalculate_spreadsheet_formulas().unwrap();
    let values = calculation
        .cells
        .iter()
        .map(|cell| (cell.cell.path(), cell.value.clone()))
        .collect::<std::collections::BTreeMap<_, _>>();
    for path in ["/Sheet1/E1", "/Sheet1/F1"] {
        assert_eq!(
            values[path],
            SpreadsheetFormulaValue::Number { value: "5".into() },
            "{path}"
        );
    }
    assert_eq!(
        values["/Sheet1/G1"],
        SpreadsheetFormulaValue::Text {
            value: "Sales[Qty]=5".into()
        }
    );

    let artifact = NativeOfficeReplayArtifact::dump(&editor.snapshot().unwrap(), "/").unwrap();
    let mut restored =
        NativeOfficeEditor::create(temp.path().join("table-formula-rewrite-restored.xlsx"))
            .await
            .unwrap();
    restored.apply_replay(&artifact).unwrap();
    assert_eq!(
        restored.package().content_sha256(),
        editor.package().content_sha256()
    );
}

#[tokio::test]
async fn table_mutation_rejects_referenced_removal_and_unsafe_local_geometry_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let mut referenced = NativeOfficeEditor::create(temp.path().join("referenced-table.xlsx"))
        .await
        .unwrap();
    referenced
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C4", ["Item", "Qty", "Price"]),
        )
        .unwrap();
    referenced
        .set_cell_value("/Sheet1/E1", formula("SUM(Sales[Qty])"))
        .unwrap();
    let before = referenced.package().content_sha256();
    let error = referenced.remove("/Sheet1/table[1]").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_table_referenced");
    assert_eq!(referenced.package().content_sha256(), before);

    let mut local = NativeOfficeEditor::create(temp.path().join("local-table-geometry.xlsx"))
        .await
        .unwrap();
    local
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:C4", ["Item", "Qty", "Price"]),
        )
        .unwrap();
    local
        .set_cell_value("/Sheet1/A2", formula("[@Qty]"))
        .unwrap();
    let before = local.package().content_sha256();
    let error = local
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Sales", "A3:C6", ["Item", "Qty", "Price"]),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_table_formula_rewrite_unsupported"
    );
    assert_eq!(local.package().content_sha256(), before);
}

#[tokio::test]
async fn table_mutation_rewrites_chart_and_other_table_formula_carriers() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("table-formula-carriers.xlsx"))
        .await
        .unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:B3", ["Item", "Qty"]),
        )
        .unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Summary", "D1:D3", ["Derived"]),
        )
        .unwrap();
    let chart = editor
        .add_part("/Sheet1", NativeOfficePartType::Chart)
        .unwrap();
    editor
        .replace_xml_part(
            &chart.part,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Sales[Qty]</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )
        .unwrap();

    let mut package = editor.package().clone();
    let table_part = "xl/tables/table2.xml";
    let table_xml = std::str::from_utf8(package.part(table_part).unwrap()).unwrap();
    let table_xml = table_xml.replace(
        r#"<tableColumn id="1" name="Derived"/>"#,
        r#"<tableColumn id="1" name="Derived"><calculatedColumnFormula>SUM(Sales[Qty])</calculatedColumnFormula></tableColumn>"#,
    );
    assert!(table_xml.contains("calculatedColumnFormula"));
    package
        .set_part(table_part, table_xml.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Orders", "A1:B3", ["Product", "Units"]),
        )
        .unwrap();

    let chart_xml = std::str::from_utf8(
        editor
            .package()
            .part(chart.part.trim_start_matches('/'))
            .unwrap(),
    )
    .unwrap();
    assert!(chart_xml.contains("<c:f>Orders[Units]</c:f>"));
    let table_xml = std::str::from_utf8(editor.package().part(table_part).unwrap()).unwrap();
    assert!(
        table_xml.contains("<calculatedColumnFormula>SUM(Orders[Units])</calculatedColumnFormula>")
    );
}

#[tokio::test]
async fn table_geometry_changes_clear_formula_and_chart_caches() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("table-geometry-caches.xlsx"))
        .await
        .unwrap();
    editor
        .add_spreadsheet_table(
            "/Sheet1",
            NativeSpreadsheetTable::new("Sales", "A1:B3", ["Item", "Qty"]),
        )
        .unwrap();
    editor.set_cell_value("/Sheet1/B2", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/D1", formula("SUM(Sales[Qty])"))
        .unwrap();
    editor.recalculate_spreadsheet_formulas().unwrap();
    let chart = editor
        .add_part("/Sheet1", NativeOfficePartType::Chart)
        .unwrap();
    editor
        .replace_xml_part(
            &chart.part,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Sales[Qty]</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )
        .unwrap();

    editor
        .set_spreadsheet_table(
            "/Sheet1/table[1]",
            NativeSpreadsheetTable::new("Sales", "A1:B4", ["Item", "Qty"]),
        )
        .unwrap();

    let worksheet = editor
        .package()
        .xml_part("xl/worksheets/sheet1.xml")
        .unwrap();
    let worksheet = crate::xml_edit::index_xml(&worksheet).unwrap();
    let mut cells = Vec::new();
    worksheet.descendants_named("c", &mut cells);
    let formula_cell = cells
        .into_iter()
        .find(|cell| cell.attributes.get("r").map(String::as_str) == Some("D1"))
        .unwrap();
    assert!(formula_cell
        .children
        .iter()
        .any(|child| child.local_name == "f"));
    assert!(!formula_cell
        .children
        .iter()
        .any(|child| child.local_name == "v"));

    let chart_xml = std::str::from_utf8(
        editor
            .package()
            .part(chart.part.trim_start_matches('/'))
            .unwrap(),
    )
    .unwrap();
    assert!(chart_xml.contains("<c:f>Sales[Qty]</c:f>"));
    assert!(!chart_xml.contains("numCache"));
}
