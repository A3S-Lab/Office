use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeSpreadsheetNamedRange, SpreadsheetCellValue,
    SpreadsheetFormulaDependencyGraph, SpreadsheetFormulaUnresolvedReferenceKind,
    MAX_SPREADSHEET_FORMULA_DEPTH, MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS,
};

fn number(value: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Number {
        value: value.to_string(),
    }
}

fn formula(expression: &str) -> SpreadsheetCellValue {
    SpreadsheetCellValue::Formula {
        expression: expression.to_string(),
    }
}

#[tokio::test]
async fn dependency_graph_orders_cross_sheet_ranges_and_named_references() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("graph.xlsx"))
        .await
        .unwrap();
    editor.add_worksheet("Data").unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/B1", formula("A1*2"))
        .unwrap();
    editor
        .set_cell_value("/Data/A1", formula("Sheet1!B1+1"))
        .unwrap();
    editor
        .add_named_range(NativeSpreadsheetNamedRange::new(
            "Inputs",
            "'Sheet1'!$B$1:$B$3",
        ))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/B3", formula("B1+1"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/C1", formula("SUM(Inputs)+Data!A1"))
        .unwrap();

    let graph = editor
        .snapshot()
        .unwrap()
        .formula_dependency_graph()
        .unwrap();
    assert_eq!(graph.nodes.len(), 4);
    assert!(graph.cycles.is_empty());
    assert_eq!(
        graph
            .calculation_order
            .iter()
            .map(|cell| cell.path())
            .collect::<Vec<_>>(),
        ["/Sheet1/B1", "/Sheet1/B3", "/Data/A1", "/Sheet1/C1"]
    );
    let target = graph
        .nodes
        .iter()
        .find(|node| node.cell.path() == "/Sheet1/C1")
        .unwrap();
    assert_eq!(
        target
            .dependencies
            .iter()
            .map(|cell| cell.path())
            .collect::<Vec<_>>(),
        ["/Sheet1/B1", "/Sheet1/B3", "/Data/A1"]
    );
    assert!(target.unresolved_references.is_empty());
}

#[tokio::test]
async fn dependency_graph_reports_stable_cycles_and_unresolved_references() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("cycles.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("B1+Missing!A1"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/B1", formula("A1+UnknownName"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/C1", formula("[Book.xlsx]Data!A1"))
        .unwrap();

    let graph = editor
        .snapshot()
        .unwrap()
        .formula_dependency_graph()
        .unwrap();
    assert_eq!(
        graph
            .cycles
            .iter()
            .map(|cycle| cycle.iter().map(|cell| cell.path()).collect::<Vec<_>>())
            .collect::<Vec<_>>(),
        [vec!["/Sheet1/A1", "/Sheet1/B1"]]
    );
    assert_eq!(
        graph
            .nodes
            .iter()
            .find(|node| node.cell.path() == "/Sheet1/A1")
            .unwrap()
            .unresolved_references[0]
            .kind,
        SpreadsheetFormulaUnresolvedReferenceKind::MissingWorksheet
    );
    assert_eq!(
        graph
            .nodes
            .iter()
            .find(|node| node.cell.path() == "/Sheet1/B1")
            .unwrap()
            .unresolved_references[0]
            .kind,
        SpreadsheetFormulaUnresolvedReferenceKind::UndefinedName
    );
    assert_eq!(
        graph
            .nodes
            .iter()
            .find(|node| node.cell.path() == "/Sheet1/C1")
            .unwrap()
            .unresolved_references[0]
            .kind,
        SpreadsheetFormulaUnresolvedReferenceKind::ExternalWorkbook
    );
}

#[tokio::test]
async fn dependency_graph_and_calculation_bound_nested_named_references() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("named-depth.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/B1", number("1")).unwrap();
    let mutations = (0..=MAX_SPREADSHEET_FORMULA_DEPTH)
        .map(|index| {
            let reference = if index == MAX_SPREADSHEET_FORMULA_DEPTH {
                "'Sheet1'!$B$1".to_string()
            } else {
                format!("Chain_{}", index + 1)
            };
            NativeOfficeMutation::AddNamedRange {
                named_range: NativeSpreadsheetNamedRange::new(format!("Chain_{index}"), reference),
            }
        })
        .collect::<Vec<_>>();
    editor.apply_batch(&mutations).unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("Chain_0"))
        .unwrap();

    let graph = editor
        .snapshot()
        .unwrap()
        .formula_dependency_graph()
        .unwrap();
    assert_eq!(
        graph.nodes[0].unresolved_references[0].kind,
        SpreadsheetFormulaUnresolvedReferenceKind::NamedRangeDepth
    );
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_formula_named_reference_depth"
    );
    assert_eq!(error.details["namedRange"], "Chain_128");
}

#[tokio::test]
async fn dependency_graph_bounds_overlapping_reference_candidate_visits() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("reference-visits.xlsx"))
        .await
        .unwrap();
    const FORMULAS: usize = 4_000;
    editor
        .set_cell_value(
            format!("/Sheet1/A1:A{FORMULAS}"),
            SpreadsheetCellValue::Formula {
                expression: "1".into(),
            },
        )
        .unwrap();
    let area_count = MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS / FORMULAS + 1;
    let areas = (0..area_count)
        .map(|offset| format!("A1:A{}", FORMULAS + offset))
        .collect::<Vec<_>>()
        .join(",");
    editor
        .set_cell_value("/Sheet1/B1", formula(&format!("SUM({areas})")))
        .unwrap();

    let error = editor
        .snapshot()
        .unwrap()
        .formula_dependency_graph()
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_formula_reference_visit_limit"
    );
    assert_eq!(
        error.details["visits"],
        MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS + 1
    );
}

#[test]
fn dependency_graph_contract_is_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SpreadsheetFormulaDependencyGraph>();
}
