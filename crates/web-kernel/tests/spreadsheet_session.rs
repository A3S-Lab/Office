use a3s_office_web_kernel::{
    calculate_spreadsheet_session, SpreadsheetCalculationSession,
    SpreadsheetCalculationSessionCellChange, SpreadsheetCalculationSessionRequest,
    SpreadsheetCalculationSessionScope, SpreadsheetCalculationSessionUpdate, SpreadsheetCoordinate,
    SpreadsheetInputCell, SpreadsheetInputSheet, SpreadsheetInputTable, SpreadsheetValue,
    OFFICE_KERNEL_PROTOCOL_VERSION,
};

#[test]
fn recalculates_only_the_dirty_dependency_subgraph() {
    assert_send_sync::<SpreadsheetCalculationSession>();
    let mut session = SpreadsheetCalculationSession::default();
    let initial = calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    assert_eq!(
        initial
            .cells
            .iter()
            .map(|cell| (cell.column, cell.value.clone()))
            .collect::<Vec<_>>(),
        vec![(1, number(4.0)), (2, number(5.0)), (3, number(10.0)),]
    );
    assert_eq!(initial.stats.formula_cell_count, 3);
    assert_eq!(initial.stats.dirty_formula_cell_count, 3);
    assert_eq!(initial.stats.evaluated_formula_cell_count, 3);
    assert_eq!(initial.stats.reused_formula_cell_count, 0);
    assert_eq!(initial.stats.dependency_edge_count, 2);

    let patched =
        calculate_spreadsheet_session(&mut session, &input_patch_request(1, 2, 3.0)).unwrap();

    assert_eq!(
        patched
            .cells
            .iter()
            .map(|cell| (cell.column, cell.value.clone()))
            .collect::<Vec<_>>(),
        vec![(1, number(6.0)), (2, number(7.0))]
    );
    assert_eq!(patched.stats.formula_cell_count, 3);
    assert_eq!(patched.stats.dirty_formula_cell_count, 2);
    assert_eq!(patched.stats.evaluated_formula_cell_count, 2);
    assert_eq!(patched.stats.reused_formula_cell_count, 0);
    assert_eq!(patched.stats.dependency_edge_count, 2);
}

#[test]
fn keeps_unrelated_edits_out_of_the_formula_evaluator() {
    let mut session = SpreadsheetCalculationSession::default();
    calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    let result = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 2,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sheet-1".into(),
                    row: 0,
                    column: 4,
                    formula: None,
                    value: number(99.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap();

    assert!(result.cells.is_empty());
    assert!(result.calculation_order.is_empty());
    assert_eq!(result.stats.formula_cell_count, 3);
    assert_eq!(result.stats.dirty_formula_cell_count, 0);
    assert_eq!(result.stats.evaluated_formula_cell_count, 0);
    assert_eq!(result.stats.dependency_edge_count, 2);
}

#[test]
fn rewires_formula_dependencies_without_rebuilding_the_workbook() {
    let mut session = SpreadsheetCalculationSession::default();
    calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    let rewired = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 2,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sheet-1".into(),
                    row: 0,
                    column: 1,
                    formula: Some("=D1*2".into()),
                    value: number(4.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap();

    assert_eq!(
        rewired
            .cells
            .iter()
            .map(|cell| (cell.column, cell.value.clone()))
            .collect::<Vec<_>>(),
        vec![(1, number(20.0)), (2, number(21.0))]
    );
    assert_eq!(rewired.stats.dirty_formula_cell_count, 2);
    assert_eq!(rewired.stats.evaluated_formula_cell_count, 2);
    assert_eq!(rewired.stats.reused_formula_cell_count, 1);
    assert_eq!(rewired.stats.dependency_edge_count, 2);

    let old_input =
        calculate_spreadsheet_session(&mut session, &input_patch_request(2, 3, 8.0)).unwrap();
    assert!(old_input.cells.is_empty());
    assert_eq!(old_input.stats.evaluated_formula_cell_count, 0);

    let new_input = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 4,
            revision: 4,
            document_revision: 4,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 3,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sheet-1".into(),
                    row: 0,
                    column: 3,
                    formula: Some("=11".into()),
                    value: number(10.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap();
    assert_eq!(
        new_input
            .cells
            .iter()
            .map(|cell| (cell.column, cell.value.clone()))
            .collect::<Vec<_>>(),
        vec![(3, number(11.0)), (1, number(22.0)), (2, number(23.0)),]
    );
    assert_eq!(new_input.stats.evaluated_formula_cell_count, 3);
}

#[test]
fn rejects_stale_patches_without_mutating_the_session() {
    let mut session = SpreadsheetCalculationSession::default();
    calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    let mut stale = input_patch_request(0, 2, 9.0);
    stale.document_revision = 2;
    let error = calculate_spreadsheet_session(&mut session, &stale).unwrap_err();
    assert_eq!(
        error.code,
        "office.kernel.spreadsheet.session_revision_mismatch"
    );

    let valid =
        calculate_spreadsheet_session(&mut session, &input_patch_request(1, 2, 3.0)).unwrap();
    assert_eq!(
        valid.cells.last().map(|cell| &cell.value),
        Some(&number(7.0))
    );
}

#[test]
fn tracks_references_to_cells_that_were_blank_when_the_graph_was_built() {
    let mut session = SpreadsheetCalculationSession::default();
    let initial = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 1,
            revision: 1,
            document_revision: 1,
            update: SpreadsheetCalculationSessionUpdate::Replace {
                sheets: vec![SpreadsheetInputSheet {
                    id: "sheet-1".into(),
                    name: "Sheet 1".into(),
                    cells: vec![input_cell(0, 1, Some("=A1+1"), SpreadsheetValue::Blank)],
                    tables: vec![],
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Workbook,
        },
    )
    .unwrap();
    assert_eq!(initial.cells[0].value, number(1.0));
    assert_eq!(initial.stats.dependency_edge_count, 1);

    let patched =
        calculate_spreadsheet_session(&mut session, &input_patch_request(1, 2, 4.0)).unwrap();
    assert_eq!(patched.cells[0].value, number(5.0));
    assert_eq!(patched.stats.dirty_formula_cell_count, 1);
}

#[test]
fn keeps_unresolved_formulas_and_their_dependents_conservatively_dirty() {
    let mut session = SpreadsheetCalculationSession::default();
    let initial = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 1,
            revision: 1,
            document_revision: 1,
            update: SpreadsheetCalculationSessionUpdate::Replace {
                sheets: vec![SpreadsheetInputSheet {
                    id: "sheet-1".into(),
                    name: "Sheet 1".into(),
                    cells: vec![
                        input_cell(0, 0, None, number(2.0)),
                        input_cell(0, 1, Some("=A3S_UNKNOWN(A1)"), number(41.0)),
                        input_cell(0, 2, Some("=B1+1"), number(42.0)),
                    ],
                    tables: vec![],
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Workbook,
        },
    )
    .unwrap();
    assert_eq!(initial.issues.len(), 2);

    let patched = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 2,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sheet-1".into(),
                    row: 0,
                    column: 3,
                    formula: None,
                    value: number(99.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap();
    assert_eq!(patched.stats.dirty_formula_cell_count, 2);
    assert_eq!(patched.stats.evaluated_formula_cell_count, 2);
    assert_eq!(patched.issues.len(), 2);
}

#[test]
fn empty_patches_cannot_move_the_session_revision_backwards() {
    let mut session = SpreadsheetCalculationSession::default();
    calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    let error = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 0,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap_err();
    assert_eq!(
        error.code,
        "office.kernel.spreadsheet.session_revision_invalid"
    );

    let valid =
        calculate_spreadsheet_session(&mut session, &input_patch_request(1, 2, 3.0)).unwrap();
    assert_eq!(
        valid.cells.last().map(|cell| &cell.value),
        Some(&number(7.0))
    );
}

#[test]
fn target_calculation_preserves_unvisited_dirty_dependents() {
    let mut session = SpreadsheetCalculationSession::default();
    calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    let targeted = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 2,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sheet-1".into(),
                    row: 0,
                    column: 0,
                    formula: None,
                    value: number(3.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Targets {
                targets: vec![coordinate(0, 1)],
            },
        },
    )
    .unwrap();
    assert_eq!(
        targeted
            .cells
            .iter()
            .map(|cell| (cell.column, cell.value.clone()))
            .collect::<Vec<_>>(),
        vec![(1, number(6.0))]
    );

    let remaining = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 3,
            revision: 3,
            document_revision: 3,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 2,
                changes: vec![],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap();
    assert_eq!(
        remaining
            .cells
            .iter()
            .map(|cell| (cell.column, cell.value.clone()))
            .collect::<Vec<_>>(),
        vec![(2, number(7.0))]
    );
}

#[test]
fn retains_table_catalog_across_incremental_cell_patches() {
    let mut session = SpreadsheetCalculationSession::default();
    let initial = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 1,
            revision: 1,
            document_revision: 1,
            update: SpreadsheetCalculationSessionUpdate::Replace {
                sheets: vec![SpreadsheetInputSheet {
                    id: "sales".into(),
                    name: "Sales".into(),
                    cells: vec![
                        input_cell(0, 0, None, number(2.0)),
                        input_cell(1, 0, None, number(3.0)),
                        input_cell(0, 1, Some("=SUM(Sales[Quantity])"), SpreadsheetValue::Blank),
                    ],
                    tables: vec![SpreadsheetInputTable {
                        name: "Sales".into(),
                        display_name: None,
                        start_row: 0,
                        end_row: 1,
                        start_column: 0,
                        end_column: 0,
                        columns: vec!["Quantity".into()],
                        header_row: false,
                        totals_row: false,
                    }],
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Workbook,
        },
    )
    .unwrap();
    assert_eq!(initial.cells[0].value, number(5.0));
    assert_eq!(initial.stats.dependency_edge_count, 2);

    let patched = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 2,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sales".into(),
                    row: 0,
                    column: 0,
                    formula: None,
                    value: number(4.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Dirty,
        },
    )
    .unwrap();
    assert_eq!(patched.cells[0].value, number(7.0));
    assert_eq!(patched.stats.dependency_edge_count, 2);
}

#[test]
fn invalid_targets_do_not_advance_or_replace_the_session() {
    let mut session = SpreadsheetCalculationSession::default();
    calculate_spreadsheet_session(&mut session, &replace_request()).unwrap();

    let error = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: 2,
            revision: 2,
            document_revision: 2,
            update: SpreadsheetCalculationSessionUpdate::Patch {
                base_document_revision: 1,
                changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                    sheet_id: "sheet-1".into(),
                    row: 0,
                    column: 0,
                    formula: None,
                    value: number(99.0),
                }],
            },
            calculation: SpreadsheetCalculationSessionScope::Targets {
                targets: vec![SpreadsheetCoordinate {
                    sheet_id: "missing".into(),
                    row: 0,
                    column: 0,
                }],
            },
        },
    )
    .unwrap_err();
    assert_eq!(error.code, "office.kernel.spreadsheet.target_invalid");

    let valid =
        calculate_spreadsheet_session(&mut session, &input_patch_request(1, 2, 3.0)).unwrap();
    assert_eq!(
        valid.cells.last().map(|cell| &cell.value),
        Some(&number(7.0))
    );
}

fn replace_request() -> SpreadsheetCalculationSessionRequest {
    SpreadsheetCalculationSessionRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "spreadsheetSessionCalculation".into(),
        request_id: 1,
        revision: 1,
        document_revision: 1,
        update: SpreadsheetCalculationSessionUpdate::Replace {
            sheets: vec![SpreadsheetInputSheet {
                id: "sheet-1".into(),
                name: "Sheet 1".into(),
                cells: vec![
                    input_cell(0, 0, None, number(2.0)),
                    input_cell(0, 1, Some("=A1*2"), SpreadsheetValue::Blank),
                    input_cell(0, 2, Some("=B1+1"), SpreadsheetValue::Blank),
                    input_cell(0, 3, Some("=10"), SpreadsheetValue::Blank),
                ],
                tables: vec![],
            }],
        },
        calculation: SpreadsheetCalculationSessionScope::Workbook,
    }
}

fn input_patch_request(
    base_document_revision: u64,
    document_revision: u64,
    value: f64,
) -> SpreadsheetCalculationSessionRequest {
    SpreadsheetCalculationSessionRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "spreadsheetSessionCalculation".into(),
        request_id: u32::try_from(document_revision).unwrap(),
        revision: u32::try_from(document_revision).unwrap(),
        document_revision,
        update: SpreadsheetCalculationSessionUpdate::Patch {
            base_document_revision,
            changes: vec![SpreadsheetCalculationSessionCellChange::Upsert {
                sheet_id: "sheet-1".into(),
                row: 0,
                column: 0,
                formula: None,
                value: number(value),
            }],
        },
        calculation: SpreadsheetCalculationSessionScope::Dirty,
    }
}

fn input_cell(
    row: u32,
    column: u32,
    formula: Option<&str>,
    value: SpreadsheetValue,
) -> SpreadsheetInputCell {
    SpreadsheetInputCell {
        row,
        column,
        formula: formula.map(str::to_owned),
        value,
    }
}

fn number(value: f64) -> SpreadsheetValue {
    SpreadsheetValue::Number { value }
}

fn coordinate(row: u32, column: u32) -> SpreadsheetCoordinate {
    SpreadsheetCoordinate {
        sheet_id: "sheet-1".into(),
        row,
        column,
    }
}

fn assert_send_sync<T: Send + Sync>() {}
