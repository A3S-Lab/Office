use std::path::Path;

use serde_json::{json, Map as JsonMap, Value as JsonValue};
use yrs::{Any, Array, Map, Out, Transact};

use super::*;

const YJS_SPREADSHEET_UPDATE_BASE64: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../tests/fixtures/browser-spreadsheet-collaboration-update.base64"
));

#[test]
fn typed_spreadsheet_cells_merge_leaves_preserve_projection_and_survive_restart() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-replica");
    let store = initialized_spreadsheet_store(&root);
    let original = json!({
        "v": 10,
        "m": "10",
        "ct": { "fa": "0.00", "t": "n" },
    });
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), Some(10.0));

    let before_stale_create = store.inspect().unwrap();
    let stale_create = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-create-over-existing-cell",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: None,
                next_cell: original.clone(),
            },
        ))
        .unwrap_err();
    assert_eq!(
        stale_create.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_stale_create.document_state_sha256
    );

    let value_request = spreadsheet_mutation_request(
        "spreadsheet-set-value-1",
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: "sheet-data".to_owned(),
            row: 1,
            column: 0,
            expected_cell: Some(original.clone()),
            next_cell: json!({
                "v": 12,
                "m": "12",
                "f": "=6*2",
                "ct": { "fa": "0.00", "t": "n" },
            }),
        },
    );
    let value = store.mutate(value_request.clone()).unwrap();
    assert!(value.state_changed);
    assert_eq!(value.sequence, Some(2));

    let style = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-style-1",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: Some(original.clone()),
                next_cell: json!({
                    "v": 10,
                    "m": "10",
                    "bg": "#DBEAFE",
                    "ct": { "fa": "0.00", "t": "n" },
                }),
            },
        ))
        .unwrap();
    assert!(style.state_changed);
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), Some(12.0));
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 0, &["bg"]),
        Some("#DBEAFE".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 0, &["ct", "fa"]),
        Some("0.00".to_owned())
    );

    let before_conflict = store.inspect().unwrap();
    let conflict = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-conflict",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: Some(original.clone()),
                next_cell: json!({
                    "v": 99,
                    "m": "99",
                    "ct": { "fa": "0.00", "t": "n" },
                }),
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_conflict.document_state_sha256
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-create-dense-1",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 1,
                expected_cell: None,
                next_cell: json!({ "v": 20, "m": "20" }),
            },
        ))
        .unwrap();
    assert_eq!(
        spreadsheet_mode(&store, "sheet-data"),
        Some("data".to_owned())
    );
    assert_eq!(spreadsheet_row_lengths(&store, "sheet-data"), vec![1, 2]);

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-extend-dense-1",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 3,
                column: 4,
                expected_cell: None,
                next_cell: json!({ "v": "far", "m": "far" }),
            },
        ))
        .unwrap();
    assert_eq!(
        spreadsheet_row_lengths(&store, "sheet-data"),
        vec![1, 2, 0, 5]
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-create-empty-sheet-cell-1",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-empty".to_owned(),
                row: 100,
                column: 5,
                expected_cell: None,
                next_cell: json!({
                    "v": "sparse native",
                    "m": "sparse native",
                    "ps": { "value": "Agent note", "isShow": false },
                }),
            },
        ))
        .unwrap();
    assert_eq!(
        spreadsheet_mode(&store, "sheet-empty"),
        Some("celldata".to_owned())
    );
    assert!(spreadsheet_row_lengths(&store, "sheet-empty").is_empty());
    assert_eq!(
        cell_string(&store, "sheet-empty", 100, 5, &["ps", "value"]),
        Some("Agent note".to_owned())
    );

    let current = json!({
        "v": 12,
        "m": "12",
        "f": "=6*2",
        "bg": "#DBEAFE",
        "ct": { "fa": "0.00", "t": "n" },
    });
    let deleted = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-cell-1",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: current.clone(),
            },
        ))
        .unwrap();
    assert!(deleted.state_changed);
    assert!(!spreadsheet_cell_present(&store, "sheet-data", 1, 0));
    assert_eq!(
        spreadsheet_row_lengths(&store, "sheet-data"),
        vec![1, 2, 0, 5]
    );
    let delete_retry = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-cell-retry",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: current,
            },
        ))
        .unwrap();
    assert!(!delete_retry.state_changed);
    assert_eq!(delete_retry.sequence, None);

    let replay = store.mutate(value_request).unwrap();
    assert!(replay.duplicate);
    assert_eq!(replay.sequence, Some(2));
    drop(store);

    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert!(!spreadsheet_cell_present(&reopened, "sheet-data", 1, 0));
    assert_eq!(
        cell_string(&reopened, "sheet-empty", 100, 5, &["v"]),
        Some("sparse native".to_owned())
    );
}

#[test]
fn typed_spreadsheet_batch_cells_commit_one_dense_gesture_and_survive_restart() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-replica");
    let store = initialized_spreadsheet_store(&root);
    let original = json!({
        "v": 10,
        "m": "10",
        "ct": { "fa": "0.00", "t": "n" },
    });

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-concurrent-number-format",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: Some(original.clone()),
                next_cell: json!({
                    "v": 10,
                    "m": "10",
                    "ct": { "fa": "0.000", "t": "n" },
                }),
            },
        ))
        .unwrap();
    let disposable = json!({ "v": "remove", "m": "remove" });
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-create-disposable-dense-cell",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 0,
                column: 1,
                expected_cell: None,
                next_cell: disposable.clone(),
            },
        ))
        .unwrap();

    let request = spreadsheet_mutation_request(
        "spreadsheet-batch-dense-gesture",
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet-data".to_owned(),
            changes: vec![
                spreadsheet_cell_change(
                    1,
                    0,
                    Some(original),
                    Some(json!({
                        "v": 12,
                        "m": "12",
                        "f": "=6*2",
                        "ct": { "fa": "0.00", "t": "s" },
                    })),
                ),
                spreadsheet_cell_change(3, 4, None, Some(json!({ "v": "far", "m": "far" }))),
                spreadsheet_cell_change(0, 1, Some(disposable), None),
            ],
        },
    );
    let result = store.mutate(request.clone()).unwrap();
    assert!(result.state_changed);
    assert_eq!(result.sequence, Some(4));
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), Some(12.0));
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 0, &["ct", "fa"]),
        Some("0.000".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 0, &["ct", "t"]),
        Some("s".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 3, 4, &["v"]),
        Some("far".to_owned())
    );
    assert!(!spreadsheet_cell_present(&store, "sheet-data", 0, 1));
    assert_eq!(
        spreadsheet_row_lengths(&store, "sheet-data"),
        vec![2, 2, 0, 5]
    );
    drop(store);

    let reopened = NativeOfficeCollaborationStore::open(&root).unwrap();
    assert_eq!(
        cell_string(&reopened, "sheet-data", 1, 0, &["ct", "fa"]),
        Some("0.000".to_owned())
    );
    assert_eq!(
        spreadsheet_row_lengths(&reopened, "sheet-data"),
        vec![2, 2, 0, 5]
    );
    let replay = reopened.mutate(request).unwrap();
    assert!(replay.duplicate);
    assert_eq!(replay.sequence, Some(4));
}

#[test]
fn typed_spreadsheet_batch_cells_preserve_sparse_mode_and_fail_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-replica");
    let store = initialized_spreadsheet_store(&root);
    let sparse = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-batch-sparse-create",
            NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
                sheet_id: "sheet-empty".to_owned(),
                changes: vec![
                    spreadsheet_cell_change(
                        100,
                        5,
                        None,
                        Some(json!({ "v": "first", "m": "first" })),
                    ),
                    spreadsheet_cell_change(
                        200,
                        7,
                        None,
                        Some(json!({ "v": "second", "m": "second" })),
                    ),
                ],
            },
        ))
        .unwrap();
    assert!(sparse.state_changed);
    assert_eq!(
        spreadsheet_mode(&store, "sheet-empty"),
        Some("celldata".to_owned())
    );
    assert!(spreadsheet_row_lengths(&store, "sheet-empty").is_empty());

    let before = store.inspect().unwrap();
    let conflict = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-batch-atomic-conflict",
            NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
                sheet_id: "sheet-empty".to_owned(),
                changes: vec![
                    spreadsheet_cell_change(300, 9, None, Some(json!({ "v": "must-not-appear" }))),
                    spreadsheet_cell_change(
                        100,
                        5,
                        Some(json!({ "v": "stale", "m": "stale" })),
                        Some(json!({ "v": "conflict", "m": "conflict" })),
                    ),
                ],
            },
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);
    assert!(!spreadsheet_cell_present(&store, "sheet-empty", 300, 9));
    assert_eq!(
        cell_string(&store, "sheet-empty", 100, 5, &["v"]),
        Some("first".to_owned())
    );
}

#[test]
fn typed_spreadsheet_batch_cell_contract_rejects_ambiguous_or_unbounded_changes() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-replica");
    let store = initialized_spreadsheet_store(&root);
    let too_many = (0..4_097)
        .map(|row| spreadsheet_cell_change(row, 0, None, Some(json!({ "v": row }))))
        .collect();
    let invalid_values = vec![
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet-empty".to_owned(),
            changes: vec![],
        },
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet-empty".to_owned(),
            changes: vec![
                spreadsheet_cell_change(0, 0, None, Some(json!({ "v": 1 }))),
                spreadsheet_cell_change(0, 0, None, Some(json!({ "v": 2 }))),
            ],
        },
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet-empty".to_owned(),
            changes: vec![spreadsheet_cell_change(0, 0, None, None)],
        },
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet-empty".to_owned(),
            changes: vec![spreadsheet_cell_change(
                1_048_576,
                0,
                None,
                Some(json!({ "v": 1 })),
            )],
        },
        NativeOfficeCollaborationMutation::SpreadsheetBatchCells {
            sheet_id: "sheet-empty".to_owned(),
            changes: too_many,
        },
    ];
    let before = store.inspect().unwrap();
    assert!(
        serde_json::from_value::<NativeOfficeCollaborationMutation>(json!({
            "type": "spreadsheet-batch-cells",
            "sheetId": "sheet-empty",
            "changes": [{
                "row": 0,
                "column": 0,
                "expectedCell": { "v": "must-not-delete" }
            }]
        }))
        .is_err()
    );
    for (index, mutation) in invalid_values.into_iter().enumerate() {
        let error = store
            .mutate(spreadsheet_mutation_request(
                &format!("spreadsheet-invalid-batch-{index}"),
                mutation,
            ))
            .unwrap_err();
        assert!(
            matches!(
                error.code.as_str(),
                "office.collaboration.mutation_invalid"
                    | "office.collaboration.mutation_range_invalid"
            ),
            "unexpected error: {error:?}"
        );
    }
    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);
}

#[test]
fn typed_spreadsheet_cell_contract_is_bounded_kind_safe_and_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-replica");
    let store = initialized_spreadsheet_store(&root);
    let before = store.inspect().unwrap();

    let invalid_values = [
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: "sheet-data".to_owned(),
            row: 1_048_576,
            column: 0,
            expected_cell: None,
            next_cell: json!({ "v": 1 }),
        },
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: " sheet-data".to_owned(),
            row: 0,
            column: 0,
            expected_cell: None,
            next_cell: json!({ "v": 1 }),
        },
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: "sheet-data".to_owned(),
            row: 0,
            column: 0,
            expected_cell: Some(json!([])),
            next_cell: json!({ "v": 1 }),
        },
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: "sheet-data".to_owned(),
            row: 0,
            column: 0,
            expected_cell: None,
            next_cell: json!({ "__proto__": { "polluted": true } }),
        },
        NativeOfficeCollaborationMutation::SpreadsheetSetCell {
            sheet_id: "sheet-data".to_owned(),
            row: 0,
            column: 0,
            expected_cell: None,
            next_cell: deeply_nested_cell(),
        },
    ];
    for (index, mutation) in invalid_values.into_iter().enumerate() {
        let error = store
            .mutate(spreadsheet_mutation_request(
                &format!("spreadsheet-invalid-{index}"),
                mutation,
            ))
            .unwrap_err();
        assert!(
            matches!(
                error.code.as_str(),
                "office.collaboration.mutation_invalid"
                    | "office.collaboration.mutation_range_invalid"
            ),
            "unexpected error: {error:?}"
        );
    }
    let missing = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-missing-sheet",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "missing-sheet".to_owned(),
                row: 0,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": 1 }),
            },
        ))
        .unwrap_err();
    assert_eq!(missing.code, "office.collaboration.mutation_match_conflict");
    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);

    let markdown_root = temp.path().join("markdown-replica");
    let markdown = NativeOfficeCollaborationStore::create(create_request(&markdown_root)).unwrap();
    let error = markdown
        .mutate(mutation_request(
            "spreadsheet-on-markdown",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 0,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": 1 }),
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.mutation_kind_mismatch");
}

#[test]
fn typed_spreadsheet_mutations_reject_malformed_browser_roots_before_writing() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-replica");
    let store = initialized_spreadsheet_store(&root);
    let malformed = malformed_cell_mode_update();
    store
        .apply(spreadsheet_apply_request(
            "spreadsheet-malformed-browser-root",
            malformed,
        ))
        .unwrap();
    let before = store.inspect().unwrap();

    let error = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-after-malformed-root",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 1,
                expected_cell: None,
                next_cell: json!({ "v": 20 }),
            },
        ))
        .unwrap_err();
    assert_eq!(error.code, "office.collaboration.content_invalid");
    let after = store.inspect().unwrap();
    assert_eq!(after.current_sequence, before.current_sequence);
    assert_eq!(after.document_state_sha256, before.document_state_sha256);
}

fn initialized_spreadsheet_store(root: &Path) -> NativeOfficeCollaborationStore {
    let store = NativeOfficeCollaborationStore::create(spreadsheet_create_request(root)).unwrap();
    store
        .apply(spreadsheet_apply_request(
            "bootstrap-browser-spreadsheet",
            STANDARD
                .decode(YJS_SPREADSHEET_UPDATE_BASE64.trim())
                .unwrap(),
        ))
        .unwrap();
    store
}

fn spreadsheet_create_request(root: &Path) -> NativeOfficeCollaborationCreateRequest {
    NativeOfficeCollaborationCreateRequest {
        store: root.to_path_buf(),
        artifact_id: "fixture-spreadsheet".to_owned(),
        kind: NativeOfficeCollaborationArtifactKind::Spreadsheet,
        actor_id: "agent-spreadsheet".to_owned(),
        actor_kind: NativeOfficeCollaborationActorKind::Agent,
        mode: NativeOfficeCollaborationMode::Edit,
        operation_id: "create-spreadsheet-1".to_owned(),
        namespace: None,
        client_id: Some(900_003),
        initial_update: None,
    }
}

fn spreadsheet_apply_request(
    operation_id: &str,
    update: Vec<u8>,
) -> NativeOfficeCollaborationApplyRequest {
    NativeOfficeCollaborationApplyRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-spreadsheet".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-spreadsheet".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Spreadsheet,
        update,
        if_state_vector: None,
        origin: None,
    }
}

fn spreadsheet_mutation_request(
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> NativeOfficeCollaborationMutationRequest {
    NativeOfficeCollaborationMutationRequest {
        operation_id: operation_id.to_owned(),
        actor_id: "agent-spreadsheet".to_owned(),
        mode: NativeOfficeCollaborationMode::Edit,
        expected_artifact_id: "fixture-spreadsheet".to_owned(),
        expected_kind: NativeOfficeCollaborationArtifactKind::Spreadsheet,
        mutation,
        if_state_vector: None,
    }
}

fn spreadsheet_cell_change(
    row: u32,
    column: u32,
    expected_cell: Option<JsonValue>,
    next_cell: Option<JsonValue>,
) -> NativeOfficeCollaborationSpreadsheetCellChange {
    NativeOfficeCollaborationSpreadsheetCellChange {
        row,
        column,
        expected_cell,
        next_cell,
    }
}

fn replica_document(store: &NativeOfficeCollaborationStore) -> Doc {
    let exported = store.synchronize(None).unwrap();
    let peer = Doc::with_client_id(818_183);
    peer.transact_mut()
        .apply_update(Update::decode_v1(&exported.update).unwrap())
        .unwrap();
    peer
}

fn cell_field(
    store: &NativeOfficeCollaborationStore,
    sheet_id: &str,
    row: u32,
    column: u32,
    kind: &str,
    path: &[&str],
) -> Option<Any> {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let record = match sheets.get(&transaction, sheet_id) {
        Some(Out::YMap(record)) => record,
        value => panic!("unexpected sheet record: {value:?}"),
    };
    let fields = match record.get(&transaction, "cells") {
        Some(Out::YMap(fields)) => fields,
        value => panic!("unexpected cell field map: {value:?}"),
    };
    let mut identity = vec![kind.to_owned()];
    identity.extend(path.iter().map(|value| (*value).to_owned()));
    let flat_key = serde_json::to_string(&identity).unwrap();
    let encoded = serde_json::to_string(&(row, column, flat_key)).unwrap();
    match fields.get(&transaction, encoded.as_str()) {
        Some(Out::Any(value)) => Some(value),
        None => None,
        value => panic!("unexpected cell field value: {value:?}"),
    }
}

fn cell_number(
    store: &NativeOfficeCollaborationStore,
    sheet_id: &str,
    row: u32,
    column: u32,
    path: &[&str],
) -> Option<f64> {
    match cell_field(store, sheet_id, row, column, "value", path) {
        Some(Any::Number(value)) => Some(value),
        None => None,
        value => panic!("unexpected numeric cell field: {value:?}"),
    }
}

fn cell_string(
    store: &NativeOfficeCollaborationStore,
    sheet_id: &str,
    row: u32,
    column: u32,
    path: &[&str],
) -> Option<String> {
    match cell_field(store, sheet_id, row, column, "value", path) {
        Some(Any::String(value)) => Some(value.to_string()),
        None => None,
        value => panic!("unexpected string cell field: {value:?}"),
    }
}

fn spreadsheet_cell_present(
    store: &NativeOfficeCollaborationStore,
    sheet_id: &str,
    row: u32,
    column: u32,
) -> bool {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let Some(Out::YMap(record)) = sheets.get(&transaction, sheet_id) else {
        return false;
    };
    let Some(Out::YMap(presence)) = record.get(&transaction, "cellPresence") else {
        return false;
    };
    matches!(
        presence.get(&transaction, format!("{row}:{column}").as_str()),
        Some(Out::Any(Any::Bool(true)))
    )
}

fn spreadsheet_mode(store: &NativeOfficeCollaborationStore, sheet_id: &str) -> Option<String> {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let Some(Out::YMap(record)) = sheets.get(&transaction, sheet_id) else {
        return None;
    };
    match record.get(&transaction, "cellMode") {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        None => None,
        value => panic!("unexpected cell mode: {value:?}"),
    }
}

fn spreadsheet_row_lengths(store: &NativeOfficeCollaborationStore, sheet_id: &str) -> Vec<u32> {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let Some(Out::YMap(record)) = sheets.get(&transaction, sheet_id) else {
        return Vec::new();
    };
    let Some(Out::YArray(lengths)) = record.get(&transaction, "dataRowLengths") else {
        return Vec::new();
    };
    (0..lengths.len(&transaction))
        .map(|index| match lengths.get(&transaction, index) {
            Some(Out::Any(Any::Number(value))) => value as u32,
            value => panic!("unexpected row length: {value:?}"),
        })
        .collect()
}

fn deeply_nested_cell() -> JsonValue {
    let mut value = JsonValue::String("too deep".to_owned());
    for index in 0..130 {
        let mut object = JsonMap::new();
        object.insert(format!("level{index}"), value);
        value = JsonValue::Object(object);
    }
    json!({ "v": value })
}

fn malformed_cell_mode_update() -> Vec<u8> {
    let document = Doc::with_client_id(424_245);
    document
        .transact_mut()
        .apply_update(
            Update::decode_v1(
                &STANDARD
                    .decode(YJS_SPREADSHEET_UPDATE_BASE64.trim())
                    .unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    let before = document.transact().state_vector();
    let sheets = document.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = document.transact();
    let record = match sheets.get(&transaction, "sheet-data") {
        Some(Out::YMap(record)) => record,
        value => panic!("unexpected sheet record: {value:?}"),
    };
    drop(transaction);
    record.insert(&mut document.transact_mut(), "cellMode", "invalid");
    let update = document.transact().encode_state_as_update_v1(&before);
    update
}
