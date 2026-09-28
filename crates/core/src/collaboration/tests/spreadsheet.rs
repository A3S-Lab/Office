use std::collections::BTreeMap;
use std::path::Path;

use serde_json::{json, Map as JsonMap, Value as JsonValue};
use yrs::{Any, Array, Map, Out, Transact};

use super::*;

const YJS_SPREADSHEET_UPDATE_BASE64: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../tests/fixtures/browser-spreadsheet-collaboration-update.base64"
));

#[test]
fn spreadsheet_replica_export_writes_every_supported_cell_or_returns_no_package() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("export-replica"));
    store
        .mutate(spreadsheet_mutation_request(
            "export-seed-formula",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 1,
                expected_cell: None,
                next_cell: json!({"v": 2, "m": "2", "f": "=A2"}),
            },
        ))
        .unwrap();

    let data = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Old</t></is></c><c r=\"C1\"><v>99</v></c></row><row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"B2\"><f>=A1</f><v>1</v></c></row></sheetData></worksheet>";
    let sparse = "<worksheet><sheetData><row r=\"6\"><c r=\"D6\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>";
    let shared = "<sst><si><t>Hello</t></si><si><t>Other</t></si></sst>";
    let mut parts = BTreeMap::new();
    parts.insert(
        "xl/worksheets/sheet1.xml".to_owned(),
        data.as_bytes().to_vec(),
    );
    parts.insert(
        "xl/worksheets/sheet2.xml".to_owned(),
        sparse.as_bytes().to_vec(),
    );
    parts.insert(
        "xl/sharedStrings.xml".to_owned(),
        shared.as_bytes().to_vec(),
    );
    parts.insert("xl/styles.xml".to_owned(), b"<styles/>".to_vec());
    let sheet_parts = [
        ("sheet-data", "xl/worksheets/sheet1.xml"),
        ("sheet-sparse", "xl/worksheets/sheet2.xml"),
    ];

    let written = export_spreadsheet_replica(&store, parts.clone(), &sheet_parts).unwrap();
    let data_xml = String::from_utf8(written["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let sparse_xml = String::from_utf8(written["xl/worksheets/sheet2.xml"].clone()).unwrap();
    let shared_xml = String::from_utf8(written["xl/sharedStrings.xml"].clone()).unwrap();
    assert!(data_xml.contains("<t>Header</t>"));
    assert!(data_xml.contains("<v>10</v>"));
    assert!(data_xml.contains("<f>=A2</f>"));
    assert!(data_xml.contains("<v>2</v>"));
    assert!(data_xml.contains("<v>99</v>"));
    assert!(sparse_xml.contains("t=\"s\""));
    assert!(sparse_xml.contains("<v>0</v>"));
    assert!(shared_xml.contains("<t>edge</t>"));
    assert!(shared_xml.contains("<t>Other</t>"));
    assert!(!shared_xml.contains("<t>Hello</t>"));
    assert_eq!(written["xl/styles.xml"], parts["xl/styles.xml"]);

    let mut shared_twice = parts.clone();
    shared_twice.insert(
        "xl/worksheets/sheet2.xml".to_owned(),
        b"<worksheet><sheetData><row r=\"6\"><c r=\"D6\" t=\"s\"><v>0</v></c><c r=\"E6\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>".to_vec(),
    );
    assert!(export_spreadsheet_replica(&store, shared_twice, &sheet_parts).is_err());

    let mut mismatched = parts.clone();
    mismatched.insert(
        "xl/worksheets/sheet1.xml".to_owned(),
        data.replace(
            "<c r=\"A2\"><v>1</v></c>",
            "<c r=\"A2\" t=\"inlineStr\"><is><t>1</t></is></c>",
        )
        .into_bytes(),
    );
    assert!(export_spreadsheet_replica(&store, mismatched, &sheet_parts).is_err());
    assert!(export_spreadsheet_replica(
        &store,
        parts,
        &[("sheet-data", "xl/worksheets/sheet1.xml")],
    )
    .is_err());
}

#[test]
fn reordered_spreadsheet_updates_converge_and_reopen_the_merged_export() {
    let temp = tempfile::tempdir().unwrap();
    let original = json!({
        "v": 10,
        "m": "10",
        "ct": { "fa": "0.00", "t": "n" },
    });
    let updates = [
        offline_spreadsheet_update(
            &temp,
            910_111,
            "offline-value",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: Some(original.clone()),
                next_cell: json!({
                    "v": 12,
                    "m": "12",
                    "ct": { "fa": "0.00", "t": "n" },
                }),
            },
        ),
        offline_spreadsheet_update(
            &temp,
            910_112,
            "offline-style",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                expected_cell: Some(original),
                next_cell: json!({
                    "v": 10,
                    "m": "10",
                    "bg": "#DBEAFE",
                    "ct": { "fa": "0.00", "t": "n" },
                }),
            },
        ),
        offline_spreadsheet_update(
            &temp,
            910_113,
            "offline-formula",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 1,
                expected_cell: None,
                next_cell: json!({"v": 2, "m": "2", "f": "=A2"}),
            },
        ),
    ];

    let mut hashes = Vec::new();
    for (permutation_index, permutation) in three_update_permutations().into_iter().enumerate() {
        let ordered = [
            updates[permutation[0]].clone(),
            updates[permutation[1]].clone(),
            updates[permutation[2]].clone(),
        ];
        let store = deliver_spreadsheet_updates(
            &temp.path().join(format!("order-{permutation_index}")),
            910_200 + permutation_index as u64,
            &ordered,
        );
        assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), Some(12.0));
        assert_eq!(
            cell_string(&store, "sheet-data", 1, 0, &["bg"]),
            Some("#DBEAFE".to_owned())
        );
        assert_eq!(
            cell_string(&store, "sheet-data", 1, 1, &["f"]),
            Some("=A2".to_owned())
        );
        assert_eq!(
            cell_string(&store, "sheet-data", 0, 0, &["v"]),
            Some("Header".to_owned())
        );
        hashes.push(store.inspect().unwrap().document_state_sha256);
    }
    assert!(hashes.iter().all(|hash| hash == &hashes[0]));

    let repeated = [
        updates[0].clone(),
        updates[0].clone(),
        updates[1].clone(),
        updates[2].clone(),
    ];
    let replayed = deliver_spreadsheet_updates(&temp.path().join("repeated"), 910_220, &repeated);
    assert_eq!(replayed.inspect().unwrap().document_state_sha256, hashes[0]);

    let data = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Old</t></is></c><c r=\"C1\"><v>99</v></c></row><row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"B2\"><f>=A1</f><v>1</v></c></row></sheetData></worksheet>";
    let sparse = "<worksheet><sheetData><row r=\"6\"><c r=\"D6\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>";
    let shared = "<sst><si><t>Hello</t></si><si><t>Other</t></si></sst>";
    let mut parts = BTreeMap::new();
    parts.insert(
        "xl/worksheets/sheet1.xml".to_owned(),
        data.as_bytes().to_vec(),
    );
    parts.insert(
        "xl/worksheets/sheet2.xml".to_owned(),
        sparse.as_bytes().to_vec(),
    );
    parts.insert(
        "xl/sharedStrings.xml".to_owned(),
        shared.as_bytes().to_vec(),
    );
    parts.insert("xl/styles.xml".to_owned(), b"<styles/>".to_vec());
    let written = export_spreadsheet_replica(
        &replayed,
        parts.clone(),
        &[
            ("sheet-data", "xl/worksheets/sheet1.xml"),
            ("sheet-sparse", "xl/worksheets/sheet2.xml"),
        ],
    )
    .unwrap();
    assert_eq!(written["xl/styles.xml"], parts["xl/styles.xml"]);
    let data_xml = String::from_utf8(written["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let number = reopen_exported_worksheet_cell(&data_xml, 1, 0).unwrap();
    assert_eq!(number.formula, None);
    assert_eq!(number.cached.as_deref(), Some("12"));
    let formula = reopen_exported_worksheet_cell(&data_xml, 1, 1).unwrap();
    assert_eq!(formula.formula.as_deref(), Some("=A2"));
    assert_eq!(formula.cached.as_deref(), Some("2"));
    let untouched = reopen_exported_worksheet_cell(&data_xml, 0, 2).unwrap();
    assert_eq!(untouched.formula, None);
    assert_eq!(untouched.cached.as_deref(), Some("99"));
    let header = reopen_exported_worksheet_cell(&data_xml, 0, 0).unwrap();
    assert_eq!(header.inline.as_deref(), Some("Header"));
}

fn offline_spreadsheet_update(
    temp: &tempfile::TempDir,
    client_id: u64,
    operation_id: &str,
    mutation: NativeOfficeCollaborationMutation,
) -> Vec<u8> {
    let store = spreadsheet_store_with_client(&temp.path().join(operation_id), client_id);
    let base = store.synchronize(None).unwrap().state_vector;
    store
        .mutate(spreadsheet_mutation_request(operation_id, mutation))
        .unwrap();
    store.synchronize(Some(&base)).unwrap().update
}

fn deliver_spreadsheet_updates(
    root: &Path,
    client_id: u64,
    updates: &[Vec<u8>],
) -> NativeOfficeCollaborationStore {
    let store = spreadsheet_store_with_client(root, client_id);
    for (index, update) in updates.iter().enumerate() {
        store
            .apply(spreadsheet_apply_request(
                &format!("deliver-{client_id}-{index}"),
                update.clone(),
            ))
            .unwrap();
    }
    store
}

fn spreadsheet_store_with_client(root: &Path, client_id: u64) -> NativeOfficeCollaborationStore {
    let mut request = spreadsheet_create_request(root);
    request.client_id = Some(client_id);
    request.operation_id = format!("create-spreadsheet-{client_id}");
    let store = NativeOfficeCollaborationStore::create(request).unwrap();
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

fn three_update_permutations() -> Vec<[usize; 3]> {
    let mut result = Vec::with_capacity(6);
    for first in 0..3 {
        for second in 0..3 {
            for third in 0..3 {
                let permutation = [first, second, third];
                if permutation
                    .iter()
                    .copied()
                    .collect::<std::collections::HashSet<_>>()
                    .len()
                    == 3
                {
                    result.push(permutation);
                }
            }
        }
    }
    assert_eq!(result.len(), 6);
    result
}

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
    assert_eq!(
        value.caret,
        Some(NativeOfficeCollaborationFrameCaret::Spreadsheet {
            sheet_id: "sheet-data".to_owned(),
            row: 1,
            column: 0,
            index_utf16: None,
        })
    );
    assert!(store
        .inspect()
        .unwrap()
        .root_names
        .iter()
        .all(|name| !name.contains("caret")));

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
            next_cell: json!({ "f": "=1+1" }),
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
    let formula_only = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-formula-without-cache",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 0,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "f": "=2+2" }),
            },
        ))
        .unwrap_err();
    assert_eq!(formula_only.code, "office.collaboration.mutation_invalid");
    assert_eq!(
        formula_only.suggestion.as_deref(),
        Some(
            "Write explicit cached \"v\" / \"m\" with \"f\" on spreadsheet-set-cell or spreadsheet-batch-cells. Do not call recalculate-spreadsheet-formulas on collaboration replicas."
        )
    );
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

#[test]
fn find_lists_spreadsheet_display_text_by_coordinate() {
    let temp = tempfile::tempdir().unwrap();
    let root = temp.path().join("spreadsheet-find");
    let store = initialized_spreadsheet_store(&root);
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-find-seed",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-empty".to_owned(),
                row: 2,
                column: 3,
                expected_cell: None,
                next_cell: json!({ "v": "Draft note", "m": "Draft note" }),
            },
        ))
        .unwrap();

    let (kind, found) = store.find_text("Draft", 10).unwrap();
    assert_eq!(kind, NativeOfficeCollaborationArtifactKind::Spreadsheet);
    assert_eq!(found.match_count, 1);
    assert!(!found.truncated);
    assert_eq!(found.matches[0].occurrence, 1);
    assert_eq!(found.matches[0].sheet_id.as_deref(), Some("sheet-empty"));
    assert_eq!(found.matches[0].row, Some(2));
    assert_eq!(found.matches[0].column, Some(3));
    assert_eq!(found.matches[0].index_utf16, 0);

    let limited = store.find_text("Draft", 1).unwrap().1;
    assert_eq!(limited.match_count, 1);
    assert!(!limited.truncated);
}

#[test]
fn spreadsheet_splice_is_cell_local_and_keeps_formula_cells_field_addressed() {
    let temp = tempfile::tempdir().unwrap();
    let first = initialized_spreadsheet_store(&temp.path().join("first"));
    let text = json!({ "v": "Hello", "m": "Hello" });
    first
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-create-plain-text",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 6,
                column: 0,
                expected_cell: None,
                next_cell: text.clone(),
            },
        ))
        .unwrap();

    let mut peer_request = spreadsheet_create_request(&temp.path().join("peer"));
    peer_request.client_id = Some(900_073);
    peer_request.operation_id = "create-spreadsheet-peer".to_owned();
    peer_request.initial_update = Some(first.synchronize(None).unwrap().update);
    let second = NativeOfficeCollaborationStore::create(peer_request).unwrap();

    let inserted = first
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-splice-cjk",
            NativeOfficeCollaborationMutation::SpreadsheetSplice {
                sheet_id: "sheet-data".to_owned(),
                row: 6,
                column: 0,
                index_utf16: 5,
                delete_utf16: 0,
                expected_slice: String::new(),
                insert: "中".to_owned(),
            },
        ))
        .unwrap();
    assert_eq!(
        inserted.caret,
        Some(NativeOfficeCollaborationFrameCaret::Spreadsheet {
            sheet_id: "sheet-data".to_owned(),
            row: 6,
            column: 0,
            index_utf16: Some(6),
        })
    );
    assert!(first
        .inspect()
        .unwrap()
        .root_names
        .iter()
        .all(|name| !name.contains("caret")));

    second
        .apply(spreadsheet_apply_request(
            "apply-first-splice",
            first.synchronize(None).unwrap().update,
        ))
        .unwrap();
    second
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-splice-follow",
            NativeOfficeCollaborationMutation::SpreadsheetSplice {
                sheet_id: "sheet-data".to_owned(),
                row: 6,
                column: 0,
                index_utf16: 6,
                delete_utf16: 0,
                expected_slice: String::new(),
                insert: "文".to_owned(),
            },
        ))
        .unwrap();
    first
        .apply(spreadsheet_apply_request(
            "apply-second-splice",
            second
                .synchronize(Some(&first.synchronize(None).unwrap().state_vector))
                .unwrap()
                .update,
        ))
        .unwrap();
    assert_eq!(projected_cell_display(&first, 6, 0), "Hello中文");
    assert_eq!(projected_cell_display(&second, 6, 0), "Hello中文");

    first
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-style-after-splice",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 6,
                column: 0,
                expected_cell: Some(json!({ "v": "Hello中文", "m": "Hello中文" })),
                next_cell: json!({ "v": "Hello中文", "m": "Hello中文", "bg": "#DBEAFE" }),
            },
        ))
        .unwrap();
    assert_eq!(projected_cell_display(&first, 6, 0), "Hello中文");
    assert_eq!(
        cell_string(&first, "sheet-data", 6, 0, &["bg"]),
        Some("#DBEAFE".to_owned())
    );

    let stale = first
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-splice-stale",
            NativeOfficeCollaborationMutation::SpreadsheetSplice {
                sheet_id: "sheet-data".to_owned(),
                row: 6,
                column: 0,
                index_utf16: 0,
                delete_utf16: 1,
                expected_slice: "X".to_owned(),
                insert: "Y".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    assert_eq!(projected_cell_display(&first, 6, 0), "Hello中文");

    let formula = first
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-splice-formula",
            NativeOfficeCollaborationMutation::SpreadsheetSplice {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 0,
                index_utf16: 0,
                delete_utf16: 0,
                expected_slice: String::new(),
                insert: "中".to_owned(),
            },
        ))
        .unwrap_err();
    assert_eq!(formula.code, "office.collaboration.mutation_match_conflict");
}

#[test]
fn spreadsheet_structure_rewrites_formula_references_in_the_same_frame() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("structure"));
    let before = store.inspect().unwrap().document_state_sha256;
    let zero = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-zero-rows",
            NativeOfficeCollaborationMutation::SpreadsheetInsertRows {
                sheet_id: "sheet-data".to_owned(),
                at: 12,
                count: 0,
            },
        ))
        .unwrap_err();
    assert_eq!(zero.code, "office.collaboration.mutation_invalid");
    assert_eq!(store.inspect().unwrap().document_state_sha256, before);

    let three_dimensional = json!({
        "v": 1,
        "m": "1",
        "f": "=Sheet1:Sheet3!A1",
    });
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-three-dimensional-formula",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 30,
                column: 0,
                expected_cell: None,
                next_cell: three_dimensional.clone(),
            },
        ))
        .unwrap();
    let before_structure = store.inspect().unwrap().document_state_sha256;
    let rejected = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-rows-three-dimensional",
            NativeOfficeCollaborationMutation::SpreadsheetInsertRows {
                sheet_id: "sheet-data".to_owned(),
                at: 12,
                count: 1,
            },
        ))
        .unwrap_err();
    assert!(!rejected.code.is_empty());
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_structure
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 30, 0, &["f"]).as_deref(),
        Some("=Sheet1:Sheet3!A1")
    );
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-three-dimensional-formula",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteCell {
                sheet_id: "sheet-data".to_owned(),
                row: 30,
                column: 0,
                expected_cell: three_dimensional,
            },
        ))
        .unwrap();

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-structure-watcher",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 10,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": 1, "m": "1", "f": "=A13+B13+Other!A13" }),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-structure-moved-text",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 12,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": "moved", "m": "moved" }),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-splice-structure-moved-text",
            NativeOfficeCollaborationMutation::SpreadsheetSplice {
                sheet_id: "sheet-data".to_owned(),
                row: 12,
                column: 0,
                index_utf16: 5,
                delete_utf16: 0,
                expected_slice: String::new(),
                insert: "!".to_owned(),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-structure-stable-formula",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 14,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": 1, "m": "1", "f": "=A1" }),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-structure-tail",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 18,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": "tail", "m": "tail" }),
            },
        ))
        .unwrap();

    let inserted = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-rows",
            NativeOfficeCollaborationMutation::SpreadsheetInsertRows {
                sheet_id: "sheet-data".to_owned(),
                at: 12,
                count: 1,
            },
        ))
        .unwrap();
    assert!(inserted.state_changed);
    assert_eq!(
        inserted.caret,
        Some(NativeOfficeCollaborationFrameCaret::Spreadsheet {
            sheet_id: "sheet-data".to_owned(),
            row: 12,
            column: 0,
            index_utf16: None,
        })
    );
    store.project().unwrap();
    assert!(!spreadsheet_cell_present(&store, "sheet-data", 12, 0));
    assert_eq!(
        cell_string(&store, "sheet-data", 10, 0, &["f"]).as_deref(),
        Some("=A14+B14+Other!A13")
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 13, 0, &["m"]).as_deref(),
        Some("moved!")
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 15, 0, &["f"]).as_deref(),
        Some("=A1")
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 19, 0, &["m"]).as_deref(),
        Some("tail")
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-columns",
            NativeOfficeCollaborationMutation::SpreadsheetInsertColumns {
                sheet_id: "sheet-data".to_owned(),
                at: 1,
                count: 1,
            },
        ))
        .unwrap();
    store.project().unwrap();
    assert_eq!(
        cell_string(&store, "sheet-data", 10, 0, &["f"]).as_deref(),
        Some("=A14+C14+Other!A13")
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-rows",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteRows {
                sheet_id: "sheet-data".to_owned(),
                at: 13,
                count: 1,
            },
        ))
        .unwrap();
    store.project().unwrap();
    assert!(!spreadsheet_cell_present(&store, "sheet-data", 13, 0));
    assert_eq!(
        cell_string(&store, "sheet-data", 10, 0, &["f"]).as_deref(),
        Some("=#REF!+#REF!+Other!A13")
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 14, 0, &["f"]).as_deref(),
        Some("=A1")
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 18, 0, &["m"]).as_deref(),
        Some("tail")
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-set-structure-column-formula",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 20,
                column: 0,
                expected_cell: None,
                next_cell: json!({ "v": 1, "m": "1", "f": "=A1+B1" }),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-columns",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteColumns {
                sheet_id: "sheet-data".to_owned(),
                at: 1,
                count: 1,
            },
        ))
        .unwrap();
    store.project().unwrap();
    assert_eq!(
        cell_string(&store, "sheet-data", 20, 0, &["f"]).as_deref(),
        Some("=A1+#REF!")
    );
}

#[test]
fn spreadsheet_structure_rewrites_region_addresses_in_the_same_frame() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("regions"));
    seed_region_addresses(&store, true);
    let before = store.inspect().unwrap().document_state_sha256;
    let rejected = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-rows-chart-three-dimensional",
            NativeOfficeCollaborationMutation::SpreadsheetInsertRows {
                sheet_id: "sheet-data".to_owned(),
                at: 2,
                count: 1,
            },
        ))
        .unwrap_err();
    assert!(!rejected.code.is_empty());
    assert_eq!(store.inspect().unwrap().document_state_sha256, before);
    assert_eq!(
        region_number(&store, &["config"], &["merge", "2_0", "r"]),
        Some(2.0)
    );
    assert_eq!(
        region_string(&store, &["charts", "chart-1"], &["titleReference"]).as_deref(),
        Some("Sheet1:Sheet3!A1")
    );

    set_chart_title_reference(&store, "A1");
    let inserted = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-region-rows",
            NativeOfficeCollaborationMutation::SpreadsheetInsertRows {
                sheet_id: "sheet-data".to_owned(),
                at: 2,
                count: 1,
            },
        ))
        .unwrap();
    assert!(inserted.state_changed);
    store.project().unwrap();
    assert_eq!(
        region_number(&store, &["config"], &["merge", "2_0", "r"]),
        None
    );
    assert_eq!(
        region_number(&store, &["config"], &["merge", "3_0", "r"]),
        Some(3.0)
    );
    assert_eq!(
        region_number(&store, &["config"], &["merge", "3_0", "rs"]),
        Some(2.0)
    );
    assert_eq!(
        region_number(&store, &["config"], &["merge", "0_1", "rs"]),
        Some(5.0)
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-1"], &["range", "row"]),
        Some((3, 5))
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-2"], &["range", "row"]),
        Some((3, 3))
    );
    assert_eq!(
        region_string(&store, &["charts", "chart-1"], &["categoryReference"]).as_deref(),
        Some("A4:A6+Other!A3")
    );
    assert_eq!(series_values_reference(&store).as_deref(), Some("B4:B6"));
    assert_eq!(
        column_formula(&store, 0, "calculatedFormula").as_deref(),
        Some("=A4")
    );
    assert_eq!(
        column_formula(&store, 1, "totalsFormula").as_deref(),
        Some("=SUM(B4:B6)")
    );
    assert_eq!(
        region_number(&store, &["charts", "chart-1"], &["left"]),
        Some(12.0)
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-region-rows",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteRows {
                sheet_id: "sheet-data".to_owned(),
                at: 3,
                count: 2,
            },
        ))
        .unwrap();
    store.project().unwrap();
    assert_eq!(
        region_number(&store, &["config"], &["merge", "3_0", "r"]),
        None
    );
    assert_eq!(
        region_number(&store, &["config"], &["merge", "0_1", "rs"]),
        Some(3.0)
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-1"], &["range", "row"]),
        Some((3, 3))
    );
    assert!(region_any(&store, &["tables", "table-2"], "value", &["id"]).is_none());
    assert!(!table_ids(&store).contains(&"table-2".to_owned()));
    assert!(table_ids(&store).contains(&"table-1".to_owned()));
    assert_eq!(
        region_string(&store, &["charts", "chart-1"], &["categoryReference"]).as_deref(),
        Some("A4:A4+Other!A3")
    );
    assert_eq!(series_values_reference(&store).as_deref(), Some("B4:B4"));
    assert_eq!(
        column_formula(&store, 0, "calculatedFormula").as_deref(),
        Some("=#REF!")
    );
    assert_eq!(
        region_number(&store, &["charts", "chart-1"], &["left"]),
        Some(12.0)
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-insert-region-columns",
            NativeOfficeCollaborationMutation::SpreadsheetInsertColumns {
                sheet_id: "sheet-data".to_owned(),
                at: 1,
                count: 1,
            },
        ))
        .unwrap();
    store.project().unwrap();
    assert_eq!(
        region_number(&store, &["config"], &["merge", "0_1", "c"]),
        None
    );
    assert_eq!(
        region_number(&store, &["config"], &["merge", "0_2", "c"]),
        Some(2.0)
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-1"], &["range", "column"]),
        Some((0, 2))
    );
    assert_eq!(column_count(&store), Some(3));
    assert_eq!(series_values_reference(&store).as_deref(), Some("C4:C4"));
    assert_eq!(
        region_number(&store, &["charts", "chart-1"], &["top"]),
        Some(8.0)
    );

    store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-delete-region-columns",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteColumns {
                sheet_id: "sheet-data".to_owned(),
                at: 2,
                count: 1,
            },
        ))
        .unwrap();
    store.project().unwrap();
    assert_eq!(
        region_number(&store, &["config"], &["merge", "0_2", "c"]),
        None
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-1"], &["range", "column"]),
        Some((0, 1))
    );
    assert_eq!(column_count(&store), Some(2));
    assert_eq!(series_values_reference(&store).as_deref(), Some("#REF!"));
    assert_eq!(
        region_number(&store, &["charts", "chart-1"], &["left"]),
        Some(12.0)
    );
}

fn projected_cell_display(store: &NativeOfficeCollaborationStore, row: u32, column: u32) -> String {
    let NativeOfficeCollaborationProjectedContent::Spreadsheet { sheets } =
        store.project().unwrap().content
    else {
        panic!("spreadsheet projection");
    };
    let cell = sheets
        .iter()
        .find(|sheet| sheet.sheet_id == "sheet-data")
        .and_then(|sheet| {
            sheet
                .cells
                .iter()
                .find(|cell| cell.row == row && cell.column == column)
        })
        .expect("projected cell");
    cell.cell["m"].as_str().unwrap().to_owned()
}

#[test]
fn spreadsheet_sort_permutes_rows_inside_a_rectangle_or_writes_nothing() {
    let top_left = json!({"v": 1, "m": "1", "f": "=A1"});
    let top_right = json!({"v": 2, "m": "2"});
    let bottom_left = json!({"v": 3, "m": "3"});
    let bottom_right = json!({"v": 4, "m": "4", "f": "=B2"});
    let outside_formula = json!({"v": 9, "m": "9", "f": "=C3"});
    let below = json!({"v": 8, "m": "8"});
    let expected = vec![
        Some(top_left.clone()),
        Some(top_right.clone()),
        Some(bottom_left.clone()),
        Some(bottom_right.clone()),
    ];

    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("sort-rows"));
    seed_sort_rectangle(
        &store,
        &top_left,
        &top_right,
        &bottom_left,
        &bottom_right,
        &outside_formula,
        &below,
    );
    seed_distant_merge(&store);
    let moved = store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-sort-rows",
            sort_rows(expected.clone(), vec![1, 0]),
        ))
        .unwrap();
    assert!(moved.state_changed);
    assert_eq!(cell_number(&store, "sheet-data", 5, 0, &["v"]), Some(3.0));
    assert_eq!(cell_number(&store, "sheet-data", 5, 1, &["v"]), Some(4.0));
    assert_eq!(
        cell_string(&store, "sheet-data", 5, 1, &["f"]).as_deref(),
        Some("=B2")
    );
    assert_eq!(cell_number(&store, "sheet-data", 6, 0, &["v"]), Some(1.0));
    assert_eq!(
        cell_string(&store, "sheet-data", 6, 0, &["f"]).as_deref(),
        Some("=A1")
    );
    assert_eq!(cell_number(&store, "sheet-data", 6, 1, &["v"]), Some(2.0));
    assert_eq!(
        cell_string(&store, "sheet-data", 4, 0, &["f"]).as_deref(),
        Some("=C3")
    );
    assert_eq!(cell_number(&store, "sheet-data", 7, 0, &["v"]), Some(8.0));

    let identity = tempfile::tempdir().unwrap();
    let identity_store = initialized_spreadsheet_store(&identity.path().join("sort-identity"));
    seed_sort_rectangle(
        &identity_store,
        &top_left,
        &top_right,
        &bottom_left,
        &bottom_right,
        &outside_formula,
        &below,
    );
    let before_identity = identity_store.inspect().unwrap();
    let same = identity_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-sort-identity",
            sort_rows(expected.clone(), vec![0, 1]),
        ))
        .unwrap();
    assert!(!same.state_changed);
    assert_eq!(
        identity_store.inspect().unwrap().document_state_sha256,
        before_identity.document_state_sha256
    );

    let stale = tempfile::tempdir().unwrap();
    let stale_store = initialized_spreadsheet_store(&stale.path().join("sort-stale"));
    seed_sort_rectangle(
        &stale_store,
        &top_left,
        &top_right,
        &bottom_left,
        &bottom_right,
        &outside_formula,
        &below,
    );
    let before_stale = stale_store.inspect().unwrap();
    let mut stale_expected = expected.clone();
    stale_expected[0] = Some(json!({"v": 99, "m": "99"}));
    let conflict = stale_store
        .mutate(spreadsheet_mutation_request(
            "spreadsheet-sort-stale",
            sort_rows(stale_expected, vec![1, 0]),
        ))
        .unwrap_err();
    assert_eq!(
        conflict.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        stale_store.inspect().unwrap().document_state_sha256,
        before_stale.document_state_sha256
    );
    assert_eq!(
        cell_number(&stale_store, "sheet-data", 5, 0, &["v"]),
        Some(1.0)
    );

    for (name, seed) in [
        ("merge", SortBlocker::Merge),
        ("table", SortBlocker::Table),
        ("pivot", SortBlocker::Pivot),
    ] {
        let temp = tempfile::tempdir().unwrap();
        let store = initialized_spreadsheet_store(&temp.path().join(format!("sort-{name}")));
        seed_sort_rectangle(
            &store,
            &top_left,
            &top_right,
            &bottom_left,
            &bottom_right,
            &outside_formula,
            &below,
        );
        seed_sort_blocker(&store, seed);
        let before = store.inspect().unwrap();
        let blocked = store
            .mutate(spreadsheet_mutation_request(
                &format!("spreadsheet-sort-{name}"),
                sort_rows(expected.clone(), vec![1, 0]),
            ))
            .unwrap_err();
        assert_eq!(blocked.code, "office.collaboration.mutation_invalid");
        assert_eq!(
            store.inspect().unwrap().document_state_sha256,
            before.document_state_sha256
        );
        assert_eq!(cell_number(&store, "sheet-data", 5, 0, &["v"]), Some(1.0));
        assert_eq!(
            cell_string(&store, "sheet-data", 4, 0, &["f"]).as_deref(),
            Some("=C3")
        );
    }
}

fn sort_rows(
    expected_cells: Vec<Option<JsonValue>>,
    source_rows: Vec<u32>,
) -> NativeOfficeCollaborationMutation {
    NativeOfficeCollaborationMutation::SpreadsheetSortRows {
        sheet_id: "sheet-data".to_owned(),
        row: 5,
        column: 0,
        row_count: 2,
        column_count: 2,
        source_rows,
        expected_cells,
    }
}

fn seed_sort_rectangle(
    store: &NativeOfficeCollaborationStore,
    top_left: &JsonValue,
    top_right: &JsonValue,
    bottom_left: &JsonValue,
    bottom_right: &JsonValue,
    outside_formula: &JsonValue,
    below: &JsonValue,
) {
    for (operation, row, column, cell) in [
        ("sort-seed-outside", 4, 0, outside_formula),
        ("sort-seed-top-left", 5, 0, top_left),
        ("sort-seed-top-right", 5, 1, top_right),
        ("sort-seed-bottom-left", 6, 0, bottom_left),
        ("sort-seed-bottom-right", 6, 1, bottom_right),
        ("sort-seed-below", 7, 0, below),
    ] {
        store
            .mutate(spreadsheet_mutation_request(
                operation,
                NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                    sheet_id: "sheet-data".to_owned(),
                    row,
                    column,
                    expected_cell: None,
                    next_cell: cell.clone(),
                },
            ))
            .unwrap();
    }
}

#[derive(Clone, Copy)]
enum SortBlocker {
    Merge,
    Table,
    Pivot,
}

fn seed_distant_merge(store: &NativeOfficeCollaborationStore) {
    let document = replica_document(store);
    let before = document.transact().state_vector();
    let sheets = document.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let record = {
        let transaction = document.transact();
        match sheets.get(&transaction, "sheet-data") {
            Some(Out::YMap(record)) => record,
            value => panic!("unexpected sheet record: {value:?}"),
        }
    };
    let config = ensure_map(&document, &record, "config");
    put_flat(&document, &config, "object", &["merge"], Any::Bool(true));
    put_merge(&document, &config, 0, 0, 1, 1);
    let update = document.transact().encode_state_as_update_v1(&before);
    store
        .apply(spreadsheet_apply_request("seed-sort-distant-merge", update))
        .unwrap();
}

fn seed_sort_blocker(store: &NativeOfficeCollaborationStore, blocker: SortBlocker) {
    let document = replica_document(store);
    let before = document.transact().state_vector();
    let sheets = document.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let record = {
        let transaction = document.transact();
        match sheets.get(&transaction, "sheet-data") {
            Some(Out::YMap(record)) => record,
            value => panic!("unexpected sheet record: {value:?}"),
        }
    };
    match blocker {
        SortBlocker::Merge => {
            let config = ensure_map(&document, &record, "config");
            put_flat(&document, &config, "object", &["merge"], Any::Bool(true));
            put_merge(&document, &config, 5, 0, 1, 1);
            put_merge(&document, &config, 0, 0, 1, 1);
        }
        SortBlocker::Table => {
            let tables = ensure_map(&document, &record, "tables");
            let order = ensure_array(&document, &record, "tableOrder");
            put_table(
                &document,
                &tables,
                &order,
                "table-sort",
                (5, 6),
                (0, 1),
                vec![table_column("Amount", None), table_column("Note", None)],
            );
        }
        SortBlocker::Pivot => {
            let pivots = ensure_map(&document, &record, "pivotTables");
            pivots.insert(
                &mut document.transact_mut(),
                "pivot-1",
                yrs::MapPrelim::default(),
            );
        }
    }
    let update = document.transact().encode_state_as_update_v1(&before);
    store
        .apply(spreadsheet_apply_request("seed-sort-blocker", update))
        .unwrap();
}

#[test]
fn spreadsheet_table_frame_writes_the_shared_record_or_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("table-frame"));
    let kept = cell_number(&store, "sheet-data", 1, 0, &["v"]);
    let table = sales_table("Sales", (0, 2), (0, 1));
    let before_create = store.synchronize(None).unwrap().state_vector;
    store
        .mutate(spreadsheet_mutation_request(
            "create-sales",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: table.clone(),
            },
        ))
        .unwrap();
    let created = store.synchronize(Some(&before_create)).unwrap().update;
    assert_eq!(
        region_any(&store, &["tables", "table-sales"], "value", &["name"]),
        Some(Any::String(std::sync::Arc::from("Sales")))
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-sales"], &["range", "row"]),
        Some((0, 2))
    );
    assert_eq!(table_order(&store), vec!["table-sales".to_owned()]);
    assert!(table_claims(&store).iter().any(|claim| {
        claim.contains("\"kind\":\"table\"")
            && claim.contains("\"id\":\"table-sales\"")
            && claim.contains("\"parentId\":\"sheet-data\"")
    }));
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), kept);

    let before_conflict = store.inspect().unwrap();
    let stale = store
        .mutate(spreadsheet_mutation_request(
            "stale-sales",
            NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: sales_table("Other", (0, 2), (0, 1)),
                next_table: sales_table("Renamed", (0, 2), (0, 1)),
            },
        ))
        .unwrap_err();
    assert_eq!(stale.code, "office.collaboration.mutation_match_conflict");
    let overlap = store
        .mutate(spreadsheet_mutation_request(
            "overlap-sales",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: sales_table_with_id("table-other", "Other", (1, 3), (0, 1)),
            },
        ))
        .unwrap_err();
    assert_eq!(overlap.code, "office.collaboration.mutation_invalid");
    let duplicate_name = store
        .mutate(spreadsheet_mutation_request(
            "duplicate-sales",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: sales_table_with_id("table-copy", "sales", (4, 6), (0, 1)),
            },
        ))
        .unwrap_err();
    assert_eq!(duplicate_name.code, "office.collaboration.mutation_invalid");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_conflict.document_state_sha256
    );
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), kept);

    let renamed = sales_table("Renamed", (0, 2), (0, 1));
    store
        .mutate(spreadsheet_mutation_request(
            "rename-sales",
            NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: table.clone(),
                next_table: renamed.clone(),
            },
        ))
        .unwrap();
    assert_eq!(
        region_any(&store, &["tables", "table-sales"], "value", &["name"]),
        Some(Any::String(std::sync::Arc::from("Renamed")))
    );

    let stale_delete = store
        .mutate(spreadsheet_mutation_request(
            "stale-delete-sales",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: table,
            },
        ))
        .unwrap_err();
    assert_eq!(
        stale_delete.code,
        "office.collaboration.mutation_match_conflict"
    );
    assert_eq!(
        region_any(&store, &["tables", "table-sales"], "value", &["name"]),
        Some(Any::String(std::sync::Arc::from("Renamed")))
    );

    store
        .mutate(spreadsheet_mutation_request(
            "delete-sales",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: renamed,
            },
        ))
        .unwrap();
    assert!(region_any(&store, &["tables", "table-sales"], "value", &["name"]).is_none());
    assert!(table_order(&store).is_empty());
    assert!(table_claims(&store)
        .iter()
        .any(|claim| { claim.contains("\"id\":\"table-sales\"") }));
    let reused = store
        .mutate(spreadsheet_mutation_request(
            "reuse-sales",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: sales_table("Fresh", (0, 2), (0, 1)),
            },
        ))
        .unwrap_err();
    assert_eq!(reused.code, "office.collaboration.mutation_invalid");
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), kept);

    let peer = initialized_spreadsheet_store(&temp.path().join("table-peer"));
    peer.apply(spreadsheet_apply_request("apply-sales-create", created))
        .unwrap();
    assert_eq!(
        region_any(&peer, &["tables", "table-sales"], "value", &["name"]),
        Some(Any::String(std::sync::Arc::from("Sales")))
    );
}

#[test]
fn spreadsheet_table_frame_rewrites_structured_references_or_writes_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("table-formulas"));
    let kept = cell_number(&store, "sheet-data", 1, 0, &["v"]);
    let table = sales_table("Sales", (0, 2), (0, 1));
    store
        .mutate(spreadsheet_mutation_request(
            "create-formula-sales",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: table.clone(),
            },
        ))
        .unwrap();
    let qualified = "='[Book.xlsx]Sheet1'!Sales[Column0]";
    let literal = r#"="Sales[Column0]""#;
    for (operation, row, column, formula) in [
        ("formula-qualified", 8, 0, "=Sales[Column0]"),
        ("formula-literal", 8, 1, literal),
        ("formula-external", 8, 2, qualified),
        ("formula-unrelated", 8, 3, "=A1"),
        ("formula-local", 1, 1, "=[@Column0]"),
    ] {
        store
            .mutate(spreadsheet_mutation_request(
                operation,
                NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                    sheet_id: "sheet-data".to_owned(),
                    row,
                    column,
                    expected_cell: None,
                    next_cell: formula_cell(formula),
                },
            ))
            .unwrap();
    }
    store
        .mutate(spreadsheet_mutation_request(
            "text-literal",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 8,
                column: 4,
                expected_cell: None,
                next_cell: json!({ "v": "Sales[Column0]", "m": "Sales[Column0]" }),
            },
        ))
        .unwrap();

    let renamed = sales_table("Renamed", (0, 2), (0, 1));
    store
        .mutate(spreadsheet_mutation_request(
            "rename-formula-sales",
            NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: table,
                next_table: renamed.clone(),
            },
        ))
        .unwrap();
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 0, &["f"]),
        Some("=Renamed[Column0]".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 1, &["f"]),
        Some(literal.to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 2, &["f"]),
        Some(qualified.to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 3, &["f"]),
        Some("=A1".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 1, &["f"]),
        Some("=[@Column0]".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 4, &["v"]),
        Some("Sales[Column0]".to_owned())
    );
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), kept);

    let amount = table_with_columns("Renamed", (0, 2), (0, 1), &["Amount", "Column1"]);
    store
        .mutate(spreadsheet_mutation_request(
            "rename-formula-column",
            NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: renamed,
                next_table: amount.clone(),
            },
        ))
        .unwrap();
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 0, &["f"]),
        Some("=Renamed[Amount]".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 1, &["f"]),
        Some("=[@Amount]".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 2, &["f"]),
        Some(qualified.to_owned())
    );

    set_unscoped_local_name(&store, true);
    let before_unscoped = store.inspect().unwrap();
    let units = table_with_columns("Renamed", (0, 2), (0, 1), &["Units", "Column1"]);
    let unscoped = store
        .mutate(spreadsheet_mutation_request(
            "rename-unscoped-local",
            NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: amount.clone(),
                next_table: units,
            },
        ))
        .unwrap_err();
    assert_eq!(
        unscoped.code,
        "use.office.spreadsheet_table_formula_rewrite_unsupported"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_unscoped.document_state_sha256
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 0, &["f"]),
        Some("=Renamed[Amount]".to_owned())
    );
    set_unscoped_local_name(&store, false);

    let before_geometry = store.inspect().unwrap();
    let grown = table_with_columns("Renamed", (0, 4), (0, 1), &["Amount", "Column1"]);
    let geometry = store
        .mutate(spreadsheet_mutation_request(
            "grow-formula-sales",
            NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: amount.clone(),
                next_table: grown,
            },
        ))
        .unwrap_err();
    assert_eq!(
        geometry.code,
        "use.office.spreadsheet_table_formula_rewrite_unsupported"
    );
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_geometry.document_state_sha256
    );
    assert_eq!(
        region_pair(&store, &["tables", "table-sales"], &["range", "row"]),
        Some((0, 2))
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 1, 1, &["f"]),
        Some("=[@Amount]".to_owned())
    );

    let before_delete = store.inspect().unwrap();
    let referenced = store
        .mutate(spreadsheet_mutation_request(
            "delete-referenced-sales",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: amount.clone(),
            },
        ))
        .unwrap_err();
    assert_eq!(referenced.code, "use.office.spreadsheet_table_referenced");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_delete.document_state_sha256
    );
    assert_eq!(
        region_any(&store, &["tables", "table-sales"], "value", &["name"]),
        Some(Any::String(std::sync::Arc::from("Renamed")))
    );

    store
        .mutate(spreadsheet_mutation_request(
            "clear-qualified",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 8,
                column: 0,
                expected_cell: Some(json!({
                    "f": "=Renamed[Amount]",
                    "v": "=Sales[Column0]",
                    "m": "=Sales[Column0]",
                })),
                next_cell: json!({ "f": "=1", "v": "1", "m": "1" }),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "clear-local",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 1,
                column: 1,
                expected_cell: Some(json!({
                    "f": "=[@Amount]",
                    "v": "=[@Column0]",
                    "m": "=[@Column0]",
                })),
                next_cell: json!({ "v": 1, "m": "1" }),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "delete-unreferenced-sales",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-sales".to_owned(),
                expected_table: amount,
            },
        ))
        .unwrap();
    assert!(region_any(&store, &["tables", "table-sales"], "value", &["name"]).is_none());
    assert!(table_order(&store).is_empty());
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 1, &["f"]),
        Some(literal.to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 2, &["f"]),
        Some(qualified.to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 4, &["v"]),
        Some("Sales[Column0]".to_owned())
    );
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), kept);
}

#[test]
fn spreadsheet_table_conversion_paints_existing_cells_or_writes_nothing() {
    let temp = tempfile::tempdir().unwrap();
    let store = initialized_spreadsheet_store(&temp.path().join("table-conversion"));
    let kept = cell_number(&store, "sheet-data", 1, 0, &["v"]);
    let table = painted_table();
    store
        .mutate(spreadsheet_mutation_request(
            "create-paint",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: table.clone(),
            },
        ))
        .unwrap();
    for (operation, row, column, cell) in [
        ("paint-header", 4, 0, json!({ "v": "H", "m": "H" })),
        ("paint-body", 5, 0, formula_cell("=1+1")),
        ("paint-stripe", 6, 0, json!({ "v": "S", "m": "S" })),
        (
            "paint-outside",
            8,
            0,
            json!({ "v": "Out", "m": "Out", "bg": "#112233" }),
        ),
        ("paint-reference", 9, 0, formula_cell("=Paint[Column0]")),
    ] {
        store
            .mutate(spreadsheet_mutation_request(
                operation,
                NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                    sheet_id: "sheet-data".to_owned(),
                    row,
                    column,
                    expected_cell: None,
                    next_cell: cell,
                },
            ))
            .unwrap();
    }

    let before_referenced = store.inspect().unwrap();
    let borders_before = border_count(&store);
    let referenced = store
        .mutate(spreadsheet_mutation_request(
            "delete-referenced-paint",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-paint".to_owned(),
                expected_table: table.clone(),
            },
        ))
        .unwrap_err();
    assert_eq!(referenced.code, "use.office.spreadsheet_table_referenced");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_referenced.document_state_sha256
    );
    assert!(cell_string(&store, "sheet-data", 4, 0, &["bg"]).is_none());
    assert_eq!(border_count(&store), borders_before);

    store
        .mutate(spreadsheet_mutation_request(
            "clear-paint-reference",
            NativeOfficeCollaborationMutation::SpreadsheetSetCell {
                sheet_id: "sheet-data".to_owned(),
                row: 9,
                column: 0,
                expected_cell: Some(formula_cell("=Paint[Column0]")),
                next_cell: formula_cell("=1"),
            },
        ))
        .unwrap();
    store
        .mutate(spreadsheet_mutation_request(
            "delete-paint",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-paint".to_owned(),
                expected_table: table,
            },
        ))
        .unwrap();

    assert!(region_any(&store, &["tables", "table-paint"], "value", &["name"]).is_none());
    assert!(table_order(&store).is_empty());
    assert!(table_claims(&store)
        .iter()
        .any(|claim| claim.contains("\"id\":\"table-paint\"")));
    assert_eq!(
        cell_string(&store, "sheet-data", 4, 0, &["bg"]),
        Some("#e6e9ec".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 4, 0, &["fc"]),
        Some("#445166".to_owned())
    );
    assert_eq!(cell_number(&store, "sheet-data", 4, 0, &["bl"]), Some(1.0));
    assert_eq!(
        cell_string(&store, "sheet-data", 4, 0, &["v"]),
        Some("H".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 5, 0, &["bg"]),
        Some("#ffffff".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 5, 0, &["fc"]),
        Some("#1f2937".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 5, 0, &["f"]),
        Some("=1+1".to_owned())
    );
    assert!(cell_number(&store, "sheet-data", 5, 0, &["bl"]).is_none());
    assert_eq!(
        cell_string(&store, "sheet-data", 6, 0, &["bg"]),
        Some("#f0f1f3".to_owned())
    );
    assert_eq!(
        cell_string(&store, "sheet-data", 8, 0, &["bg"]),
        Some("#112233".to_owned())
    );
    assert!(!spreadsheet_cell_present(&store, "sheet-data", 5, 1));
    assert_eq!(cell_number(&store, "sheet-data", 1, 0, &["v"]), kept);
    assert_eq!(border_count(&store), borders_before + 1);
    assert!(border_records(&store).iter().any(|record| {
        record.get("borderType").is_some_and(
            |value| matches!(value, Any::String(text) if text.as_ref() == "border-all"),
        ) && record
            .get("color")
            .is_some_and(|value| matches!(value, Any::String(text) if text.as_ref() == "#b2bac5"))
            && record
                .get("style")
                .is_some_and(|value| matches!(value, Any::String(text) if text.as_ref() == "1"))
    }));

    let mut large = sales_table_with_id("table-large", "Large", (0, 100_000), (2, 2));
    large["style"] = json!({ "family": "light", "number": 1 });
    store
        .mutate(spreadsheet_mutation_request(
            "create-large",
            NativeOfficeCollaborationMutation::SpreadsheetCreateTable {
                sheet_id: "sheet-data".to_owned(),
                table: large.clone(),
            },
        ))
        .unwrap();
    let before_large = store.inspect().unwrap();
    let oversized = store
        .mutate(spreadsheet_mutation_request(
            "delete-large",
            NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                sheet_id: "sheet-data".to_owned(),
                table_id: "table-large".to_owned(),
                expected_table: large,
            },
        ))
        .unwrap_err();
    assert_eq!(oversized.code, "office.collaboration.mutation_invalid");
    assert_eq!(
        store.inspect().unwrap().document_state_sha256,
        before_large.document_state_sha256
    );
    assert_eq!(
        region_any(&store, &["tables", "table-large"], "value", &["name"]),
        Some(Any::String(std::sync::Arc::from("Large")))
    );
}

fn painted_table() -> JsonValue {
    let mut table = sales_table_with_id("table-paint", "Paint", (4, 6), (0, 1));
    table["style"] = json!({ "family": "light", "number": 1 });
    table["showRowStripes"] = json!(true);
    table
}

fn border_count(store: &NativeOfficeCollaborationStore) -> usize {
    border_values(store).len()
}

fn border_records(
    store: &NativeOfficeCollaborationStore,
) -> Vec<std::sync::Arc<std::collections::HashMap<String, Any>>> {
    border_values(store)
        .into_iter()
        .filter_map(|value| match value {
            Any::Map(record) => Some(record),
            _ => None,
        })
        .collect()
}

fn border_values(store: &NativeOfficeCollaborationStore) -> Vec<Any> {
    match region_any(store, &["config"], "value", &["borderInfo"]) {
        Some(Any::Array(items)) => items.to_vec(),
        _ => Vec::new(),
    }
}

fn set_unscoped_local_name(store: &NativeOfficeCollaborationStore, present: bool) {
    let document = replica_document(store);
    let before = document.transact().state_vector();
    let named = document.get_or_insert_map("a3s.office.spreadsheet.named-ranges");
    if present {
        let record = ensure_map(&document, &named, "range-local");
        put_flat(
            &document,
            &record,
            "value",
            &["reference"],
            Any::String(std::sync::Arc::from("[@Amount]")),
        );
    } else {
        named.remove(&mut document.transact_mut(), "range-local");
    }
    let update = document.transact().encode_state_as_update_v1(&before);
    store
        .apply(spreadsheet_apply_request(
            if present {
                "seed-unscoped-local"
            } else {
                "drop-unscoped-local"
            },
            update,
        ))
        .unwrap();
}

fn formula_cell(formula: &str) -> JsonValue {
    json!({ "f": formula, "v": formula, "m": formula })
}

fn table_with_columns(
    name: &str,
    row: (u32, u32),
    column: (u32, u32),
    names: &[&str],
) -> JsonValue {
    let mut table = sales_table(name, row, column);
    table["columns"] = JsonValue::Array(names.iter().map(|name| json!({ "name": name })).collect());
    table
}

fn sales_table(name: &str, row: (u32, u32), column: (u32, u32)) -> JsonValue {
    sales_table_with_id("table-sales", name, row, column)
}

fn sales_table_with_id(id: &str, name: &str, row: (u32, u32), column: (u32, u32)) -> JsonValue {
    let width = column.1 - column.0 + 1;
    let columns = (0..width)
        .map(|index| json!({ "name": format!("Column{index}") }))
        .collect::<Vec<_>>();
    json!({
        "id": id,
        "name": name,
        "range": { "row": [row.0, row.1], "column": [column.0, column.1] },
        "columns": columns,
        "filters": [],
        "headerRow": true,
        "totalsRow": false,
        "style": { "family": "none" },
        "showFirstColumn": false,
        "showLastColumn": false,
        "showRowStripes": false,
        "showColumnStripes": false,
    })
}

fn table_order(store: &NativeOfficeCollaborationStore) -> Vec<String> {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let Some(Out::YMap(record)) = sheets.get(&transaction, "sheet-data") else {
        return Vec::new();
    };
    let Some(Out::YArray(order)) = record.get(&transaction, "tableOrder") else {
        return Vec::new();
    };
    (0..order.len(&transaction))
        .filter_map(|index| match order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => Some(value.to_string()),
            _ => None,
        })
        .collect()
}

fn table_claims(store: &NativeOfficeCollaborationStore) -> Vec<String> {
    let peer = replica_document(store);
    let claims = peer.get_or_insert_array("a3s.office.spreadsheet.record-claims");
    let transaction = peer.transact();
    (0..claims.len(&transaction))
        .filter_map(|index| match claims.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => Some(value.to_string()),
            _ => None,
        })
        .collect()
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

fn seed_region_addresses(store: &NativeOfficeCollaborationStore, three_dimensional: bool) {
    let document = replica_document(store);
    let before = document.transact().state_vector();
    let sheets = document.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let record = {
        let transaction = document.transact();
        match sheets.get(&transaction, "sheet-data") {
            Some(Out::YMap(record)) => record,
            value => panic!("unexpected sheet record: {value:?}"),
        }
    };
    let config = ensure_map(&document, &record, "config");
    put_flat(&document, &config, "object", &["merge"], Any::Bool(true));
    put_merge(&document, &config, 2, 0, 2, 1);
    put_merge(&document, &config, 0, 1, 4, 1);

    let charts = ensure_map(&document, &record, "charts");
    let chart = charts.insert(
        &mut document.transact_mut(),
        "chart-1",
        yrs::MapPrelim::default(),
    );
    let title = if three_dimensional {
        "Sheet1:Sheet3!A1"
    } else {
        "A1"
    };
    put_flat(
        &document,
        &chart,
        "value",
        &["titleReference"],
        Any::String(std::sync::Arc::from(title)),
    );
    put_flat(
        &document,
        &chart,
        "value",
        &["categoryReference"],
        Any::String(std::sync::Arc::from("A3:A5+Other!A3")),
    );
    put_flat(
        &document,
        &chart,
        "value",
        &["series"],
        Any::Array(std::sync::Arc::from([Any::Map(std::sync::Arc::new({
            let mut series = std::collections::HashMap::new();
            series.insert("name".to_owned(), Any::String(std::sync::Arc::from("S")));
            series.insert(
                "values".to_owned(),
                Any::Array(std::sync::Arc::from(Vec::<Any>::new())),
            );
            series.insert(
                "valuesReference".to_owned(),
                Any::String(std::sync::Arc::from("B3:B5")),
            );
            series
        }))])),
    );
    put_flat(&document, &chart, "value", &["left"], Any::Number(12.0));
    put_flat(&document, &chart, "value", &["top"], Any::Number(8.0));

    let tables = ensure_map(&document, &record, "tables");
    let order = ensure_array(&document, &record, "tableOrder");
    put_table(
        &document,
        &tables,
        &order,
        "table-1",
        (2, 4),
        (0, 1),
        vec![
            table_column("Amount", Some(("calculatedFormula", "=A3"))),
            table_column("Total", Some(("totalsFormula", "=SUM(B3:B5)"))),
        ],
    );
    put_table(
        &document,
        &tables,
        &order,
        "table-2",
        (2, 2),
        (0, 0),
        vec![table_column("Only", None)],
    );

    let update = document.transact().encode_state_as_update_v1(&before);
    store
        .apply(spreadsheet_apply_request("seed-region-addresses", update))
        .unwrap();
}

fn set_chart_title_reference(store: &NativeOfficeCollaborationStore, title: &str) {
    let document = replica_document(store);
    let before = document.transact().state_vector();
    let sheets = document.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let chart = {
        let transaction = document.transact();
        let record = match sheets.get(&transaction, "sheet-data") {
            Some(Out::YMap(record)) => record,
            value => panic!("unexpected sheet record: {value:?}"),
        };
        let charts = match record.get(&transaction, "charts") {
            Some(Out::YMap(charts)) => charts,
            value => panic!("unexpected charts: {value:?}"),
        };
        match charts.get(&transaction, "chart-1") {
            Some(Out::YMap(chart)) => chart,
            value => panic!("unexpected chart: {value:?}"),
        }
    };
    put_flat(
        &document,
        &chart,
        "value",
        &["titleReference"],
        Any::String(std::sync::Arc::from(title)),
    );
    let update = document.transact().encode_state_as_update_v1(&before);
    store
        .apply(spreadsheet_apply_request(
            "set-chart-title-reference",
            update,
        ))
        .unwrap();
}

fn put_merge(
    document: &Doc,
    config: &yrs::MapRef,
    row: u32,
    column: u32,
    row_span: u32,
    column_span: u32,
) {
    let id = format!("{row}_{column}");
    put_flat(
        document,
        config,
        "object",
        &["merge", id.as_str()],
        Any::Bool(true),
    );
    for (field, number) in [
        ("r", row),
        ("c", column),
        ("rs", row_span),
        ("cs", column_span),
    ] {
        put_flat(
            document,
            config,
            "value",
            &["merge", id.as_str(), field],
            Any::Number(f64::from(number)),
        );
    }
}

fn put_table(
    document: &Doc,
    tables: &yrs::MapRef,
    order: &yrs::ArrayRef,
    id: &str,
    row: (u32, u32),
    column: (u32, u32),
    columns: Vec<Any>,
) {
    let table = tables.insert(&mut document.transact_mut(), id, yrs::MapPrelim::default());
    put_flat(
        document,
        &table,
        "value",
        &["id"],
        Any::String(std::sync::Arc::from(id)),
    );
    put_flat(document, &table, "object", &["range"], Any::Bool(true));
    put_flat(
        document,
        &table,
        "value",
        &["range", "row"],
        Any::Array(std::sync::Arc::from([
            Any::Number(f64::from(row.0)),
            Any::Number(f64::from(row.1)),
        ])),
    );
    put_flat(
        document,
        &table,
        "value",
        &["range", "column"],
        Any::Array(std::sync::Arc::from([
            Any::Number(f64::from(column.0)),
            Any::Number(f64::from(column.1)),
        ])),
    );
    put_flat(
        document,
        &table,
        "value",
        &["columns"],
        Any::Array(std::sync::Arc::from(columns)),
    );
    order.push_back(&mut document.transact_mut(), id);
}

fn table_column(name: &str, formula: Option<(&str, &str)>) -> Any {
    let mut column = std::collections::HashMap::new();
    column.insert("name".to_owned(), Any::String(std::sync::Arc::from(name)));
    if let Some((key, formula)) = formula {
        column.insert(key.to_owned(), Any::String(std::sync::Arc::from(formula)));
    }
    Any::Map(std::sync::Arc::new(column))
}

fn ensure_map(document: &Doc, parent: &yrs::MapRef, key: &str) -> yrs::MapRef {
    if let Some(Out::YMap(map)) = parent.get(&document.transact(), key) {
        return map;
    }
    parent.insert(&mut document.transact_mut(), key, yrs::MapPrelim::default())
}

fn ensure_array(document: &Doc, parent: &yrs::MapRef, key: &str) -> yrs::ArrayRef {
    if let Some(Out::YArray(array)) = parent.get(&document.transact(), key) {
        return array;
    }
    parent.insert(
        &mut document.transact_mut(),
        key,
        yrs::ArrayPrelim::default(),
    )
}

fn put_flat(document: &Doc, map: &yrs::MapRef, kind: &str, path: &[&str], value: Any) {
    let mut identity = vec![kind.to_owned()];
    identity.extend(path.iter().copied().map(str::to_owned));
    let key = serde_json::to_string(&identity).unwrap();
    map.insert(&mut document.transact_mut(), key, value);
}

fn region_number(
    store: &NativeOfficeCollaborationStore,
    parents: &[&str],
    path: &[&str],
) -> Option<f64> {
    match region_any(store, parents, "value", path)? {
        Any::Number(value) => Some(value),
        _ => None,
    }
}

fn region_string(
    store: &NativeOfficeCollaborationStore,
    parents: &[&str],
    path: &[&str],
) -> Option<String> {
    match region_any(store, parents, "value", path)? {
        Any::String(value) => Some(value.to_string()),
        _ => None,
    }
}

fn region_pair(
    store: &NativeOfficeCollaborationStore,
    parents: &[&str],
    path: &[&str],
) -> Option<(u32, u32)> {
    let Any::Array(items) = region_any(store, parents, "value", path)? else {
        return None;
    };
    if items.len() != 2 {
        return None;
    }
    Some((any_u32(&items[0])?, any_u32(&items[1])?))
}

fn region_any(
    store: &NativeOfficeCollaborationStore,
    parents: &[&str],
    kind: &str,
    path: &[&str],
) -> Option<Any> {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let mut current = match sheets.get(&transaction, "sheet-data") {
        Some(Out::YMap(record)) => record,
        _ => return None,
    };
    for parent in parents {
        current = match current.get(&transaction, parent) {
            Some(Out::YMap(map)) => map,
            _ => return None,
        };
    }
    let mut identity = vec![kind.to_owned()];
    identity.extend(path.iter().copied().map(str::to_owned));
    let key = serde_json::to_string(&identity).unwrap();
    match current.get(&transaction, &key) {
        Some(Out::Any(value)) => Some(value),
        _ => None,
    }
}

fn series_values_reference(store: &NativeOfficeCollaborationStore) -> Option<String> {
    let Any::Array(items) = region_any(store, &["charts", "chart-1"], "value", &["series"])? else {
        return None;
    };
    let Any::Map(series) = items.first()? else {
        return None;
    };
    match series.get("valuesReference")? {
        Any::String(value) => Some(value.to_string()),
        _ => None,
    }
}

fn column_formula(
    store: &NativeOfficeCollaborationStore,
    index: usize,
    field: &str,
) -> Option<String> {
    let Any::Array(columns) = region_any(store, &["tables", "table-1"], "value", &["columns"])?
    else {
        return None;
    };
    let Any::Map(column) = columns.get(index)? else {
        return None;
    };
    match column.get(field)? {
        Any::String(value) => Some(value.to_string()),
        _ => None,
    }
}

fn column_count(store: &NativeOfficeCollaborationStore) -> Option<usize> {
    let Any::Array(columns) = region_any(store, &["tables", "table-1"], "value", &["columns"])?
    else {
        return None;
    };
    Some(columns.len())
}

fn table_ids(store: &NativeOfficeCollaborationStore) -> Vec<String> {
    let peer = replica_document(store);
    let sheets = peer.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = peer.transact();
    let record = match sheets.get(&transaction, "sheet-data") {
        Some(Out::YMap(record)) => record,
        _ => return Vec::new(),
    };
    let Some(Out::YArray(order)) = record.get(&transaction, "tableOrder") else {
        return Vec::new();
    };
    let mut ids = Vec::new();
    for index in 0..order.len(&transaction) {
        if let Some(Out::Any(Any::String(id))) = order.get(&transaction, index) {
            ids.push(id.to_string());
        }
    }
    ids
}

fn any_u32(value: &Any) -> Option<u32> {
    let Any::Number(number) = value else {
        return None;
    };
    if !number.is_finite() || number.fract() != 0.0 || *number < 0.0 {
        return None;
    }
    let integer = *number as u32;
    (f64::from(integer) == *number).then_some(integer)
}
