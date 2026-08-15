use std::collections::HashSet;

use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use serde_json::{Map as JsonMap, Value as JsonValue};
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, Transact, Update};

const SPREADSHEET_COLLABORATION_FIXTURE_BASE64: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../tests/fixtures/browser-spreadsheet-collaboration-update.base64"
));

pub fn spreadsheet_collaboration_fixture() -> Vec<u8> {
    STANDARD
        .decode(SPREADSHEET_COLLABORATION_FIXTURE_BASE64.trim())
        .unwrap()
}

pub fn spreadsheet_cell(update: &[u8], sheet_id: &str, row: u32, column: u32) -> Option<JsonValue> {
    let document = Doc::with_client_id(710_014);
    document
        .transact_mut()
        .apply_update(Update::decode_v1(update).unwrap())
        .unwrap();
    let sheets = document.get_or_insert_map("a3s.office.spreadsheet.sheets");
    let transaction = document.transact();
    let record = match sheets.get(&transaction, sheet_id) {
        Some(Out::YMap(record)) => record,
        value => panic!("unexpected Spreadsheet sheet: {value:?}"),
    };
    let presence = match record.get(&transaction, "cellPresence") {
        Some(Out::YMap(presence)) => presence,
        value => panic!("unexpected Spreadsheet cell presence: {value:?}"),
    };
    if !matches!(
        presence.get(&transaction, format!("{row}:{column}").as_str()),
        Some(Out::Any(Any::Bool(true)))
    ) {
        return None;
    }
    let fields = match record.get(&transaction, "cells") {
        Some(Out::YMap(fields)) => fields,
        value => panic!("unexpected Spreadsheet cell fields: {value:?}"),
    };
    let mut entries = Vec::new();
    for (encoded, value) in fields.iter(&transaction) {
        let (field_row, field_column, flat_key) =
            serde_json::from_str::<(u32, u32, String)>(encoded).unwrap();
        if field_row != row || field_column != column {
            continue;
        }
        let identity = serde_json::from_str::<Vec<String>>(&flat_key).unwrap();
        let kind = identity[0].clone();
        let value = match (kind.as_str(), value) {
            ("object", Out::Any(Any::Bool(true))) => JsonValue::Bool(true),
            ("value", Out::Any(value)) => serde_json::to_value(value).unwrap(),
            _ => panic!("unexpected Spreadsheet cell field"),
        };
        entries.push((kind, identity[1..].to_vec(), value));
    }
    entries.sort_by(|left, right| {
        left.1.len().cmp(&right.1.len()).then_with(|| {
            if left.0 == right.0 {
                std::cmp::Ordering::Equal
            } else if left.0 == "object" {
                std::cmp::Ordering::Less
            } else {
                std::cmp::Ordering::Greater
            }
        })
    });
    let object_paths = entries
        .iter()
        .filter(|(kind, _, _)| kind == "object")
        .map(|(_, path, _)| path.clone())
        .collect::<HashSet<_>>();
    for (_, path, _) in &entries {
        for depth in 1..path.len() {
            assert!(object_paths.contains(&path[..depth]));
        }
    }
    let mut cell = JsonMap::new();
    for (kind, path, value) in entries {
        if kind == "object" {
            ensure_object_path(&mut cell, &path);
        } else {
            set_value(&mut cell, &path, value);
        }
    }
    Some(JsonValue::Object(cell))
}

fn ensure_object_path(object: &mut JsonMap<String, JsonValue>, path: &[String]) {
    let (key, rest) = path.split_first().expect("non-empty cell field path");
    let child = object
        .entry(key.clone())
        .or_insert_with(|| JsonValue::Object(JsonMap::new()))
        .as_object_mut()
        .expect("Spreadsheet object marker must not overlap a value");
    if !rest.is_empty() {
        ensure_object_path(child, rest);
    }
}

fn set_value(object: &mut JsonMap<String, JsonValue>, path: &[String], value: JsonValue) {
    let (key, rest) = path.split_first().expect("non-empty cell field path");
    if rest.is_empty() {
        assert!(object.insert(key.clone(), value).is_none());
        return;
    }
    let child = object
        .get_mut(key)
        .and_then(JsonValue::as_object_mut)
        .expect("Spreadsheet value must have an object-marker parent");
    set_value(child, rest, value);
}
