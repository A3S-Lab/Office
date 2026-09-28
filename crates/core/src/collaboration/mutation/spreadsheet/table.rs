//! A table is one record in the sheet's `tables` map.
//!
//! Create, update, and delete write that record and `tableOrder`. Create also
//! appends the browser record claim. A stale observed record, an overlapping
//! range, or a repeated name writes nothing. Deleting the record leaves the
//! claim, so the same ID cannot later name a different table. A name, column,
//! or geometry change rewrites structured references in that same transaction.
//! A reference that cannot be rewritten, or a delete that is still named by
//! one, writes nothing.

use std::collections::BTreeMap;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Any, Array, ArrayPrelim, Map, MapPrelim, Out, ReadTxn, Transact};

use super::json::{
    any_to_json, decode_flat_json_key, flattened_cell, json_equal, json_to_any, reconstruct_cell,
    DecodedFlatJsonEntry, FlatJsonEntryKind,
};
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, spreadsheet_match_conflict,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation, MAX_SPREADSHEET_COLUMNS,
    MAX_SPREADSHEET_ROWS,
};
use crate::collaboration::mutation::document::comment::canonical_json;

const TABLES_KEY: &str = "tables";
const TABLE_ORDER_KEY: &str = "tableOrder";

pub(super) fn is_table_mutation(mutation: &NativeOfficeCollaborationMutation) -> bool {
    matches!(
        mutation,
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { .. }
            | NativeOfficeCollaborationMutation::SpreadsheetUpdateTable { .. }
            | NativeOfficeCollaborationMutation::SpreadsheetDeleteTable { .. }
    )
}

pub(in crate::collaboration) fn table_origin(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<(String, u32, u32)> {
    let (sheet_id, table) = match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { sheet_id, table } => {
            (sheet_id, table)
        }
        NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
            sheet_id,
            next_table,
            ..
        } => (sheet_id, next_table),
        NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
            sheet_id,
            expected_table,
            ..
        } => (sheet_id, expected_table),
        _ => {
            return Err(invalid_spreadsheet_mutation(
                "The supplied mutation is not a Spreadsheet table frame.",
            ))
        }
    };
    let id = table_id_of(mutation)?;
    let table = checked_table(table, &id)?;
    let (row, _) = axis(table.get("range"), "row")?;
    let (column, _) = axis(table.get("range"), "column")?;
    Ok((sheet_id.clone(), row, column))
}

pub(super) fn validate_table_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { sheet_id, table } => {
            require_sheet_id(sheet_id)?;
            let id = required_identifier(table.get("id"), "Spreadsheet table ID")?;
            checked_table(table, &id)?;
        }
        NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
            sheet_id,
            table_id,
            expected_table,
            next_table,
        } => {
            require_sheet_id(sheet_id)?;
            require_identifier(table_id, "Spreadsheet table ID")?;
            checked_table(expected_table, table_id)?;
            checked_table(next_table, table_id)?;
        }
        NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
            sheet_id,
            table_id,
            expected_table,
        } => {
            require_sheet_id(sheet_id)?;
            require_identifier(table_id, "Spreadsheet table ID")?;
            checked_table(expected_table, table_id)?;
        }
        _ => {
            return Err(invalid_spreadsheet_mutation(
                "The supplied mutation is not a Spreadsheet table frame.",
            ))
        }
    }
    Ok(())
}

pub(super) fn apply_table_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let sheets = doc.get_or_insert_map(format!("{}.spreadsheet.sheets", manifest.namespace));
    let claims =
        doc.get_or_insert_array(format!("{}.spreadsheet.record-claims", manifest.namespace));
    let named_root = format!("{}.spreadsheet.named-ranges", manifest.namespace);
    let transaction = doc.transact();
    let has_named = transaction.root_refs().any(|(name, _)| name == named_root);
    drop(transaction);
    let named_ranges = has_named.then(|| doc.get_or_insert_map(named_root));
    let transaction = doc.transact();
    let sheet_id = sheet_id_of(mutation);
    let record = match sheets.get(&transaction, sheet_id) {
        Some(Out::YMap(record)) => record,
        Some(_) => {
            return Err(invalid_shared_spreadsheet(format!(
                "Shared Spreadsheet sheet '{sheet_id}' is not a typed map."
            )))
        }
        None => {
            return Err(spreadsheet_match_conflict(format!(
                "Spreadsheet sheet ID '{sheet_id}' does not exist."
            )))
        }
    };
    let live = read_tables(&record, &transaction)?;
    let reserved_names = reserved_names(
        &sheets,
        named_ranges.as_ref(),
        &transaction,
        sheet_id,
        &live,
    )?;
    let sheet_name = super::table_appearance::sheet_display_name(&record, &transaction, sheet_id)?;
    let mut plan = plan_table_write(mutation, &live, &reserved_names, &claims, &transaction)?;
    drop(transaction);
    let (rewritten_next, writes, border) = super::table_appearance::prepare_table_frame(
        doc,
        manifest,
        mutation,
        sheet_id,
        &sheet_name,
        &plan,
    )?;
    if let Some(next) = rewritten_next {
        if let TablePlan::Update { fields, .. } = &mut plan {
            *fields = flat_fields(&next)?;
        }
    }
    commit_table(doc, &record, &claims, plan, writes, border)
}

fn sheet_id_of(mutation: &NativeOfficeCollaborationMutation) -> &str {
    match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { sheet_id, .. }
        | NativeOfficeCollaborationMutation::SpreadsheetUpdateTable { sheet_id, .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteTable { sheet_id, .. } => sheet_id,
        _ => "",
    }
}

fn table_id_of(mutation: &NativeOfficeCollaborationMutation) -> UseResult<String> {
    match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { table, .. } => {
            required_identifier(table.get("id"), "Spreadsheet table ID")
        }
        NativeOfficeCollaborationMutation::SpreadsheetUpdateTable { table_id, .. }
        | NativeOfficeCollaborationMutation::SpreadsheetDeleteTable { table_id, .. } => {
            Ok(table_id.clone())
        }
        _ => Err(invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet table frame.",
        )),
    }
}

struct LiveTable {
    value: JsonValue,
}

pub(super) enum TablePlan {
    Create {
        id: String,
        fields: Vec<(String, Any)>,
        claim: String,
    },
    Update {
        id: String,
        fields: Vec<(String, Any)>,
    },
    Delete {
        id: String,
    },
    Unchanged,
}

fn plan_table_write<T: ReadTxn>(
    mutation: &NativeOfficeCollaborationMutation,
    live: &BTreeMap<String, LiveTable>,
    reserved_names: &BTreeMap<String, String>,
    claims: &yrs::ArrayRef,
    transaction: &T,
) -> UseResult<TablePlan> {
    match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetCreateTable { table, sheet_id } => {
            let id = required_identifier(table.get("id"), "Spreadsheet table ID")?;
            let table = checked_table(table, &id)?;
            refuse_name_and_overlap(&id, &table, live, reserved_names)?;
            let fingerprint = canonical_json(&table)?;
            if let Some(existing) = live.get(&id) {
                if json_equal(&existing.value, &table) {
                    return Ok(TablePlan::Unchanged);
                }
                return Err(invalid_spreadsheet_mutation(format!(
                    "Spreadsheet table ID '{id}' already names a different record."
                )));
            }
            if let Some(claimed) = claim_fingerprint(claims, transaction, sheet_id, &id)? {
                if claimed != fingerprint {
                    return Err(invalid_spreadsheet_mutation(format!(
                        "Spreadsheet table ID '{id}' is reserved for a different record."
                    )));
                }
            }
            Ok(TablePlan::Create {
                id,
                fields: flat_fields(&table)?,
                claim: table_claim(
                    sheet_id,
                    table
                        .get("id")
                        .and_then(JsonValue::as_str)
                        .unwrap_or_default(),
                    &fingerprint,
                )?,
            })
        }
        NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
            table_id,
            expected_table,
            next_table,
            ..
        } => {
            let expected = checked_table(expected_table, table_id)?;
            let next = checked_table(next_table, table_id)?;
            let Some(existing) = live.get(table_id) else {
                return Err(spreadsheet_match_conflict(format!(
                    "Spreadsheet table '{table_id}' does not exist."
                )));
            };
            if !json_equal(&existing.value, &expected) {
                return Err(spreadsheet_match_conflict(format!(
                    "Spreadsheet table '{table_id}' changed before this update."
                )));
            }
            if json_equal(&expected, &next) {
                return Ok(TablePlan::Unchanged);
            }
            refuse_name_and_overlap(table_id, &next, live, reserved_names)?;
            Ok(TablePlan::Update {
                id: table_id.clone(),
                fields: flat_fields(&next)?,
            })
        }
        NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
            table_id,
            expected_table,
            ..
        } => {
            let expected = checked_table(expected_table, table_id)?;
            let Some(existing) = live.get(table_id) else {
                return Err(spreadsheet_match_conflict(format!(
                    "Spreadsheet table '{table_id}' does not exist."
                )));
            };
            if !json_equal(&existing.value, &expected) {
                return Err(spreadsheet_match_conflict(format!(
                    "Spreadsheet table '{table_id}' changed before it was deleted."
                )));
            }
            Ok(TablePlan::Delete {
                id: table_id.clone(),
            })
        }
        _ => Err(invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet table frame.",
        )),
    }
}

fn commit_table(
    doc: &yrs::Doc,
    record: &yrs::MapRef,
    claims: &yrs::ArrayRef,
    plan: TablePlan,
    writes: Vec<super::state::SpreadsheetReferenceWrite>,
    border: Option<super::table_appearance::BorderWrite>,
) -> UseResult<()> {
    match plan {
        TablePlan::Unchanged => Ok(()),
        TablePlan::Create { id, fields, claim } => {
            ready_container(doc, record, TABLES_KEY, true)?;
            ready_container(doc, record, TABLE_ORDER_KEY, false)?;
            let mut transaction = doc.transact_mut();
            let tables = match record.get(&transaction, TABLES_KEY) {
                Some(Out::YMap(map)) => map,
                _ => record.insert(&mut transaction, TABLES_KEY, MapPrelim::default()),
            };
            let order = match record.get(&transaction, TABLE_ORDER_KEY) {
                Some(Out::YArray(array)) => array,
                _ => record.insert(&mut transaction, TABLE_ORDER_KEY, ArrayPrelim::default()),
            };
            let table = tables.insert(&mut transaction, id.clone(), MapPrelim::default());
            for (key, value) in fields {
                table.insert(&mut transaction, key, value);
            }
            order.push_back(&mut transaction, id);
            if !claim_present(claims, &transaction, &claim) {
                claims.push_back(&mut transaction, claim);
            }
            apply_formula_writes(&mut transaction, writes);
            Ok(())
        }
        TablePlan::Update { id, fields } => {
            let table = {
                let transaction = doc.transact();
                let tables = required_map(record, &transaction, TABLES_KEY)?;
                match tables.get(&transaction, &id) {
                    Some(Out::YMap(table)) => table,
                    _ => {
                        return Err(invalid_shared_spreadsheet(
                            "A shared Spreadsheet table is not a typed map.",
                        ))
                    }
                }
            };
            let mut transaction = doc.transact_mut();
            let existing = table
                .keys(&transaction)
                .map(str::to_owned)
                .collect::<Vec<_>>();
            let written = fields
                .iter()
                .map(|(key, _)| key.clone())
                .collect::<std::collections::BTreeSet<_>>();
            for key in existing {
                if !written.contains(&key) {
                    table.remove(&mut transaction, &key);
                }
            }
            for (key, value) in fields {
                table.insert(&mut transaction, key, value);
            }
            apply_formula_writes(&mut transaction, writes);
            Ok(())
        }
        TablePlan::Delete { id } => {
            let indexes = {
                let transaction = doc.transact();
                required_map(record, &transaction, TABLES_KEY)?;
                let order = optional_array(record, &transaction, TABLE_ORDER_KEY)?;
                let mut indexes = Vec::new();
                if let Some(order) = order {
                    for index in 0..order.len(&transaction) {
                        if let Some(Out::Any(Any::String(value))) = order.get(&transaction, index) {
                            if value.as_ref() == id {
                                indexes.push(index);
                            }
                        }
                    }
                }
                if border.is_some() {
                    match record.get(&transaction, "config") {
                        None | Some(Out::YMap(_)) => {}
                        Some(_) => {
                            return Err(invalid_shared_spreadsheet(
                                "Shared Spreadsheet 'config' is not a typed map.",
                            ));
                        }
                    }
                }
                indexes
            };
            let mut transaction = doc.transact_mut();
            if let Some(Out::YMap(tables)) = record.get(&transaction, TABLES_KEY) {
                tables.remove(&mut transaction, &id);
            }
            if let Some(Out::YArray(order)) = record.get(&transaction, TABLE_ORDER_KEY) {
                for index in indexes.into_iter().rev() {
                    order.remove(&mut transaction, index);
                }
            }
            apply_formula_writes(&mut transaction, writes);
            if let Some(border) = border {
                match record.get(&transaction, "config") {
                    Some(Out::YMap(config)) => {
                        config.insert(&mut transaction, border.key, border.value);
                    }
                    None => {
                        let config =
                            record.insert(&mut transaction, "config", MapPrelim::default());
                        config.insert(&mut transaction, border.key, border.value);
                    }
                    Some(_) => {}
                }
            }
            Ok(())
        }
    }
}

fn apply_formula_writes(
    transaction: &mut yrs::TransactionMut<'_>,
    writes: Vec<super::state::SpreadsheetReferenceWrite>,
) {
    for write in writes {
        write.map.insert(transaction, write.key, write.value);
    }
}

fn ready_container(doc: &yrs::Doc, record: &yrs::MapRef, key: &str, map: bool) -> UseResult<()> {
    let transaction = doc.transact();
    match record.get(&transaction, key) {
        None => Ok(()),
        Some(Out::YMap(_)) if map => Ok(()),
        Some(Out::YArray(_)) if !map => Ok(()),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet '{key}' has the wrong type."
        ))),
    }
}

fn claim_present<T: ReadTxn>(claims: &yrs::ArrayRef, transaction: &T, claim: &str) -> bool {
    (0..claims.len(transaction)).any(|index| {
        matches!(
            claims.get(transaction, index),
            Some(Out::Any(Any::String(value))) if value.as_ref() == claim
        )
    })
}

fn required_map<T: ReadTxn>(
    record: &yrs::MapRef,
    transaction: &T,
    key: &str,
) -> UseResult<yrs::MapRef> {
    match record.get(transaction, key) {
        Some(Out::YMap(map)) => Ok(map),
        _ => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet '{key}' is not a typed map."
        ))),
    }
}

fn optional_array<T: ReadTxn>(
    record: &yrs::MapRef,
    transaction: &T,
    key: &str,
) -> UseResult<Option<yrs::ArrayRef>> {
    match record.get(transaction, key) {
        None => Ok(None),
        Some(Out::YArray(array)) => Ok(Some(array)),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet '{key}' is not a typed array."
        ))),
    }
}

fn read_tables<T: ReadTxn>(
    record: &yrs::MapRef,
    transaction: &T,
) -> UseResult<BTreeMap<String, LiveTable>> {
    let Some(Out::YMap(tables)) = record.get(transaction, TABLES_KEY) else {
        if record.get(transaction, TABLE_ORDER_KEY).is_some() {
            return Err(invalid_shared_spreadsheet(
                "Shared Spreadsheet table order exists without table records.",
            ));
        }
        return Ok(BTreeMap::new());
    };
    let mut live = BTreeMap::new();
    for (id, value) in tables.iter(transaction) {
        let Out::YMap(map) = value else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet table is not a typed map.",
            ));
        };
        live.insert(
            id.to_owned(),
            LiveTable {
                value: read_flat_record(&map, transaction)?,
            },
        );
    }
    Ok(live)
}

fn read_flat_record<T: ReadTxn>(map: &yrs::MapRef, transaction: &T) -> UseResult<JsonValue> {
    let mut fields = Vec::new();
    for (key, value) in map.iter(transaction) {
        let (kind, path) = decode_flat_json_key(&key)?;
        let value = match (kind, value) {
            (FlatJsonEntryKind::Object, Out::Any(Any::Bool(true))) => JsonValue::Bool(true),
            (FlatJsonEntryKind::Value, Out::Any(value)) => any_to_json(value, 0)?,
            _ => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet table field is not a flat JSON value.",
                ))
            }
        };
        fields.push(DecodedFlatJsonEntry { kind, path, value });
    }
    reconstruct_cell(fields)
}

fn flat_fields(value: &JsonValue) -> UseResult<Vec<(String, Any)>> {
    flattened_cell(value)?
        .into_iter()
        .map(|(key, value)| Ok((key, json_to_any(&value, 0)?)))
        .collect()
}

fn reserved_names<T: ReadTxn>(
    sheets: &yrs::MapRef,
    named_ranges: Option<&yrs::MapRef>,
    transaction: &T,
    sheet_id: &str,
    live: &BTreeMap<String, LiveTable>,
) -> UseResult<BTreeMap<String, String>> {
    let mut names = BTreeMap::new();
    for (id, table) in live {
        remember_table_names(&mut names, id, &table.value)?;
    }
    for (other_sheet, value) in sheets.iter(transaction) {
        if other_sheet == sheet_id {
            continue;
        }
        let Out::YMap(record) = value else {
            continue;
        };
        for (id, table) in read_tables(&record, transaction)? {
            remember_table_names(&mut names, &id, &table.value)?;
        }
    }
    if let Some(named_ranges) = named_ranges {
        for (_id, value) in named_ranges.iter(transaction) {
            let Out::YMap(record) = value else {
                continue;
            };
            let value = read_flat_record(&record, transaction)?;
            if let Some(name) = value.get("name").and_then(JsonValue::as_str) {
                names
                    .entry(name.to_lowercase())
                    .or_insert_with(|| format!("named range '{name}'"));
            }
        }
    }
    Ok(names)
}

fn remember_table_names(
    names: &mut BTreeMap<String, String>,
    id: &str,
    table: &JsonValue,
) -> UseResult<()> {
    for key in ["name", "displayName"] {
        if let Some(name) = table.get(key).and_then(JsonValue::as_str) {
            names
                .entry(name.to_lowercase())
                .or_insert_with(|| format!("table '{id}'"));
        }
    }
    Ok(())
}

fn refuse_name_and_overlap(
    id: &str,
    table: &JsonValue,
    live: &BTreeMap<String, LiveTable>,
    reserved_names: &BTreeMap<String, String>,
) -> UseResult<()> {
    for key in ["name", "displayName"] {
        let Some(name) = table.get(key).and_then(JsonValue::as_str) else {
            continue;
        };
        if let Some(owner) = reserved_names.get(&name.to_lowercase()) {
            if owner != &format!("table '{id}'") {
                return Err(invalid_spreadsheet_mutation(format!(
                    "Spreadsheet table name '{name}' is already used by {owner}."
                )));
            }
        }
    }
    let (row_start, row_end) = axis(table.get("range"), "row")?;
    let (column_start, column_end) = axis(table.get("range"), "column")?;
    for (other_id, other) in live {
        if other_id == id {
            continue;
        }
        let (other_row_start, other_row_end) = axis(other.value.get("range"), "row")?;
        let (other_column_start, other_column_end) = axis(other.value.get("range"), "column")?;
        let rows = row_start <= other_row_end && row_end >= other_row_start;
        let columns = column_start <= other_column_end && column_end >= other_column_start;
        if rows && columns {
            return Err(invalid_spreadsheet_mutation(format!(
                "Spreadsheet table '{id}' overlaps table '{other_id}'."
            )));
        }
    }
    Ok(())
}

fn claim_fingerprint<T: ReadTxn>(
    claims: &yrs::ArrayRef,
    transaction: &T,
    sheet_id: &str,
    id: &str,
) -> UseResult<Option<String>> {
    let mut found = None;
    for index in 0..claims.len(transaction) {
        let Some(Out::Any(Any::String(raw))) = claims.get(transaction, index) else {
            continue;
        };
        let Ok(value) = serde_json::from_str::<JsonValue>(raw.as_ref()) else {
            continue;
        };
        if value.get("kind").and_then(JsonValue::as_str) != Some("table") {
            continue;
        }
        if value.get("id").and_then(JsonValue::as_str) != Some(id) {
            continue;
        }
        if value.get("parentId").and_then(JsonValue::as_str) != Some(sheet_id) {
            continue;
        }
        let Some(fingerprint) = value.get("fingerprint").and_then(JsonValue::as_str) else {
            continue;
        };
        if found
            .as_deref()
            .is_some_and(|existing| existing != fingerprint)
        {
            return Err(invalid_spreadsheet_mutation(format!(
                "Spreadsheet table ID '{id}' has conflicting creation claims."
            )));
        }
        found = Some(fingerprint.to_owned());
    }
    Ok(found)
}

fn table_claim(sheet_id: &str, id: &str, fingerprint: &str) -> UseResult<String> {
    canonical_json(&serde_json::json!({
        "fingerprint": fingerprint,
        "id": id,
        "kind": "table",
        "parentId": sheet_id,
    }))
}

fn checked_table<'a>(value: &'a JsonValue, id: &str) -> UseResult<&'a JsonValue> {
    let object = value.as_object().ok_or_else(|| {
        invalid_spreadsheet_mutation("A Spreadsheet table must be a JSON object.")
    })?;
    let record_id = required_identifier(object.get("id"), "Spreadsheet table ID")?;
    if record_id != id {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet table ID does not match the record.",
        ));
    }
    valid_table_name(object.get("name"))?;
    if object.contains_key("displayName") {
        valid_table_name(object.get("displayName"))?;
    }
    if let Some(ooxml_id) = object.get("ooxmlId") {
        if ooxml_id.as_u64().is_none_or(|value| value < 1) {
            return Err(invalid_spreadsheet_mutation(
                "A Spreadsheet table OOXML ID must be a positive integer.",
            ));
        }
    }
    let (row_start, row_end) = axis(object.get("range"), "row")?;
    let (column_start, column_end) = axis(object.get("range"), "column")?;
    let header = required_bool(object.get("headerRow"), "header row")?;
    let totals = required_bool(object.get("totalsRow"), "totals row")?;
    let height = row_end - row_start + 1;
    if height <= u32::from(header) + u32::from(totals) {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet table must include at least one body row.",
        ));
    }
    let width = column_end - column_start + 1;
    checked_columns(object.get("columns"), width)?;
    let filter_count = checked_filters(object.get("filters"), width)?;
    if !header && filter_count > 0 {
        return Err(invalid_spreadsheet_mutation(
            "Spreadsheet table filters require a header row.",
        ));
    }
    let style_none = checked_style(object.get("style"))?;
    for key in [
        "showFirstColumn",
        "showLastColumn",
        "showRowStripes",
        "showColumnStripes",
    ] {
        let enabled = required_bool(object.get(key), key)?;
        if style_none && enabled {
            return Err(invalid_spreadsheet_mutation(
                "Spreadsheet table style flags require a built-in style.",
            ));
        }
    }
    let _ = (row_start, column_start);
    Ok(value)
}

fn checked_columns(value: Option<&JsonValue>, width: u32) -> UseResult<()> {
    let columns = value.and_then(JsonValue::as_array).ok_or_else(|| {
        invalid_spreadsheet_mutation("A Spreadsheet table needs one column per worksheet column.")
    })?;
    if columns.len() != width as usize {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet table needs one column per worksheet column.",
        ));
    }
    let mut names = std::collections::BTreeSet::new();
    for column in columns {
        let name = column
            .get("name")
            .and_then(JsonValue::as_str)
            .ok_or_else(|| {
                invalid_spreadsheet_mutation("A Spreadsheet table column needs a name.")
            })?;
        if name.is_empty()
            || name != name.trim()
            || name.chars().count() > 255
            || name.chars().any(|character| character.is_control())
        {
            return Err(invalid_spreadsheet_mutation(
                "A Spreadsheet table column name is invalid.",
            ));
        }
        if !names.insert(name.to_lowercase()) {
            return Err(invalid_spreadsheet_mutation(
                "Spreadsheet table column names must be unique.",
            ));
        }
    }
    Ok(())
}

fn checked_filters(value: Option<&JsonValue>, width: u32) -> UseResult<usize> {
    let filters = value
        .and_then(JsonValue::as_array)
        .ok_or_else(|| invalid_spreadsheet_mutation("A Spreadsheet table needs a filter array."))?;
    let mut columns = std::collections::BTreeSet::new();
    for filter in filters {
        let column = filter
            .get("column")
            .and_then(JsonValue::as_u64)
            .ok_or_else(|| {
                invalid_spreadsheet_mutation("A Spreadsheet table filter needs a column.")
            })?;
        if column >= u64::from(width) || !columns.insert(column) {
            return Err(invalid_spreadsheet_mutation(
                "Spreadsheet table filter columns must be unique and inside the table.",
            ));
        }
        if !filter.get("criteria").is_some_and(JsonValue::is_object) {
            return Err(invalid_spreadsheet_mutation(
                "A Spreadsheet table filter needs criteria.",
            ));
        }
    }
    Ok(filters.len())
}

fn checked_style(value: Option<&JsonValue>) -> UseResult<bool> {
    let style = value
        .and_then(JsonValue::as_object)
        .ok_or_else(|| invalid_spreadsheet_mutation("A Spreadsheet table needs a style."))?;
    match style.get("family").and_then(JsonValue::as_str) {
        Some("none") if style.len() == 1 => Ok(true),
        Some(family @ ("light" | "medium" | "dark")) if style.len() == 2 => {
            let maximum = match family {
                "light" => 21,
                "medium" => 28,
                _ => 11,
            };
            let number = style
                .get("number")
                .and_then(JsonValue::as_u64)
                .ok_or_else(|| {
                    invalid_spreadsheet_mutation("A Spreadsheet table style number is invalid.")
                })?;
            if (1..=maximum).contains(&number) {
                Ok(false)
            } else {
                Err(invalid_spreadsheet_mutation(
                    "A Spreadsheet table style number is invalid.",
                ))
            }
        }
        _ => Err(invalid_spreadsheet_mutation(
            "A Spreadsheet table style is invalid.",
        )),
    }
}

fn axis(range: Option<&JsonValue>, key: &str) -> UseResult<(u32, u32)> {
    let pair = range
        .and_then(|range| range.get(key))
        .and_then(JsonValue::as_array)
        .filter(|pair| pair.len() == 2)
        .ok_or_else(|| {
            invalid_spreadsheet_mutation(format!(
                "A Spreadsheet table needs an ordered {key} range."
            ))
        })?;
    let start = pair[0].as_u64().ok_or_else(|| {
        invalid_spreadsheet_mutation(format!("A Spreadsheet table {key} range is invalid."))
    })?;
    let end = pair[1].as_u64().ok_or_else(|| {
        invalid_spreadsheet_mutation(format!("A Spreadsheet table {key} range is invalid."))
    })?;
    let maximum = if key == "row" {
        u64::from(MAX_SPREADSHEET_ROWS)
    } else {
        u64::from(MAX_SPREADSHEET_COLUMNS)
    };
    if start > end || end >= maximum {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet table {key} range is outside the sheet."
        )));
    }
    Ok((start as u32, end as u32))
}

fn valid_table_name(value: Option<&JsonValue>) -> UseResult<()> {
    let name = value
        .and_then(JsonValue::as_str)
        .ok_or_else(|| invalid_spreadsheet_mutation("A Spreadsheet table needs a name."))?;
    let characters = name.chars().count();
    let mut chars = name.chars();
    let first = chars.next();
    let body_ok = chars.all(|character| {
        character.is_alphanumeric() || character == '_' || character == '.' || character == '\\'
    });
    if name != name.trim()
        || characters == 0
        || characters > 255
        || name.eq_ignore_ascii_case("r")
        || name.eq_ignore_ascii_case("c")
        || name.to_ascii_lowercase().starts_with("_xlnm.")
        || !first.is_some_and(|character| {
            character.is_alphabetic() || character == '_' || character == '\\'
        })
        || !body_ok
        || looks_like_cell(name)
        || looks_like_rc(name)
    {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet table name is invalid.",
        ));
    }
    Ok(())
}

fn looks_like_cell(name: &str) -> bool {
    let mut letters = 0;
    let mut digits = 0;
    for character in name.chars() {
        if digits == 0 && character.is_ascii_alphabetic() {
            letters += 1;
            if letters > 3 {
                return false;
            }
        } else if letters > 0 && character.is_ascii_digit() {
            if digits == 0 && character == '0' {
                return false;
            }
            digits += 1;
        } else {
            return false;
        }
    }
    letters > 0 && digits > 0
}

fn looks_like_rc(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    let mut parts = lower.split('c');
    let Some(row) = parts.next() else {
        return false;
    };
    let Some(column) = parts.next() else {
        return false;
    };
    parts.next().is_none()
        && row.len() > 1
        && row.starts_with('r')
        && row.as_bytes()[1..].iter().all(u8::is_ascii_digit)
        && !column.is_empty()
        && column.bytes().all(|byte| byte.is_ascii_digit())
}

fn required_bool(value: Option<&JsonValue>, label: &str) -> UseResult<bool> {
    value.and_then(JsonValue::as_bool).ok_or_else(|| {
        invalid_spreadsheet_mutation(format!("A Spreadsheet table needs a Boolean {label}."))
    })
}

fn required_identifier(value: Option<&JsonValue>, label: &str) -> UseResult<String> {
    require_identifier(value.and_then(JsonValue::as_str).unwrap_or(""), label)
}

fn require_sheet_id(value: &str) -> UseResult<()> {
    require_identifier(value, "Spreadsheet sheet ID").map(|_| ())
}

fn require_identifier(value: &str, label: &str) -> UseResult<String> {
    if value.is_empty() || value != value.trim() || value.chars().count() > 256 {
        return Err(invalid_spreadsheet_mutation(format!("{label} is invalid.")));
    }
    Ok(value.to_owned())
}
