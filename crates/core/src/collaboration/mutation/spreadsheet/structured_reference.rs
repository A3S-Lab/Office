//! Rewrite structured references in the same transaction as a table update.
//! A geometry change or an unprovable local reference plans nothing. Deleting
//! a table that a structured reference still names plans nothing.

use std::collections::BTreeMap;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{Any, Map, Out, ReadTxn, Transact};

use super::json::{
    any_to_json, decode_flat_json_key, encode_flat_json_key, json_equal, json_to_any,
    FlatJsonEntryKind,
};
use super::state::{
    encode_cell_field_key, ordered_sheet_ids, read_sheet_state, SpreadsheetReferenceWrite,
    SpreadsheetSheetState,
};
use super::{invalid_shared_spreadsheet, NativeOfficeCollaborationManifest};
use crate::spreadsheet_formula::{LocalStructuredReferenceContext, StructuredReferenceRewritePlan};

const CHARTS_KEY: &str = "charts";
const TABLES_KEY: &str = "tables";
const FORMULA_METADATA_KEY: &str = "formulaMetadata";
const FORMULA_FIELDS: &[&str] = &[
    "f",
    "formula",
    "reference",
    "calculatedFormula",
    "totalsFormula",
    "titleReference",
    "categoryReference",
    "nameReference",
    "valuesReference",
    "xValuesReference",
    "bubbleSizesReference",
];

#[derive(Clone, Copy, PartialEq, Eq)]
struct TableSpan {
    row_start: u32,
    row_end: u32,
    column_start: u32,
    column_end: u32,
}

pub(super) fn rewrite_table_update(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    sheet_name: &str,
    table_id: &str,
    previous: &JsonValue,
    next: &JsonValue,
) -> UseResult<(Option<JsonValue>, Vec<SpreadsheetReferenceWrite>)> {
    let (plan, required) = rename_plan(previous, next, sheet_name)?;
    if !required {
        return Ok((None, Vec::new()));
    }
    let span = table_span(previous)?;
    let mut rewritten = next.clone();
    rewrite_json(
        &mut rewritten,
        &plan,
        LocalStructuredReferenceContext::Applies,
    )?;
    let writes = collect_formula_writes(doc, manifest, sheet_id, table_id, &span, &plan)?;
    let next = (!json_equal(next, &rewritten)).then_some(rewritten);
    Ok((next, writes))
}

pub(super) fn refuse_referenced_table(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    sheet_name: &str,
    table_id: &str,
    previous: &JsonValue,
) -> UseResult<()> {
    let plan = removal_plan(previous, sheet_name)?;
    let span = table_span(previous)?;
    let mut owned = previous.clone();
    rewrite_json(&mut owned, &plan, LocalStructuredReferenceContext::Applies)?;
    collect_formula_writes(doc, manifest, sheet_id, table_id, &span, &plan)?;
    Ok(())
}

fn collect_formula_writes(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    table_sheet_id: &str,
    table_id: &str,
    span: &TableSpan,
    plan: &StructuredReferenceRewritePlan,
) -> UseResult<Vec<SpreadsheetReferenceWrite>> {
    let mut writes = Vec::new();
    for sheet_id in ordered_sheet_ids(doc, manifest)? {
        let state = read_sheet_state(doc, manifest, &sheet_id)?;
        rewrite_cells(&state, &sheet_id, table_sheet_id, span, plan, &mut writes)?;
        let transaction = doc.transact();
        if let Some(Out::YMap(charts)) = state.record.get(&transaction, CHARTS_KEY) {
            for chart in child_maps(&charts, &transaction)? {
                rewrite_map(
                    &chart,
                    &transaction,
                    plan,
                    LocalStructuredReferenceContext::Unknown,
                    &mut writes,
                )?;
            }
        }
        if let Some(Out::YMap(tables)) = state.record.get(&transaction, TABLES_KEY) {
            for (id, table) in tables.iter(&transaction) {
                if sheet_id == table_sheet_id && id == table_id {
                    continue;
                }
                let Out::YMap(table) = table else {
                    return Err(invalid_shared_spreadsheet(
                        "A shared Spreadsheet table is not a typed map.",
                    ));
                };
                rewrite_map(
                    &table,
                    &transaction,
                    plan,
                    LocalStructuredReferenceContext::DoesNotApply,
                    &mut writes,
                )?;
            }
        }
        if let Some(Out::YMap(metadata)) = state.record.get(&transaction, FORMULA_METADATA_KEY) {
            rewrite_metadata(&metadata, &transaction, plan, &mut writes)?;
        }
    }
    rewrite_named_ranges(doc, manifest, table_sheet_id, plan, &mut writes)?;
    Ok(writes)
}

fn rewrite_cells(
    state: &SpreadsheetSheetState,
    sheet_id: &str,
    table_sheet_id: &str,
    span: &TableSpan,
    plan: &StructuredReferenceRewritePlan,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    for ((row, column), cell) in &state.cells {
        let Some(formula) = cell.get("f") else {
            continue;
        };
        let JsonValue::String(formula) = formula else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet cell formula is not text.",
            ));
        };
        let context = if sheet_id == table_sheet_id && span.contains(*row, *column) {
            LocalStructuredReferenceContext::Applies
        } else {
            LocalStructuredReferenceContext::DoesNotApply
        };
        let rewritten = plan.rewrite(formula, context)?;
        if rewritten.formula == *formula {
            continue;
        }
        let flat = encode_flat_json_key(FlatJsonEntryKind::Value, &["f".to_owned()])?;
        let key = encode_cell_field_key(*row, *column, &flat)?;
        writes.push(SpreadsheetReferenceWrite {
            map: state.fields.clone(),
            key,
            value: json_to_any(&JsonValue::String(rewritten.formula), 0)?,
        });
    }
    Ok(())
}

fn rewrite_named_ranges(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    table_sheet_id: &str,
    plan: &StructuredReferenceRewritePlan,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    let root = format!("{}.spreadsheet.named-ranges", manifest.namespace);
    let present = {
        let transaction = doc.transact();
        transaction.root_refs().any(|(name, _)| name == root)
    };
    if !present {
        return Ok(());
    }
    let named = doc.get_or_insert_map(root);
    let transaction = doc.transact();
    for (_, value) in named.iter(&transaction) {
        let Out::YMap(record) = value else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet named range is not a typed map.",
            ));
        };
        let scope = flat_string(&record, &transaction, "scopeSheetId")?;
        let context = match scope.as_deref() {
            Some(scope) if scope != table_sheet_id => LocalStructuredReferenceContext::DoesNotApply,
            _ => LocalStructuredReferenceContext::Unknown,
        };
        if let Some(reference) = flat_string(&record, &transaction, "reference")? {
            push_rewritten_string(&record, "reference", &reference, plan, context, writes)?;
        }
    }
    Ok(())
}

fn rewrite_metadata<T: ReadTxn>(
    map: &yrs::MapRef,
    transaction: &T,
    plan: &StructuredReferenceRewritePlan,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    for (key, value) in map.iter(transaction) {
        let (kind, path) = decode_flat_json_key(&key)?;
        if kind != FlatJsonEntryKind::Value {
            continue;
        }
        let formula_path = path
            .last()
            .is_some_and(|part| part == "reference" || part == "formula" || part == "f")
            || path.first().is_some_and(|part| part == "sourceFormulas");
        if !formula_path {
            continue;
        }
        let Out::Any(any) = value else {
            continue;
        };
        let JsonValue::String(text) = any_to_json(any, 0)? else {
            continue;
        };
        let rewritten = plan.rewrite(&text, LocalStructuredReferenceContext::Unknown)?;
        if rewritten.formula == text {
            continue;
        }
        writes.push(SpreadsheetReferenceWrite {
            map: map.clone(),
            key: key.to_owned(),
            value: json_to_any(&JsonValue::String(rewritten.formula), 0)?,
        });
    }
    Ok(())
}

fn rewrite_map<T: ReadTxn>(
    map: &yrs::MapRef,
    transaction: &T,
    plan: &StructuredReferenceRewritePlan,
    context: LocalStructuredReferenceContext,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    for (key, value) in map.iter(transaction) {
        let (kind, path) = decode_flat_json_key(&key)?;
        if kind != FlatJsonEntryKind::Value {
            continue;
        }
        let Out::Any(any) = value else {
            continue;
        };
        let json = any_to_json(any, 0)?;
        if formula_path(&path) {
            let JsonValue::String(text) = json else {
                continue;
            };
            let rewritten = plan.rewrite(&text, context)?;
            if rewritten.formula == text {
                continue;
            }
            writes.push(SpreadsheetReferenceWrite {
                map: map.clone(),
                key: key.to_owned(),
                value: json_to_any(&JsonValue::String(rewritten.formula), 0)?,
            });
            continue;
        }
        if !(json.is_array() || json.is_object()) {
            continue;
        }
        let mut next = json.clone();
        rewrite_json(&mut next, plan, context)?;
        if json_equal(&json, &next) {
            continue;
        }
        writes.push(SpreadsheetReferenceWrite {
            map: map.clone(),
            key: key.to_owned(),
            value: json_to_any(&next, 0)?,
        });
    }
    Ok(())
}

fn rewrite_json(
    value: &mut JsonValue,
    plan: &StructuredReferenceRewritePlan,
    context: LocalStructuredReferenceContext,
) -> UseResult<()> {
    match value {
        JsonValue::Object(object) => {
            let keys = object.keys().cloned().collect::<Vec<_>>();
            for key in keys {
                if FORMULA_FIELDS.contains(&key.as_str()) {
                    let Some(JsonValue::String(text)) = object.get(&key) else {
                        continue;
                    };
                    let text = text.clone();
                    let rewritten = plan.rewrite(&text, context)?;
                    if rewritten.formula != text {
                        object.insert(key, JsonValue::String(rewritten.formula));
                    }
                    continue;
                }
                if let Some(child) = object.get_mut(&key) {
                    rewrite_json(child, plan, context)?;
                }
            }
        }
        JsonValue::Array(items) => {
            for item in items {
                rewrite_json(item, plan, context)?;
            }
        }
        _ => {}
    }
    Ok(())
}

fn push_rewritten_string(
    map: &yrs::MapRef,
    field: &str,
    text: &str,
    plan: &StructuredReferenceRewritePlan,
    context: LocalStructuredReferenceContext,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    let rewritten = plan.rewrite(text, context)?;
    if rewritten.formula == text {
        return Ok(());
    }
    let key = encode_flat_json_key(FlatJsonEntryKind::Value, &[field.to_owned()])?;
    writes.push(SpreadsheetReferenceWrite {
        map: map.clone(),
        key,
        value: json_to_any(&JsonValue::String(rewritten.formula), 0)?,
    });
    Ok(())
}

fn flat_string<T: ReadTxn>(
    map: &yrs::MapRef,
    transaction: &T,
    field: &str,
) -> UseResult<Option<String>> {
    let key = encode_flat_json_key(FlatJsonEntryKind::Value, &[field.to_owned()])?;
    match map.get(transaction, key.as_str()) {
        Some(Out::Any(Any::String(value))) => Ok(Some(value.to_string())),
        None => Ok(None),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet field '{field}' is not text."
        ))),
    }
}

fn child_maps<T: ReadTxn>(map: &yrs::MapRef, transaction: &T) -> UseResult<Vec<yrs::MapRef>> {
    let mut children = Vec::new();
    for (_, value) in map.iter(transaction) {
        match value {
            Out::YMap(child) => children.push(child),
            _ => {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet chart is not a typed map.",
                ))
            }
        }
    }
    Ok(children)
}

fn rename_plan(
    previous: &JsonValue,
    next: &JsonValue,
    sheet_name: &str,
) -> UseResult<(StructuredReferenceRewritePlan, bool)> {
    let old_name = required_name(previous, "name")?;
    let new_name = required_name(next, "name")?;
    let old_display = text_field(previous, "displayName").unwrap_or_else(|| old_name.clone());
    let new_display = text_field(next, "displayName").unwrap_or_else(|| new_name.clone());
    let mut aliases = BTreeMap::new();
    aliases.insert(old_name.to_lowercase(), new_name.clone());
    aliases.insert(old_display.to_lowercase(), new_display.clone());
    let old_columns = column_names(previous)?;
    let new_columns = column_names(next)?;
    let mut columns = BTreeMap::new();
    for (index, old_column) in old_columns.iter().enumerate() {
        let replacement = new_columns.get(index).cloned();
        if replacement.as_deref() != Some(old_column.as_str()) {
            columns.insert(old_column.to_lowercase(), replacement);
        }
    }
    let aliases_changed = if old_name.eq_ignore_ascii_case(&old_display) {
        old_display != new_display
    } else {
        old_name != new_name || old_display != new_display
    };
    let geometry_changed = table_span(previous)? != table_span(next)?
        || flag(previous, "headerRow")? != flag(next, "headerRow")?
        || flag(previous, "totalsRow")? != flag(next, "totalsRow")?;
    let required = aliases_changed || !columns.is_empty() || geometry_changed;
    Ok((
        StructuredReferenceRewritePlan::rename(
            old_name,
            sheet_name,
            aliases,
            columns,
            geometry_changed,
        ),
        required,
    ))
}

fn removal_plan(
    previous: &JsonValue,
    sheet_name: &str,
) -> UseResult<StructuredReferenceRewritePlan> {
    let name = required_name(previous, "name")?;
    let display = text_field(previous, "displayName").unwrap_or_else(|| name.clone());
    Ok(StructuredReferenceRewritePlan::removal(
        name.clone(),
        sheet_name,
        [name, display],
    ))
}

fn table_span(table: &JsonValue) -> UseResult<TableSpan> {
    let range = table.get("range").ok_or_else(|| {
        invalid_shared_spreadsheet("A shared Spreadsheet table is missing its range.")
    })?;
    let (row_start, row_end) = axis(range, "row")?;
    let (column_start, column_end) = axis(range, "column")?;
    Ok(TableSpan {
        row_start,
        row_end,
        column_start,
        column_end,
    })
}

impl TableSpan {
    fn contains(&self, row: u32, column: u32) -> bool {
        (self.row_start..=self.row_end).contains(&row)
            && (self.column_start..=self.column_end).contains(&column)
    }
}

fn axis(range: &JsonValue, key: &str) -> UseResult<(u32, u32)> {
    let pair = range
        .get(key)
        .and_then(JsonValue::as_array)
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table range is not a pair.")
        })?;
    if pair.len() != 2 {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is not a pair.",
        ));
    }
    Ok((index(&pair[0])?, index(&pair[1])?))
}

fn index(value: &JsonValue) -> UseResult<u32> {
    value
        .as_u64()
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table range is not an integer.")
        })
}

fn column_names(table: &JsonValue) -> UseResult<Vec<String>> {
    let columns = table
        .get("columns")
        .and_then(JsonValue::as_array)
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table is missing its columns.")
        })?;
    columns
        .iter()
        .map(|column| {
            column
                .get("name")
                .and_then(JsonValue::as_str)
                .map(str::to_owned)
                .ok_or_else(|| {
                    invalid_shared_spreadsheet(
                        "A shared Spreadsheet table column is missing its name.",
                    )
                })
        })
        .collect()
}

fn required_name(table: &JsonValue, key: &str) -> UseResult<String> {
    text_field(table, key).ok_or_else(|| {
        invalid_shared_spreadsheet("A shared Spreadsheet table is missing its name.")
    })
}

fn text_field(table: &JsonValue, key: &str) -> Option<String> {
    table
        .get(key)
        .and_then(JsonValue::as_str)
        .map(str::to_owned)
}

fn flag(table: &JsonValue, key: &str) -> UseResult<bool> {
    table.get(key).and_then(JsonValue::as_bool).ok_or_else(|| {
        invalid_shared_spreadsheet("A shared Spreadsheet table flag is not a boolean.")
    })
}

fn formula_path(path: &[String]) -> bool {
    path.last()
        .is_some_and(|part| FORMULA_FIELDS.contains(&part.as_str()))
}
