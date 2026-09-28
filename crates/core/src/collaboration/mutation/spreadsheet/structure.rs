use std::collections::HashMap;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::{GetString, Map, Out, ReadTxn, Transact};

use super::cell::validate_sheet_id;
use super::json::{
    cell_field_patches, decode_flat_json_key, encode_flat_json_key, json_to_any, FlatJsonEntryKind,
};
use super::state::{
    commit_sheet_structure, encode_coordinate, extended_dense_row_lengths, ordered_sheet_ids,
    read_sheet_state, SheetStructureCommit, SpreadsheetCellPresenceChange, SpreadsheetCellWrite,
    SpreadsheetReferenceWrite, SpreadsheetSheetState, SpreadsheetTextMove, CELL_TEXT_KEY,
};
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMutation, MAX_SPREADSHEET_COLUMNS, MAX_SPREADSHEET_ROWS,
};
use crate::spreadsheet_formula::{rewrite_formula_references, ReferenceAxis, ReferenceEdit};

const MAX_STRUCTURE_COUNT: u32 = 1_024;
const FORMULA_METADATA_KEY: &str = "formulaMetadata";
const PIVOT_TABLES_KEY: &str = "pivotTables";

struct AxisEdit {
    sheet_id: String,
    at: u32,
    count: u32,
    axis: ReferenceAxis,
    insert: bool,
}

pub(super) fn is_structure_mutation(mutation: &NativeOfficeCollaborationMutation) -> bool {
    decode(mutation).is_some()
}

pub(super) fn validate_structure_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let edit = decode(mutation).ok_or_else(|| {
        invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet structure mutation.",
        )
    })?;
    validate_sheet_id(&edit.sheet_id)?;
    validate_span(edit.at, edit.count, edit.axis)
}

pub(super) fn apply_structure_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    let edit = decode(mutation).ok_or_else(|| {
        invalid_spreadsheet_mutation(
            "The supplied mutation is not a Spreadsheet structure mutation.",
        )
    })?;
    let reference_edit = reference_edit(edit.at, edit.count, edit.insert)?;
    let sheet_ids = ordered_sheet_ids(doc, manifest)?;
    if !sheet_ids.iter().any(|id| id == &edit.sheet_id) {
        return Err(super::spreadsheet_match_conflict(format!(
            "Spreadsheet sheet ID '{}' does not exist.",
            edit.sheet_id
        )));
    }

    let mut loaded = Vec::with_capacity(sheet_ids.len());
    for sheet_id in sheet_ids {
        let state = read_sheet_state(doc, manifest, &sheet_id)?;
        let name = sheet_display_name(doc, &state, &sheet_id)?;
        loaded.push((sheet_id, name, state));
    }
    let target_name = loaded
        .iter()
        .find(|(id, _, _)| id == &edit.sheet_id)
        .map(|(_, name, _)| name.clone())
        .ok_or_else(|| {
            super::spreadsheet_match_conflict(format!(
                "Spreadsheet sheet ID '{}' does not exist.",
                edit.sheet_id
            ))
        })?;
    let names = loaded
        .iter()
        .map(|(id, name, _)| (id.clone(), name.clone()))
        .collect::<HashMap<_, _>>();

    if collection_len(doc, &loaded, &edit.sheet_id, PIVOT_TABLES_KEY)? > 0 {
        return Err(invalid_spreadsheet_mutation(
            "Row and column structural edits on worksheets containing pivot tables are not yet lossless.",
        ));
    }

    let references = rewrite_sidecar_references(
        doc,
        manifest,
        &loaded,
        &names,
        target_name.as_str(),
        edit.axis,
        reference_edit,
    )?;
    let addresses = super::addresses::rewrite_region_addresses(
        doc,
        &loaded,
        edit.sheet_id.as_str(),
        target_name.as_str(),
        edit.axis,
        reference_edit,
        edit.at,
        edit.count,
        edit.insert,
    )?;

    let mut commits = Vec::with_capacity(loaded.len());
    let mut text_rekeys = Vec::new();
    for (sheet_id, name, state) in loaded {
        let target = sheet_id == edit.sheet_id;
        let (commit, rekeys) = plan_sheet(
            state,
            target,
            name.as_str(),
            target_name.as_str(),
            edit.axis,
            reference_edit,
            edit.at,
            edit.count,
            edit.insert,
        )?;
        text_rekeys.push((commit.state.record.clone(), rekeys));
        commits.push(commit);
    }

    let text_moves = read_text_moves(doc, &text_rekeys)?;
    let dirty = commits
        .iter()
        .any(|commit| !commit.changes.is_empty() || commit.next_row_lengths.is_some())
        || !text_moves.is_empty()
        || !references.is_empty()
        || !addresses.writes.is_empty()
        || !addresses.removals.is_empty()
        || !addresses.array_removals.is_empty();
    if dirty {
        let mut references = references;
        references.extend(addresses.writes);
        commit_sheet_structure(
            doc,
            commits,
            text_moves,
            references,
            addresses.removals,
            addresses.array_removals,
        )?;
    }
    Ok(())
}

fn plan_sheet(
    state: SpreadsheetSheetState,
    target: bool,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
    at: u32,
    count: u32,
    insert: bool,
) -> UseResult<(SheetStructureCommit, Vec<(String, Option<String>)>)> {
    let empty = JsonValue::Object(Default::default());
    let mut removes = Vec::new();
    let mut keeps = Vec::new();
    let mut inserts = Vec::new();
    let mut rekeys = Vec::new();
    let mut surviving = Vec::new();

    for ((row, column), cell) in &state.cells {
        let mapped = if target {
            map_coordinate(*row, *column, axis, at, count, insert)?
        } else {
            Some((*row, *column))
        };
        let Some((next_row, next_column)) = mapped else {
            removes.push(cell_write(
                *row,
                *column,
                cell_field_patches(cell, cell, &empty)?,
                SpreadsheetCellPresenceChange::Remove,
            ));
            rekeys.push((encode_coordinate(*row, *column), None));
            continue;
        };
        let rewritten = rewrite_cell_formula(cell, current_sheet, target_sheet, axis, edit)?;
        if (next_row, next_column) == (*row, *column) {
            if rewritten != *cell {
                keeps.push(cell_write(
                    *row,
                    *column,
                    cell_field_patches(cell, cell, &rewritten)?,
                    SpreadsheetCellPresenceChange::Keep,
                ));
            }
        } else {
            removes.push(cell_write(
                *row,
                *column,
                cell_field_patches(cell, cell, &empty)?,
                SpreadsheetCellPresenceChange::Remove,
            ));
            inserts.push(cell_write(
                next_row,
                next_column,
                cell_field_patches(&empty, &empty, &rewritten)?,
                SpreadsheetCellPresenceChange::Insert,
            ));
            rekeys.push((
                encode_coordinate(*row, *column),
                Some(encode_coordinate(next_row, next_column)),
            ));
        }
        surviving.push((next_row, next_column));
    }

    let next_row_lengths =
        if target && matches!(state.mode, Some(super::state::SpreadsheetCellMode::Data)) {
            let shifted = shift_row_lengths(&state.row_lengths, at, count, axis, insert)?;
            let covered = extended_dense_row_lengths(&shifted, surviving)?;
            if covered == state.row_lengths {
                None
            } else {
                Some(covered)
            }
        } else {
            None
        };
    let mut changes = removes;
    changes.extend(keeps);
    changes.extend(inserts);
    Ok((
        SheetStructureCommit {
            state,
            changes,
            next_row_lengths,
        },
        rekeys,
    ))
}

fn rewrite_cell_formula(
    cell: &JsonValue,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<JsonValue> {
    let Some(formula) = cell.get("f") else {
        return Ok(cell.clone());
    };
    let JsonValue::String(formula) = formula else {
        return Err(invalid_spreadsheet_mutation(
            "A collaborative Spreadsheet formula must be a string.",
        ));
    };
    let rewritten =
        rewrite_formula_references(formula, Some(current_sheet), target_sheet, axis, edit)?;
    if rewritten == *formula {
        return Ok(cell.clone());
    }
    let mut next = cell.clone();
    next.as_object_mut()
        .expect("spreadsheet cell is an object")
        .insert("f".to_owned(), JsonValue::String(rewritten));
    Ok(next)
}

fn rewrite_sidecar_references(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    loaded: &[(String, String, SpreadsheetSheetState)],
    names: &HashMap<String, String>,
    target_sheet_name: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Vec<SpreadsheetReferenceWrite>> {
    let named_root = format!("{}.spreadsheet.named-ranges", manifest.namespace);
    let area_root = format!("{}.spreadsheet.print-areas", manifest.namespace);
    let (has_named, has_areas) = {
        let transaction = doc.transact();
        let mut has_named = false;
        let mut has_areas = false;
        for (name, _) in transaction.root_refs() {
            if name == named_root {
                has_named = true;
            } else if name == area_root {
                has_areas = true;
            }
        }
        (has_named, has_areas)
    };
    let named = has_named.then(|| doc.get_or_insert_map(named_root));
    let areas = has_areas.then(|| doc.get_or_insert_map(area_root));
    let mut writes = Vec::new();
    let transaction = doc.transact();
    for (_, sheet_name, state) in loaded {
        if let Some(Out::YMap(metadata)) = state.record.get(&transaction, FORMULA_METADATA_KEY) {
            writes.extend(rewrite_metadata_fields(
                &metadata,
                &transaction,
                sheet_name,
                target_sheet_name,
                axis,
                edit,
            )?);
        }
    }
    if let Some(named) = named {
        for (_, child) in named.iter(&transaction) {
            let Out::YMap(record) = child else {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet named range is not a typed map.",
                ));
            };
            let scope = flat_string(&record, &transaction, "scopeSheetId")?;
            let current = match scope.as_deref() {
                Some(scope_id) => names
                    .get(scope_id)
                    .cloned()
                    .unwrap_or_else(|| scope_id.to_owned()),
                None => String::new(),
            };
            push_reference_rewrite(
                &mut writes,
                &record,
                &transaction,
                current.as_str(),
                target_sheet_name,
                axis,
                edit,
            )?;
        }
    }
    if let Some(areas) = areas {
        for (sheet_id, child) in areas.iter(&transaction) {
            let Out::YMap(record) = child else {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet print area is not a typed map.",
                ));
            };
            let sheet_key = sheet_id.to_owned();
            let current = names
                .get(&sheet_key)
                .cloned()
                .unwrap_or_else(|| sheet_key.clone());
            push_reference_rewrite(
                &mut writes,
                &record,
                &transaction,
                current.as_str(),
                target_sheet_name,
                axis,
                edit,
            )?;
        }
    }
    Ok(writes)
}

fn rewrite_metadata_fields<T: ReadTxn>(
    map: &yrs::MapRef,
    transaction: &T,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Vec<SpreadsheetReferenceWrite>> {
    let mut writes = Vec::new();
    for (key, value) in map.iter(transaction) {
        let (kind, path) = decode_flat_json_key(&key)?;
        if kind != FlatJsonEntryKind::Value {
            continue;
        }
        let rewrite = path
            .last()
            .is_some_and(|part| part == "reference" || part == "formula" || part == "f")
            || path.first().is_some_and(|part| part == "sourceFormulas");
        if !rewrite {
            continue;
        }
        let Out::Any(yrs::Any::String(text)) = value else {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet formula metadata reference is not text.",
            ));
        };
        if let Some(next) = changed_reference(&text, current_sheet, target_sheet, axis, edit)? {
            writes.push(SpreadsheetReferenceWrite {
                map: map.clone(),
                key: key.to_owned(),
                value: json_to_any(&JsonValue::String(next), 0)?,
            });
        }
    }
    Ok(writes)
}

fn push_reference_rewrite<T: ReadTxn>(
    writes: &mut Vec<SpreadsheetReferenceWrite>,
    map: &yrs::MapRef,
    transaction: &T,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<()> {
    let Some(reference) = flat_string(map, transaction, "reference")? else {
        return Ok(());
    };
    let Some(next) = changed_reference(&reference, current_sheet, target_sheet, axis, edit)? else {
        return Ok(());
    };
    let key = encode_flat_json_key(FlatJsonEntryKind::Value, &["reference".to_owned()])?;
    writes.push(SpreadsheetReferenceWrite {
        map: map.clone(),
        key,
        value: json_to_any(&JsonValue::String(next), 0)?,
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
        Some(Out::Any(yrs::Any::String(value))) => Ok(Some(value.to_string())),
        None => Ok(None),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet field '{field}' is not text."
        ))),
    }
}

fn changed_reference(
    formula: &str,
    current_sheet: &str,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Option<String>> {
    let current = if current_sheet.is_empty() {
        None
    } else {
        Some(current_sheet)
    };
    let rewritten = rewrite_formula_references(formula, current, target_sheet, axis, edit)?;
    if rewritten == formula {
        Ok(None)
    } else {
        Ok(Some(rewritten))
    }
}

fn read_text_moves(
    doc: &yrs::Doc,
    rekeys: &[(yrs::MapRef, Vec<(String, Option<String>)>)],
) -> UseResult<Vec<SpreadsheetTextMove>> {
    let transaction = doc.transact();
    let mut moves = Vec::new();
    for (record, rekeys) in rekeys {
        let Some(Out::YMap(texts)) = record.get(&transaction, CELL_TEXT_KEY) else {
            continue;
        };
        for (from, to) in rekeys {
            match texts.get(&transaction, from.as_str()) {
                Some(Out::YText(text)) => moves.push(SpreadsheetTextMove {
                    texts: texts.clone(),
                    from: from.clone(),
                    to: to.clone(),
                    content: text.get_string(&transaction),
                }),
                None => {}
                Some(_) => {
                    return Err(invalid_shared_spreadsheet(
                        "The spreadsheet cell text is not a collaborative text node.",
                    ));
                }
            }
        }
    }
    Ok(moves)
}

fn sheet_display_name(
    doc: &yrs::Doc,
    state: &SpreadsheetSheetState,
    sheet_id: &str,
) -> UseResult<String> {
    let transaction = doc.transact();
    match state.record.get(&transaction, "name") {
        Some(Out::Any(yrs::Any::String(value))) if !value.is_empty() => Ok(value.to_string()),
        None | Some(Out::Any(yrs::Any::String(_))) => Ok(sheet_id.to_owned()),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet sheet '{sheet_id}' has a name that is not text."
        ))),
    }
}

fn collection_len(
    doc: &yrs::Doc,
    loaded: &[(String, String, SpreadsheetSheetState)],
    sheet_id: &str,
    key: &str,
) -> UseResult<u32> {
    let Some((_, _, state)) = loaded.iter().find(|(id, _, _)| id == sheet_id) else {
        return Ok(0);
    };
    let transaction = doc.transact();
    match state.record.get(&transaction, key) {
        Some(Out::YMap(map)) => Ok(map.len(&transaction)),
        None => Ok(0),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet '{key}' is not a typed map."
        ))),
    }
}

fn map_coordinate(
    row: u32,
    column: u32,
    axis: ReferenceAxis,
    at: u32,
    count: u32,
    insert: bool,
) -> UseResult<Option<(u32, u32)>> {
    match axis {
        ReferenceAxis::Row => map_index(row, at, count, insert, MAX_SPREADSHEET_ROWS)
            .map(|row| row.map(|row| (row, column))),
        ReferenceAxis::Column => map_index(column, at, count, insert, MAX_SPREADSHEET_COLUMNS)
            .map(|column| column.map(|column| (row, column))),
    }
}

pub(super) fn map_index(
    index: u32,
    at: u32,
    count: u32,
    insert: bool,
    limit: u32,
) -> UseResult<Option<u32>> {
    if insert {
        if index < at {
            return Ok(Some(index));
        }
        let next = index.checked_add(count).ok_or_else(structure_too_large)?;
        if next >= limit {
            return Err(structure_too_large());
        }
        return Ok(Some(next));
    }
    let end = at.checked_add(count).ok_or_else(structure_too_large)?;
    if index < at {
        Ok(Some(index))
    } else if index < end {
        Ok(None)
    } else {
        Ok(Some(index - count))
    }
}

fn shift_row_lengths(
    lengths: &[u32],
    at: u32,
    count: u32,
    axis: ReferenceAxis,
    insert: bool,
) -> UseResult<Vec<u32>> {
    match axis {
        ReferenceAxis::Column => lengths
            .iter()
            .map(|length| adjust_width(*length, at, count, insert))
            .collect(),
        ReferenceAxis::Row if insert => {
            let mut rows = lengths.to_vec();
            let index = at as usize;
            if index > rows.len() {
                if index > MAX_SPREADSHEET_ROWS as usize {
                    return Err(structure_too_large());
                }
                rows.resize(index, 0);
            }
            rows.splice(index..index, std::iter::repeat(0).take(count as usize));
            if rows.len() > MAX_SPREADSHEET_ROWS as usize {
                return Err(structure_too_large());
            }
            Ok(rows)
        }
        ReferenceAxis::Row => {
            let mut rows = lengths.to_vec();
            let start = at as usize;
            if start < rows.len() {
                let end = start.saturating_add(count as usize).min(rows.len());
                rows.drain(start..end);
            }
            Ok(rows)
        }
    }
}

fn adjust_width(length: u32, at: u32, count: u32, insert: bool) -> UseResult<u32> {
    if insert {
        if length <= at {
            return Ok(length);
        }
        let next = length.checked_add(count).ok_or_else(structure_too_large)?;
        if next > MAX_SPREADSHEET_COLUMNS {
            return Err(structure_too_large());
        }
        return Ok(next);
    }
    let end = at.saturating_add(count);
    if length <= at {
        Ok(length)
    } else if length <= end {
        Ok(at)
    } else {
        Ok(length - count)
    }
}

fn cell_write(
    row: u32,
    column: u32,
    patches: Vec<super::json::FlatJsonPatch>,
    presence: SpreadsheetCellPresenceChange,
) -> SpreadsheetCellWrite {
    SpreadsheetCellWrite {
        row,
        column,
        patches,
        presence,
    }
}

fn decode(mutation: &NativeOfficeCollaborationMutation) -> Option<AxisEdit> {
    let (sheet_id, at, count, axis, insert) = match mutation {
        NativeOfficeCollaborationMutation::SpreadsheetInsertRows {
            sheet_id,
            at,
            count,
        } => (sheet_id, *at, *count, ReferenceAxis::Row, true),
        NativeOfficeCollaborationMutation::SpreadsheetDeleteRows {
            sheet_id,
            at,
            count,
        } => (sheet_id, *at, *count, ReferenceAxis::Row, false),
        NativeOfficeCollaborationMutation::SpreadsheetInsertColumns {
            sheet_id,
            at,
            count,
        } => (sheet_id, *at, *count, ReferenceAxis::Column, true),
        NativeOfficeCollaborationMutation::SpreadsheetDeleteColumns {
            sheet_id,
            at,
            count,
        } => (sheet_id, *at, *count, ReferenceAxis::Column, false),
        _ => return None,
    };
    Some(AxisEdit {
        sheet_id: sheet_id.clone(),
        at,
        count,
        axis,
        insert,
    })
}

fn reference_edit(at: u32, count: u32, insert: bool) -> UseResult<ReferenceEdit> {
    let excel = at.checked_add(1).ok_or_else(structure_too_large)?;
    Ok(if insert {
        ReferenceEdit::Insert { at: excel, count }
    } else {
        ReferenceEdit::Delete {
            start: excel,
            count,
        }
    })
}

fn validate_span(at: u32, count: u32, axis: ReferenceAxis) -> UseResult<()> {
    let limit = match axis {
        ReferenceAxis::Row => MAX_SPREADSHEET_ROWS,
        ReferenceAxis::Column => MAX_SPREADSHEET_COLUMNS,
    };
    if count == 0 || count > MAX_STRUCTURE_COUNT {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet structure edit must move 1 to {MAX_STRUCTURE_COUNT} rows or columns."
        ))
        .with_detail("count", u64::from(count)));
    }
    if at >= limit {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet structure edit starts outside the sheet.",
        )
        .with_detail("at", u64::from(at)));
    }
    let end = at.checked_add(count).ok_or_else(structure_too_large)?;
    if end > limit {
        return Err(structure_too_large());
    }
    Ok(())
}

pub(super) fn structure_too_large() -> a3s_use_core::UseError {
    super::super::super::collaboration_error(
        "office.collaboration.mutation_too_large",
        "The Spreadsheet structure edit exceeds the sheet bounds.",
    )
}
