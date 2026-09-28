//! A deleted table with a built-in style leaves that style on the cells that
//! already exist. Empty positions stay empty. `family: none` has nothing to
//! paint. A range larger than the browser conversion limit writes nothing.

use std::sync::Arc;

use a3s_use_core::UseResult;
use serde_json::{json, Value as JsonValue};
use yrs::{Any, Map, Out, ReadTxn, Transact};

use super::json::{encode_flat_json_key, json_to_any, FlatJsonEntryKind};
use super::state::{encode_cell_field_key, read_sheet_state, SpreadsheetReferenceWrite};
use super::table::TablePlan;
use super::{
    invalid_shared_spreadsheet, invalid_spreadsheet_mutation, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationMutation,
};

const MAX_TABLE_CONVERSION_CELLS: u64 = 100_000;
const CONFIG_KEY: &str = "config";

const LIGHT_ACCENTS: [&str; 7] = [
    "#64748b", "#2563eb", "#0f766e", "#7c3aed", "#c2410c", "#be123c", "#15803d",
];
const DARK_ACCENTS: [&str; 11] = [
    "#1e293b", "#1e3a8a", "#134e4a", "#4c1d95", "#7c2d12", "#881337", "#14532d", "#0f172a",
    "#312e81", "#3f3f46", "#422006",
];

pub(super) struct BorderWrite {
    pub(super) key: String,
    pub(super) value: Any,
}

pub(super) struct TableConversion {
    pub(super) cells: Vec<SpreadsheetReferenceWrite>,
    pub(super) border: Option<BorderWrite>,
}

pub(super) fn prepare_table_frame(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
    sheet_id: &str,
    sheet_name: &str,
    plan: &TablePlan,
) -> UseResult<(
    Option<JsonValue>,
    Vec<SpreadsheetReferenceWrite>,
    Option<BorderWrite>,
)> {
    match plan {
        TablePlan::Create { .. } | TablePlan::Unchanged => Ok((None, Vec::new(), None)),
        TablePlan::Update { id, .. } => {
            let NativeOfficeCollaborationMutation::SpreadsheetUpdateTable {
                expected_table,
                next_table,
                ..
            } = mutation
            else {
                return Err(invalid_spreadsheet_mutation(
                    "The supplied mutation is not a Spreadsheet table frame.",
                ));
            };
            let (next, writes) = super::structured_reference::rewrite_table_update(
                doc,
                manifest,
                sheet_id,
                sheet_name,
                id,
                expected_table,
                next_table,
            )?;
            Ok((next, writes, None))
        }
        TablePlan::Delete { id } => {
            let NativeOfficeCollaborationMutation::SpreadsheetDeleteTable {
                expected_table, ..
            } = mutation
            else {
                return Err(invalid_spreadsheet_mutation(
                    "The supplied mutation is not a Spreadsheet table frame.",
                ));
            };
            super::structured_reference::refuse_referenced_table(
                doc,
                manifest,
                sheet_id,
                sheet_name,
                id,
                expected_table,
            )?;
            let conversion = conversion_writes(doc, manifest, sheet_id, expected_table)?;
            Ok((None, conversion.cells, conversion.border))
        }
    }
}

pub(super) fn sheet_display_name<T: ReadTxn>(
    record: &yrs::MapRef,
    transaction: &T,
    sheet_id: &str,
) -> UseResult<String> {
    match record.get(transaction, "name") {
        Some(Out::Any(Any::String(value))) if !value.is_empty() => Ok(value.to_string()),
        None | Some(Out::Any(Any::String(_))) => Ok(sheet_id.to_owned()),
        Some(_) => Err(invalid_shared_spreadsheet(format!(
            "Shared Spreadsheet sheet '{sheet_id}' has a name that is not text."
        ))),
    }
}

pub(super) fn conversion_writes(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    sheet_id: &str,
    table: &JsonValue,
) -> UseResult<TableConversion> {
    let Some(style) = paint_style(table)? else {
        return Ok(TableConversion {
            cells: Vec::new(),
            border: None,
        });
    };
    let span = span_of(table)?;
    let area = u64::from(span.rows) * u64::from(span.columns);
    if area > MAX_TABLE_CONVERSION_CELLS {
        return Err(invalid_spreadsheet_mutation(
            "Spreadsheet table appearance is too large to materialize.",
        ));
    }
    let state = read_sheet_state(doc, manifest, sheet_id)?;
    let mut cells = Vec::new();
    for ((row, column), _) in &state.cells {
        if !span.contains(*row, *column) {
            continue;
        }
        let paint = cell_paint(&style, &span, *row, *column);
        push_cell_field(
            &state.fields,
            *row,
            *column,
            "bg",
            &paint.background,
            &mut cells,
        )?;
        push_cell_field(&state.fields, *row, *column, "fc", &paint.text, &mut cells)?;
        if paint.bold {
            let key = cell_key(*row, *column, "bl")?;
            cells.push(SpreadsheetReferenceWrite {
                map: state.fields.clone(),
                key,
                value: json_to_any(&json!(1), 0)?,
            });
        }
    }
    let border = border_write(doc, &state.record, &style, &span)?;
    Ok(TableConversion {
        cells,
        border: Some(border),
    })
}

struct Span {
    row_start: u32,
    row_end: u32,
    column_start: u32,
    column_end: u32,
    rows: u32,
    columns: u32,
    header: bool,
    totals: bool,
    row_stripes: bool,
    column_stripes: bool,
    first_column: bool,
    last_column: bool,
}

struct PaintStyle {
    border: String,
    header: String,
    header_text: String,
    primary_row: String,
    secondary_row: String,
    stripe_column: String,
    text: String,
    total: String,
    total_text: String,
}

struct CellPaint {
    background: String,
    text: String,
    bold: bool,
}

fn paint_style(table: &JsonValue) -> UseResult<Option<PaintStyle>> {
    let style = table
        .get("style")
        .and_then(JsonValue::as_object)
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table is missing its style.")
        })?;
    let family = style
        .get("family")
        .and_then(JsonValue::as_str)
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table style is missing its family.")
        })?;
    if family == "none" {
        return Ok(None);
    }
    let number = style
        .get("number")
        .and_then(JsonValue::as_u64)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table style number is invalid.")
        })?;
    let palette = match family {
        "light" => light_palette(number),
        "medium" => medium_palette(number),
        "dark" => dark_palette(number),
        _ => {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet table style is invalid.",
            ))
        }
    };
    Ok(Some(palette))
}

fn cell_paint(style: &PaintStyle, span: &Span, row: u32, column: u32) -> CellPaint {
    let header = span.header && row == span.row_start;
    let totals = span.totals && row == span.row_end;
    let data_start = span.row_start + u32::from(span.header);
    let data_index = row.saturating_sub(data_start);
    let secondary = span.row_stripes && data_index % 2 == 1;
    let mut background = if header {
        style.header.clone()
    } else if totals {
        style.total.clone()
    } else if secondary {
        style.secondary_row.clone()
    } else {
        style.primary_row.clone()
    };
    if !header && !totals && span.column_stripes && (column - span.column_start) % 2 == 1 {
        background = mix_hex(&background, &style.stripe_column, 0.42);
    }
    let text = if header {
        style.header_text.clone()
    } else if totals {
        style.total_text.clone()
    } else {
        style.text.clone()
    };
    let bold = header
        || totals
        || (span.first_column && column == span.column_start)
        || (span.last_column && column == span.column_end);
    CellPaint {
        background,
        text,
        bold,
    }
}

fn border_write(
    doc: &yrs::Doc,
    record: &yrs::MapRef,
    style: &PaintStyle,
    span: &Span,
) -> UseResult<BorderWrite> {
    let key = encode_flat_json_key(FlatJsonEntryKind::Value, &["borderInfo".to_owned()])?;
    let transaction = doc.transact();
    let mut items = match record.get(&transaction, CONFIG_KEY) {
        None => Vec::new(),
        Some(Out::YMap(config)) => match config.get(&transaction, key.as_str()) {
            None => Vec::new(),
            Some(Out::Any(Any::Array(items))) => items.to_vec(),
            Some(_) => {
                return Err(invalid_shared_spreadsheet(
                    "Shared Spreadsheet border information is not an array.",
                ))
            }
        },
        Some(_) => {
            return Err(invalid_shared_spreadsheet(
                "Shared Spreadsheet 'config' is not a typed map.",
            ))
        }
    };
    items.push(json_to_any(
        &json!({
            "rangeType": "range",
            "borderType": "border-all",
            "color": style.border,
            "style": "1",
            "range": [{
                "row": [span.row_start, span.row_end],
                "column": [span.column_start, span.column_end],
            }],
        }),
        0,
    )?);
    Ok(BorderWrite {
        key,
        value: Any::Array(Arc::from(items)),
    })
}

fn push_cell_field(
    map: &yrs::MapRef,
    row: u32,
    column: u32,
    field: &str,
    text: &str,
    writes: &mut Vec<SpreadsheetReferenceWrite>,
) -> UseResult<()> {
    writes.push(SpreadsheetReferenceWrite {
        map: map.clone(),
        key: cell_key(row, column, field)?,
        value: json_to_any(&JsonValue::String(text.to_owned()), 0)?,
    });
    Ok(())
}

fn cell_key(row: u32, column: u32, field: &str) -> UseResult<String> {
    let flat = encode_flat_json_key(FlatJsonEntryKind::Value, &[field.to_owned()])?;
    encode_cell_field_key(row, column, &flat)
}

fn span_of(table: &JsonValue) -> UseResult<Span> {
    let range = table.get("range").ok_or_else(|| {
        invalid_shared_spreadsheet("A shared Spreadsheet table is missing its range.")
    })?;
    let (row_start, row_end) = pair(range, "row")?;
    let (column_start, column_end) = pair(range, "column")?;
    Ok(Span {
        row_start,
        row_end,
        column_start,
        column_end,
        rows: row_end - row_start + 1,
        columns: column_end - column_start + 1,
        header: flag(table, "headerRow")?,
        totals: flag(table, "totalsRow")?,
        row_stripes: flag(table, "showRowStripes")?,
        column_stripes: flag(table, "showColumnStripes")?,
        first_column: flag(table, "showFirstColumn")?,
        last_column: flag(table, "showLastColumn")?,
    })
}

impl Span {
    fn contains(&self, row: u32, column: u32) -> bool {
        (self.row_start..=self.row_end).contains(&row)
            && (self.column_start..=self.column_end).contains(&column)
    }
}

fn light_palette(number: u32) -> PaintStyle {
    let accent = LIGHT_ACCENTS[((number - 1) as usize) % LIGHT_ACCENTS.len()];
    let variant = (number - 1) / LIGHT_ACCENTS.len() as u32;
    let header_weight = [0.84, 0.73, 0.62]
        .get(variant as usize)
        .copied()
        .unwrap_or(0.72);
    PaintStyle {
        border: mix_hex(accent, "#ffffff", 0.5),
        header: mix_hex(accent, "#ffffff", header_weight),
        header_text: mix_hex(accent, "#0f172a", 0.38),
        primary_row: "#ffffff".to_owned(),
        secondary_row: mix_hex(accent, "#ffffff", 0.9 - f64::from(variant) * 0.025),
        stripe_column: mix_hex(accent, "#ffffff", 0.82),
        text: "#1f2937".to_owned(),
        total: mix_hex(accent, "#ffffff", 0.78),
        total_text: "#172033".to_owned(),
    }
}

fn medium_palette(number: u32) -> PaintStyle {
    let accent = LIGHT_ACCENTS[((number - 1) as usize) % LIGHT_ACCENTS.len()];
    let variant = (number - 1) / LIGHT_ACCENTS.len() as u32;
    let header = mix_hex(accent, "#111827", f64::from(variant) * 0.08);
    PaintStyle {
        border: mix_hex(accent, "#ffffff", 0.34),
        header,
        header_text: "#ffffff".to_owned(),
        primary_row: "#ffffff".to_owned(),
        secondary_row: mix_hex(accent, "#ffffff", 0.86 - f64::from(variant) * 0.02),
        stripe_column: mix_hex(accent, "#ffffff", 0.76),
        text: "#172033".to_owned(),
        total: mix_hex(accent, "#ffffff", 0.72),
        total_text: mix_hex(accent, "#111827", 0.48),
    }
}

fn dark_palette(number: u32) -> PaintStyle {
    let accent = DARK_ACCENTS
        .get((number as usize).saturating_sub(1))
        .copied()
        .unwrap_or(DARK_ACCENTS[0]);
    PaintStyle {
        border: mix_hex(accent, "#ffffff", 0.32),
        header: accent.to_owned(),
        header_text: "#ffffff".to_owned(),
        primary_row: mix_hex(accent, "#ffffff", 0.9),
        secondary_row: mix_hex(accent, "#ffffff", 0.78),
        stripe_column: mix_hex(accent, "#ffffff", 0.68),
        text: "#172033".to_owned(),
        total: mix_hex(accent, "#ffffff", 0.62),
        total_text: mix_hex(accent, "#111827", 0.34),
    }
}

fn mix_hex(left: &str, right: &str, right_weight: f64) -> String {
    let weight = right_weight.clamp(0.0, 1.0);
    let left = parse_hex(left);
    let right = parse_hex(right);
    let mut hex = String::from("#");
    for index in 0..3 {
        let value = (f64::from(left[index]) * (1.0 - weight) + f64::from(right[index]) * weight)
            .round() as u8;
        hex.push_str(&format!("{value:02x}"));
    }
    hex
}

fn parse_hex(value: &str) -> [u8; 3] {
    let hex = value.trim_start_matches('#');
    let mut channels = [0; 3];
    for (index, channel) in channels.iter_mut().enumerate() {
        let start = index * 2;
        *channel = u8::from_str_radix(&hex[start..start + 2], 16).unwrap_or(0);
    }
    channels
}

fn pair(range: &JsonValue, key: &str) -> UseResult<(u32, u32)> {
    let values = range
        .get(key)
        .and_then(JsonValue::as_array)
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table range is not a pair.")
        })?;
    if values.len() != 2 {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet table range is not a pair.",
        ));
    }
    Ok((index(&values[0])?, index(&values[1])?))
}

fn index(value: &JsonValue) -> UseResult<u32> {
    value
        .as_u64()
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| {
            invalid_shared_spreadsheet("A shared Spreadsheet table range is not an integer.")
        })
}

fn flag(table: &JsonValue, key: &str) -> UseResult<bool> {
    table.get(key).and_then(JsonValue::as_bool).ok_or_else(|| {
        invalid_shared_spreadsheet("A shared Spreadsheet table flag is not a boolean.")
    })
}

#[cfg(test)]
mod tests {
    use super::light_palette;

    #[test]
    fn light_one_palette_matches_the_browser_mix() {
        let palette = light_palette(1);
        assert_eq!(palette.border, "#b2bac5");
        assert_eq!(palette.header, "#e6e9ec");
        assert_eq!(palette.header_text, "#445166");
        assert_eq!(palette.primary_row, "#ffffff");
        assert_eq!(palette.secondary_row, "#f0f1f3");
        assert_eq!(palette.text, "#1f2937");
    }
}
