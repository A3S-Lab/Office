use a3s_office::{
    DocumentNode, NativeSpreadsheetTable, NativeSpreadsheetTableColumn, NativeSpreadsheetTableStyle,
};
use a3s_use_core::UseResult;

use super::arguments::ParsedArguments;
use super::spreadsheet_filter;
use super::usage_error;

pub(super) fn build_new(parsed: &ParsedArguments) -> UseResult<NativeSpreadsheetTable> {
    let name = parsed
        .name
        .as_deref()
        .ok_or_else(|| usage_error("native Spreadsheet table add requires --name <name>"))?;
    let range = one_range(parsed, true)?
        .ok_or_else(|| usage_error("native Spreadsheet table add requires --range <A1-range>"))?;
    if parsed.table_columns.is_empty() {
        return Err(usage_error(
            "native Spreadsheet table add requires one --table-column <name> per range column",
        ));
    }
    let mut table = NativeSpreadsheetTable::new(name, range, parsed.table_columns.clone());
    table.filters = spreadsheet_filter::new_columns(parsed)?;
    if let Some(display_name) = &parsed.table_display_name {
        table.display_name = Some(display_name.clone());
    }
    apply_updates(&mut table, parsed)?;
    normalize_none_style(&mut table, parsed);
    Ok(table)
}

pub(super) fn merge_existing(
    node: &DocumentNode,
    parsed: &ParsedArguments,
) -> UseResult<NativeSpreadsheetTable> {
    let mut table = NativeSpreadsheetTable::from_semantic_node(node)?;
    if let Some(name) = &parsed.name {
        table.name = name.clone();
    }
    if let Some(display_name) = &parsed.table_display_name {
        table.display_name = if display_name.eq_ignore_ascii_case("none") {
            None
        } else {
            Some(display_name.clone())
        };
    }
    if let Some(range) = one_range(parsed, false)? {
        table.range = range;
    }
    if !parsed.table_columns.is_empty() {
        table.columns = parsed
            .table_columns
            .iter()
            .cloned()
            .map(NativeSpreadsheetTableColumn::new)
            .collect();
    }
    spreadsheet_filter::update_existing_columns(&mut table.filters, parsed)?;
    apply_updates(&mut table, parsed)?;
    normalize_none_style(&mut table, parsed);
    Ok(table)
}

fn apply_updates(table: &mut NativeSpreadsheetTable, parsed: &ParsedArguments) -> UseResult<()> {
    if let Some(value) = &parsed.table_header_row {
        table.header_row = boolean("--header-row", value)?;
    }
    if let Some(value) = &parsed.table_totals_row {
        table.totals_row = boolean("--totals-row", value)?;
    }
    if let Some(value) = &parsed.table_style {
        table.style = parse_style(value)?;
    }
    if let Some(value) = &parsed.table_show_first_column {
        table.show_first_column = boolean("--show-first-column", value)?;
    }
    if let Some(value) = &parsed.table_show_last_column {
        table.show_last_column = boolean("--show-last-column", value)?;
    }
    if let Some(value) = &parsed.table_show_row_stripes {
        table.show_row_stripes = boolean("--show-row-stripes", value)?;
    }
    if let Some(value) = &parsed.table_show_column_stripes {
        table.show_column_stripes = boolean("--show-column-stripes", value)?;
    }
    Ok(())
}

fn normalize_none_style(table: &mut NativeSpreadsheetTable, parsed: &ParsedArguments) {
    if table.style == NativeSpreadsheetTableStyle::None {
        if parsed.table_show_first_column.is_none() {
            table.show_first_column = false;
        }
        if parsed.table_show_last_column.is_none() {
            table.show_last_column = false;
        }
        if parsed.table_show_row_stripes.is_none() {
            table.show_row_stripes = false;
        }
        if parsed.table_show_column_stripes.is_none() {
            table.show_column_stripes = false;
        }
    }
}

fn one_range(parsed: &ParsedArguments, required: bool) -> UseResult<Option<String>> {
    if parsed.validation_ranges.len() > 1 {
        return Err(usage_error(
            "native Spreadsheet tables accept exactly one --range value",
        ));
    }
    let range = parsed.validation_ranges.first().cloned();
    if required && range.is_none() {
        return Err(usage_error(
            "native Spreadsheet table add requires --range <A1-range>",
        ));
    }
    Ok(range)
}

fn parse_style(value: &str) -> UseResult<NativeSpreadsheetTableStyle> {
    let normalized = value.to_ascii_lowercase();
    if normalized == "none" {
        return Ok(NativeSpreadsheetTableStyle::None);
    }
    let (family, number) = normalized.split_once(':').ok_or_else(|| {
        usage_error("--style requires none, light:<1-21>, medium:<1-28>, or dark:<1-11>")
    })?;
    let number = number.parse::<u8>().map_err(|_| {
        usage_error("--style requires none, light:<1-21>, medium:<1-28>, or dark:<1-11>")
    })?;
    let style = match family {
        "light" => NativeSpreadsheetTableStyle::Light { number },
        "medium" => NativeSpreadsheetTableStyle::Medium { number },
        "dark" => NativeSpreadsheetTableStyle::Dark { number },
        _ => {
            return Err(usage_error(
                "--style requires none, light:<1-21>, medium:<1-28>, or dark:<1-11>",
            ))
        }
    };
    Ok(style)
}

fn boolean(option: &str, value: &str) -> UseResult<bool> {
    match value.to_ascii_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(usage_error(format!(
            "{option} requires true or false, received '{value}'"
        ))),
    }
}
