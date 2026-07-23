use a3s_office::{DocumentNode, NativeSpreadsheetAutoFilter, NativeSpreadsheetFilterColumn};
use a3s_use_core::UseResult;

use super::{usage_error, ParsedArguments};

pub(super) fn canonical_path(path: &str) -> Option<String> {
    let (sheet, segment) = path.rsplit_once('/')?;
    matches!(
        segment.to_ascii_lowercase().as_str(),
        "autofilter" | "auto-filter" | "filter"
    )
    .then(|| format!("{sheet}/autofilter"))
}

pub(super) fn build_new(parsed: &ParsedArguments) -> UseResult<NativeSpreadsheetAutoFilter> {
    let range = one_range(parsed, true)?
        .ok_or_else(|| usage_error("native Spreadsheet AutoFilter add requires --range"))?;
    Ok(NativeSpreadsheetAutoFilter {
        range,
        columns: new_columns(parsed)?,
    })
}

pub(super) fn merge_existing(
    node: &DocumentNode,
    parsed: &ParsedArguments,
) -> UseResult<NativeSpreadsheetAutoFilter> {
    let mut filter = NativeSpreadsheetAutoFilter::from_semantic_node(node).map_err(|error| {
        usage_error(format!(
            "Spreadsheet AutoFilter node is not editable: {error}"
        ))
    })?;
    if let Some(range) = one_range(parsed, false)? {
        filter.range = range;
    }
    update_existing_columns(&mut filter.columns, parsed)?;
    Ok(filter)
}

pub(super) fn new_columns(
    parsed: &ParsedArguments,
) -> UseResult<Vec<NativeSpreadsheetFilterColumn>> {
    if parsed.clear_filters {
        return Err(usage_error(
            "--clear-filters is available only when setting an existing Spreadsheet table or AutoFilter",
        ));
    }
    parse_columns(parsed)
}

pub(super) fn update_existing_columns(
    columns: &mut Vec<NativeSpreadsheetFilterColumn>,
    parsed: &ParsedArguments,
) -> UseResult<()> {
    if parsed.clear_filters && !parsed.spreadsheet_filters.is_empty() {
        return Err(usage_error(
            "--clear-filters cannot be combined with --filter",
        ));
    }
    if parsed.clear_filters {
        columns.clear();
    } else if !parsed.spreadsheet_filters.is_empty() {
        *columns = parse_columns(parsed)?;
    }
    Ok(())
}

fn parse_columns(parsed: &ParsedArguments) -> UseResult<Vec<NativeSpreadsheetFilterColumn>> {
    parsed
        .spreadsheet_filters
        .iter()
        .enumerate()
        .map(|(index, value)| {
            serde_json::from_str(value).map_err(|error| {
                usage_error(format!(
                    "--filter value {} must be a strict JSON filter-column object: {error}",
                    index + 1
                ))
            })
        })
        .collect()
}

fn one_range(parsed: &ParsedArguments, required: bool) -> UseResult<Option<String>> {
    if parsed.validation_ranges.len() > 1 {
        return Err(usage_error(
            "native Spreadsheet AutoFilters accept exactly one --range value",
        ));
    }
    let range = parsed.validation_ranges.first().cloned();
    if required && range.is_none() {
        return Err(usage_error(
            "native Spreadsheet AutoFilter add requires --range <A1-range>",
        ));
    }
    Ok(range)
}
