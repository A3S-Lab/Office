use a3s_use_core::UseResult;

use super::{node_not_found, sort_error, state};
use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::index_xml;
use crate::NativeOfficePackage;

pub(super) struct ResolvedSort {
    pub(super) sheet_path: String,
    pub(super) part_name: String,
    pub(super) range: CellRange,
}

pub(super) fn resolve(document: &NativeOfficeDocument, requested: &str) -> UseResult<ResolvedSort> {
    let (sheet_request, explicit_range) = requested
        .rsplit_once('/')
        .filter(|(sheet, range)| !sheet.is_empty() && range.contains(|c: char| c.is_ascii_digit()))
        .map_or((requested, None), |(sheet, range)| (sheet, Some(range)));
    if !sheet_request.starts_with('/') || sheet_request[1..].contains('/') {
        return Err(sort_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet sorting requires /SheetName or /SheetName/A1:D100.",
        ));
    }
    let sheet = document
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path.eq_ignore_ascii_case(sheet_request)
        })
        .ok_or_else(|| node_not_found(requested))?;
    let part_name = sheet.format.get("part").cloned().ok_or_else(|| {
        sort_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{}' has no source part.", sheet.path),
        )
    })?;
    let range = explicit_range.map_or_else(|| used_range(sheet), CellRange::parse)?;
    Ok(ResolvedSort {
        sheet_path: sheet.path.clone(),
        part_name,
        range,
    })
}

fn used_range(sheet: &DocumentNode) -> UseResult<CellRange> {
    let mut cells = sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Row)
        .flat_map(|row| row.children.iter())
        .filter(|node| node.node_type == OfficeNodeType::Cell)
        .filter_map(|cell| cell.path.rsplit_once('/').map(|(_, reference)| reference))
        .map(CellReference::parse)
        .collect::<UseResult<Vec<_>>>()?;
    if cells.is_empty() {
        return CellRange::parse("A1");
    }
    cells.sort_unstable();
    let start_column = cells.iter().map(|cell| cell.column).min().unwrap_or(1);
    let end_column = cells.iter().map(|cell| cell.column).max().unwrap_or(1);
    let start_row = cells.iter().map(|cell| cell.row).min().unwrap_or(1);
    let end_row = cells.iter().map(|cell| cell.row).max().unwrap_or(1);
    Ok(CellRange {
        start: CellReference {
            column: start_column,
            row: start_row,
        },
        end: CellReference {
            column: end_column,
            row: end_row,
        },
    })
}

pub(super) fn validate_geometry(
    package: &NativeOfficePackage,
    sheet: &DocumentNode,
    resolved: &ResolvedSort,
    header: bool,
) -> UseResult<CellRange> {
    let mut data_range = CellRange {
        start: CellReference {
            column: resolved.range.start.column,
            row: resolved.range.start.row + u32::from(header),
        },
        end: resolved.range.end,
    };
    if data_range.start.row > data_range.end.row {
        return Err(sort_error(
            "use.office.spreadsheet_sort_range_empty",
            format!(
                "Spreadsheet sort range '{}' contains no data row after excluding its header.",
                resolved.range.a1()
            ),
        ));
    }

    for table in sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Table)
    {
        let table_range = table
            .format
            .get("ref")
            .ok_or_else(|| {
                sort_error(
                    "use.office.spreadsheet_table_invalid",
                    "Spreadsheet table has no ref.",
                )
            })
            .and_then(|value| CellRange::parse(value))?;
        if !resolved.range.intersects(table_range) {
            continue;
        }
        if table.format.get("nativeMutable").map(String::as_str) != Some("true") {
            return Err(sort_error(
                "use.office.spreadsheet_sort_table_unsupported",
                format!(
                    "Spreadsheet sort range intersects non-mutable table '{}'.",
                    table.path
                ),
            ));
        }
        let table_header = format_bool(table, "headerRow")?;
        if format_bool(table, "totalsRow")? {
            return Err(sort_error(
                "use.office.spreadsheet_sort_table_totals_unsupported",
                "Native Spreadsheet sorting does not yet reorder a table with a totals row.",
            ));
        }
        let table_data = CellRange {
            start: CellReference {
                column: table_range.start.column,
                row: table_range.start.row + u32::from(table_header),
            },
            end: table_range.end,
        };
        let covers_table = (resolved.range == table_range && header == table_header)
            || (resolved.range == table_data && !header);
        if covers_table {
            data_range = table_data;
        } else {
            return Err(sort_error(
                "use.office.spreadsheet_sort_table_partial",
                format!(
                    "Sort range '{}' partially intersects table '{}' at '{}'; sort the complete table or its exact data range.",
                    resolved.range.a1(),
                    table.path,
                    table_range.a1()
                ),
            ));
        }
    }

    for filter in sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::AutoFilter)
    {
        let filter_range = filter
            .format
            .get("ref")
            .ok_or_else(|| {
                sort_error(
                    "use.office.spreadsheet_filter_invalid",
                    "Spreadsheet AutoFilter has no ref.",
                )
            })
            .and_then(|value| CellRange::parse(value))?;
        if !resolved.range.intersects(filter_range) {
            continue;
        }
        if filter.format.get("nativeMutable").map(String::as_str) != Some("true") {
            return Err(sort_error(
                "use.office.spreadsheet_sort_filter_unsupported",
                format!(
                    "Spreadsheet sort range intersects non-mutable AutoFilter '{}'.",
                    filter.path
                ),
            ));
        }
        let filter_data = CellRange {
            start: CellReference {
                column: filter_range.start.column,
                row: filter_range.start.row.saturating_add(1),
            },
            end: filter_range.end,
        };
        let covers_filter = (resolved.range == filter_range && header)
            || (resolved.range == filter_data && !header);
        if covers_filter {
            data_range = filter_data;
        } else {
            return Err(sort_error(
                "use.office.spreadsheet_sort_filter_partial",
                format!(
                    "Sort range '{}' partially intersects AutoFilter '{}' at '{}'; sort the complete filter range with a header or its exact data range.",
                    resolved.range.a1(),
                    filter.path,
                    filter_range.a1()
                ),
            ));
        }
    }

    let worksheet = package.xml_part(&resolved.part_name)?;
    let root = index_xml(&worksheet)?;
    for collection in root
        .children
        .iter()
        .filter(|child| child.local_name == "mergeCells" && child.namespace == root.namespace)
    {
        for merged in collection
            .children
            .iter()
            .filter(|child| child.local_name == "mergeCell" && child.namespace == root.namespace)
        {
            if let Some(reference) = merged.attributes.get("ref") {
                let merged_range = CellRange::parse(reference)?;
                if resolved.range.intersects(merged_range) {
                    return Err(sort_error(
                        "use.office.spreadsheet_sort_merge_overlap",
                        format!(
                            "Spreadsheet sort range '{}' overlaps merged range '{}'.",
                            resolved.range.a1(),
                            merged_range.a1()
                        ),
                    ));
                }
            }
        }
    }
    Ok(data_range)
}

pub(super) fn validate_existing_state(sheet: &DocumentNode) -> UseResult<()> {
    if let Some(sort_state) = sheet
        .children
        .iter()
        .find(|node| node.node_type == OfficeNodeType::SortState)
    {
        state::require_mutable(sort_state)?;
    }
    Ok(())
}

pub(super) fn validate_formula_boundary(package: &NativeOfficePackage) -> UseResult<()> {
    let parts = package
        .part_names()
        .filter(|part| {
            (part.starts_with("xl/worksheets/") || part.starts_with("xl/tables/"))
                && part.ends_with(".xml")
        })
        .map(str::to_string)
        .collect::<Vec<_>>();
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut formulas = Vec::new();
        for name in ["f", "calculatedColumnFormula", "totalsRowFormula"] {
            index.descendants_named(name, &mut formulas);
        }
        if !formulas.is_empty() {
            return Err(sort_error(
                "use.office.spreadsheet_sort_formula_unsupported",
                "Native Spreadsheet sorting rejects workbooks with cell or table formulas until permutation-safe formula rewriting is available.",
            )
            .with_detail("part", part_name));
        }
    }
    Ok(())
}

fn format_bool(node: &DocumentNode, key: &str) -> UseResult<bool> {
    match node.format.get(key).map(String::as_str) {
        Some("true") => Ok(true),
        Some("false") => Ok(false),
        _ => Err(sort_error(
            "use.office.spreadsheet_sort_metadata_invalid",
            format!(
                "Spreadsheet node '{}' has invalid {key} metadata.",
                node.path
            ),
        )),
    }
}
