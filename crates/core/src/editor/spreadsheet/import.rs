use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::{
    editor_error, expanded_element, indexed_cells, indexed_cells_in_row, indexed_rows,
    mark_workbook_for_recalculation, prefix, qualified, update_dimension,
};
use crate::editor::{
    NativeSpreadsheetAutoFilter, NativeSpreadsheetCellFormat, NativeSpreadsheetDelimitedImport,
    NativeSpreadsheetFrozenPane, NativeSpreadsheetImportResult,
    MAX_NATIVE_SPREADSHEET_IMPORT_CELLS,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference, MAX_COLUMNS, MAX_ROWS};
use crate::xml_edit::{apply_patches, index_xml, insert_child, IndexedXmlElement, XmlPatch};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod parse;

pub(super) fn apply(
    package: &mut NativeOfficePackage,
    requested_sheet: &str,
    request: &NativeSpreadsheetDelimitedImport,
) -> UseResult<NativeSpreadsheetImportResult> {
    let sheet = resolve_sheet(package, requested_sheet)?;
    let start = request.validate()?;
    let parsed = parse::parse(request, workbook_uses_1904_date_system(package)?)?;
    if parsed.rows.is_empty() {
        return Ok(NativeSpreadsheetImportResult {
            path: sheet.path.clone(),
            sheet: sheet.path,
            start_cell: start.a1(),
            range: None,
            format: request.format,
            row_count: 0,
            column_count: 0,
            header: request.header,
            changed: false,
            filter_path: None,
            freeze_path: None,
        });
    }
    let range = target_range(start, parsed.rows.len(), parsed.max_columns)?;
    let worksheet = package.xml_part(&sheet.part)?;
    let root = index_xml(&worksheet)?;
    let sheet_data = root
        .descendant("sheetData")
        .ok_or_else(|| import_part_error(&sheet.part, "has no sheetData element"))?;
    let edited = write_rows(package, &worksheet, sheet_data, start, &parsed.rows)?;
    let edited = update_dimension(&sheet.part, edited)?;
    package.set_part(&sheet.part, edited)?;
    mark_workbook_for_recalculation(package)?;

    let (filter_path, freeze_path) = if request.header {
        let filter = NativeSpreadsheetAutoFilter::new(range.a1());
        let snapshot = NativeOfficeDocument::from_package(package.clone())?;
        let existing = snapshot
            .get(&sheet.path, 1)?
            .children
            .into_iter()
            .find(|node| node.node_type == OfficeNodeType::AutoFilter);
        let filter_path = if let Some(existing) = existing {
            super::auto_filter::set(package, &existing.path, &filter)?
        } else {
            super::auto_filter::add(package, &sheet.path, &filter)?
        };
        let top_row = start
            .row
            .checked_add(1)
            .filter(|row| *row <= MAX_ROWS)
            .ok_or_else(|| {
                import_error(
                    "use.office.spreadsheet_import_row_limit",
                    "A header import cannot freeze below Excel's final worksheet row.",
                )
            })?;
        let top_left = CellReference {
            column: start.column,
            row: top_row,
        }
        .a1();
        let freeze_path = super::view::set(
            package,
            &sheet.path,
            &NativeSpreadsheetFrozenPane::new(start.row, 0, top_left),
        )?;
        (Some(filter_path), Some(freeze_path))
    } else {
        (None, None)
    };

    let range_name = range.a1();
    Ok(NativeSpreadsheetImportResult {
        path: format!("{}/{}", sheet.path, range_name),
        sheet: sheet.path,
        start_cell: start.a1(),
        range: Some(range_name),
        format: request.format,
        row_count: parsed.rows.len(),
        column_count: parsed.max_columns,
        header: request.header,
        changed: true,
        filter_path,
        freeze_path,
    })
}

struct ResolvedSheet {
    path: String,
    part: String,
}

fn resolve_sheet(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedSheet> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(import_error(
            "use.office.mutation_type_unsupported",
            "Delimited import is available only for Spreadsheet documents.",
        ));
    }
    if !requested.starts_with('/')
        || requested.len() < 2
        || requested.trim_start_matches('/').contains('/')
        || requested.chars().any(char::is_control)
    {
        return Err(import_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet import requires a worksheet path such as /Sheet1.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet && node.path.eq_ignore_ascii_case(requested)
        })
        .ok_or_else(|| {
            import_error(
                "use.office.node_not_found",
                format!("Office semantic path '{requested}' does not exist."),
            )
        })?;
    Ok(ResolvedSheet {
        path: sheet.path.clone(),
        part: sheet.format.get("part").cloned().ok_or_else(|| {
            import_error(
                "use.office.spreadsheet_sheet_invalid",
                format!("Worksheet '{}' has no source part.", sheet.path),
            )
        })?,
    })
}

fn target_range(start: CellReference, rows: usize, columns: usize) -> UseResult<CellRange> {
    let row_count = u32::try_from(rows).map_err(|_| row_limit())?;
    let column_count = u32::try_from(columns).map_err(|_| column_limit())?;
    let end_row = start
        .row
        .checked_add(row_count.saturating_sub(1))
        .filter(|row| *row <= MAX_ROWS)
        .ok_or_else(row_limit)?;
    let end_column = start
        .column
        .checked_add(column_count.saturating_sub(1))
        .filter(|column| *column <= MAX_COLUMNS)
        .ok_or_else(column_limit)?;
    let cells = rows.checked_mul(columns).ok_or_else(cell_limit)?;
    if cells > MAX_NATIVE_SPREADSHEET_IMPORT_CELLS {
        return Err(cell_limit().with_detail("cells", cells));
    }
    Ok(CellRange {
        start,
        end: CellReference {
            column: end_column,
            row: end_row,
        },
    })
}

fn write_rows(
    package: &mut NativeOfficePackage,
    worksheet: &LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    start: CellReference,
    source_rows: &[Vec<parse::ParsedField>],
) -> UseResult<Vec<u8>> {
    let rows = indexed_rows(sheet_data);
    let row_map = rows.iter().copied().collect::<BTreeMap<_, _>>();
    let existing_cells = indexed_cells(sheet_data)
        .into_iter()
        .map(|(reference, _, cell)| (reference, cell))
        .collect::<BTreeMap<_, _>>();
    let mut date_base_styles = BTreeSet::new();
    for (row_offset, fields) in source_rows.iter().enumerate() {
        for (column_offset, field) in fields.iter().enumerate() {
            if !field.value().is_some_and(|(_, date)| date) {
                continue;
            }
            let reference = source_reference(start, row_offset, column_offset)?;
            let base = existing_cells
                .get(&reference)
                .map(|cell| super::style::cell_style_index(cell))
                .transpose()?
                .unwrap_or(0);
            date_base_styles.insert(base);
        }
    }
    let date_styles = if date_base_styles.is_empty() {
        BTreeMap::new()
    } else {
        super::style::derived_cell_style_indexes(
            package,
            &date_base_styles,
            &NativeSpreadsheetCellFormat {
                number_format: Some("date".into()),
                ..NativeSpreadsheetCellFormat::default()
            },
        )?
    };

    if sheet_data.empty {
        let prefix = prefix(&sheet_data.qualified_name);
        let row_tag = qualified(prefix, "row");
        let mut rows = String::new();
        for (row_offset, fields) in source_rows.iter().enumerate() {
            let cells = new_cells(prefix, start, row_offset, fields, &date_styles)?;
            if cells.is_empty() {
                continue;
            }
            let row_number = start
                .row
                .checked_add(u32::try_from(row_offset).map_err(|_| row_limit())?)
                .ok_or_else(row_limit)?;
            rows.push_str(&format!(
                "<{row_tag} r=\"{row_number}\">{cells}</{row_tag}>"
            ));
        }
        return insert_child(worksheet, sheet_data, rows);
    }

    let mut patches = Vec::new();
    let mut insertions = BTreeMap::<usize, Vec<(u32, u32, String)>>::new();
    let sheet_prefix = prefix(&sheet_data.qualified_name);
    for (row_offset, fields) in source_rows.iter().enumerate() {
        let row_number = start
            .row
            .checked_add(u32::try_from(row_offset).map_err(|_| row_limit())?)
            .ok_or_else(row_limit)?;
        let Some(row) = row_map.get(&row_number).copied() else {
            let cells = new_cells(sheet_prefix, start, row_offset, fields, &date_styles)?;
            if cells.is_empty() {
                continue;
            }
            let row_tag = qualified(sheet_prefix, "row");
            let fragment = format!("<{row_tag} r=\"{row_number}\">{cells}</{row_tag}>");
            let position = rows
                .iter()
                .find(|(existing, _)| *existing > row_number)
                .map_or(sheet_data.content_range.end, |(_, next)| {
                    next.full_range.start
                });
            insertions
                .entry(position)
                .or_default()
                .push((row_number, 0, fragment));
            continue;
        };
        if row.empty {
            let cells = new_cells(
                prefix(&row.qualified_name),
                start,
                row_offset,
                fields,
                &date_styles,
            )?;
            if !cells.is_empty() {
                patches.push(XmlPatch::new(
                    row.full_range.clone(),
                    expanded_element(row, &cells),
                ));
            }
            continue;
        }
        let cells = indexed_cells_in_row(row_number, row);
        for (column_offset, field) in fields.iter().enumerate() {
            let reference = source_reference(start, row_offset, column_offset)?;
            if let Some((_, cell)) = cells
                .iter()
                .find(|(existing, _)| existing.column == reference.column)
            {
                patches.push(XmlPatch::new(
                    cell.full_range.clone(),
                    existing_cell_fragment(worksheet, cell, reference, field, &date_styles)?,
                ));
            } else if !field.is_empty() {
                let position = cells
                    .iter()
                    .find(|(existing, _)| existing.column > reference.column)
                    .map(|(_, next)| next.full_range.start)
                    .or_else(|| {
                        row.children
                            .iter()
                            .find(|child| child.local_name != "c")
                            .map(|child| child.full_range.start)
                    })
                    .unwrap_or(row.content_range.end);
                insertions.entry(position).or_default().push((
                    row_number,
                    reference.column,
                    new_cell(prefix(&row.qualified_name), reference, field, &date_styles)?,
                ));
            }
        }
    }
    for (position, mut fragments) in insertions {
        fragments.sort_by_key(|(row, column, _)| (*row, *column));
        patches.push(XmlPatch::new(
            position..position,
            fragments
                .into_iter()
                .map(|(_, _, fragment)| fragment)
                .collect::<String>(),
        ));
    }
    apply_patches(worksheet, patches)
}

fn new_cells(
    prefix: Option<&str>,
    start: CellReference,
    row_offset: usize,
    fields: &[parse::ParsedField],
    date_styles: &BTreeMap<usize, usize>,
) -> UseResult<String> {
    fields
        .iter()
        .enumerate()
        .filter(|(_, field)| !field.is_empty())
        .map(|(column_offset, field)| {
            let reference = source_reference(start, row_offset, column_offset)?;
            new_cell(prefix, reference, field, date_styles)
        })
        .collect()
}

fn new_cell(
    prefix: Option<&str>,
    reference: CellReference,
    field: &parse::ParsedField,
    date_styles: &BTreeMap<usize, usize>,
) -> UseResult<String> {
    let (value, date) = field.value().ok_or_else(import_cell_invalid)?;
    let tag = qualified(prefix, "c");
    let (value_type, content) = super::cell_content(prefix, value);
    let style = if date {
        format!(" s=\"{}\"", required_date_style(date_styles, 0)?)
    } else {
        String::new()
    };
    let value_type = value_type
        .map(|value_type| format!(" t=\"{value_type}\""))
        .unwrap_or_default();
    Ok(format!(
        "<{tag} r=\"{}\"{style}{value_type}>{content}</{tag}>",
        reference.a1()
    ))
}

fn existing_cell_fragment(
    part: &LosslessXmlPart,
    cell: &IndexedXmlElement,
    reference: CellReference,
    field: &parse::ParsedField,
    date_styles: &BTreeMap<usize, usize>,
) -> UseResult<String> {
    let mut attributes = cell.qualified_attributes.clone();
    attributes.insert("r".into(), reference.a1());
    let mut content = String::new();
    if let Some((value, date)) = field.value() {
        let (value_type, value_content) = super::cell_content(prefix(&cell.qualified_name), value);
        if let Some(value_type) = value_type {
            attributes.insert("t".into(), value_type.into());
        } else {
            attributes.remove("t");
        }
        if date {
            let base = super::style::cell_style_index(cell)?;
            attributes.insert(
                "s".into(),
                required_date_style(date_styles, base)?.to_string(),
            );
        }
        content.push_str(&value_content);
    } else {
        attributes.remove("t");
    }
    for child in &cell.children {
        let owned_value = child.namespace == cell.namespace
            && matches!(child.local_name.as_str(), "f" | "v" | "is");
        if !owned_value {
            let bytes = &part.parse_bytes()[child.full_range.clone()];
            content.push_str(std::str::from_utf8(bytes).map_err(|error| {
                import_error(
                    "use.office.spreadsheet_import_cell_invalid",
                    format!("Spreadsheet cell extension content is not UTF-8: {error}"),
                )
            })?);
        }
    }
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", crate::xml_edit::escape_attribute(&value)))
        .collect::<String>();
    let tag = qualified(prefix(&cell.qualified_name), "c");
    Ok(format!("<{tag}{attributes}>{content}</{tag}>"))
}

fn source_reference(
    start: CellReference,
    row_offset: usize,
    column_offset: usize,
) -> UseResult<CellReference> {
    Ok(CellReference {
        column: start
            .column
            .checked_add(u32::try_from(column_offset).map_err(|_| column_limit())?)
            .ok_or_else(column_limit)?,
        row: start
            .row
            .checked_add(u32::try_from(row_offset).map_err(|_| row_limit())?)
            .ok_or_else(row_limit)?,
    })
}

fn required_date_style(styles: &BTreeMap<usize, usize>, base: usize) -> UseResult<usize> {
    styles.get(&base).copied().ok_or_else(|| {
        import_error(
            "use.office.spreadsheet_styles_invalid",
            format!("Spreadsheet import could not derive a date style from style {base}."),
        )
    })
}

fn workbook_uses_1904_date_system(package: &NativeOfficePackage) -> UseResult<bool> {
    let workbook = package.xml_part("xl/workbook.xml")?;
    let root = index_xml(&workbook)?;
    let properties = root
        .children
        .iter()
        .filter(|child| child.local_name == "workbookPr" && child.namespace == root.namespace)
        .collect::<Vec<_>>();
    if properties.len() > 1 {
        return Err(import_error(
            "use.office.spreadsheet_import_date_system_invalid",
            "Spreadsheet workbook contains multiple workbookPr elements.",
        ));
    }
    match properties
        .first()
        .and_then(|properties| properties.attributes.get("date1904"))
        .map(String::as_str)
    {
        None | Some("0" | "false") => Ok(false),
        Some("1" | "true") => Ok(true),
        Some(value) => Err(import_error(
            "use.office.spreadsheet_import_date_system_invalid",
            format!("Spreadsheet workbook has invalid date1904='{value}'."),
        )),
    }
}

fn row_limit() -> a3s_use_core::UseError {
    import_error(
        "use.office.spreadsheet_import_row_limit",
        format!("Spreadsheet import cannot exceed Excel's {MAX_ROWS} rows."),
    )
}

fn column_limit() -> a3s_use_core::UseError {
    import_error(
        "use.office.spreadsheet_import_column_limit",
        format!("Spreadsheet import cannot exceed Excel's {MAX_COLUMNS} columns."),
    )
}

fn cell_limit() -> a3s_use_core::UseError {
    import_error(
        "use.office.spreadsheet_import_cell_count_limit",
        format!(
            "Spreadsheet import accepts at most {MAX_NATIVE_SPREADSHEET_IMPORT_CELLS} rectangular target cells."
        ),
    )
}

fn import_cell_invalid() -> a3s_use_core::UseError {
    import_error(
        "use.office.spreadsheet_import_cell_invalid",
        "Spreadsheet import attempted to materialize an empty source field.",
    )
}

fn import_part_error(part: &str, reason: &str) -> a3s_use_core::UseError {
    import_error(
        "use.office.spreadsheet_import_part_invalid",
        format!("Spreadsheet worksheet part '{part}' {reason}."),
    )
    .with_detail("part", part)
}

fn import_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(code, message)
}
