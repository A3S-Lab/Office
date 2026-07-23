use a3s_use_core::UseResult;

use super::table_xml::{
    cell_xml, drawing_namespace, graphic_frame_xml, grid_column_xml, insert_ordered_child,
    ordered_child_insertion, replace_attribute_patch, row_xml,
};
use super::{editor_error, locate_path, node_not_found, prefix};
use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::xml_edit::{apply_patches, index_xml, IndexedXmlElement, XmlPatch};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod column;

const DEFAULT_ROW_HEIGHT_EMU: u64 = 370_840;
const MAX_TABLE_ROWS: usize = 5_000;
const MAX_TABLE_COLUMNS: usize = 5_000;
const MAX_TABLE_CELLS: usize = 100_000;

pub(super) fn is_column_path(path: &str) -> bool {
    column::is_path(path)
}

pub(super) fn move_column(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&crate::editor::NativeOfficeInsertPosition>,
) -> UseResult<String> {
    column::move_column(package, path, target_parent, position)
}

pub(super) fn copy_column(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&crate::editor::NativeOfficeInsertPosition>,
) -> UseResult<String> {
    column::copy_column(package, path, target_parent, position)
}

pub(super) fn swap_columns(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<crate::editor::NativeOfficeSwapResult> {
    column::swap_columns(package, path, with)
}

pub(super) fn set_column_width(
    package: &mut NativeOfficePackage,
    path: &str,
    width_emu: u64,
) -> UseResult<()> {
    column::set_width(package, path, width_emu)
}

pub(super) fn add_table(
    package: &mut NativeOfficePackage,
    parent: &str,
    rows: usize,
    columns: usize,
) -> UseResult<String> {
    require_presentation(package, "add-table")?;
    validate_dimensions(rows, columns)?;

    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide = snapshot.get(parent, 0)?;
    if slide.node_type != OfficeNodeType::Slide {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Presentation tables require a slide parent such as /slide[1].",
        ));
    }
    let part_name = source_part(&slide)?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let shape_tree = index
        .descendant("spTree")
        .ok_or_else(|| node_not_found(parent))?;
    let position = shape_tree
        .children
        .iter()
        .filter(|child| child.local_name == "graphicFrame" && child.descendant("tbl").is_some())
        .count()
        + 1;
    let id = next_non_visual_id(shape_tree)?;
    let height = DEFAULT_ROW_HEIGHT_EMU
        .checked_mul(u64::try_from(rows).map_err(|_| table_limit("Table row count overflowed."))?)
        .ok_or_else(|| table_limit("Presentation table height overflowed."))?;
    let fragment = graphic_frame_xml(
        id,
        position,
        rows,
        columns,
        height,
        prefix(&shape_tree.qualified_name),
        drawing_namespace(&part),
    );
    let edited = insert_ordered_child(&part, shape_tree, fragment)?;
    package.set_part(part_name, edited)?;
    Ok(format!("{}/table[{position}]", slide.path))
}

pub(super) fn add_row(
    package: &mut NativeOfficePackage,
    parent: &str,
    columns: Option<usize>,
) -> UseResult<String> {
    require_presentation(package, "add-table-row")?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(parent, 0)?;
    if requested.node_type != OfficeNodeType::Table {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Presentation table rows require a table parent such as /slide[1]/table[1].",
        ));
    }
    let part_name = slide_part_for_path(&snapshot, parent)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let frame = locate_path(&index, parent)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(parent, "has no DrawingML table element"))?;
    let grid_columns = table_grid_columns(table);
    if grid_columns == 0 {
        return Err(invalid_table(parent, "has no grid columns"));
    }
    if let Some(columns) = columns {
        if columns != grid_columns {
            return Err(editor_error(
                "use.office.presentation_table_row_grid_mismatch",
                format!(
                    "Presentation table row requested {columns} columns, but the table grid has {grid_columns}."
                ),
            ));
        }
    }
    let existing_rows = direct_child_count(table, "tr");
    validate_dimensions(existing_rows.saturating_add(1), grid_columns)?;
    let height = total_row_height(table)?
        .checked_add(DEFAULT_ROW_HEIGHT_EMU)
        .ok_or_else(|| table_limit("Presentation table height overflowed."))?;
    let extents = frame_extents(frame, parent)?;
    let insertion = ordered_child_insertion(table);
    let patches = vec![
        XmlPatch::new(
            insertion..insertion,
            row_xml(grid_columns, prefix(&table.qualified_name)),
        ),
        replace_attribute_patch(extents, "cy", height.to_string()),
    ];
    let edited = apply_patches(&part, patches)?;
    package.set_part(&part_name, edited)?;
    Ok(format!("{parent}/tr[{}]", existing_rows + 1))
}

pub(super) fn add_cell(
    package: &mut NativeOfficePackage,
    parent: &str,
    text: &str,
) -> UseResult<String> {
    require_presentation(package, "add-table-cell")?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(parent, 0)?;
    if requested.node_type != OfficeNodeType::TableRow {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Presentation table cells require a row parent such as /slide[1]/table[1]/tr[1].",
        ));
    }
    let table_path = parent_path(parent).ok_or_else(|| node_not_found(parent))?;
    let part_name = slide_part_for_path(&snapshot, parent)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let frame = locate_path(&index, table_path)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(table_path, "has no DrawingML table element"))?;
    let grid_columns = table_grid_columns(table);
    if grid_columns == 0 {
        return Err(invalid_table(table_path, "has no grid columns"));
    }
    validate_dimensions(direct_child_count(table, "tr"), grid_columns)?;
    let row = locate_path(&index, parent)?;
    let occupied = logical_cell_count(row)?;
    if occupied >= grid_columns {
        return Err(editor_error(
            "use.office.presentation_table_cell_grid_full",
            format!(
                "Presentation table row '{parent}' already occupies all {grid_columns} grid columns. Add a column to the parent table instead; the row was not changed."
            ),
        ));
    }
    let physical_position = direct_child_count(row, "tc") + 1;
    let edited = insert_ordered_child(&part, row, cell_xml(text, prefix(&row.qualified_name)))?;
    package.set_part(&part_name, edited)?;
    Ok(format!("{parent}/tc[{physical_position}]"))
}

pub(super) fn add_column(
    package: &mut NativeOfficePackage,
    parent: &str,
    index: Option<usize>,
    text: &str,
) -> UseResult<String> {
    require_presentation(package, "add-table-column")?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(parent, 0)?;
    if requested.node_type != OfficeNodeType::Table {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Presentation table columns require a table parent such as /slide[1]/table[1].",
        ));
    }
    let part_name = slide_part_for_path(&snapshot, parent)?;
    let part = package.xml_part(&part_name)?;
    let xml = index_xml(&part)?;
    let frame = locate_path(&xml, parent)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(parent, "has no DrawingML table element"))?;
    ensure_no_merges(table, "column insertion")?;
    let grid = table
        .child("tblGrid", 1)
        .ok_or_else(|| invalid_table(parent, "has no table grid"))?;
    let columns = direct_children(grid, "gridCol");
    let rows = direct_children(table, "tr");
    let column_count = columns.len();
    validate_dimensions(rows.len(), column_count)?;
    ensure_rectangular_rows(parent, &rows, column_count)?;
    let insertion_index = index.unwrap_or(column_count);
    if insertion_index > column_count {
        return Err(editor_error(
            "use.office.presentation_table_column_index_invalid",
            format!(
                "Presentation table column index {insertion_index} is outside the zero-based insertion range 0..={column_count}."
            ),
        ));
    }
    validate_dimensions(rows.len(), column_count.saturating_add(1))?;
    let width = average_grid_width(parent, &columns)?;
    let total_width = total_grid_width(parent, &columns)?
        .checked_add(width)
        .filter(|total| *total <= i64::MAX as u64)
        .ok_or_else(|| table_limit("Presentation table width overflowed."))?;
    let extents = frame_extents(frame, parent)?;
    let mut patches = Vec::with_capacity(rows.len().saturating_add(2));
    patches.push(XmlPatch::new(
        insertion_before_or_end(grid, "gridCol", insertion_index)
            ..insertion_before_or_end(grid, "gridCol", insertion_index),
        grid_column_xml(width, prefix(&grid.qualified_name)),
    ));
    for row in rows {
        let insertion = insertion_before_or_end(row, "tc", insertion_index);
        patches.push(XmlPatch::new(
            insertion..insertion,
            cell_xml(text, prefix(&row.qualified_name)),
        ));
    }
    patches.push(replace_attribute_patch(
        extents,
        "cx",
        total_width.to_string(),
    ));
    let edited = apply_patches(&part, patches)?;
    package.set_part(&part_name, edited)?;
    Ok(format!("{parent}/col[{}]", insertion_index + 1))
}

pub(super) fn remove(
    package: &mut NativeOfficePackage,
    snapshot: &NativeOfficeDocument,
    requested: &DocumentNode,
) -> UseResult<()> {
    let part_name = slide_part_for_path(snapshot, &requested.path)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let edited =
        match requested.node_type {
            OfficeNodeType::Table => {
                let frame = locate_path(&index, &requested.path)?;
                apply_patches(
                    &part,
                    vec![XmlPatch::new(frame.full_range.clone(), Vec::new())],
                )?
            }
            OfficeNodeType::TableRow => remove_row(&part, &index, &requested.path)?,
            OfficeNodeType::TableColumn => remove_column(&part, &index, &requested.path)?,
            OfficeNodeType::TableCell => remove_cell(&part, &index, &requested.path)?,
            _ => return Err(editor_error(
                "use.office.mutation_type_unsupported",
                "Native Presentation table removal requires a table, row, column, or cell path.",
            )),
        };
    package.set_part(&part_name, edited)
}

fn remove_column(
    part: &LosslessXmlPart,
    index: &IndexedXmlElement,
    path: &str,
) -> UseResult<Vec<u8>> {
    let table_path = parent_path(path).ok_or_else(|| node_not_found(path))?;
    let position = table_column_position(path).ok_or_else(|| node_not_found(path))?;
    let frame = locate_path(index, table_path)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(table_path, "has no DrawingML table element"))?;
    ensure_no_merges(table, "column removal")?;
    let grid = table
        .child("tblGrid", 1)
        .ok_or_else(|| invalid_table(table_path, "has no table grid"))?;
    let columns = direct_children(grid, "gridCol");
    let rows = direct_children(table, "tr");
    validate_dimensions(rows.len(), columns.len())?;
    ensure_rectangular_rows(table_path, &rows, columns.len())?;
    if columns.len() <= 1 {
        return Err(editor_error(
            "use.office.presentation_last_table_column",
            "A Presentation table must retain at least one column; remove the table instead.",
        ));
    }
    let column_index = position - 1;
    let column = columns
        .get(column_index)
        .ok_or_else(|| node_not_found(path))?;
    let removed_width = grid_column_width(table_path, column)?;
    let width = total_grid_width(table_path, &columns)?
        .checked_sub(removed_width)
        .ok_or_else(|| invalid_table(table_path, "has inconsistent column widths"))?;
    let extents = frame_extents(frame, table_path)?;
    let mut patches = Vec::with_capacity(rows.len().saturating_add(2));
    patches.push(XmlPatch::new(column.full_range.clone(), Vec::new()));
    for row in rows {
        let cell = direct_children(row, "tc")
            .get(column_index)
            .copied()
            .ok_or_else(|| invalid_table(table_path, "has a row shorter than its table grid"))?;
        patches.push(XmlPatch::new(cell.full_range.clone(), Vec::new()));
    }
    patches.push(replace_attribute_patch(extents, "cx", width.to_string()));
    apply_patches(part, patches)
}

fn remove_row(part: &LosslessXmlPart, index: &IndexedXmlElement, path: &str) -> UseResult<Vec<u8>> {
    let table_path = parent_path(path).ok_or_else(|| node_not_found(path))?;
    let frame = locate_path(index, table_path)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(table_path, "has no DrawingML table element"))?;
    ensure_no_merges(table, "row removal")?;
    if direct_child_count(table, "tr") <= 1 {
        return Err(editor_error(
            "use.office.presentation_last_table_row",
            "A Presentation table must retain at least one row; remove the table instead.",
        ));
    }
    let row = locate_path(index, path)?;
    let height = total_row_height(table)?
        .checked_sub(row_height(row)?)
        .ok_or_else(|| invalid_table(table_path, "has inconsistent row heights"))?;
    let extents = frame_extents(frame, table_path)?;
    apply_patches(
        part,
        vec![
            XmlPatch::new(row.full_range.clone(), Vec::new()),
            replace_attribute_patch(extents, "cy", height.to_string()),
        ],
    )
}

fn remove_cell(
    part: &LosslessXmlPart,
    index: &IndexedXmlElement,
    path: &str,
) -> UseResult<Vec<u8>> {
    let row_path = parent_path(path).ok_or_else(|| node_not_found(path))?;
    let table_path = parent_path(row_path).ok_or_else(|| node_not_found(path))?;
    let frame = locate_path(index, table_path)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(table_path, "has no DrawingML table element"))?;
    ensure_no_merges(table, "cell removal")?;
    let grid_columns = table_grid_columns(table);
    if grid_columns == 0 {
        return Err(invalid_table(table_path, "has no grid columns"));
    }
    let row = locate_path(index, row_path)?;
    let cell = locate_path(index, path)?;
    let occupied = logical_cell_count(row)?;
    let removed = cell_grid_span(cell)?;
    let retained = occupied.saturating_sub(removed);
    if retained < grid_columns {
        return Err(editor_error(
            "use.office.presentation_table_cell_grid_invalid",
            format!(
                "Removing '{path}' would leave {retained} occupied columns for a {grid_columns}-column Presentation table grid. Remove the corresponding parent-table column instead; the row was not changed."
            ),
        ));
    }
    apply_patches(
        part,
        vec![XmlPatch::new(cell.full_range.clone(), Vec::new())],
    )
}

fn next_non_visual_id(shape_tree: &IndexedXmlElement) -> UseResult<u32> {
    let mut properties = Vec::new();
    shape_tree.descendants_named("cNvPr", &mut properties);
    properties
        .into_iter()
        .filter_map(|element| element.attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(1)
        .checked_add(1)
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_shape_limit",
                "Presentation non-visual object IDs are exhausted.",
            )
        })
}

fn source_part(node: &DocumentNode) -> UseResult<&str> {
    node.format.get("part").map(String::as_str).ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })
}

fn slide_part_for_path(snapshot: &NativeOfficeDocument, path: &str) -> UseResult<String> {
    let slide_segment = path
        .trim_start_matches('/')
        .split('/')
        .find(|segment| segment.starts_with("slide["))
        .ok_or_else(|| node_not_found(path))?;
    let slide = snapshot.get(&format!("/{slide_segment}"), 0)?;
    source_part(&slide).map(str::to_string)
}

fn frame_extents<'a>(frame: &'a IndexedXmlElement, path: &str) -> UseResult<&'a IndexedXmlElement> {
    frame
        .child("xfrm", 1)
        .and_then(|transform| transform.child("ext", 1))
        .ok_or_else(|| invalid_table(path, "has no graphic-frame extents"))
}

fn validate_dimensions(rows: usize, columns: usize) -> UseResult<()> {
    if rows == 0 || columns == 0 {
        return Err(editor_error(
            "use.office.presentation_table_dimensions_invalid",
            "Native Presentation table dimensions must be positive integers.",
        ));
    }
    if rows > MAX_TABLE_ROWS || columns > MAX_TABLE_COLUMNS {
        return Err(table_limit(format!(
            "Native Presentation tables support at most {MAX_TABLE_ROWS} rows and {MAX_TABLE_COLUMNS} columns."
        )));
    }
    let cells = rows
        .checked_mul(columns)
        .ok_or_else(|| table_limit("Presentation table dimensions overflowed."))?;
    if cells > MAX_TABLE_CELLS {
        return Err(table_limit(format!(
            "Native Presentation table creation is limited to {MAX_TABLE_CELLS} cells."
        )));
    }
    Ok(())
}

fn table_grid_columns(table: &IndexedXmlElement) -> usize {
    table
        .child("tblGrid", 1)
        .map_or(0, |grid| direct_child_count(grid, "gridCol"))
}

fn ensure_rectangular_rows(
    path: &str,
    rows: &[&IndexedXmlElement],
    columns: usize,
) -> UseResult<()> {
    if rows
        .iter()
        .any(|row| direct_child_count(row, "tc") != columns)
    {
        return Err(editor_error(
            "use.office.presentation_table_grid_mismatch",
            format!(
                "Presentation table '{path}' has a row whose physical cell count does not match its {columns}-column grid. Repair the row before changing columns."
            ),
        ));
    }
    Ok(())
}

fn average_grid_width(path: &str, columns: &[&IndexedXmlElement]) -> UseResult<u64> {
    let total = total_grid_width(path, columns)?;
    let count = u64::try_from(columns.len())
        .map_err(|_| table_limit("Presentation table column count overflowed."))?;
    total
        .checked_div(count)
        .filter(|width| *width > 0)
        .ok_or_else(|| invalid_table(path, "has no positive-width grid columns"))
}

fn total_grid_width(path: &str, columns: &[&IndexedXmlElement]) -> UseResult<u64> {
    columns.iter().try_fold(0_u64, |total, column| {
        total
            .checked_add(grid_column_width(path, column)?)
            .filter(|width| *width <= i64::MAX as u64)
            .ok_or_else(|| table_limit("Presentation table width overflowed."))
    })
}

fn grid_column_width(path: &str, column: &IndexedXmlElement) -> UseResult<u64> {
    column
        .attributes
        .get("w")
        .and_then(|value| value.parse::<u64>().ok())
        .filter(|width| *width > 0 && *width <= i64::MAX as u64)
        .ok_or_else(|| {
            invalid_table(
                path,
                "has a grid column outside the positive signed-64-bit width range",
            )
        })
}

fn insertion_before_or_end(parent: &IndexedXmlElement, name: &str, index: usize) -> usize {
    direct_children(parent, name).get(index).map_or_else(
        || ordered_child_insertion(parent),
        |child| child.full_range.start,
    )
}

fn direct_children<'a>(element: &'a IndexedXmlElement, name: &str) -> Vec<&'a IndexedXmlElement> {
    element
        .children
        .iter()
        .filter(|child| child.local_name == name)
        .collect()
}

fn logical_cell_count(row: &IndexedXmlElement) -> UseResult<usize> {
    row.children
        .iter()
        .filter(|child| child.local_name == "tc")
        .try_fold(0_usize, |total, cell| {
            total.checked_add(cell_grid_span(cell)?).ok_or_else(|| {
                invalid_table("Presentation table row", "has an overflowing grid span")
            })
        })
}

fn cell_grid_span(cell: &IndexedXmlElement) -> UseResult<usize> {
    match cell.attributes.get("gridSpan") {
        None => Ok(1),
        Some(value) => value
            .parse::<usize>()
            .ok()
            .filter(|value| *value > 0)
            .ok_or_else(|| invalid_table("Presentation table cell", "has an invalid gridSpan")),
    }
}

fn total_row_height(table: &IndexedXmlElement) -> UseResult<u64> {
    table
        .children
        .iter()
        .filter(|child| child.local_name == "tr")
        .try_fold(0_u64, |total, row| {
            total
                .checked_add(row_height(row)?)
                .ok_or_else(|| table_limit("Presentation table height overflowed."))
        })
}

fn row_height(row: &IndexedXmlElement) -> UseResult<u64> {
    row.attributes
        .get("h")
        .map_or(Ok(DEFAULT_ROW_HEIGHT_EMU), |value| {
            value
                .parse::<u64>()
                .map_err(|_| invalid_table("Presentation table row", "has an invalid height"))
        })
}

fn ensure_no_merges(table: &IndexedXmlElement, operation: &str) -> UseResult<()> {
    let mut cells = Vec::new();
    table.descendants_named("tc", &mut cells);
    let merged = cells.into_iter().any(|cell| {
        cell.attributes
            .get("gridSpan")
            .and_then(|value| value.parse::<usize>().ok())
            .is_some_and(|value| value > 1)
            || cell
                .attributes
                .get("rowSpan")
                .and_then(|value| value.parse::<usize>().ok())
                .is_some_and(|value| value > 1)
            || attribute_is_true(cell, "hMerge")
            || attribute_is_true(cell, "vMerge")
    });
    if merged {
        Err(editor_error(
            "use.office.presentation_table_merge_unsupported",
            format!(
                "Native Presentation {operation} does not yet rewrite merged-cell spans; the table was not changed."
            ),
        ))
    } else {
        Ok(())
    }
}

fn attribute_is_true(element: &IndexedXmlElement, name: &str) -> bool {
    element
        .attributes
        .get(name)
        .is_some_and(|value| matches!(value.as_str(), "1" | "true" | "on"))
}

fn direct_child_count(element: &IndexedXmlElement, name: &str) -> usize {
    element
        .children
        .iter()
        .filter(|child| child.local_name == name)
        .count()
}

fn parent_path(path: &str) -> Option<&str> {
    path.rsplit_once('/').map(|(parent, _)| parent)
}

fn table_column_position(path: &str) -> Option<usize> {
    path.rsplit_once("/col[")
        .and_then(|(_, position)| position.strip_suffix(']'))
        .and_then(|position| position.parse::<usize>().ok())
        .filter(|position| *position > 0)
}

fn require_presentation(package: &NativeOfficePackage, operation: &str) -> UseResult<()> {
    if package.kind() == DocumentKind::Presentation {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_type_unsupported",
            format!("Native {operation} is available only for Presentation documents."),
        ))
    }
}

fn invalid_table(path: &str, problem: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.presentation_table_invalid",
        format!("Presentation table '{path}' {problem}."),
    )
}

fn table_limit(message: impl Into<String>) -> a3s_use_core::UseError {
    editor_error("use.office.presentation_table_limit", message)
}
