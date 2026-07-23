use a3s_use_core::UseResult;

use super::{
    direct_children, ensure_no_merges, ensure_rectangular_rows, frame_extents, invalid_table,
    node_not_found, ordered_child_insertion, require_presentation, slide_part_for_path,
    table_limit, total_grid_width, validate_dimensions,
};
use crate::editor::{NativeOfficeInsertPosition, NativeOfficeSwapResult};
use crate::xml_edit::{apply_patches, element_fragment, index_xml, IndexedXmlElement, XmlPatch};
use crate::{LosslessXmlPart, NativeOfficePackage};

#[derive(Debug, Clone, PartialEq, Eq)]
struct ColumnPath {
    table: String,
    position: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Operation {
    Move,
    Copy,
}

pub(super) fn is_path(path: &str) -> bool {
    parse_path(path).is_some()
}

pub(super) fn set_width(
    package: &mut NativeOfficePackage,
    path: &str,
    width_emu: u64,
) -> UseResult<()> {
    require_presentation(package, "set-table-column-width")?;
    if width_emu == 0 || width_emu > i64::MAX as u64 {
        return Err(super::editor_error(
            "use.office.presentation_table_width_invalid",
            "Presentation table-column width must be a positive signed 64-bit EMU value.",
        ));
    }
    let source = parse_path(path).ok_or_else(|| node_not_found(path))?;
    let part_name = slide_part(package, &source.table)?;
    let part = package.xml_part(&part_name)?;
    let xml = index_xml(&part)?;
    let frame = super::locate_path(&xml, &source.table)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(&source.table, "has no DrawingML table element"))?;
    let grid = table
        .child("tblGrid", 1)
        .ok_or_else(|| invalid_table(&source.table, "has no table grid"))?;
    let columns = direct_children(grid, "gridCol");
    let rows = direct_children(table, "tr");
    validate_dimensions(rows.len(), columns.len())?;
    let source_index = checked_index(path, source.position, columns.len())?;
    let old_width = super::grid_column_width(&source.table, columns[source_index])?;
    let width = total_grid_width(&source.table, &columns)?
        .checked_sub(old_width)
        .and_then(|remaining| remaining.checked_add(width_emu))
        .filter(|total| *total <= i64::MAX as u64)
        .ok_or_else(|| {
            super::editor_error(
                "use.office.presentation_table_width_invalid",
                "Presentation table width exceeds the signed 64-bit OOXML coordinate range.",
            )
        })?;
    let extents = frame_extents(frame, &source.table)?;
    let edited = apply_patches(
        &part,
        vec![
            super::replace_attribute_patch(columns[source_index], "w", width_emu.to_string()),
            super::replace_attribute_patch(extents, "cx", width.to_string()),
        ],
    )?;
    package.set_part(&part_name, edited)
}

pub(super) fn move_column(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    require_presentation(package, "move-table-column")?;
    let source = parse_path(path).ok_or_else(|| node_not_found(path))?;
    require_same_table(&source, target_parent)?;
    let part_name = slide_part(package, &source.table)?;
    let part = package.xml_part(&part_name)?;
    let xml = index_xml(&part)?;
    let frame = super::locate_path(&xml, &source.table)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(&source.table, "has no DrawingML table element"))?;
    let (grid, columns, rows) = validated_structure(table, &source.table)?;
    let source_index = checked_index(path, source.position, columns.len())?;
    let slot = resolve_slot(
        &source,
        columns.len(),
        source_index,
        position,
        Operation::Move,
    )?;
    if slot == source_index {
        return Ok(path.to_string());
    }

    let mut patches = Vec::with_capacity(rows.len().saturating_mul(2).saturating_add(2));
    append_move_patches(&part, grid, &columns, source_index, slot, &mut patches)?;
    for row in rows {
        let cells = direct_children(row, "tc");
        append_move_patches(&part, row, &cells, source_index, slot, &mut patches)?;
    }
    let edited = apply_patches(&part, patches)?;
    package.set_part(&part_name, edited)?;
    Ok(format!("{}/col[{}]", source.table, slot + 1))
}

pub(super) fn copy_column(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    require_presentation(package, "copy-table-column")?;
    let source = parse_path(path).ok_or_else(|| node_not_found(path))?;
    require_same_table(&source, target_parent)?;
    let part_name = slide_part(package, &source.table)?;
    let part = package.xml_part(&part_name)?;
    let xml = index_xml(&part)?;
    let frame = super::locate_path(&xml, &source.table)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(&source.table, "has no DrawingML table element"))?;
    let (grid, columns, rows) = validated_structure(table, &source.table)?;
    let source_index = checked_index(path, source.position, columns.len())?;
    let slot = resolve_slot(
        &source,
        columns.len(),
        source_index,
        position,
        Operation::Copy,
    )?;
    validate_dimensions(rows.len(), columns.len().saturating_add(1))?;
    let source_width = super::grid_column_width(&source.table, columns[source_index])?;
    let width = total_grid_width(&source.table, &columns)?
        .checked_add(source_width)
        .filter(|total| *total <= i64::MAX as u64)
        .ok_or_else(|| table_limit("Presentation table width overflowed."))?;
    let extents = frame_extents(frame, &source.table)?;

    let mut patches = Vec::with_capacity(rows.len().saturating_add(2));
    append_copy_patch(&part, grid, &columns, source_index, slot, &mut patches)?;
    for row in rows {
        let cells = direct_children(row, "tc");
        append_copy_patch(&part, row, &cells, source_index, slot, &mut patches)?;
    }
    patches.push(super::replace_attribute_patch(
        extents,
        "cx",
        width.to_string(),
    ));
    let edited = apply_patches(&part, patches)?;
    package.set_part(&part_name, edited)?;
    Ok(format!("{}/col[{}]", source.table, slot + 1))
}

pub(super) fn swap_columns(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    require_presentation(package, "swap-table-columns")?;
    let first = parse_path(path).ok_or_else(|| node_not_found(path))?;
    let second = parse_path(with).ok_or_else(|| node_not_found(with))?;
    if first.table != second.table {
        return Err(super::editor_error(
            "use.office.mutation_parent_unsupported",
            "Presentation table-column swap requires both columns to belong to the same table.",
        ));
    }
    let part_name = slide_part(package, &first.table)?;
    let part = package.xml_part(&part_name)?;
    let xml = index_xml(&part)?;
    let frame = super::locate_path(&xml, &first.table)?;
    let table = frame
        .descendant("tbl")
        .ok_or_else(|| invalid_table(&first.table, "has no DrawingML table element"))?;
    let (_grid, columns, rows) = validated_structure(table, &first.table)?;
    let first_index = checked_index(path, first.position, columns.len())?;
    let second_index = checked_index(with, second.position, columns.len())?;
    if first_index == second_index {
        return Ok(NativeOfficeSwapResult {
            first: path.to_string(),
            second: with.to_string(),
        });
    }

    let mut patches = Vec::with_capacity(rows.len().saturating_mul(2).saturating_add(2));
    append_swap_patches(&part, &columns, first_index, second_index, &mut patches)?;
    for row in rows {
        let cells = direct_children(row, "tc");
        append_swap_patches(&part, &cells, first_index, second_index, &mut patches)?;
    }
    let edited = apply_patches(&part, patches)?;
    package.set_part(&part_name, edited)?;
    Ok(NativeOfficeSwapResult {
        first: format!("{}/col[{}]", first.table, second_index + 1),
        second: format!("{}/col[{}]", first.table, first_index + 1),
    })
}

fn validated_structure<'a>(
    table: &'a IndexedXmlElement,
    path: &str,
) -> UseResult<(
    &'a IndexedXmlElement,
    Vec<&'a IndexedXmlElement>,
    Vec<&'a IndexedXmlElement>,
)> {
    ensure_no_merges(table, "column arrangement")?;
    let grid = table
        .child("tblGrid", 1)
        .ok_or_else(|| invalid_table(path, "has no table grid"))?;
    let columns = direct_children(grid, "gridCol");
    let rows = direct_children(table, "tr");
    validate_dimensions(rows.len(), columns.len())?;
    ensure_rectangular_rows(path, &rows, columns.len())?;
    total_grid_width(path, &columns)?;
    Ok((grid, columns, rows))
}

fn append_move_patches(
    part: &LosslessXmlPart,
    parent: &IndexedXmlElement,
    elements: &[&IndexedXmlElement],
    source_index: usize,
    slot: usize,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let source = elements[source_index];
    let remaining = elements
        .iter()
        .copied()
        .enumerate()
        .filter_map(|(index, element)| (index != source_index).then_some(element))
        .collect::<Vec<_>>();
    let insertion = remaining.get(slot).map_or_else(
        || ordered_child_insertion(parent),
        |element| element.full_range.start,
    );
    patches.push(XmlPatch::new(source.full_range.clone(), Vec::new()));
    patches.push(XmlPatch::new(
        insertion..insertion,
        element_fragment(part, source)?.to_vec(),
    ));
    Ok(())
}

fn append_copy_patch(
    part: &LosslessXmlPart,
    parent: &IndexedXmlElement,
    elements: &[&IndexedXmlElement],
    source_index: usize,
    slot: usize,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let insertion = elements.get(slot).map_or_else(
        || ordered_child_insertion(parent),
        |element| element.full_range.start,
    );
    patches.push(XmlPatch::new(
        insertion..insertion,
        element_fragment(part, elements[source_index])?.to_vec(),
    ));
    Ok(())
}

fn append_swap_patches(
    part: &LosslessXmlPart,
    elements: &[&IndexedXmlElement],
    first: usize,
    second: usize,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    patches.push(XmlPatch::new(
        elements[first].full_range.clone(),
        element_fragment(part, elements[second])?.to_vec(),
    ));
    patches.push(XmlPatch::new(
        elements[second].full_range.clone(),
        element_fragment(part, elements[first])?.to_vec(),
    ));
    Ok(())
}

fn resolve_slot(
    source: &ColumnPath,
    count: usize,
    source_index: usize,
    position: Option<&NativeOfficeInsertPosition>,
    operation: Operation,
) -> UseResult<usize> {
    let remaining = (0..count)
        .filter(|index| operation == Operation::Copy || *index != source_index)
        .collect::<Vec<_>>();
    match position {
        Some(NativeOfficeInsertPosition::Index { index }) => {
            if *index > remaining.len() {
                return Err(super::editor_error(
                    "use.office.position_invalid",
                    format!(
                        "Presentation table-column insertion index {index} is outside 0-{}.",
                        remaining.len()
                    ),
                ));
            }
            Ok(*index)
        }
        Some(NativeOfficeInsertPosition::Before { path }) => anchor_slot(
            source,
            count,
            source_index,
            &remaining,
            path,
            false,
            operation,
        ),
        Some(NativeOfficeInsertPosition::After { path }) => anchor_slot(
            source,
            count,
            source_index,
            &remaining,
            path,
            true,
            operation,
        ),
        None if operation == Operation::Copy => Ok(source_index + 1),
        None => Ok(remaining.len()),
    }
}

fn anchor_slot(
    source: &ColumnPath,
    count: usize,
    source_index: usize,
    remaining: &[usize],
    anchor: &str,
    after: bool,
    operation: Operation,
) -> UseResult<usize> {
    let anchor_path = parse_path(anchor).ok_or_else(|| node_not_found(anchor))?;
    if anchor_path.table != source.table {
        return Err(super::editor_error(
            "use.office.mutation_parent_unsupported",
            "Presentation table-column placement anchors must belong to the source table.",
        ));
    }
    let anchor_index = checked_index(anchor, anchor_path.position, count)?;
    if operation == Operation::Move && anchor_index == source_index {
        return Ok(source_index.min(remaining.len()));
    }
    let slot = remaining
        .iter()
        .position(|index| *index == anchor_index)
        .ok_or_else(|| node_not_found(anchor))?;
    Ok(slot + usize::from(after))
}

fn require_same_table(source: &ColumnPath, target_parent: Option<&str>) -> UseResult<()> {
    if target_parent.is_none_or(|parent| parent == source.table) {
        Ok(())
    } else {
        Err(super::editor_error(
            "use.office.mutation_parent_unsupported",
            format!(
                "Native Presentation table-column move/copy currently requires the source table '{}'.",
                source.table
            ),
        ))
    }
}

fn checked_index(path: &str, position: usize, count: usize) -> UseResult<usize> {
    position
        .checked_sub(1)
        .filter(|index| *index < count)
        .ok_or_else(|| node_not_found(path))
}

fn slide_part(package: &NativeOfficePackage, table_path: &str) -> UseResult<String> {
    let snapshot = crate::semantic::NativeOfficeDocument::from_package(package.clone())?;
    snapshot.get(table_path, 0)?;
    slide_part_for_path(&snapshot, table_path)
}

fn parse_path(path: &str) -> Option<ColumnPath> {
    let (table, position) = path.rsplit_once("/col[")?;
    if !table.starts_with("/slide[") || !table.contains("/table[") {
        return None;
    }
    let position = position
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|position| *position > 0)?;
    Some(ColumnPath {
        table: table.to_string(),
        position,
    })
}
