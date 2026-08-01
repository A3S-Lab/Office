use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode};
use crate::xml_tree::XmlElement;

const MAX_LOGICAL_COLUMNS: usize = 1_000_000;
const MAX_TABLE_GRID_AREA: usize = 1_000_000;
const WORD_TABLE_GRID_ERROR: &str = "use.office.word_table_grid_invalid";
const PRESENTATION_TABLE_GRID_ERROR: &str = "use.office.presentation_table_grid_invalid";

#[derive(Debug)]
struct CellAnnotation {
    row: usize,
    column: usize,
    row_span: usize,
    column_span: usize,
    merge_anchor: bool,
    merge_anchor_path: Option<String>,
}

#[derive(Debug, Clone)]
struct WordActiveMerge {
    anchor_row: usize,
    anchor_cell: usize,
    anchor_path: String,
    column_span: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WordVerticalMerge {
    None,
    Restart,
    Continue,
}

pub(super) fn annotate_word_table(table: &XmlElement, node: &mut DocumentNode) -> UseResult<()> {
    let rows = table.children_named("tr").collect::<Vec<_>>();
    ensure_semantic_shape(node, &rows, WORD_TABLE_GRID_ERROR)?;

    let declared_columns = table
        .child("tblGrid")
        .map(|grid| grid.children_named("gridCol").count())
        .filter(|count| *count > 0);
    if declared_columns.is_some_and(|count| count > MAX_LOGICAL_COLUMNS) {
        return Err(word_grid_error(format!(
            "Word table '{}' exceeds the supported logical column limit of {MAX_LOGICAL_COLUMNS}.",
            node.path
        )));
    }

    let mut annotations: Vec<Vec<CellAnnotation>> = Vec::with_capacity(rows.len());
    let mut active_merges = BTreeMap::<usize, WordActiveMerge>::new();

    for (row_offset, row) in rows.iter().enumerate() {
        let row_number = row_offset + 1;
        let row_path = &node.children[row_offset].path;
        let cells = row.children_named("tc").collect::<Vec<_>>();
        let grid_before = word_grid_offset(row, "gridBefore", row_path)?;
        let grid_after = word_grid_offset(row, "gridAfter", row_path)?;
        let mut column = grid_before.checked_add(1).ok_or_else(|| {
            word_grid_error(format!(
                "Word table row '{row_path}' has an overflowing gridBefore value."
            ))
        })?;
        require_column_bound(column.saturating_sub(1), WORD_TABLE_GRID_ERROR, row_path)?;

        let mut row_annotations = Vec::with_capacity(cells.len());
        let mut next_active_merges = BTreeMap::new();

        for (cell_offset, cell) in cells.iter().enumerate() {
            let cell_path = &node.children[row_offset].children[cell_offset].path;
            let column_span = word_grid_span(cell, cell_path)?;
            let end_column =
                checked_column_end(column, column_span, WORD_TABLE_GRID_ERROR, cell_path)?;
            if let Some(count) = declared_columns {
                if end_column > count {
                    return Err(word_grid_error(format!(
                        "Word table cell '{cell_path}' ends at logical column {end_column}, beyond the declared {count}-column grid."
                    )));
                }
            }

            let vertical_merge = word_vertical_merge(cell, cell_path)?;
            let annotation = match vertical_merge {
                WordVerticalMerge::None => CellAnnotation {
                    row: row_number,
                    column,
                    row_span: 1,
                    column_span,
                    merge_anchor: true,
                    merge_anchor_path: None,
                },
                WordVerticalMerge::Restart => {
                    let active = WordActiveMerge {
                        anchor_row: row_offset,
                        anchor_cell: cell_offset,
                        anchor_path: cell_path.clone(),
                        column_span,
                    };
                    next_active_merges.insert(column, active);
                    CellAnnotation {
                        row: row_number,
                        column,
                        row_span: 1,
                        column_span,
                        merge_anchor: true,
                        merge_anchor_path: None,
                    }
                }
                WordVerticalMerge::Continue => {
                    let active = active_merges.get(&column).ok_or_else(|| {
                        word_grid_error(format!(
                            "Word table cell '{cell_path}' continues a vertical merge with no anchor in the preceding row."
                        ))
                    })?;
                    if active.column_span != column_span {
                        return Err(word_grid_error(format!(
                            "Word table cell '{cell_path}' continues merge anchor '{}' with column span {}, but declares column span {column_span}.",
                            active.anchor_path, active.column_span
                        )));
                    }
                    let anchor = annotations
                        .get_mut(active.anchor_row)
                        .and_then(|row| row.get_mut(active.anchor_cell))
                        .ok_or_else(|| {
                            word_grid_error(format!(
                                "Word table cell '{cell_path}' references an unavailable merge anchor '{}'.",
                                active.anchor_path
                            ))
                        })?;
                    anchor.row_span = anchor.row_span.checked_add(1).ok_or_else(|| {
                        word_grid_error(format!(
                            "Word table merge anchor '{}' has an overflowing row span.",
                            active.anchor_path
                        ))
                    })?;
                    next_active_merges.insert(column, active.clone());
                    CellAnnotation {
                        row: row_number,
                        column,
                        row_span: 1,
                        column_span,
                        merge_anchor: false,
                        merge_anchor_path: Some(active.anchor_path.clone()),
                    }
                }
            };
            row_annotations.push(annotation);
            column = end_column.checked_add(1).ok_or_else(|| {
                word_grid_error(format!(
                    "Word table cell '{cell_path}' has an overflowing logical column extent."
                ))
            })?;
        }

        let row_width = column
            .checked_sub(1)
            .and_then(|used| used.checked_add(grid_after))
            .ok_or_else(|| {
                word_grid_error(format!(
                    "Word table row '{row_path}' has an overflowing gridAfter value."
                ))
            })?;
        require_column_bound(row_width, WORD_TABLE_GRID_ERROR, row_path)?;
        if let Some(count) = declared_columns {
            if row_width > count {
                return Err(word_grid_error(format!(
                    "Word table row '{row_path}' occupies {row_width} logical columns, beyond its declared {count}-column grid."
                )));
            }
        }
        annotations.push(row_annotations);
        active_merges = next_active_merges;
    }

    apply_annotations(node, annotations, WORD_TABLE_GRID_ERROR)?;
    Ok(())
}

#[derive(Debug)]
struct PresentationCell {
    row_span: usize,
    column_span: usize,
    horizontal_merge: bool,
    vertical_merge: bool,
}

impl PresentationCell {
    fn is_covered(&self) -> bool {
        self.horizontal_merge || self.vertical_merge
    }
}

#[derive(Debug)]
struct PresentationAnchor {
    row: usize,
    column: usize,
    path: String,
}

pub(super) fn annotate_presentation_table(
    table: &XmlElement,
    node: &mut DocumentNode,
) -> UseResult<()> {
    let rows = table.children_named("tr").collect::<Vec<_>>();
    ensure_semantic_shape(node, &rows, PRESENTATION_TABLE_GRID_ERROR)?;

    let mut parsed_rows = Vec::with_capacity(rows.len());
    for (row_offset, row) in rows.iter().enumerate() {
        let mut parsed_cells = Vec::new();
        for (cell_offset, cell) in row.children_named("tc").enumerate() {
            let path = &node.children[row_offset].children[cell_offset].path;
            parsed_cells.push(PresentationCell {
                row_span: positive_attribute(
                    cell,
                    "rowSpan",
                    1,
                    PRESENTATION_TABLE_GRID_ERROR,
                    path,
                )?,
                column_span: positive_attribute(
                    cell,
                    "gridSpan",
                    1,
                    PRESENTATION_TABLE_GRID_ERROR,
                    path,
                )?,
                horizontal_merge: on_off_attribute(
                    cell,
                    "hMerge",
                    PRESENTATION_TABLE_GRID_ERROR,
                    path,
                )?,
                vertical_merge: on_off_attribute(
                    cell,
                    "vMerge",
                    PRESENTATION_TABLE_GRID_ERROR,
                    path,
                )?,
            });
        }
        parsed_rows.push(parsed_cells);
    }

    let declared_columns = table
        .child("tblGrid")
        .map(|grid| grid.children_named("gridCol").count())
        .filter(|count| *count > 0);
    let columns = match declared_columns {
        Some(columns) => columns,
        None => infer_presentation_columns(&parsed_rows)?,
    };
    require_column_bound(columns, PRESENTATION_TABLE_GRID_ERROR, &node.path)?;
    let area = rows.len().checked_mul(columns).ok_or_else(|| {
        presentation_grid_error(format!(
            "Presentation table '{}' has an overflowing logical grid.",
            node.path
        ))
    })?;
    if area > MAX_TABLE_GRID_AREA {
        return Err(presentation_grid_error(format!(
            "Presentation table '{}' logical grid contains {area} cells, exceeding the supported limit of {MAX_TABLE_GRID_AREA}.",
            node.path
        )));
    }

    let mut annotations = Vec::with_capacity(rows.len());
    let mut anchors = Vec::<PresentationAnchor>::new();
    let mut occupancy = BTreeMap::<(usize, usize), usize>::new();

    for (row_offset, cells) in parsed_rows.iter().enumerate() {
        let row_number = row_offset + 1;
        let uses_explicit_covered_cells = cells.iter().any(PresentationCell::is_covered);
        let mut compact_column = 1_usize;
        let mut row_annotations = Vec::with_capacity(cells.len());

        for (cell_offset, cell) in cells.iter().enumerate() {
            let cell_path = &node.children[row_offset].children[cell_offset].path;
            let column = if uses_explicit_covered_cells {
                cell_offset + 1
            } else {
                while compact_column <= columns
                    && occupancy.contains_key(&(row_number, compact_column))
                {
                    compact_column += 1;
                }
                compact_column
            };
            if column == 0 || column > columns {
                return Err(presentation_grid_error(format!(
                    "Presentation table cell '{cell_path}' has logical column {column}, beyond the {columns}-column grid."
                )));
            }

            if cell.is_covered() {
                if cell.row_span != 1 || cell.column_span != 1 {
                    return Err(presentation_grid_error(format!(
                        "Covered presentation table cell '{cell_path}' cannot declare rowSpan or gridSpan."
                    )));
                }
                let anchor_index = occupancy
                    .get(&(row_number, column))
                    .copied()
                    .ok_or_else(|| {
                        presentation_grid_error(format!(
                            "Presentation table cell '{cell_path}' is marked as covered but has no merge anchor."
                        ))
                    })?;
                let anchor = &anchors[anchor_index];
                let expected_horizontal = column > anchor.column;
                let expected_vertical = row_number > anchor.row;
                if cell.horizontal_merge != expected_horizontal
                    || cell.vertical_merge != expected_vertical
                {
                    return Err(presentation_grid_error(format!(
                        "Presentation table cell '{cell_path}' has merge flags hMerge={} and vMerge={}, inconsistent with anchor '{}'.",
                        cell.horizontal_merge, cell.vertical_merge, anchor.path
                    )));
                }
                row_annotations.push(CellAnnotation {
                    row: row_number,
                    column,
                    row_span: 1,
                    column_span: 1,
                    merge_anchor: false,
                    merge_anchor_path: Some(anchor.path.clone()),
                });
                if !uses_explicit_covered_cells {
                    compact_column = column + 1;
                }
                continue;
            }

            let end_row = checked_extent(
                row_number,
                cell.row_span,
                PRESENTATION_TABLE_GRID_ERROR,
                cell_path,
                "row",
            )?;
            let end_column = checked_column_end(
                column,
                cell.column_span,
                PRESENTATION_TABLE_GRID_ERROR,
                cell_path,
            )?;
            if end_row > rows.len() || end_column > columns {
                return Err(presentation_grid_error(format!(
                    "Presentation table cell '{cell_path}' spans through row {end_row}, column {end_column}, beyond the {rows_len}x{columns} logical grid.",
                    rows_len = rows.len()
                )));
            }
            for occupied_row in row_number..=end_row {
                for occupied_column in column..=end_column {
                    if let Some(existing) = occupancy.get(&(occupied_row, occupied_column)) {
                        return Err(presentation_grid_error(format!(
                            "Presentation table cell '{cell_path}' overlaps merge anchor '{}' at row {occupied_row}, column {occupied_column}.",
                            anchors[*existing].path
                        )));
                    }
                }
            }

            let anchor_index = anchors.len();
            anchors.push(PresentationAnchor {
                row: row_number,
                column,
                path: cell_path.clone(),
            });
            for occupied_row in row_number..=end_row {
                for occupied_column in column..=end_column {
                    occupancy.insert((occupied_row, occupied_column), anchor_index);
                }
            }
            row_annotations.push(CellAnnotation {
                row: row_number,
                column,
                row_span: cell.row_span,
                column_span: cell.column_span,
                merge_anchor: true,
                merge_anchor_path: None,
            });
            if !uses_explicit_covered_cells {
                compact_column = end_column.checked_add(1).ok_or_else(|| {
                    presentation_grid_error(format!(
                        "Presentation table cell '{cell_path}' has an overflowing logical column extent."
                    ))
                })?;
            }
        }
        annotations.push(row_annotations);
    }

    apply_annotations(node, annotations, PRESENTATION_TABLE_GRID_ERROR)?;
    node.format
        .entry("columns".into())
        .or_insert_with(|| columns.to_string());
    Ok(())
}

fn infer_presentation_columns(rows: &[Vec<PresentationCell>]) -> UseResult<usize> {
    let mut columns = 0_usize;
    for cells in rows {
        let row_columns = if cells.iter().any(PresentationCell::is_covered) {
            cells
                .iter()
                .enumerate()
                .try_fold(cells.len(), |width, (offset, cell)| {
                    let end = checked_column_end(
                        offset + 1,
                        cell.column_span,
                        PRESENTATION_TABLE_GRID_ERROR,
                        "inferred presentation table cell",
                    )?;
                    Ok::<_, a3s_use_core::UseError>(width.max(end))
                })?
        } else {
            cells.iter().try_fold(0_usize, |width, cell| {
                width.checked_add(cell.column_span).ok_or_else(|| {
                    presentation_grid_error(
                        "Presentation table has an overflowing inferred column count.",
                    )
                })
            })?
        };
        require_column_bound(
            row_columns,
            PRESENTATION_TABLE_GRID_ERROR,
            "presentation table",
        )?;
        columns = columns.max(row_columns);
    }
    Ok(columns)
}

fn ensure_semantic_shape(
    node: &DocumentNode,
    rows: &[&XmlElement],
    error_code: &str,
) -> UseResult<()> {
    if node.children.len() != rows.len() {
        return Err(semantic_error(
            error_code,
            format!(
                "Table '{}' semantic row count does not match its source XML.",
                node.path
            ),
        ));
    }
    for (offset, row) in rows.iter().enumerate() {
        let cell_count = row.children_named("tc").count();
        if node.children[offset].children.len() != cell_count {
            return Err(semantic_error(
                error_code,
                format!(
                    "Table row '{}' semantic cell count does not match its source XML.",
                    node.children[offset].path
                ),
            ));
        }
    }
    Ok(())
}

fn apply_annotations(
    node: &mut DocumentNode,
    annotations: Vec<Vec<CellAnnotation>>,
    error_code: &str,
) -> UseResult<()> {
    if node.children.len() != annotations.len() {
        return Err(semantic_error(
            error_code,
            format!("Table '{}' has incomplete grid annotations.", node.path),
        ));
    }
    for (row, row_annotations) in node.children.iter_mut().zip(annotations) {
        if row.children.len() != row_annotations.len() {
            return Err(semantic_error(
                error_code,
                format!("Table row '{}' has incomplete grid annotations.", row.path),
            ));
        }
        for (cell, annotation) in row.children.iter_mut().zip(row_annotations) {
            cell.format.insert("row".into(), annotation.row.to_string());
            cell.format
                .insert("column".into(), annotation.column.to_string());
            cell.format
                .insert("rowSpan".into(), annotation.row_span.to_string());
            cell.format
                .insert("columnSpan".into(), annotation.column_span.to_string());
            cell.format
                .insert("mergeAnchor".into(), annotation.merge_anchor.to_string());
            if let Some(anchor_path) = annotation.merge_anchor_path {
                cell.format.insert("mergeAnchorPath".into(), anchor_path);
            } else {
                cell.format.remove("mergeAnchorPath");
            }
        }
    }
    Ok(())
}

fn word_grid_offset(row: &XmlElement, name: &str, path: &str) -> UseResult<usize> {
    let Some(offset) = row
        .child("trPr")
        .and_then(|properties| properties.child(name))
    else {
        return Ok(0);
    };
    let value = offset.attribute("val").ok_or_else(|| {
        word_grid_error(format!(
            "Word table row '{path}' has {name} without a val attribute."
        ))
    })?;
    let value = value.parse::<usize>().map_err(|_| {
        word_grid_error(format!(
            "Word table row '{path}' has invalid {name} value '{value}'."
        ))
    })?;
    require_column_bound(value, WORD_TABLE_GRID_ERROR, path)?;
    Ok(value)
}

fn word_grid_span(cell: &XmlElement, path: &str) -> UseResult<usize> {
    let Some(span) = cell
        .child("tcPr")
        .and_then(|properties| properties.child("gridSpan"))
    else {
        return Ok(1);
    };
    positive_attribute(span, "val", 1, WORD_TABLE_GRID_ERROR, path)
}

fn word_vertical_merge(cell: &XmlElement, path: &str) -> UseResult<WordVerticalMerge> {
    let Some(merge) = cell
        .child("tcPr")
        .and_then(|properties| properties.child("vMerge"))
    else {
        return Ok(WordVerticalMerge::None);
    };
    match merge.attribute("val") {
        None | Some("continue") => Ok(WordVerticalMerge::Continue),
        Some("restart") => Ok(WordVerticalMerge::Restart),
        Some(value) => Err(word_grid_error(format!(
            "Word table cell '{path}' has invalid vMerge value '{value}'."
        ))),
    }
}

fn positive_attribute(
    element: &XmlElement,
    attribute: &str,
    default: usize,
    error_code: &str,
    path: &str,
) -> UseResult<usize> {
    let Some(value) = element.attribute(attribute) else {
        return Ok(default);
    };
    value
        .parse::<usize>()
        .ok()
        .filter(|value| *value > 0)
        .ok_or_else(|| {
            semantic_error(
                error_code,
                format!("Table cell '{path}' has invalid {attribute} value '{value}'."),
            )
        })
}

fn on_off_attribute(
    element: &XmlElement,
    attribute: &str,
    error_code: &str,
    path: &str,
) -> UseResult<bool> {
    match element.attribute(attribute) {
        None | Some("0" | "false" | "off") => Ok(false),
        Some("1" | "true" | "on") => Ok(true),
        Some(value) => Err(semantic_error(
            error_code,
            format!("Table cell '{path}' has invalid {attribute} value '{value}'."),
        )),
    }
}

fn checked_column_end(start: usize, span: usize, error_code: &str, path: &str) -> UseResult<usize> {
    let end = checked_extent(start, span, error_code, path, "column")?;
    require_column_bound(end, error_code, path)?;
    Ok(end)
}

fn checked_extent(
    start: usize,
    span: usize,
    error_code: &str,
    path: &str,
    axis: &str,
) -> UseResult<usize> {
    start.checked_add(span.saturating_sub(1)).ok_or_else(|| {
        semantic_error(
            error_code,
            format!("Table cell '{path}' has an overflowing {axis} span."),
        )
    })
}

fn require_column_bound(value: usize, error_code: &str, path: &str) -> UseResult<()> {
    if value > MAX_LOGICAL_COLUMNS {
        return Err(semantic_error(
            error_code,
            format!(
                "Table '{path}' logical column {value} exceeds the supported limit of {MAX_LOGICAL_COLUMNS}."
            ),
        ));
    }
    Ok(())
}

fn word_grid_error(message: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(WORD_TABLE_GRID_ERROR, message)
}

fn presentation_grid_error(message: impl Into<String>) -> a3s_use_core::UseError {
    semantic_error(PRESENTATION_TABLE_GRID_ERROR, message)
}
