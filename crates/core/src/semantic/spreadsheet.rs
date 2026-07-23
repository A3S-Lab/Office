use std::{
    cmp::Reverse,
    collections::{BTreeMap, BinaryHeap},
};

use a3s_use_core::UseResult;

use super::{semantic_error, DocumentNode, OfficeNodeType};
use crate::spreadsheet_reference::{
    column_name as a1_column_name, first_intersecting_ranges, parse_column, CellRange,
    CellReference, MAX_COLUMNS,
};
use crate::xml_tree::{parse_xml_tree, XmlElement, XmlNode};
use crate::{NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget};

mod auto_filter;
mod conditional_formatting;
mod data_validation;
mod named_range;
mod sort_state;
mod style;
mod table;
mod view;

use style::{read_differential_formats, read_styles};

const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEET_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const SPREADSHEET_DRAWING_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_MERGED_RANGES: usize = 100_000;

#[derive(Debug, Clone, Copy)]
struct MergedRange {
    range: CellRange,
}

#[derive(Debug, Clone, Copy)]
struct ObservedCell {
    reference: CellReference,
    row_index: usize,
    cell_index: usize,
}

#[derive(Debug, Clone, Copy)]
struct ValidationArea {
    range: CellRange,
    validation_index: usize,
}

pub(super) fn read(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
) -> UseResult<DocumentNode> {
    let workbook_part = package.xml_part("xl/workbook.xml")?;
    let workbook = parse_xml_tree(&workbook_part)?;
    require_spreadsheet_element(&workbook, "workbook", workbook_part.name())?;
    let shared_strings = read_shared_strings(package)?;
    let styles = read_styles(package)?;
    let differential_formats = read_differential_formats(package)?;
    let source = RelationshipSource::Part {
        part_name: "xl/workbook.xml".to_string(),
    };

    let mut root = DocumentNode::new("/", "workbook", OfficeNodeType::Workbook);
    let sheets = workbook.child("sheets").ok_or_else(|| {
        semantic_error(
            "use.office.spreadsheet_sheets_missing",
            "Spreadsheet workbook has no sheets collection.",
        )
    })?;
    let mut sheet_names = Vec::new();
    for sheet in sheets.children_named("sheet") {
        let name = sheet.attribute("name").ok_or_else(|| {
            semantic_error(
                "use.office.spreadsheet_sheet_invalid",
                "Spreadsheet sheet is missing its name.",
            )
        })?;
        validate_sheet_name(name)?;
        sheet_names.push(name.to_string());
        let relationship_id = relationship_attribute(sheet, "id").ok_or_else(|| {
            semantic_error(
                "use.office.spreadsheet_sheet_invalid",
                format!("Spreadsheet sheet '{name}' has no relationship ID."),
            )
        })?;
        let relationship = opc
            .relationships()
            .relationship(&source, relationship_id)
            .ok_or_else(|| {
                semantic_error(
                    "use.office.spreadsheet_sheet_missing",
                    format!(
                        "Spreadsheet sheet '{name}' references missing relationship '{relationship_id}'."
                    ),
                )
            })?;
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            return Err(semantic_error(
                "use.office.spreadsheet_sheet_invalid",
                format!("Spreadsheet sheet '{name}' cannot use an external relationship."),
            ));
        };
        let mut sheet_node = read_worksheet(
            package,
            opc,
            part_name,
            name,
            &shared_strings,
            &styles,
            &differential_formats,
        )?;
        if let Some(sheet_id) = sheet.attribute("sheetId") {
            sheet_node.format.insert("sheetId".into(), sheet_id.into());
        }
        if let Some(state) = sheet.attribute("state") {
            sheet_node.format.insert("state".into(), state.into());
        }
        root.children.push(sheet_node);
    }
    if let Some(collection) = named_range::read(&workbook, &sheet_names)? {
        root.children.push(collection);
    }
    root.text = root
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|sheet| sheet.text.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    Ok(root)
}

pub(super) fn virtual_get(
    root: &DocumentNode,
    path: &str,
    depth: usize,
) -> UseResult<Option<DocumentNode>> {
    if let Some(node) = named_range::virtual_get(root, path, depth)? {
        return Ok(Some(node));
    }
    let Some((requested_sheet, target)) =
        path.strip_prefix('/').and_then(|path| path.split_once('/'))
    else {
        return Ok(None);
    };
    let Some(sheet) = root.children.iter().find(|node| {
        node.node_type == OfficeNodeType::Worksheet
            && node.path[1..].eq_ignore_ascii_case(requested_sheet)
    }) else {
        return Ok(None);
    };
    if !target.contains(':') {
        if let Ok(reference) = CellReference::parse(target) {
            let merged = merged_ranges(sheet).find(|merged| merged.range.contains(reference));
            let validation = validation_areas(sheet)
                .into_iter()
                .find(|(range, _)| range.contains(reference));
            if merged.is_some() || validation.is_some() {
                let mut node = DocumentNode::new(
                    format!("{}/{}", sheet.path, reference.a1()),
                    "cell",
                    OfficeNodeType::Cell,
                );
                node.format
                    .insert("column".into(), column_name(reference.column)?);
                node.format.insert("row".into(), reference.row.to_string());
                node.format.insert("empty".into(), "true".into());
                if let Some(merged) = merged {
                    annotate_cell_merge(&mut node, reference, &merged);
                }
                if let Some((_, validation)) = validation {
                    annotate_cell_validation(&mut node, validation);
                }
                return Ok(Some(node));
            }
        }
    }
    if let Some(column) = target
        .strip_prefix("col[")
        .and_then(|value| value.strip_suffix(']'))
    {
        return virtual_column(sheet, column, depth).map(Some);
    }
    let Some((start, end)) = target.split_once(':') else {
        return Ok(None);
    };
    virtual_range(sheet, start, end, depth).map(Some)
}

fn virtual_column(sheet: &DocumentNode, requested: &str, depth: usize) -> UseResult<DocumentNode> {
    let column = if requested
        .chars()
        .all(|character| character.is_ascii_digit())
    {
        let number = requested.parse::<u32>().map_err(|error| {
            semantic_error(
                "use.office.spreadsheet_column_invalid",
                format!("Spreadsheet column '{requested}' is invalid: {error}"),
            )
        })?;
        column_name(number)?
    } else {
        let normalized = requested.to_ascii_uppercase();
        if column_number(&format!("{normalized}1")).is_none() {
            return Err(semantic_error(
                "use.office.spreadsheet_column_invalid",
                format!("Spreadsheet column '{requested}' is outside A:XFD."),
            ));
        }
        normalized
    };
    let path = format!("{}/col[{column}]", sheet.path);
    let mut node = DocumentNode::new(path, "col", OfficeNodeType::Column);
    node.format.insert("column".into(), column.clone());
    let cells = cells(sheet)
        .into_iter()
        .filter(|cell| {
            cell.format
                .get("column")
                .is_some_and(|value| value == &column)
        })
        .collect::<Vec<_>>();
    node.child_count = cells.len();
    node.text = cells
        .iter()
        .map(|cell| cell.text.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    if depth > 0 {
        node.children = cells
            .into_iter()
            .map(|cell| cell.clone_to_depth(depth - 1))
            .collect();
    }
    Ok(node)
}

fn virtual_range(
    sheet: &DocumentNode,
    start: &str,
    end: &str,
    depth: usize,
) -> UseResult<DocumentNode> {
    let (start_column, start_row, start_reference) = cell_coordinates(start)?;
    let (end_column, end_row, end_reference) = cell_coordinates(end)?;
    let min_column = start_column.min(end_column);
    let max_column = start_column.max(end_column);
    let min_row = start_row.min(end_row);
    let max_row = start_row.max(end_row);
    let path = format!("{}/{}:{}", sheet.path, start_reference, end_reference);
    let mut node = DocumentNode::new(path, "range", OfficeNodeType::Range);
    node.format.insert(
        "normalizedRef".into(),
        format!(
            "{}{}:{}{}",
            column_name(min_column)?,
            min_row,
            column_name(max_column)?,
            max_row
        ),
    );
    let requested_range = CellRange {
        start: CellReference {
            column: min_column,
            row: min_row,
        },
        end: CellReference {
            column: max_column,
            row: max_row,
        },
    };
    node.format.insert(
        "merge".into(),
        merged_ranges(sheet)
            .any(|merged| merged.range == requested_range)
            .to_string(),
    );
    if let Some((_, validation)) = validation_areas(sheet)
        .into_iter()
        .find(|(range, _)| *range == requested_range)
    {
        annotate_cell_validation(&mut node, validation);
    }
    let matching = cells(sheet)
        .into_iter()
        .filter(|cell| {
            cell.path
                .rsplit('/')
                .next()
                .and_then(|reference| cell_coordinates(reference).ok())
                .is_some_and(|(column, row, _)| {
                    (min_column..=max_column).contains(&column)
                        && (min_row..=max_row).contains(&row)
                })
        })
        .collect::<Vec<_>>();
    node.child_count = matching.len();
    node.text = matching
        .iter()
        .map(|cell| format!("{}={}", cell.path, cell.text))
        .collect::<Vec<_>>()
        .join("\n");
    if depth > 0 {
        node.children = matching
            .into_iter()
            .map(|cell| cell.clone_to_depth(depth - 1))
            .collect();
    }
    Ok(node)
}

fn cells(sheet: &DocumentNode) -> Vec<&DocumentNode> {
    sheet
        .children
        .iter()
        .flat_map(|row| row.children.iter())
        .filter(|node| node.node_type == OfficeNodeType::Cell)
        .collect()
}

fn cell_coordinates(reference: &str) -> UseResult<(u32, u32, String)> {
    let reference = CellReference::parse(reference)?;
    Ok((reference.column, reference.row, reference.a1()))
}

fn read_worksheet(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    part_name: &str,
    sheet_name: &str,
    shared_strings: &[String],
    styles: &[BTreeMap<String, String>],
    differential_formats: &[style::DifferentialFormat],
) -> UseResult<DocumentNode> {
    let part = package.xml_part(part_name)?;
    let worksheet = parse_xml_tree(&part)?;
    require_spreadsheet_element(&worksheet, "worksheet", part.name())?;
    let sheet_path = format!("/{sheet_name}");
    let mut sheet_node = DocumentNode::new(&sheet_path, "sheet", OfficeNodeType::Worksheet);
    sheet_node.format.insert("part".into(), part_name.into());
    let merged = read_merged_ranges(&worksheet, part_name)?;
    if !merged.is_empty() {
        sheet_node
            .format
            .insert("mergeCount".into(), merged.len().to_string());
    }
    let mut observed_cells = Vec::new();
    if let Some(sheet_data) = worksheet.child("sheetData") {
        let mut inferred_row = 0_u32;
        for row in sheet_data.children_named("row") {
            let row_number = row
                .attribute("r")
                .and_then(|value| value.parse::<u32>().ok())
                .unwrap_or_else(|| inferred_row.saturating_add(1));
            if row_number == 0 {
                return Err(semantic_error(
                    "use.office.spreadsheet_row_invalid",
                    format!("Worksheet '{sheet_name}' contains row zero."),
                ));
            }
            inferred_row = row_number;
            let row_path = format!("{sheet_path}/row[{row_number}]");
            let mut row_node = DocumentNode::new(&row_path, "row", OfficeNodeType::Row);
            copy_attribute(row, "ht", "height", &mut row_node);
            copy_attribute(row, "hidden", "hidden", &mut row_node);
            copy_attribute(row, "outlineLevel", "outlineLevel", &mut row_node);
            let mut inferred_column = 0_u32;
            for cell in row.children_named("c") {
                let reference = match cell.attribute("r") {
                    Some(reference) => normalize_cell_reference(reference, row_number)?,
                    None => {
                        inferred_column = inferred_column.saturating_add(1);
                        format!("{}{row_number}", column_name(inferred_column)?)
                    }
                };
                inferred_column = column_number(&reference).unwrap_or(inferred_column);
                let cell_node = read_cell(cell, &sheet_path, &reference, shared_strings, styles)?;
                let cell_reference = CellReference::parse(&reference)?;
                observed_cells.push(ObservedCell {
                    reference: cell_reference,
                    row_index: sheet_node.children.len(),
                    cell_index: row_node.children.len(),
                });
                row_node.children.push(cell_node);
            }
            row_node.text = row_node
                .children
                .iter()
                .map(|cell| cell.text.as_str())
                .collect::<Vec<_>>()
                .join("\t");
            sheet_node.children.push(row_node);
        }
    }
    annotate_merged_cells(&mut sheet_node, &observed_cells, &merged);
    for (offset, merged) in merged.iter().enumerate() {
        let mut node = DocumentNode::new(
            format!("{sheet_path}/mergeCell[{}]", offset + 1),
            "mergeCell",
            OfficeNodeType::Range,
        );
        node.text = merged.range.a1();
        node.format.insert("ref".into(), merged.range.a1());
        node.format.insert("merge".into(), "true".into());
        sheet_node.children.push(node);
    }
    let conditional_formats =
        conditional_formatting::read(&worksheet, part_name, &sheet_path, differential_formats)?;
    if !conditional_formats.is_empty() {
        sheet_node.format.insert(
            "conditionalFormatCount".into(),
            conditional_formats.len().to_string(),
        );
        sheet_node.children.extend(conditional_formats);
    }
    let validations = data_validation::read(&worksheet, part_name, &sheet_path)?;
    if !validations.is_empty() {
        annotate_data_validations(&mut sheet_node, &observed_cells, &validations);
        sheet_node
            .format
            .insert("dataValidationCount".into(), validations.len().to_string());
        sheet_node.children.extend(validations);
    }
    if let Some(filter) = auto_filter::read(&worksheet, part_name, &sheet_path)? {
        sheet_node
            .format
            .insert("autoFilterRef".into(), filter.format["ref"].clone());
        sheet_node.children.push(filter);
    }
    if let Some(sort) = sort_state::read(&worksheet, part_name, &sheet_path)? {
        sheet_node.children.push(sort);
    }
    if let Some(freeze) = view::read(&worksheet, &sheet_path) {
        sheet_node.children.push(freeze);
    }
    let tables = table::read(package, opc, &worksheet, part_name, &sheet_path)?;
    if !tables.is_empty() {
        sheet_node
            .format
            .insert("tableCount".into(), tables.len().to_string());
        sheet_node.children.extend(tables);
    }
    append_worksheet_hyperlinks(opc, &worksheet, part_name, &sheet_path, &mut sheet_node)?;
    append_worksheet_comments(package, opc, part_name, &sheet_path, &mut sheet_node)?;
    append_worksheet_pictures(
        package,
        opc,
        &worksheet,
        part_name,
        &sheet_path,
        &mut sheet_node,
    )?;
    sheet_node.text = sheet_node
        .children
        .iter()
        .filter(|child| child.node_type == OfficeNodeType::Row)
        .map(|row| row.text.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    Ok(sheet_node)
}

fn read_merged_ranges(worksheet: &XmlElement, part_name: &str) -> UseResult<Vec<MergedRange>> {
    let containers = worksheet
        .child_elements()
        .filter(|element| {
            element.local_name == "mergeCells" && element.namespace == worksheet.namespace
        })
        .collect::<Vec<_>>();
    if containers.len() > 1 {
        return Err(semantic_error(
            "use.office.spreadsheet_merge_invalid",
            format!("Worksheet part '{part_name}' contains multiple mergeCells collections."),
        ));
    }
    let Some(container) = containers.first() else {
        return Ok(Vec::new());
    };
    let elements = container
        .child_elements()
        .filter(|element| {
            element.local_name == "mergeCell" && element.namespace == worksheet.namespace
        })
        .collect::<Vec<_>>();
    if elements.len() > MAX_MERGED_RANGES {
        return Err(semantic_error(
            "use.office.spreadsheet_merge_limit",
            format!(
                "Worksheet part '{part_name}' contains {} merged ranges; the limit is {MAX_MERGED_RANGES}.",
                elements.len()
            ),
        ));
    }
    let mut merged = Vec::with_capacity(elements.len());
    for element in elements {
        let reference = element.attribute("ref").ok_or_else(|| {
            semantic_error(
                "use.office.spreadsheet_merge_invalid",
                format!("Worksheet part '{part_name}' contains a mergeCell without ref."),
            )
        })?;
        let range = CellRange::parse(reference).map_err(|error| {
            semantic_error(
                "use.office.spreadsheet_merge_invalid",
                format!(
                    "Worksheet part '{part_name}' contains invalid merged range '{reference}': {error}"
                ),
            )
        })?;
        merged.push(MergedRange { range });
    }
    let ranges = merged.iter().map(|merged| merged.range).collect::<Vec<_>>();
    if let Some((left, right)) = first_intersecting_ranges(&ranges) {
        return Err(semantic_error(
            "use.office.spreadsheet_merge_invalid",
            format!(
                "Worksheet part '{part_name}' contains overlapping merged ranges '{}' and '{}'.",
                ranges[left].a1(),
                ranges[right].a1()
            ),
        ));
    }
    Ok(merged)
}

fn annotate_merged_cells(
    sheet: &mut DocumentNode,
    observed_cells: &[ObservedCell],
    merged: &[MergedRange],
) {
    let mut cells = (0..observed_cells.len()).collect::<Vec<_>>();
    cells.sort_unstable_by_key(|index| {
        let reference = observed_cells[*index].reference;
        (reference.row, reference.column)
    });
    let mut ranges = (0..merged.len()).collect::<Vec<_>>();
    ranges.sort_unstable_by_key(|index| {
        let range = merged[*index].range;
        (
            range.start.row,
            range.start.column,
            range.end.row,
            range.end.column,
        )
    });

    let mut next_range = 0_usize;
    let mut active_by_column = BTreeMap::<u32, usize>::new();
    let mut expiration = BinaryHeap::<Reverse<(u32, usize)>>::new();
    for cell_index in cells {
        let observed = observed_cells[cell_index];
        while let Some(Reverse((end_row, expired))) = expiration.peek().copied() {
            if end_row >= observed.reference.row {
                break;
            }
            expiration.pop();
            let start_column = merged[expired].range.start.column;
            if active_by_column.get(&start_column).copied() == Some(expired) {
                active_by_column.remove(&start_column);
            }
        }
        while next_range < ranges.len()
            && merged[ranges[next_range]].range.start.row <= observed.reference.row
        {
            let index = ranges[next_range];
            let range = merged[index].range;
            if range.end.row >= observed.reference.row {
                active_by_column.insert(range.start.column, index);
                expiration.push(Reverse((range.end.row, index)));
            }
            next_range += 1;
        }

        let Some((_, merge_index)) = active_by_column
            .range(..=observed.reference.column)
            .next_back()
        else {
            continue;
        };
        let merged = &merged[*merge_index];
        if merged.range.end.column < observed.reference.column {
            continue;
        }
        let node = &mut sheet.children[observed.row_index].children[observed.cell_index];
        annotate_cell_merge(node, observed.reference, merged);
    }
}

fn annotate_cell_merge(node: &mut DocumentNode, reference: CellReference, merged: &MergedRange) {
    node.format.insert("merge".into(), merged.range.a1());
    node.format.insert(
        "mergeAnchor".into(),
        (reference == merged.range.start).to_string(),
    );
}

fn merged_ranges(sheet: &DocumentNode) -> impl Iterator<Item = MergedRange> + '_ {
    sheet
        .children
        .iter()
        .filter(|node| node.tag == "mergeCell" && node.node_type == OfficeNodeType::Range)
        .filter_map(|node| node.format.get("ref"))
        .filter_map(|reference| CellRange::parse(reference).ok())
        .map(|range| MergedRange { range })
}

fn annotate_data_validations(
    sheet: &mut DocumentNode,
    observed_cells: &[ObservedCell],
    validations: &[DocumentNode],
) {
    let mut areas = validations
        .iter()
        .enumerate()
        .flat_map(|(validation_index, validation)| {
            validation
                .format
                .get("ref")
                .into_iter()
                .flat_map(|reference| reference.split_ascii_whitespace())
                .filter_map(|reference| CellRange::parse(reference).ok())
                .map(move |range| ValidationArea {
                    range,
                    validation_index,
                })
        })
        .collect::<Vec<_>>();
    areas.sort_unstable_by_key(|area| {
        (
            area.range.start.row,
            area.range.start.column,
            area.range.end.row,
            area.range.end.column,
        )
    });
    let mut cells = (0..observed_cells.len()).collect::<Vec<_>>();
    cells.sort_unstable_by_key(|index| {
        let reference = observed_cells[*index].reference;
        (reference.row, reference.column)
    });

    let mut next_area = 0_usize;
    let mut active_by_column = BTreeMap::<u32, usize>::new();
    let mut expiration = BinaryHeap::<Reverse<(u32, usize)>>::new();
    for cell_index in cells {
        let observed = observed_cells[cell_index];
        while let Some(Reverse((end_row, expired))) = expiration.peek().copied() {
            if end_row >= observed.reference.row {
                break;
            }
            expiration.pop();
            let start_column = areas[expired].range.start.column;
            if active_by_column.get(&start_column).copied() == Some(expired) {
                active_by_column.remove(&start_column);
            }
        }
        while next_area < areas.len() && areas[next_area].range.start.row <= observed.reference.row
        {
            let area = areas[next_area];
            if area.range.end.row >= observed.reference.row {
                active_by_column.insert(area.range.start.column, next_area);
                expiration.push(Reverse((area.range.end.row, next_area)));
            }
            next_area += 1;
        }
        let Some((_, area_index)) = active_by_column
            .range(..=observed.reference.column)
            .next_back()
        else {
            continue;
        };
        let area = areas[*area_index];
        if area.range.end.column < observed.reference.column {
            continue;
        }
        let node = &mut sheet.children[observed.row_index].children[observed.cell_index];
        annotate_cell_validation(node, &validations[area.validation_index]);
    }
}

fn validation_areas(sheet: &DocumentNode) -> Vec<(CellRange, &DocumentNode)> {
    sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::DataValidation)
        .flat_map(|validation| {
            validation
                .format
                .get("ref")
                .into_iter()
                .flat_map(|reference| reference.split_ascii_whitespace())
                .filter_map(|reference| CellRange::parse(reference).ok())
                .map(move |range| (range, validation))
        })
        .collect()
}

fn annotate_cell_validation(node: &mut DocumentNode, validation: &DocumentNode) {
    node.format
        .insert("dataValidation".into(), validation.path.clone());
    if let Some(validation_type) = validation.format.get("type") {
        node.format
            .insert("validationType".into(), validation_type.clone());
    }
}

fn append_worksheet_comments(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    worksheet_part: &str,
    sheet_path: &str,
    sheet: &mut DocumentNode,
) -> UseResult<()> {
    let source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    let Some(relationship) = opc
        .relationships()
        .relationships_from(&source)
        .iter()
        .find(|relationship| relationship.relationship_type.ends_with("/comments"))
    else {
        return Ok(());
    };
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        return Err(semantic_error(
            "use.office.comment_relationship_invalid",
            format!("Worksheet '{sheet_path}' comments relationship must be internal."),
        ));
    };
    let part = package.xml_part(part_name)?;
    let comments = parse_xml_tree(&part)?;
    require_spreadsheet_element(&comments, "comments", part.name())?;
    let authors = comments
        .child("authors")
        .map(|authors| {
            authors
                .children_named("author")
                .map(direct_text)
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let Some(list) = comments.child("commentList") else {
        return Ok(());
    };
    for comment in list.children_named("comment") {
        let reference = comment.attribute("ref").ok_or_else(|| {
            semantic_error(
                "use.office.comment_reference_missing",
                format!("Worksheet '{sheet_path}' contains a comment without a cell reference."),
            )
        })?;
        let reference = CellReference::parse(reference)?.a1();
        let mut node = DocumentNode::new(
            format!("{sheet_path}/{reference}/comment"),
            "comment",
            OfficeNodeType::Comment,
        );
        node.text = comment
            .child("text")
            .map(spreadsheet_text)
            .unwrap_or_default();
        node.format.insert("ref".into(), reference.clone());
        node.format.insert("part".into(), part_name.clone());
        node.format
            .insert("ownerPart".into(), worksheet_part.to_string());
        if let Some(author_id) = comment.attribute("authorId") {
            node.format.insert("authorId".into(), author_id.into());
            if let Some(author) = author_id
                .parse::<usize>()
                .ok()
                .and_then(|author_id| authors.get(author_id))
            {
                node.format.insert("author".into(), author.clone());
            }
        }
        if let Some(cell) = find_cell_mut(sheet, &reference) {
            cell.children.push(node);
        } else {
            sheet.children.push(node);
        }
    }
    Ok(())
}

fn append_worksheet_hyperlinks(
    opc: &OpcPackageModel,
    worksheet: &XmlElement,
    worksheet_part: &str,
    sheet_path: &str,
    sheet: &mut DocumentNode,
) -> UseResult<()> {
    let Some(hyperlinks) = worksheet.child("hyperlinks") else {
        return Ok(());
    };
    let source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    let mut fallback_index = 0_usize;
    for hyperlink in hyperlinks.children_named("hyperlink") {
        let reference = hyperlink.attribute("ref").ok_or_else(|| {
            semantic_error(
                "use.office.spreadsheet_hyperlink_invalid",
                format!("Worksheet '{sheet_path}' contains a hyperlink without a cell reference."),
            )
        })?;
        let range = CellRange::parse(reference)?;
        let normalized_reference = range.a1();
        let single_cell = range.is_single_cell().then(|| range.start.a1());
        let cell = single_cell
            .as_deref()
            .and_then(|reference| find_cell_mut(sheet, reference));
        fallback_index += usize::from(cell.is_none());
        let path = cell.as_ref().map_or_else(
            || format!("{sheet_path}/hyperlink[{fallback_index}]"),
            |cell| {
                if cell
                    .children
                    .iter()
                    .any(|child| child.node_type == OfficeNodeType::Hyperlink)
                {
                    format!("{}/hyperlink[2]", cell.path)
                } else {
                    format!("{}/hyperlink", cell.path)
                }
            },
        );
        let mut node = DocumentNode::new(path, "hyperlink", OfficeNodeType::Hyperlink);
        node.format.insert("ref".into(), normalized_reference);
        for (attribute, key) in [("display", "display"), ("tooltip", "tooltip")] {
            if let Some(value) = hyperlink.attribute(attribute) {
                node.format.insert(key.into(), value.into());
            }
        }
        if let Some(location) = hyperlink.attribute("location") {
            node.format.insert("targetKind".into(), "internal".into());
            node.format.insert("target".into(), location.into());
        } else if let Some(id) = hyperlink.attribute("id") {
            node.format.insert("relationshipId".into(), id.into());
            if let Some(relationship) = opc
                .relationships()
                .relationship(&source, id)
                .filter(|relationship| relationship.relationship_type.ends_with("/hyperlink"))
            {
                match &relationship.target {
                    RelationshipTarget::External { uri } => {
                        node.format.insert("targetKind".into(), "external".into());
                        node.format.insert("target".into(), uri.clone());
                    }
                    RelationshipTarget::Internal {
                        part_name,
                        fragment,
                    } => {
                        node.format.insert("targetKind".into(), "internal".into());
                        let target = fragment.as_ref().map_or_else(
                            || part_name.clone(),
                            |fragment| format!("{part_name}#{fragment}"),
                        );
                        node.format.insert("target".into(), target);
                    }
                }
            }
        }
        node.text = node
            .format
            .get("display")
            .cloned()
            .or_else(|| cell.as_ref().map(|cell| cell.text.clone()))
            .unwrap_or_default();
        if let Some(cell) = cell {
            cell.children.push(node);
        } else {
            sheet.children.push(node);
        }
    }
    Ok(())
}

fn find_cell_mut<'a>(sheet: &'a mut DocumentNode, reference: &str) -> Option<&'a mut DocumentNode> {
    let path = format!("{}/{reference}", sheet.path);
    sheet
        .children
        .iter_mut()
        .flat_map(|row| row.children.iter_mut())
        .find(|cell| cell.node_type == OfficeNodeType::Cell && cell.path == path)
}

fn append_worksheet_pictures(
    package: &NativeOfficePackage,
    opc: &OpcPackageModel,
    worksheet: &XmlElement,
    worksheet_part: &str,
    sheet_path: &str,
    sheet: &mut DocumentNode,
) -> UseResult<()> {
    let drawings = worksheet.children_named("drawing").collect::<Vec<_>>();
    if drawings.len() > 1 {
        return Err(semantic_error(
            "use.office.spreadsheet_drawing_invalid",
            format!("Worksheet '{sheet_path}' contains more than one drawing element."),
        ));
    }
    let Some(drawing) = drawings.first() else {
        return Ok(());
    };
    let relationship_id = relationship_attribute(drawing, "id").ok_or_else(|| {
        semantic_error(
            "use.office.spreadsheet_drawing_invalid",
            format!("Worksheet '{sheet_path}' drawing has no relationship ID."),
        )
    })?;
    let worksheet_source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    let relationship = opc
        .relationships()
        .relationship(&worksheet_source, relationship_id)
        .filter(|relationship| relationship.relationship_type.ends_with("/drawing"))
        .ok_or_else(|| {
            semantic_error(
                "use.office.spreadsheet_drawing_invalid",
                format!("Worksheet '{sheet_path}' drawing relationship is missing or invalid."),
            )
        })?;
    let RelationshipTarget::Internal {
        part_name: drawing_part,
        ..
    } = &relationship.target
    else {
        return Err(semantic_error(
            "use.office.spreadsheet_drawing_invalid",
            "Spreadsheet drawings must use internal relationships.",
        ));
    };
    let drawing_xml = parse_xml_tree(&package.xml_part(drawing_part)?)?;
    if drawing_xml.local_name != "wsDr"
        || !matches!(
            drawing_xml.namespace.as_deref(),
            Some(SPREADSHEET_DRAWING_NAMESPACE | STRICT_SPREADSHEET_DRAWING_NAMESPACE)
        )
    {
        return Err(semantic_error(
            "use.office.spreadsheet_drawing_invalid",
            format!("Spreadsheet drawing '/{drawing_part}' has an invalid root element."),
        ));
    }
    let drawing_source = RelationshipSource::Part {
        part_name: drawing_part.clone(),
    };
    let mut picture_index = 0_usize;
    for anchor in drawing_xml.child_elements().filter(|element| {
        matches!(
            element.local_name.as_str(),
            "oneCellAnchor" | "twoCellAnchor" | "absoluteAnchor"
        )
    }) {
        let Some(picture) = find_descendant(anchor, "pic") else {
            continue;
        };
        picture_index += 1;
        let mut node = DocumentNode::new(
            format!("{sheet_path}/picture[{picture_index}]"),
            "picture",
            OfficeNodeType::Picture,
        );
        node.format
            .insert("ownerPart".into(), format!("/{drawing_part}"));
        if let Some(properties) = find_descendant(picture, "cNvPr") {
            copy_attribute(properties, "id", "id", &mut node);
            copy_attribute(properties, "name", "name", &mut node);
            if let Some(alt) = properties
                .attribute("descr")
                .or_else(|| properties.attribute("title"))
            {
                node.format.insert("alt".into(), alt.into());
            }
        }
        if let Some(blip) = find_descendant(picture, "blip") {
            if let Some(embed) = blip.attribute("embed") {
                node.format.insert("relationshipId".into(), embed.into());
                if let Some(media) = opc
                    .relationships()
                    .relationship(&drawing_source, embed)
                    .and_then(|relationship| relationship.target.internal_part_name())
                {
                    node.format.insert("part".into(), format!("/{media}"));
                }
            }
        }
        let extent = anchor
            .child("ext")
            .or_else(|| find_descendant(picture, "xfrm").and_then(|xfrm| xfrm.child("ext")));
        if let Some(extent) = extent {
            copy_pixel_extent(extent, "cx", "widthPx", &mut node);
            copy_pixel_extent(extent, "cy", "heightPx", &mut node);
        }
        if let Some(from) = anchor.child("from") {
            let column = from
                .child("col")
                .and_then(|value| direct_text(value).parse::<u32>().ok());
            let row = from
                .child("row")
                .and_then(|value| direct_text(value).parse::<u32>().ok());
            if let (Some(column), Some(row)) = (column, row) {
                if let Ok(column) = column_name(column.saturating_add(1)) {
                    node.format
                        .insert("anchorCell".into(), format!("{column}{}", row + 1));
                }
            }
        }
        sheet.children.push(node);
    }
    Ok(())
}

fn find_descendant<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a XmlElement> {
    for child in element.child_elements() {
        if child.local_name == local_name {
            return Some(child);
        }
        if let Some(found) = find_descendant(child, local_name) {
            return Some(found);
        }
    }
    None
}

fn copy_pixel_extent(element: &XmlElement, attribute: &str, key: &str, node: &mut DocumentNode) {
    if let Some(value) = element
        .attribute(attribute)
        .and_then(|value| value.parse::<u64>().ok())
    {
        node.format
            .insert(key.into(), ((value + 4_762) / 9_525).to_string());
    }
}

fn read_cell(
    cell: &XmlElement,
    sheet_path: &str,
    reference: &str,
    shared_strings: &[String],
    styles: &[BTreeMap<String, String>],
) -> UseResult<DocumentNode> {
    let mut node = DocumentNode::new(
        format!("{sheet_path}/{reference}"),
        "cell",
        OfficeNodeType::Cell,
    );
    let column = reference
        .chars()
        .take_while(|character| character.is_ascii_alphabetic())
        .collect::<String>();
    node.format.insert("column".into(), column);
    node.format.insert(
        "row".into(),
        reference
            .chars()
            .skip_while(|character| character.is_ascii_alphabetic())
            .collect(),
    );
    let value_type = cell.attribute("t").unwrap_or("n");
    let raw_value = cell.child("v").map(direct_text).unwrap_or_default();
    node.text = match value_type {
        "s" => {
            let index = raw_value.parse::<usize>().map_err(|error| {
                semantic_error(
                    "use.office.spreadsheet_shared_string_invalid",
                    format!("Cell '{reference}' has invalid shared-string index: {error}"),
                )
            })?;
            shared_strings.get(index).cloned().ok_or_else(|| {
                semantic_error(
                    "use.office.spreadsheet_shared_string_invalid",
                    format!("Cell '{reference}' references missing shared string {index}."),
                )
            })?
        }
        "inlineStr" => cell.child("is").map(spreadsheet_text).unwrap_or_default(),
        "b" => match raw_value.as_str() {
            "1" => "true".to_string(),
            "0" => "false".to_string(),
            _ => raw_value,
        },
        _ => raw_value,
    };
    node.format
        .insert("valueType".into(), value_type_name(value_type).into());
    node.format
        .insert("empty".into(), node.text.is_empty().to_string());
    node.format.insert(
        "valuePresent".into(),
        (cell.child("v").is_some() || cell.child("is").is_some()).to_string(),
    );
    if let Some(formula) = cell.child("f") {
        node.format.insert("formula".into(), direct_text(formula));
        node.format.insert(
            "formulaCached".into(),
            cell.child("v").is_some().to_string(),
        );
        if let Some(formula_type) = formula.attribute("t") {
            node.format
                .insert("formulaType".into(), formula_type.into());
        }
        if let Some(reference) = formula.attribute("ref") {
            node.format.insert("formulaRef".into(), reference.into());
        }
    }
    if let Some(style_index) = cell
        .attribute("s")
        .and_then(|value| value.parse::<usize>().ok())
    {
        node.format
            .insert("styleIndex".into(), style_index.to_string());
        let style = styles.get(style_index).ok_or_else(|| {
            semantic_error(
                "use.office.spreadsheet_style_invalid",
                format!("Cell '{reference}' references missing style {style_index}."),
            )
        })?;
        node.format.extend(style.clone());
    }
    Ok(node)
}

fn read_shared_strings(package: &NativeOfficePackage) -> UseResult<Vec<String>> {
    if !package.contains_part("xl/sharedStrings.xml") {
        return Ok(Vec::new());
    }
    let part = package.xml_part("xl/sharedStrings.xml")?;
    let root = parse_xml_tree(&part)?;
    require_spreadsheet_element(&root, "sst", part.name())?;
    Ok(root
        .children_named("si")
        .filter(|item| is_spreadsheet_namespace(item.namespace.as_deref()))
        .map(spreadsheet_text)
        .collect())
}

fn spreadsheet_text(element: &XmlElement) -> String {
    let mut output = String::new();
    append_spreadsheet_text(element, &mut output);
    output
}

fn append_spreadsheet_text(element: &XmlElement, output: &mut String) {
    let is_spreadsheet = is_spreadsheet_namespace(element.namespace.as_deref());
    if is_spreadsheet && element.local_name == "t" {
        output.push_str(&direct_text(element));
        return;
    }
    if is_spreadsheet && element.local_name == "rPh" {
        return;
    }
    for child in element.child_elements() {
        append_spreadsheet_text(child, output);
    }
}

fn is_spreadsheet_namespace(namespace: Option<&str>) -> bool {
    matches!(
        namespace,
        Some(SPREADSHEET_NAMESPACE | STRICT_SPREADSHEET_NAMESPACE)
    )
}

fn direct_text(element: &XmlElement) -> String {
    element
        .children
        .iter()
        .filter_map(|child| match child {
            XmlNode::Text(text) => Some(text.as_str()),
            XmlNode::Element(_) => None,
        })
        .collect()
}

fn normalize_cell_reference(reference: &str, expected_row: u32) -> UseResult<String> {
    let reference = reference.to_ascii_uppercase();
    let column_length = reference
        .chars()
        .take_while(|character| character.is_ascii_alphabetic())
        .count();
    if column_length == 0
        || column_length == reference.len()
        || !reference[column_length..]
            .chars()
            .all(|character| character.is_ascii_digit())
        || reference[column_length..].parse::<u32>().ok() != Some(expected_row)
        || column_number(&reference).is_none()
    {
        return Err(semantic_error(
            "use.office.spreadsheet_cell_reference_invalid",
            format!("Spreadsheet cell reference '{reference}' is invalid for row {expected_row}."),
        ));
    }
    Ok(reference)
}

fn column_number(reference: &str) -> Option<u32> {
    let column = reference
        .chars()
        .take_while(|character| character.is_ascii_alphabetic())
        .collect::<String>();
    parse_column(&column).ok()
}

fn column_name(number: u32) -> UseResult<String> {
    if !(1..=MAX_COLUMNS).contains(&number) {
        return Err(semantic_error(
            "use.office.spreadsheet_column_invalid",
            format!("Spreadsheet column number {number} is outside XFD."),
        ));
    }
    Ok(a1_column_name(number))
}

fn validate_sheet_name(name: &str) -> UseResult<()> {
    if name.is_empty()
        || name.chars().count() > 31
        || name.chars().any(char::is_control)
        || name.contains(['/', '\\', '[', ']', ':', '*', '?'])
    {
        return Err(semantic_error(
            "use.office.spreadsheet_sheet_name_invalid",
            format!("Spreadsheet sheet name '{name}' is invalid."),
        ));
    }
    Ok(())
}

fn copy_attribute(element: &XmlElement, attribute: &str, key: &str, node: &mut DocumentNode) {
    if let Some(value) = element.attribute(attribute) {
        node.format.insert(key.into(), value.into());
    }
}

fn value_type_name(value_type: &str) -> &'static str {
    match value_type {
        "s" | "inlineStr" | "str" => "String",
        "b" => "Boolean",
        "e" => "Error",
        "d" => "Date",
        _ => "Number",
    }
}

fn relationship_attribute<'a>(element: &'a XmlElement, local_name: &str) -> Option<&'a str> {
    element
        .attribute_ns(RELATIONSHIPS_NAMESPACE, local_name)
        .or_else(|| element.attribute_ns(STRICT_RELATIONSHIPS_NAMESPACE, local_name))
}

fn require_spreadsheet_element(
    element: &XmlElement,
    local_name: &str,
    part: &str,
) -> UseResult<()> {
    if element.local_name == local_name
        && matches!(
            element.namespace.as_deref(),
            Some(SPREADSHEET_NAMESPACE | STRICT_SPREADSHEET_NAMESPACE)
        )
    {
        return Ok(());
    }
    Err(semantic_error(
        "use.office.spreadsheet_xml_invalid",
        format!("Spreadsheet part '{part}' has an unexpected root element."),
    ))
}
