use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::{
    editor_error, expanded_element, indexed_cells, indexed_cells_in_row, indexed_rows,
    node_not_found, prefix, qualified, update_dimension, validate_range_size,
};
use crate::editor::{
    NativeOfficeTextFormat, NativeOfficeTextScript, NativeOfficeUnderline,
    NativeSpreadsheetCellFormat,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{
    apply_patches, element_fragment, element_with_updated_attributes, index_xml, insert_child,
    insert_ordered_child, patch_start_tag_attributes, IndexedXmlElement, XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

const STYLES_PART: &str = "xl/styles.xml";
const TRANSITIONAL_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEET: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const STYLES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const WORKBOOK_RELATIONSHIPS: &str = "xl/_rels/workbook.xml.rels";
const MAX_STYLE_RECORDS: usize = 65_000;

mod cell_format;
mod differential;

pub(super) use differential::find_or_append as find_or_append_differential_format;

const FONT_PROPERTY_ORDER: &[&str] = &[
    "name",
    "charset",
    "family",
    "b",
    "i",
    "strike",
    "outline",
    "shadow",
    "condense",
    "extend",
    "color",
    "sz",
    "u",
    "vertAlign",
    "scheme",
];

const STYLE_CHILDREN_AFTER_FONTS: &[&str] = &[
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
];

const STYLE_CHILDREN_AFTER_CELL_XFS: &[&str] =
    &["cellStyles", "dxfs", "tableStyles", "colors", "extLst"];

const DEFAULT_STYLES: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    "<fonts count=\"1\"><font>",
    "<name val=\"Aptos\"/><family val=\"2\"/><color theme=\"1\"/>",
    "<sz val=\"11\"/><scheme val=\"minor\"/>",
    "</font></fonts>",
    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>",
    "<fill><patternFill patternType=\"gray125\"/></fill></fills>",
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>",
    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>",
    "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>",
    "</styleSheet>"
);

pub(super) fn set_text_format(
    package: &mut NativeOfficePackage,
    path: &str,
    format: &NativeOfficeTextFormat,
) -> UseResult<()> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native cell text formatting is available only for Spreadsheet documents.",
        ));
    }
    if format.double_strikethrough.is_some()
        || format.text_case.is_some()
        || format.highlight.is_some()
        || format.language.is_some()
    {
        return Err(editor_error(
            "use.office.spreadsheet_run_format_unsupported",
            "Spreadsheet cells do not support double strikethrough, run text case, run highlight, or run language through the native text-format contract.",
        ));
    }
    set_format(package, path, Some(format), None)
}

pub(super) fn set_cell_format(
    package: &mut NativeOfficePackage,
    path: &str,
    format: &NativeSpreadsheetCellFormat,
) -> UseResult<()> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native cell formatting is available only for Spreadsheet documents.",
        ));
    }
    set_format(package, path, None, Some(format))
}

fn set_format(
    package: &mut NativeOfficePackage,
    path: &str,
    text_format: Option<&NativeOfficeTextFormat>,
    cell_format: Option<&NativeSpreadsheetCellFormat>,
) -> UseResult<()> {
    let (sheet_path, reference) = path.rsplit_once('/').ok_or_else(|| node_not_found(path))?;
    if sheet_path.is_empty() {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Spreadsheet formatting requires a cell path such as /Sheet1/A1.",
        ));
    }
    let range = CellRange::parse(reference)?;
    validate_range_size(range)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path.eq_ignore_ascii_case(sheet_path)
        })
        .ok_or_else(|| node_not_found(path))?;
    let part_name = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let worksheet = package.xml_part(&part_name)?;
    let worksheet_index = index_xml(&worksheet)?;
    let sheet_data = worksheet_index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(path))?;
    let existing_styles = indexed_cells(sheet_data)
        .into_iter()
        .filter(|(reference, _, _)| range.contains(*reference))
        .map(|(reference, _, cell)| Ok((reference, cell_style_index(cell)?)))
        .collect::<UseResult<BTreeMap<_, _>>>()?;
    let mut base_styles = existing_styles.values().copied().collect::<BTreeSet<_>>();
    if existing_styles.len() < range.cell_count()? {
        base_styles.insert(0);
    }

    ensure_style_collections(package)?;
    let resolved_cell_format = cell_format::resolve(package, cell_format)?;
    let mut derived_styles = BTreeMap::new();
    for base_style in base_styles {
        derived_styles.insert(
            base_style,
            style_index_for_format(
                package,
                base_style,
                text_format,
                cell_format,
                &resolved_cell_format,
            )?,
        );
    }

    let worksheet = package.xml_part(&part_name)?;
    let worksheet_index = index_xml(&worksheet)?;
    let sheet_data = worksheet_index
        .descendant("sheetData")
        .ok_or_else(|| node_not_found(path))?;
    let edited = set_range_style(
        &worksheet,
        sheet_data,
        range,
        &existing_styles,
        &derived_styles,
    )?;
    let edited = update_dimension(&part_name, edited)?;
    package.set_part(&part_name, edited)
}

pub(super) fn derived_cell_style_indexes(
    package: &mut NativeOfficePackage,
    base_styles: &BTreeSet<usize>,
    format: &NativeSpreadsheetCellFormat,
) -> UseResult<BTreeMap<usize, usize>> {
    format.validate()?;
    ensure_style_collections(package)?;
    let resolved = cell_format::resolve(package, Some(format))?;
    base_styles
        .iter()
        .copied()
        .map(|base_style| {
            style_index_for_format(package, base_style, None, Some(format), &resolved)
                .map(|derived| (base_style, derived))
        })
        .collect()
}

fn ensure_style_collections(package: &mut NativeOfficePackage) -> UseResult<()> {
    ensure_styles_part(package)?;
    ensure_collection(package, "fonts", "font", STYLE_CHILDREN_AFTER_FONTS)?;
    ensure_collection(package, "cellXfs", "xf", STYLE_CHILDREN_AFTER_CELL_XFS)
}

fn ensure_styles_part(package: &mut NativeOfficePackage) -> UseResult<()> {
    if package.contains_part(STYLES_PART) {
        return Ok(());
    }
    crate::opc_edit::add_content_type_override(package, STYLES_PART, STYLES_CONTENT_TYPE)?;
    let workbook = package.xml_part("xl/workbook.xml")?;
    let (spreadsheet_namespace, relationship_namespace) = match workbook.root().namespace.as_deref()
    {
        Some(TRANSITIONAL_SPREADSHEET) => (TRANSITIONAL_SPREADSHEET, TRANSITIONAL_RELATIONSHIPS),
        Some(STRICT_SPREADSHEET) => (STRICT_SPREADSHEET, STRICT_RELATIONSHIPS),
        _ => {
            return Err(editor_error(
                "use.office.spreadsheet_styles_dialect_unsupported",
                "Spreadsheet style creation requires a recognized transitional or strict OOXML namespace.",
            ));
        }
    };
    let styles = DEFAULT_STYLES.replace(TRANSITIONAL_SPREADSHEET, spreadsheet_namespace);
    package.set_part(STYLES_PART, styles.into_bytes())?;
    let relationships = package.xml_part(WORKBOOK_RELATIONSHIPS)?;
    let index = index_xml(&relationships)?;
    let existing = index
        .children
        .iter()
        .filter(|child| child.local_name == "Relationship")
        .find(|child| {
            child
                .attributes
                .get("Type")
                .is_some_and(|value| value.ends_with("/styles"))
        });
    if let Some(existing) = existing {
        let target = existing.attributes.get("Target").map(String::as_str);
        if !matches!(target, Some("styles.xml" | "/xl/styles.xml")) {
            return Err(editor_error(
                "use.office.spreadsheet_styles_relationship_conflict",
                "Workbook styles relationship targets a different part.",
            ));
        }
        return Ok(());
    }
    crate::opc_edit::add_relationship(
        package,
        WORKBOOK_RELATIONSHIPS,
        &format!("{relationship_namespace}/styles"),
        "styles.xml",
    )?;
    Ok(())
}

fn ensure_collection(
    package: &mut NativeOfficePackage,
    collection_name: &str,
    child_name: &str,
    later_names: &[&str],
) -> UseResult<()> {
    let part = package.xml_part(STYLES_PART)?;
    let index = index_xml(&part)?;
    let bytes = if let Some(collection) = index.child(collection_name, 1) {
        if collection.child(child_name, 1).is_some() {
            return normalize_collection_count(package, collection_name, child_name);
        }
        let default_child =
            default_collection_child(collection_name, prefix(&collection.qualified_name))?;
        insert_child(&part, collection, default_child)?
    } else {
        let collection_prefix = prefix(&index.qualified_name);
        let tag = qualified(collection_prefix, collection_name);
        let default_child = default_collection_child(collection_name, collection_prefix)?;
        insert_ordered_child(
            &part,
            &index,
            format!("<{tag} count=\"1\">{default_child}</{tag}>"),
            later_names,
        )?
    };
    package.set_part(STYLES_PART, bytes)?;
    normalize_collection_count(package, collection_name, child_name)
}

fn default_collection_child(collection_name: &str, prefix: Option<&str>) -> UseResult<String> {
    match collection_name {
        "fonts" => {
            let font = qualified(prefix, "font");
            let name = qualified(prefix, "name");
            let family = qualified(prefix, "family");
            let color = qualified(prefix, "color");
            let size = qualified(prefix, "sz");
            let scheme = qualified(prefix, "scheme");
            Ok(format!(
                "<{font}><{name} val=\"Aptos\"/><{family} val=\"2\"/><{color} theme=\"1\"/><{size} val=\"11\"/><{scheme} val=\"minor\"/></{font}>"
            ))
        }
        "cellXfs" => {
            let xf = qualified(prefix, "xf");
            Ok(format!(
                "<{xf} numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
            ))
        }
        _ => Err(styles_invalid()),
    }
}

fn normalize_collection_count(
    package: &mut NativeOfficePackage,
    collection_name: &str,
    child_name: &str,
) -> UseResult<()> {
    let part = package.xml_part(STYLES_PART)?;
    let index = index_xml(&part)?;
    let collection = index.child(collection_name, 1).ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_styles_invalid",
            format!("Spreadsheet styles have no '{collection_name}' collection."),
        )
    })?;
    let count = collection
        .children
        .iter()
        .filter(|child| child.local_name == child_name)
        .count();
    let expected = count.to_string();
    if collection.attributes.get("count") == Some(&expected) {
        return Ok(());
    }
    let edited = patch_start_tag_attributes(
        &part,
        collection,
        &BTreeMap::from([("count".to_string(), Some(expected))]),
    )?;
    package.set_part(STYLES_PART, edited)
}

fn style_index_for_format(
    package: &mut NativeOfficePackage,
    base_style: usize,
    text_format: Option<&NativeOfficeTextFormat>,
    cell_format: Option<&NativeSpreadsheetCellFormat>,
    resolved_cell_format: &cell_format::ResolvedCellFormat,
) -> UseResult<usize> {
    let mut font_id = None;
    if text_format.is_some_and(NativeOfficeTextFormat::has_character_properties) {
        let format = text_format.ok_or_else(styles_invalid)?;
        let part = package.xml_part(STYLES_PART)?;
        let index = index_xml(&part)?;
        let cell_xfs = index.child("cellXfs", 1).ok_or_else(styles_invalid)?;
        let base = cell_xfs
            .child("xf", base_style + 1)
            .ok_or_else(|| missing_style(base_style))?;
        let base_font = base.attributes.get("fontId").map_or(Ok(0), |value| {
            value.parse::<usize>().map_err(|_| styles_invalid())
        })?;
        let candidate = derive_font_fragment(&part, base_font, format)?;
        font_id = Some(find_or_append_child(package, "fonts", "font", candidate)?);
    }

    let border_id = cell_format
        .and_then(|format| format.border.as_ref())
        .map(|border| cell_format::border_index(package, base_style, border))
        .transpose()?;

    let part = package.xml_part(STYLES_PART)?;
    let candidate = cell_format::derive_xf_fragment(
        &part,
        base_style,
        font_id,
        border_id,
        text_format,
        cell_format,
        resolved_cell_format,
    )?;
    find_or_append_child(package, "cellXfs", "xf", candidate)
}

fn derive_font_fragment(
    part: &LosslessXmlPart,
    font_index: usize,
    format: &NativeOfficeTextFormat,
) -> UseResult<Vec<u8>> {
    require_collection_child(part, "fonts", "font", font_index)?;
    let mut bytes = part.raw().to_vec();
    if let Some(value) = format.bold {
        bytes = set_font_value(bytes, font_index, "b", if value { "1" } else { "0" }, &[])?;
    }
    if let Some(value) = format.italic {
        bytes = set_font_value(bytes, font_index, "i", if value { "1" } else { "0" }, &[])?;
    }
    if let Some(value) = format.strikethrough {
        bytes = set_font_value(
            bytes,
            font_index,
            "strike",
            if value { "1" } else { "0" },
            &[],
        )?;
    }
    if let Some(family) = &format.font_family {
        bytes = set_font_value(bytes, font_index, "name", family, &[])?;
        bytes = remove_font_property(bytes, font_index, "scheme")?;
    }
    if let Some(size) = format.font_size_centipoints {
        bytes = set_font_value(bytes, font_index, "sz", &format_points(size), &[])?;
    }
    if let Some(color) = format.text_color {
        bytes = set_font_value(
            bytes,
            font_index,
            "color",
            &format!("FF{}", color.hex()),
            &["theme", "indexed", "tint", "auto"],
        )?;
    }
    if let Some(underline) = format.underline {
        bytes = set_font_value(
            bytes,
            font_index,
            "u",
            spreadsheet_underline(underline),
            &[],
        )?;
    }
    if let Some(script) = format.script {
        bytes = set_font_value(
            bytes,
            font_index,
            "vertAlign",
            spreadsheet_script(script),
            &[],
        )?;
    }
    let derived = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let element = require_collection_child(&derived, "fonts", "font", font_index)?;
    Ok(element_fragment(&derived, &element)?.to_vec())
}

fn spreadsheet_underline(underline: NativeOfficeUnderline) -> &'static str {
    match underline {
        NativeOfficeUnderline::None => "none",
        NativeOfficeUnderline::Single => "single",
        NativeOfficeUnderline::Double => "double",
    }
}

fn spreadsheet_script(script: NativeOfficeTextScript) -> &'static str {
    match script {
        NativeOfficeTextScript::Baseline => "baseline",
        NativeOfficeTextScript::Superscript => "superscript",
        NativeOfficeTextScript::Subscript => "subscript",
    }
}

fn set_font_value(
    bytes: Vec<u8>,
    font_index: usize,
    name: &str,
    value: &str,
    remove_attributes: &[&str],
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let font = require_collection_child(&part, "fonts", "font", font_index)?;
    let attribute_name = if name == "color" { "rgb" } else { "val" };
    if let Some(property) = font.child(name, 1) {
        let mut updates = BTreeMap::from([(attribute_name.to_string(), Some(value.to_string()))]);
        for attribute in remove_attributes {
            updates.insert((*attribute).to_string(), None);
        }
        return patch_start_tag_attributes(&part, property, &updates);
    }
    let tag = qualified(prefix(&font.qualified_name), name);
    let position = FONT_PROPERTY_ORDER
        .iter()
        .position(|candidate| *candidate == name)
        .ok_or_else(styles_invalid)?;
    insert_ordered_child(
        &part,
        &font,
        format!(
            "<{tag} {attribute_name}=\"{}\"/>",
            crate::xml_edit::escape_attribute(value)
        ),
        &FONT_PROPERTY_ORDER[position + 1..],
    )
}

fn remove_font_property(bytes: Vec<u8>, font_index: usize, name: &str) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let font = require_collection_child(&part, "fonts", "font", font_index)?;
    let Some(property) = font.child(name, 1) else {
        return Ok(part.raw().to_vec());
    };
    apply_patches(
        &part,
        vec![XmlPatch::new(property.full_range.clone(), Vec::new())],
    )
}

fn find_or_append_child(
    package: &mut NativeOfficePackage,
    collection_name: &str,
    child_name: &str,
    candidate: Vec<u8>,
) -> UseResult<usize> {
    let part = package.xml_part(STYLES_PART)?;
    let index = index_xml(&part)?;
    let collection = index.child(collection_name, 1).ok_or_else(styles_invalid)?;
    for (position, child) in collection
        .children
        .iter()
        .filter(|child| child.local_name == child_name)
        .enumerate()
    {
        if element_fragment(&part, child)? == candidate {
            return Ok(position);
        }
    }
    let position = collection
        .children
        .iter()
        .filter(|child| child.local_name == child_name)
        .count();
    if position >= MAX_STYLE_RECORDS {
        return Err(editor_error(
            "use.office.spreadsheet_style_limit",
            format!(
                "Spreadsheet '{collection_name}' cannot exceed {MAX_STYLE_RECORDS} native records."
            ),
        ));
    }
    let edited = insert_child(&part, collection, candidate)?;
    package.set_part(STYLES_PART, edited)?;
    normalize_collection_count(package, collection_name, child_name)?;
    Ok(position)
}

fn require_collection_child(
    part: &LosslessXmlPart,
    collection_name: &str,
    child_name: &str,
    index: usize,
) -> UseResult<IndexedXmlElement> {
    let root = index_xml(part)?;
    let collection = root.child(collection_name, 1).ok_or_else(styles_invalid)?;
    collection
        .child(child_name, index + 1)
        .cloned()
        .ok_or_else(|| missing_style(index))
}

fn set_range_style(
    part: &LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    range: CellRange,
    existing_styles: &BTreeMap<CellReference, usize>,
    derived_styles: &BTreeMap<usize, usize>,
) -> UseResult<Vec<u8>> {
    if sheet_data.empty {
        let default_style = required_derived_style(derived_styles, 0)?;
        let cell_prefix = prefix(&sheet_data.qualified_name);
        let row_tag = qualified(cell_prefix, "row");
        let rows = (range.start.row..=range.end.row)
            .map(|row| {
                let cells = (range.start.column..=range.end.column)
                    .map(|column| {
                        let reference = CellReference { column, row };
                        styled_cell(cell_prefix, &reference.a1(), default_style)
                    })
                    .collect::<String>();
                format!("<{row_tag} r=\"{row}\">{cells}</{row_tag}>")
            })
            .collect::<String>();
        return insert_child(part, sheet_data, rows);
    }

    let rows = indexed_rows(sheet_data);
    let row_map = rows.iter().copied().collect::<BTreeMap<_, _>>();
    let mut patches = Vec::new();
    let mut insertions = BTreeMap::<usize, Vec<(u32, u32, String)>>::new();
    let cell_prefix = prefix(&sheet_data.qualified_name);
    let default_style = required_derived_style(derived_styles, 0).ok();
    for row_number in range.start.row..=range.end.row {
        let Some(row) = row_map.get(&row_number).copied() else {
            let default_style = default_style.ok_or_else(styles_invalid)?;
            let cells = (range.start.column..=range.end.column)
                .map(|column| {
                    let reference = CellReference {
                        column,
                        row: row_number,
                    };
                    styled_cell(cell_prefix, &reference.a1(), default_style)
                })
                .collect::<String>();
            let row_tag = qualified(cell_prefix, "row");
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
            let default_style = default_style.ok_or_else(styles_invalid)?;
            let cells = (range.start.column..=range.end.column)
                .map(|column| {
                    let reference = CellReference {
                        column,
                        row: row_number,
                    };
                    styled_cell(prefix(&row.qualified_name), &reference.a1(), default_style)
                })
                .collect::<String>();
            patches.push(XmlPatch::new(
                row.full_range.clone(),
                expanded_element(row, &cells),
            ));
            continue;
        }
        let cells = indexed_cells_in_row(row_number, row);
        for column in range.start.column..=range.end.column {
            let reference = CellReference {
                column,
                row: row_number,
            };
            let base = existing_styles.get(&reference).copied().unwrap_or(0);
            let style = required_derived_style(derived_styles, base)?;
            if let Some((_, cell)) = cells.iter().find(|(existing, _)| existing.column == column) {
                patches.push(XmlPatch::new(
                    cell.full_range.clone(),
                    element_with_updated_attributes(
                        part,
                        cell,
                        &BTreeMap::from([("s".to_string(), Some(style.to_string()))]),
                    )?,
                ));
                continue;
            }
            let position = cells
                .iter()
                .find(|(existing, _)| existing.column > column)
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
                column,
                styled_cell(prefix(&row.qualified_name), &reference.a1(), style),
            ));
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
    apply_patches(part, patches)
}

fn required_derived_style(styles: &BTreeMap<usize, usize>, base: usize) -> UseResult<usize> {
    styles.get(&base).copied().ok_or_else(styles_invalid)
}

fn styled_cell(prefix: Option<&str>, reference: &str, style: usize) -> String {
    let tag = qualified(prefix, "c");
    format!("<{tag} r=\"{reference}\" s=\"{style}\"/>")
}

pub(super) fn cell_style_index(cell: &IndexedXmlElement) -> UseResult<usize> {
    cell.attributes.get("s").map_or(Ok(0), |value| {
        value.parse::<usize>().map_err(|_| styles_invalid())
    })
}

fn format_points(centipoints: u32) -> String {
    let points = centipoints / 100;
    let fraction = centipoints % 100;
    if fraction == 0 {
        points.to_string()
    } else if fraction % 10 == 0 {
        format!("{points}.{}", fraction / 10)
    } else {
        format!("{points}.{fraction:02}")
    }
}

fn missing_style(index: usize) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_style_invalid",
        format!("Spreadsheet style index {index} does not exist."),
    )
}

fn styles_invalid() -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_styles_invalid",
        "Spreadsheet styles do not contain a valid fonts and cellXfs model.",
    )
}
