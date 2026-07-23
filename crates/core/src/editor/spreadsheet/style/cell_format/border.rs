use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use crate::editor::{NativeSpreadsheetBorder, NativeSpreadsheetBorderLine};
use crate::xml_edit::{apply_patches, XmlPatch};

use super::super::{
    element_fragment, ensure_styles_part, find_or_append_child, index_xml, insert_child,
    insert_ordered_child, normalize_collection_count, patch_start_tag_attributes, prefix,
    qualified, require_collection_child, styles_invalid, LosslessXmlPart, NativeOfficePackage,
    STYLES_PART,
};

const STYLE_CHILDREN_AFTER_BORDERS: &[&str] = &[
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
];

const BORDER_LINE_ORDER: &[&str] = &[
    "start",
    "end",
    "left",
    "right",
    "top",
    "bottom",
    "diagonal",
    "vertical",
    "horizontal",
    "extLst",
];

pub(super) fn index(
    package: &mut NativeOfficePackage,
    base_style: usize,
    update: &NativeSpreadsheetBorder,
) -> UseResult<usize> {
    ensure_borders(package)?;
    let part = package.xml_part(STYLES_PART)?;
    let style = require_collection_child(&part, "cellXfs", "xf", base_style)?;
    let base_border = style.attributes.get("borderId").map_or(Ok(0), |value| {
        value.parse::<usize>().map_err(|_| styles_invalid())
    })?;
    let candidate = derive_border_fragment(&part, base_border, update)?;
    find_or_append_child(package, "borders", "border", candidate)
}

fn derive_border_fragment(
    part: &LosslessXmlPart,
    border_index: usize,
    update: &NativeSpreadsheetBorder,
) -> UseResult<Vec<u8>> {
    require_collection_child(part, "borders", "border", border_index)?;
    let mut bytes = part.raw().to_vec();
    for (name, line) in [
        ("left", update.left),
        ("right", update.right),
        ("top", update.top),
        ("bottom", update.bottom),
        ("diagonal", update.diagonal),
    ] {
        if let Some(line) = line {
            bytes = patch_border_line(bytes, border_index, name, line)?;
        }
    }
    if update.diagonal_up.is_some() || update.diagonal_down.is_some() {
        let current = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
        let border = require_collection_child(&current, "borders", "border", border_index)?;
        let mut attributes = BTreeMap::new();
        if let Some(value) = update.diagonal_up {
            attributes.insert(
                "diagonalUp".to_string(),
                Some(if value { "1" } else { "0" }.to_string()),
            );
        }
        if let Some(value) = update.diagonal_down {
            attributes.insert(
                "diagonalDown".to_string(),
                Some(if value { "1" } else { "0" }.to_string()),
            );
        }
        bytes = patch_start_tag_attributes(&current, &border, &attributes)?;
    }
    let derived = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let border = require_collection_child(&derived, "borders", "border", border_index)?;
    Ok(element_fragment(&derived, &border)?.to_vec())
}

fn patch_border_line(
    bytes: Vec<u8>,
    border_index: usize,
    name: &str,
    line: NativeSpreadsheetBorderLine,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let border = require_collection_child(&part, "borders", "border", border_index)?;
    let Some(existing) = border.child(name, 1) else {
        let fragment = border_line_fragment(prefix(&border.qualified_name), name, line);
        let position = BORDER_LINE_ORDER
            .iter()
            .position(|candidate| *candidate == name)
            .ok_or_else(styles_invalid)?;
        return insert_ordered_child(&part, &border, fragment, &BORDER_LINE_ORDER[position + 1..]);
    };
    let style = match line {
        NativeSpreadsheetBorderLine::None => None,
        NativeSpreadsheetBorderLine::Line { style, .. } => {
            Some(style.spreadsheet_value().to_string())
        }
    };
    let bytes = patch_start_tag_attributes(
        &part,
        existing,
        &BTreeMap::from([("style".to_string(), style)]),
    )?;
    patch_border_color(bytes, border_index, name, line_color(line))
}

fn patch_border_color(
    bytes: Vec<u8>,
    border_index: usize,
    name: &str,
    color: Option<crate::editor::NativeOfficeRgbColor>,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let border = require_collection_child(&part, "borders", "border", border_index)?;
    let line = border.child(name, 1).ok_or_else(styles_invalid)?;
    match (line.child("color", 1), color) {
        (Some(existing), Some(color)) => {
            let mut attributes =
                BTreeMap::from([("rgb".to_string(), Some(format!("FF{}", color.hex())))]);
            for name in ["theme", "indexed", "tint", "auto"] {
                attributes.insert(name.to_string(), None);
            }
            patch_start_tag_attributes(&part, existing, &attributes)
        }
        (Some(existing), None) => apply_patches(
            &part,
            vec![XmlPatch::new(existing.full_range.clone(), Vec::new())],
        ),
        (None, Some(color)) => {
            let tag = qualified(prefix(&line.qualified_name), "color");
            insert_child(&part, line, format!("<{tag} rgb=\"FF{}\"/>", color.hex()))
        }
        (None, None) => Ok(part.raw().to_vec()),
    }
}

fn border_line_fragment(
    prefix: Option<&str>,
    name: &str,
    line: NativeSpreadsheetBorderLine,
) -> String {
    let tag = qualified(prefix, name);
    match line {
        NativeSpreadsheetBorderLine::None => format!("<{tag}/>"),
        NativeSpreadsheetBorderLine::Line { style, color: None } => {
            format!("<{tag} style=\"{}\"/>", style.spreadsheet_value())
        }
        NativeSpreadsheetBorderLine::Line {
            style,
            color: Some(color),
        } => {
            let color_tag = qualified(prefix, "color");
            format!(
                "<{tag} style=\"{}\"><{color_tag} rgb=\"FF{}\"/></{tag}>",
                style.spreadsheet_value(),
                color.hex()
            )
        }
    }
}

fn line_color(line: NativeSpreadsheetBorderLine) -> Option<crate::editor::NativeOfficeRgbColor> {
    match line {
        NativeSpreadsheetBorderLine::None => None,
        NativeSpreadsheetBorderLine::Line { color, .. } => color,
    }
}

fn ensure_borders(package: &mut NativeOfficePackage) -> UseResult<()> {
    ensure_styles_part(package)?;
    let part = package.xml_part(STYLES_PART)?;
    let root = index_xml(&part)?;
    let bytes = if let Some(borders) = root.child("borders", 1) {
        if borders.child("border", 1).is_some() {
            return normalize_collection_count(package, "borders", "border");
        }
        insert_child(
            &part,
            borders,
            default_border(prefix(&borders.qualified_name)),
        )?
    } else {
        let style_prefix = prefix(&root.qualified_name);
        let borders_tag = qualified(style_prefix, "borders");
        insert_ordered_child(
            &part,
            &root,
            format!(
                "<{borders_tag} count=\"1\">{}</{borders_tag}>",
                default_border(style_prefix)
            ),
            STYLE_CHILDREN_AFTER_BORDERS,
        )?
    };
    package.set_part(STYLES_PART, bytes)?;
    normalize_collection_count(package, "borders", "border")
}

fn default_border(prefix: Option<&str>) -> String {
    let border = qualified(prefix, "border");
    let left = qualified(prefix, "left");
    let right = qualified(prefix, "right");
    let top = qualified(prefix, "top");
    let bottom = qualified(prefix, "bottom");
    let diagonal = qualified(prefix, "diagonal");
    format!("<{border}><{left}/><{right}/><{top}/><{bottom}/><{diagonal}/></{border}>")
}
