use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{ensure_styles_part, styles_invalid, STYLES_PART};
use crate::editor::{NativeOfficeRgbColor, NativeSpreadsheetDifferentialFormat};
use crate::xml_edit::{
    index_xml, insert_child, insert_ordered_child, patch_start_tag_attributes, IndexedXmlElement,
};
use crate::NativeOfficePackage;

const MAX_DIFFERENTIAL_FORMATS: usize = 65_000;
const STYLE_CHILDREN_AFTER_DXFS: &[&str] = &["tableStyles", "colors", "extLst"];

pub(in crate::editor::spreadsheet) fn find_or_append(
    package: &mut NativeOfficePackage,
    format: &NativeSpreadsheetDifferentialFormat,
) -> UseResult<Option<usize>> {
    if format.is_empty() {
        return Ok(None);
    }
    ensure_styles_part(package)?;
    let part = package.xml_part(STYLES_PART)?;
    let index = index_xml(&part)?;
    let existing = index.child("dxfs", 1);
    if let Some(collection) = existing {
        let children = collection
            .children
            .iter()
            .filter(|child| child.local_name == "dxf" && child.namespace == collection.namespace)
            .collect::<Vec<_>>();
        if children.len() != collection.children.len() {
            return Err(super::editor_error(
                "use.office.spreadsheet_conditional_format_unknown_styles",
                "A differential format cannot be added without risking unknown dxfs child content.",
            )
            .with_detail("part", STYLES_PART));
        }
        if let Some(position) = children
            .iter()
            .position(|child| read_owned_dxf(child).as_ref() == Some(format))
        {
            return Ok(Some(position));
        }
        if children.len() >= MAX_DIFFERENTIAL_FORMATS {
            return Err(super::editor_error(
                "use.office.spreadsheet_conditional_format_dxf_limit",
                format!(
                    "Spreadsheet styles already contain the maximum {MAX_DIFFERENTIAL_FORMATS} differential formats."
                ),
            ));
        }
        let fragment = dxf_fragment(super::prefix(&collection.qualified_name), format);
        let inserted = insert_child(&part, collection, fragment)?;
        package.set_part(STYLES_PART, inserted)?;
        update_count(package, children.len() + 1)?;
        return Ok(Some(children.len()));
    }

    let prefix = super::prefix(&index.qualified_name);
    let collection_name = super::qualified(prefix, "dxfs");
    let fragment = dxf_fragment(prefix, format);
    let edited = insert_ordered_child(
        &part,
        &index,
        format!("<{collection_name} count=\"1\">{fragment}</{collection_name}>"),
        STYLE_CHILDREN_AFTER_DXFS,
    )?;
    package.set_part(STYLES_PART, edited)?;
    Ok(Some(0))
}

fn dxf_fragment(
    namespace_prefix: Option<&str>,
    format: &NativeSpreadsheetDifferentialFormat,
) -> String {
    let dxf = super::qualified(namespace_prefix, "dxf");
    let mut children = String::new();
    if format.font_color.is_some() || format.bold.is_some() {
        let font = super::qualified(namespace_prefix, "font");
        let color = super::qualified(namespace_prefix, "color");
        let bold = super::qualified(namespace_prefix, "b");
        children.push_str(&format!("<{font}>"));
        if let Some(value) = format.bold {
            children.push_str(&format!(
                "<{bold} val=\"{}\"/>",
                if value { "1" } else { "0" }
            ));
        }
        if let Some(value) = format.font_color {
            children.push_str(&format!("<{color} rgb=\"FF{}\"/>", value.hex()));
        }
        children.push_str(&format!("</{font}>"));
    }
    if let Some(value) = format.fill {
        let fill = super::qualified(namespace_prefix, "fill");
        let pattern = super::qualified(namespace_prefix, "patternFill");
        let foreground = super::qualified(namespace_prefix, "fgColor");
        let background = super::qualified(namespace_prefix, "bgColor");
        children.push_str(&format!(
            "<{fill}><{pattern} patternType=\"solid\"><{foreground} rgb=\"FF{}\"/><{background} indexed=\"64\"/></{pattern}></{fill}>",
            value.hex()
        ));
    }
    format!("<{dxf}>{children}</{dxf}>")
}

fn read_owned_dxf(element: &IndexedXmlElement) -> Option<NativeSpreadsheetDifferentialFormat> {
    if !element.qualified_attributes.is_empty() {
        return None;
    }
    let mut format = NativeSpreadsheetDifferentialFormat::default();
    for child in &element.children {
        if child.namespace != element.namespace {
            return None;
        }
        match child.local_name.as_str() {
            "font" => read_font(child, &mut format)?,
            "fill" => format.fill = Some(read_fill(child)?),
            _ => return None,
        }
    }
    Some(format)
}

fn read_font(
    font: &IndexedXmlElement,
    format: &mut NativeSpreadsheetDifferentialFormat,
) -> Option<()> {
    if !font.qualified_attributes.is_empty() {
        return None;
    }
    for child in &font.children {
        if child.namespace != font.namespace {
            return None;
        }
        match child.local_name.as_str() {
            "b" => format.bold = Some(parse_bool(child.attributes.get("val").map(String::as_str))?),
            "color" => {
                format.font_color = Some(parse_color(child.attributes.get("rgb")?)?);
            }
            _ => return None,
        }
    }
    Some(())
}

fn read_fill(fill: &IndexedXmlElement) -> Option<NativeOfficeRgbColor> {
    if !fill.qualified_attributes.is_empty() || fill.children.len() != 1 {
        return None;
    }
    let pattern = &fill.children[0];
    if pattern.namespace != fill.namespace
        || pattern.local_name != "patternFill"
        || pattern.attributes.get("patternType").map(String::as_str) != Some("solid")
    {
        return None;
    }
    pattern
        .children
        .iter()
        .find(|child| child.local_name == "fgColor" && child.namespace == pattern.namespace)
        .or_else(|| {
            pattern
                .children
                .iter()
                .find(|child| child.local_name == "bgColor" && child.namespace == pattern.namespace)
        })
        .and_then(|color| color.attributes.get("rgb"))
        .and_then(|value| parse_color(value))
}

fn parse_bool(value: Option<&str>) -> Option<bool> {
    match value {
        None | Some("1" | "true") => Some(true),
        Some("0" | "false") => Some(false),
        Some(_) => None,
    }
}

fn parse_color(value: &str) -> Option<NativeOfficeRgbColor> {
    let rgb = value.strip_prefix("FF").unwrap_or(value);
    if rgb.len() != 6 || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    Some(NativeOfficeRgbColor::new(
        u8::from_str_radix(&rgb[0..2], 16).ok()?,
        u8::from_str_radix(&rgb[2..4], 16).ok()?,
        u8::from_str_radix(&rgb[4..6], 16).ok()?,
    ))
}

fn update_count(package: &mut NativeOfficePackage, count: usize) -> UseResult<()> {
    let part = package.xml_part(STYLES_PART)?;
    let index = index_xml(&part)?;
    let collection = index.child("dxfs", 1).ok_or_else(styles_invalid)?;
    let edited = patch_start_tag_attributes(
        &part,
        collection,
        &BTreeMap::from([("count".to_string(), Some(count.to_string()))]),
    )?;
    package.set_part(STYLES_PART, edited)
}
