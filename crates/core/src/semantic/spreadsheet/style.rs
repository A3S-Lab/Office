use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::require_spreadsheet_element;
use crate::semantic::semantic_error;
use crate::xml_tree::{parse_xml_tree, XmlElement};
use crate::NativeOfficePackage;

#[derive(Debug, Clone)]
pub(super) struct DifferentialFormat {
    pub(super) values: BTreeMap<String, String>,
    pub(super) supported: bool,
}

pub(super) fn read_styles(
    package: &NativeOfficePackage,
) -> UseResult<Vec<BTreeMap<String, String>>> {
    if !package.contains_part("xl/styles.xml") {
        return Ok(vec![BTreeMap::new()]);
    }
    let part = package.xml_part("xl/styles.xml")?;
    let root = parse_xml_tree(&part)?;
    require_spreadsheet_element(&root, "styleSheet", part.name())?;
    let custom_formats = root
        .child("numFmts")
        .map(|formats| {
            formats
                .children_named("numFmt")
                .filter_map(|format| {
                    Some((
                        format.attribute("numFmtId")?.to_string(),
                        format.attribute("formatCode")?.to_string(),
                    ))
                })
                .collect::<BTreeMap<_, _>>()
        })
        .unwrap_or_default();
    // Some legacy producers omit the required fonts collection while still
    // using fontId=0. Retain the prior tolerant read for that default only;
    // explicit out-of-range font references fail closed.
    let fonts = root
        .child("fonts")
        .map(|fonts| fonts.children_named("font").map(read_font).collect())
        .unwrap_or_else(|| vec![BTreeMap::new()]);
    let fills = root
        .child("fills")
        .map(|fills| fills.children_named("fill").map(read_fill).collect())
        .unwrap_or_else(|| vec![None]);
    let borders = root
        .child("borders")
        .map(|borders| borders.children_named("border").map(read_border).collect())
        .unwrap_or_else(|| vec![BTreeMap::new()]);
    let mut styles = Vec::new();
    if let Some(cell_xfs) = root.child("cellXfs") {
        for style in cell_xfs.children_named("xf") {
            let mut values = BTreeMap::new();
            for (attribute, key) in [
                ("fontId", "fontId"),
                ("fillId", "fillId"),
                ("borderId", "borderId"),
                ("xfId", "baseStyleId"),
            ] {
                if let Some(value) = style.attribute(attribute) {
                    values.insert(key.to_string(), value.to_string());
                }
            }
            if let Some(font_id) = style.attribute("fontId") {
                let index = font_id.parse::<usize>().map_err(|error| {
                    semantic_error(
                        "use.office.spreadsheet_style_invalid",
                        format!("Spreadsheet style has invalid fontId '{font_id}': {error}"),
                    )
                })?;
                let font = fonts.get(index).ok_or_else(|| {
                    semantic_error(
                        "use.office.spreadsheet_style_invalid",
                        format!("Spreadsheet style references missing font {index}."),
                    )
                })?;
                values.extend(font.clone());
            }
            if let Some(fill_id) = style.attribute("fillId") {
                let index = fill_id.parse::<usize>().map_err(|error| {
                    semantic_error(
                        "use.office.spreadsheet_style_invalid",
                        format!("Spreadsheet style has invalid fillId '{fill_id}': {error}"),
                    )
                })?;
                let fill = fills.get(index).ok_or_else(|| {
                    semantic_error(
                        "use.office.spreadsheet_style_invalid",
                        format!("Spreadsheet style references missing fill {index}."),
                    )
                })?;
                if let Some(color) = fill {
                    values.insert("fill".into(), color.clone());
                }
            }
            if let Some(border_id) = style.attribute("borderId") {
                let index = border_id.parse::<usize>().map_err(|error| {
                    semantic_error(
                        "use.office.spreadsheet_style_invalid",
                        format!("Spreadsheet style has invalid borderId '{border_id}': {error}"),
                    )
                })?;
                let border = borders.get(index).ok_or_else(|| {
                    semantic_error(
                        "use.office.spreadsheet_style_invalid",
                        format!("Spreadsheet style references missing border {index}."),
                    )
                })?;
                values.extend(border.clone());
            }
            if let Some(number_format_id) = style.attribute("numFmtId") {
                values.insert("numberFormatId".into(), number_format_id.into());
                if let Some(format) = custom_formats
                    .get(number_format_id)
                    .map(String::as_str)
                    .or_else(|| built_in_number_format(number_format_id))
                {
                    values.insert("numberFormat".into(), format.into());
                }
            }
            if let Some(alignment) = style.child("alignment") {
                for (attribute, key) in [
                    ("horizontal", "alignment"),
                    ("vertical", "verticalAlignment"),
                    ("textRotation", "textRotation"),
                    ("indent", "indent"),
                ] {
                    if let Some(value) = alignment.attribute(attribute) {
                        values.insert(key.into(), value.into());
                    }
                }
                for (attribute, key) in [("wrapText", "wrapText"), ("shrinkToFit", "shrinkToFit")] {
                    if let Some(value) = alignment.attribute(attribute) {
                        values.insert(key.into(), spreadsheet_attribute_bool(value).to_string());
                    }
                }
                if let Some(value) = alignment.attribute("readingOrder") {
                    match value {
                        "1" => {
                            values.insert("readingOrder".into(), "ltr".into());
                        }
                        "2" => {
                            values.insert("readingOrder".into(), "rtl".into());
                        }
                        _ => {}
                    }
                }
            }
            styles.push(values);
        }
    }
    if styles.is_empty() {
        Ok(vec![BTreeMap::new()])
    } else {
        Ok(styles)
    }
}

pub(super) fn read_differential_formats(
    package: &NativeOfficePackage,
) -> UseResult<Vec<DifferentialFormat>> {
    if !package.contains_part("xl/styles.xml") {
        return Ok(Vec::new());
    }
    let part = package.xml_part("xl/styles.xml")?;
    let root = parse_xml_tree(&part)?;
    require_spreadsheet_element(&root, "styleSheet", part.name())?;
    let Some(dxfs) = root.child("dxfs") else {
        return Ok(Vec::new());
    };
    let rules = dxfs.children_named("dxf").collect::<Vec<_>>();
    if rules.len() > 65_000 {
        return Err(semantic_error(
            "use.office.spreadsheet_conditional_format_dxf_limit",
            "Spreadsheet styles contain more than 65000 differential formats.",
        ));
    }
    Ok(rules
        .into_iter()
        .map(|dxf| {
            let mut values = BTreeMap::new();
            let mut supported = true;
            for child in dxf.child_elements() {
                match child.local_name.as_str() {
                    "font" => {
                        let font = read_font(child);
                        if font
                            .keys()
                            .any(|key| !matches!(key.as_str(), "bold" | "color"))
                        {
                            supported = false;
                        }
                        if let Some(value) = font.get("bold") {
                            values.insert("fontBold".into(), value.clone());
                        }
                        if let Some(value) = font.get("color") {
                            values.insert("fontColor".into(), value.clone());
                        }
                    }
                    "fill" => match read_fill(child) {
                        Some(value) => {
                            values.insert("fill".into(), value);
                        }
                        None => supported = false,
                    },
                    _ => supported = false,
                }
            }
            DifferentialFormat { values, supported }
        })
        .collect())
}

fn read_border(border: &XmlElement) -> BTreeMap<String, String> {
    let mut values = BTreeMap::new();
    for (name, key) in [
        ("left", "borderLeft"),
        ("right", "borderRight"),
        ("top", "borderTop"),
        ("bottom", "borderBottom"),
        ("diagonal", "borderDiagonal"),
    ] {
        let Some(line) = border.child(name) else {
            continue;
        };
        let Some(style) = line.attribute("style").filter(|style| *style != "none") else {
            continue;
        };
        values.insert(key.into(), style.into());
        if let Some(color) = line
            .child("color")
            .and_then(|color| color.attribute("rgb"))
            .and_then(normalize_font_rgb)
        {
            values.insert(format!("{key}Color"), color);
        }
    }
    for (attribute, key) in [
        ("diagonalUp", "borderDiagonalUp"),
        ("diagonalDown", "borderDiagonalDown"),
    ] {
        if let Some(value) = border.attribute(attribute) {
            values.insert(key.into(), spreadsheet_attribute_bool(value).to_string());
        }
    }
    values
}

fn read_fill(fill: &XmlElement) -> Option<String> {
    let pattern = fill.child("patternFill")?;
    if pattern.attribute("patternType") != Some("solid") {
        return None;
    }
    pattern
        .child("fgColor")
        .and_then(|color| color.attribute("rgb"))
        .and_then(normalize_font_rgb)
}

fn read_font(font: &XmlElement) -> BTreeMap<String, String> {
    let mut values = BTreeMap::new();
    for (child_name, key) in [("b", "bold"), ("i", "italic"), ("strike", "strike")] {
        if let Some(property) = font.child(child_name) {
            values.insert(key.into(), spreadsheet_bool_value(property).to_string());
        }
    }
    if let Some(underline) = font.child("u") {
        values.insert(
            "underline".into(),
            underline.attribute("val").unwrap_or("single").into(),
        );
    }
    if let Some(script) = font
        .child("vertAlign")
        .and_then(|property| property.attribute("val"))
    {
        values.insert("script".into(), script.into());
    }
    if let Some(name) = font.child("name").and_then(|name| name.attribute("val")) {
        values.insert("font".into(), name.into());
    }
    if let Some(size) = font
        .child("sz")
        .and_then(|size| size.attribute("val"))
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| value.is_finite())
    {
        values.insert("size".into(), format!("{size}pt"));
    }
    if let Some(rgb) = font
        .child("color")
        .and_then(|color| color.attribute("rgb"))
        .and_then(normalize_font_rgb)
    {
        values.insert("color".into(), rgb);
    }
    values
}

fn spreadsheet_bool_value(element: &XmlElement) -> bool {
    !matches!(
        element
            .attribute("val")
            .map(str::to_ascii_lowercase)
            .as_deref(),
        Some("false" | "0" | "off" | "no")
    )
}

fn spreadsheet_attribute_bool(value: &str) -> bool {
    !matches!(
        value.to_ascii_lowercase().as_str(),
        "false" | "0" | "off" | "no"
    )
}

fn normalize_font_rgb(value: &str) -> Option<String> {
    if !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    match value.len() {
        6 => Some(value.to_ascii_uppercase()),
        8 => Some(value[2..].to_ascii_uppercase()),
        _ => None,
    }
}

fn built_in_number_format(id: &str) -> Option<&'static str> {
    match id {
        "0" => Some("General"),
        "1" => Some("0"),
        "2" => Some("0.00"),
        "3" => Some("#,##0"),
        "4" => Some("#,##0.00"),
        "9" => Some("0%"),
        "10" => Some("0.00%"),
        "11" => Some("0.00E+00"),
        "12" => Some("# ?/?"),
        "13" => Some("# ??/??"),
        "14" => Some("mm-dd-yy"),
        "15" => Some("d-mmm-yy"),
        "16" => Some("d-mmm"),
        "17" => Some("mmm-yy"),
        "18" => Some("h:mm AM/PM"),
        "19" => Some("h:mm:ss AM/PM"),
        "20" => Some("h:mm"),
        "21" => Some("h:mm:ss"),
        "22" => Some("m/d/yy h:mm"),
        "49" => Some("@"),
        _ => None,
    }
}
