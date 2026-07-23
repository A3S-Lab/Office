use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use crate::editor::{
    NativeOfficeHorizontalAlignment, NativeSpreadsheetBorder, NativeSpreadsheetFill,
    NativeSpreadsheetReadingOrder, NativeSpreadsheetVerticalAlignment,
};

use super::{
    editor_error, element_fragment, ensure_styles_part, find_or_append_child, index_xml,
    insert_child, insert_ordered_child, normalize_collection_count, patch_start_tag_attributes,
    prefix, qualified, require_collection_child, styles_invalid, LosslessXmlPart,
    NativeOfficePackage, NativeOfficeTextFormat, NativeSpreadsheetCellFormat, MAX_STYLE_RECORDS,
    STYLES_PART,
};

const STYLE_CHILDREN_AFTER_NUMBER_FORMATS: &[&str] = &[
    "fonts",
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

const STYLE_CHILDREN_AFTER_FILLS: &[&str] = &[
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
];

mod border;

#[derive(Debug, Default)]
pub(super) struct ResolvedCellFormat {
    number_format_id: Option<u32>,
    fill_id: Option<usize>,
}

pub(super) fn resolve(
    package: &mut NativeOfficePackage,
    format: Option<&NativeSpreadsheetCellFormat>,
) -> UseResult<ResolvedCellFormat> {
    let Some(format) = format else {
        return Ok(ResolvedCellFormat::default());
    };
    let number_format_id = format
        .normalized_number_format()?
        .as_deref()
        .map(|code| number_format_id(package, code))
        .transpose()?;
    let fill_id = format
        .fill
        .map(|fill| fill_index(package, fill))
        .transpose()?;
    Ok(ResolvedCellFormat {
        number_format_id,
        fill_id,
    })
}

pub(super) fn derive_xf_fragment(
    part: &LosslessXmlPart,
    style_index: usize,
    font_id: Option<usize>,
    border_id: Option<usize>,
    text_format: Option<&NativeOfficeTextFormat>,
    cell_format: Option<&NativeSpreadsheetCellFormat>,
    resolved: &ResolvedCellFormat,
) -> UseResult<Vec<u8>> {
    let style = require_collection_child(part, "cellXfs", "xf", style_index)?;
    let mut updates = BTreeMap::new();
    if let Some(font_id) = font_id {
        updates.insert("fontId".to_string(), Some(font_id.to_string()));
        updates.insert("applyFont".to_string(), Some("1".to_string()));
    }
    if let Some(number_format_id) = resolved.number_format_id {
        updates.insert("numFmtId".to_string(), Some(number_format_id.to_string()));
        updates.insert("applyNumberFormat".to_string(), Some("1".to_string()));
    }
    if let Some(fill_id) = resolved.fill_id {
        updates.insert("fillId".to_string(), Some(fill_id.to_string()));
        updates.insert("applyFill".to_string(), Some("1".to_string()));
    }
    if let Some(border_id) = border_id {
        updates.insert("borderId".to_string(), Some(border_id.to_string()));
        updates.insert("applyBorder".to_string(), Some("1".to_string()));
    }
    let has_alignment = text_format.is_some_and(|format| format.alignment.is_some())
        || cell_format.is_some_and(NativeSpreadsheetCellFormat::has_alignment_properties);
    if has_alignment {
        updates.insert("applyAlignment".to_string(), Some("1".to_string()));
    }
    let mut bytes = patch_start_tag_attributes(part, &style, &updates)?;
    if has_alignment {
        bytes = patch_alignment(bytes, style_index, text_format, cell_format)?;
    }
    let derived = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let element = require_collection_child(&derived, "cellXfs", "xf", style_index)?;
    Ok(element_fragment(&derived, &element)?.to_vec())
}

pub(super) fn border_index(
    package: &mut NativeOfficePackage,
    base_style: usize,
    border: &NativeSpreadsheetBorder,
) -> UseResult<usize> {
    border::index(package, base_style, border)
}

fn patch_alignment(
    bytes: Vec<u8>,
    style_index: usize,
    text_format: Option<&NativeOfficeTextFormat>,
    cell_format: Option<&NativeSpreadsheetCellFormat>,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(STYLES_PART.to_string(), bytes)?;
    let style = require_collection_child(&part, "cellXfs", "xf", style_index)?;
    let mut updates = BTreeMap::new();
    if let Some(alignment) = text_format.and_then(|format| format.alignment) {
        updates.insert(
            "horizontal".to_string(),
            Some(
                match alignment {
                    NativeOfficeHorizontalAlignment::Left => "left",
                    NativeOfficeHorizontalAlignment::Center => "center",
                    NativeOfficeHorizontalAlignment::Right => "right",
                    NativeOfficeHorizontalAlignment::Justify => "justify",
                }
                .to_string(),
            ),
        );
    }
    if let Some(format) = cell_format {
        if let Some(alignment) = format.vertical_alignment {
            updates.insert(
                "vertical".to_string(),
                Some(vertical_alignment_name(alignment).to_string()),
            );
        }
        if let Some(value) = format.wrap_text {
            updates.insert(
                "wrapText".to_string(),
                Some(if value { "1" } else { "0" }.to_string()),
            );
        }
        if let Some(value) = format.text_rotation {
            updates.insert("textRotation".to_string(), Some(value.to_string()));
        }
        if let Some(value) = format.indent {
            updates.insert("indent".to_string(), Some(value.to_string()));
        }
        if let Some(value) = format.shrink_to_fit {
            updates.insert(
                "shrinkToFit".to_string(),
                Some(if value { "1" } else { "0" }.to_string()),
            );
        }
        if let Some(value) = format.reading_order {
            updates.insert(
                "readingOrder".to_string(),
                match value {
                    NativeSpreadsheetReadingOrder::Context => None,
                    NativeSpreadsheetReadingOrder::LeftToRight => Some("1".to_string()),
                    NativeSpreadsheetReadingOrder::RightToLeft => Some("2".to_string()),
                },
            );
        }
    }
    if let Some(alignment) = style.child("alignment", 1) {
        return patch_start_tag_attributes(&part, alignment, &updates);
    }
    let tag = qualified(prefix(&style.qualified_name), "alignment");
    let attributes = updates
        .into_iter()
        .filter_map(|(name, value)| value.map(|value| format!(" {name}=\"{value}\"")))
        .collect::<String>();
    insert_ordered_child(
        &part,
        &style,
        format!("<{tag}{attributes}/>"),
        &["protection", "extLst"],
    )
}

fn vertical_alignment_name(value: NativeSpreadsheetVerticalAlignment) -> &'static str {
    match value {
        NativeSpreadsheetVerticalAlignment::Top => "top",
        NativeSpreadsheetVerticalAlignment::Center => "center",
        NativeSpreadsheetVerticalAlignment::Bottom => "bottom",
        NativeSpreadsheetVerticalAlignment::Justify => "justify",
        NativeSpreadsheetVerticalAlignment::Distributed => "distributed",
    }
}

fn fill_index(package: &mut NativeOfficePackage, fill: NativeSpreadsheetFill) -> UseResult<usize> {
    ensure_fills(package)?;
    let part = package.xml_part(STYLES_PART)?;
    let root = index_xml(&part)?;
    let fill_prefix = root
        .child("fills", 1)
        .map(|fills| prefix(&fills.qualified_name))
        .ok_or_else(styles_invalid)?;
    let fill_tag = qualified(fill_prefix, "fill");
    let pattern_tag = qualified(fill_prefix, "patternFill");
    let fragment = match fill {
        NativeSpreadsheetFill::None => {
            format!("<{fill_tag}><{pattern_tag} patternType=\"none\"/></{fill_tag}>")
        }
        NativeSpreadsheetFill::Solid { color } => {
            let foreground_tag = qualified(fill_prefix, "fgColor");
            format!(
                "<{fill_tag}><{pattern_tag} patternType=\"solid\"><{foreground_tag} rgb=\"FF{}\"/></{pattern_tag}></{fill_tag}>",
                color.hex()
            )
        }
    };
    find_or_append_child(package, "fills", "fill", fragment.into_bytes())
}

fn ensure_fills(package: &mut NativeOfficePackage) -> UseResult<()> {
    ensure_styles_part(package)?;
    let part = package.xml_part(STYLES_PART)?;
    let root = index_xml(&part)?;
    let bytes = if let Some(fills) = root.child("fills", 1) {
        if fills.child("fill", 1).is_some() {
            return normalize_collection_count(package, "fills", "fill");
        }
        insert_child(&part, fills, default_fills(prefix(&fills.qualified_name)))?
    } else {
        let style_prefix = prefix(&root.qualified_name);
        let fills_tag = qualified(style_prefix, "fills");
        insert_ordered_child(
            &part,
            &root,
            format!(
                "<{fills_tag} count=\"2\">{}</{fills_tag}>",
                default_fills(style_prefix)
            ),
            STYLE_CHILDREN_AFTER_FILLS,
        )?
    };
    package.set_part(STYLES_PART, bytes)?;
    normalize_collection_count(package, "fills", "fill")
}

fn default_fills(prefix: Option<&str>) -> String {
    let fill = qualified(prefix, "fill");
    let pattern = qualified(prefix, "patternFill");
    format!(
        "<{fill}><{pattern} patternType=\"none\"/></{fill}><{fill}><{pattern} patternType=\"gray125\"/></{fill}>"
    )
}

fn number_format_id(package: &mut NativeOfficePackage, code: &str) -> UseResult<u32> {
    if let Some(id) = builtin_number_format_id(code) {
        return Ok(id);
    }
    let part = package.xml_part(STYLES_PART)?;
    let root = index_xml(&part)?;
    if let Some(formats) = root.child("numFmts", 1) {
        for format in formats
            .children
            .iter()
            .filter(|child| child.local_name == "numFmt")
        {
            if format.attributes.get("formatCode").map(String::as_str) == Some(code) {
                let id = format
                    .attributes
                    .get("numFmtId")
                    .ok_or_else(styles_invalid)?
                    .parse::<u32>()
                    .map_err(|_| styles_invalid())?;
                if id < 164 {
                    return Err(styles_invalid());
                }
                return Ok(id);
            }
        }
    }
    append_number_format(package, code)
}

fn append_number_format(package: &mut NativeOfficePackage, code: &str) -> UseResult<u32> {
    let part = package.xml_part(STYLES_PART)?;
    let root = index_xml(&part)?;
    let existing = root.child("numFmts", 1);
    let count = existing
        .map(|formats| {
            formats
                .children
                .iter()
                .filter(|child| child.local_name == "numFmt")
                .count()
        })
        .unwrap_or(0);
    if count >= MAX_STYLE_RECORDS {
        return Err(editor_error(
            "use.office.spreadsheet_style_limit",
            format!("Spreadsheet number formats cannot exceed {MAX_STYLE_RECORDS} records."),
        ));
    }
    let mut maximum_id = 163u32;
    if let Some(formats) = existing {
        for format in formats
            .children
            .iter()
            .filter(|child| child.local_name == "numFmt")
        {
            let id = format
                .attributes
                .get("numFmtId")
                .ok_or_else(styles_invalid)?
                .parse::<u32>()
                .map_err(|_| styles_invalid())?;
            if id < 164 {
                return Err(styles_invalid());
            }
            maximum_id = maximum_id.max(id);
        }
    }
    for collection_name in ["cellStyleXfs", "cellXfs"] {
        if let Some(collection) = root.child(collection_name, 1) {
            for style in collection
                .children
                .iter()
                .filter(|child| child.local_name == "xf")
            {
                if let Some(value) = style.attributes.get("numFmtId") {
                    let id = value.parse::<u32>().map_err(|_| styles_invalid())?;
                    if id >= 164 {
                        maximum_id = maximum_id.max(id);
                    }
                }
            }
        }
    }
    let next_id = maximum_id.checked_add(1).ok_or_else(styles_invalid)?;
    let prefix = existing
        .map(|formats| prefix(&formats.qualified_name))
        .unwrap_or_else(|| prefix(&root.qualified_name));
    let format_tag = qualified(prefix, "numFmt");
    let fragment = format!(
        "<{format_tag} numFmtId=\"{next_id}\" formatCode=\"{}\"/>",
        crate::xml_edit::escape_attribute(code)
    );
    let bytes = if let Some(formats) = existing {
        insert_child(&part, formats, fragment)?
    } else {
        let formats_tag = qualified(prefix, "numFmts");
        insert_ordered_child(
            &part,
            &root,
            format!("<{formats_tag} count=\"1\">{fragment}</{formats_tag}>"),
            STYLE_CHILDREN_AFTER_NUMBER_FORMATS,
        )?
    };
    package.set_part(STYLES_PART, bytes)?;
    normalize_collection_count(package, "numFmts", "numFmt")?;
    Ok(next_id)
}

fn builtin_number_format_id(code: &str) -> Option<u32> {
    match code {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "# ?/?" => Some(12),
        "# ??/??" => Some(13),
        "mm-dd-yy" => Some(14),
        "d-mmm-yy" => Some(15),
        "d-mmm" => Some(16),
        "mmm-yy" => Some(17),
        "h:mm AM/PM" => Some(18),
        "h:mm:ss AM/PM" => Some(19),
        "h:mm" => Some(20),
        "h:mm:ss" => Some(21),
        "m/d/yy h:mm" => Some(22),
        "@" => Some(49),
        _ => None,
    }
}
