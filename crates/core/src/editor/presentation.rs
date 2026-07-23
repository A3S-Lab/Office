use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    editor_error, escape_attribute, node_not_found, parse_segments, prefix,
    preserve_space_attribute, qualified, NativeOfficeHighlightColor,
    NativeOfficeHorizontalAlignment, NativeOfficeTextCase, NativeOfficeTextFormat,
    NativeOfficeTextScript, NativeOfficeUnderline,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::xml_edit::{
    apply_patches, escape_text, index_xml, insert_child, insert_ordered_child,
    patch_start_tag_attributes, replace_text_descendants, IndexedXmlElement, XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod arrange;
mod table;
mod table_xml;
mod text;

pub(super) use arrange::{copy_node, move_node, swap_nodes};

const CHARACTER_PROPERTY_ORDER: &[&str] = &[
    "ln",
    "noFill",
    "solidFill",
    "gradFill",
    "blipFill",
    "pattFill",
    "grpFill",
    "effectLst",
    "effectDag",
    "highlight",
    "uLnTx",
    "uLn",
    "uFillTx",
    "uFill",
    "latin",
    "ea",
    "cs",
    "sym",
    "hlinkClick",
    "hlinkMouseOver",
    "rtl",
    "extLst",
];

pub(super) fn add_table(
    package: &mut NativeOfficePackage,
    parent: &str,
    rows: usize,
    columns: usize,
) -> UseResult<String> {
    table::add_table(package, parent, rows, columns)
}

pub(super) fn add_table_row(
    package: &mut NativeOfficePackage,
    parent: &str,
    columns: Option<usize>,
) -> UseResult<String> {
    table::add_row(package, parent, columns)
}

pub(super) fn add_table_column(
    package: &mut NativeOfficePackage,
    parent: &str,
    index: Option<usize>,
    text: &str,
) -> UseResult<String> {
    table::add_column(package, parent, index, text)
}

pub(super) fn set_table_column_width(
    package: &mut NativeOfficePackage,
    path: &str,
    width_emu: u64,
) -> UseResult<()> {
    table::set_column_width(package, path, width_emu)
}

pub(super) fn add_table_cell(
    package: &mut NativeOfficePackage,
    parent: &str,
    text: &str,
) -> UseResult<String> {
    table::add_cell(package, parent, text)
}

pub(super) fn add_slide(
    package: &mut NativeOfficePackage,
    parent: &str,
    title: &str,
) -> UseResult<String> {
    if package.kind() != DocumentKind::Presentation {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native add-slide is available only for Presentation documents.",
        ));
    }
    if parent != "/" {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Presentation slides can currently be added only to /.",
        ));
    }
    let layout_part = package
        .part_names()
        .find(|name| name.starts_with("ppt/slideLayouts/slideLayout") && name.ends_with(".xml"))
        .map(str::to_string)
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_layout_missing",
                "Presentation has no slide layout for a native slide mutation.",
            )
        })?;
    let number = (1..=package.limits().max_entries.saturating_add(1))
        .find(|number| !package.contains_part(&format!("ppt/slides/slide{number}.xml")))
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_slide_limit",
                "Presentation has no available native slide part number.",
            )
        })?;
    let slide_part = format!("ppt/slides/slide{number}.xml");
    let slide_relationship_part = format!("ppt/slides/_rels/slide{number}.xml.rels");
    let presentation = package.xml_part("ppt/presentation.xml")?;
    let index = index_xml(&presentation)?;
    let slide_list = index
        .child("sldIdLst", 1)
        .ok_or_else(|| node_not_found("/"))?;
    let position = slide_list
        .children
        .iter()
        .filter(|child| child.local_name == "sldId")
        .count()
        + 1;
    let slide_id = slide_list
        .children
        .iter()
        .filter(|child| child.local_name == "sldId")
        .filter_map(|child| child.qualified_attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(255)
        .max(255)
        .checked_add(1)
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_slide_limit",
                "Presentation slide IDs are exhausted.",
            )
        })?;

    crate::opc_edit::add_content_type_override(
        package,
        &slide_part,
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    )?;
    package.set_part(&slide_part, slide_xml(title).into_bytes())?;
    let layout_target = format!(
        "../{}",
        layout_part.strip_prefix("ppt/").unwrap_or(&layout_part)
    );
    crate::opc_edit::add_relationship(
        package,
        &slide_relationship_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        &layout_target,
    )?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        "ppt/_rels/presentation.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        &format!("slides/slide{number}.xml"),
    )?;
    let slide_tag = qualified(prefix(&slide_list.qualified_name), "sldId");
    let fragment = format!(
        "<{slide_tag} xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" id=\"{slide_id}\" r:id=\"{}\"/>",
        escape_attribute(&relationship_id)
    );
    let edited = insert_child(&presentation, slide_list, fragment)?;
    package.set_part("ppt/presentation.xml", edited)?;
    Ok(format!("/slide[{position}]"))
}

pub(super) fn add_shape(
    package: &mut NativeOfficePackage,
    parent: &str,
    text: &str,
) -> UseResult<String> {
    if package.kind() != DocumentKind::Presentation {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native add-shape is available only for Presentation documents.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide = snapshot.get(parent, 0)?;
    if slide.node_type != OfficeNodeType::Slide {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Native Presentation shapes require a slide parent.",
        ));
    }
    let part_name = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let shape_tree = index
        .descendant("spTree")
        .ok_or_else(|| node_not_found(parent))?;
    let position = shape_tree
        .children
        .iter()
        .filter(|child| child.local_name == "sp")
        .count()
        + 1;
    let mut non_visual = Vec::new();
    shape_tree.descendants_named("cNvPr", &mut non_visual);
    let id = non_visual
        .into_iter()
        .filter_map(|element| element.attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(1)
        .checked_add(1)
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_shape_limit",
                "Presentation shape IDs are exhausted.",
            )
        })?;
    let offset = u32::try_from(position.saturating_sub(1)).map_err(|_| {
        editor_error(
            "use.office.presentation_shape_limit",
            "Presentation shape position does not fit the supported coordinate range.",
        )
    })? % 4;
    let y = 1_828_800_i64 + i64::from(offset) * 1_143_000;
    let fragment = text_shape_xml(TextShape {
        id,
        name: &format!("TextBox {id}"),
        text,
        x: 914_400,
        y,
        width: 10_363_200,
        height: 914_400,
        font_size: 2_000,
    });
    let edited = insert_child(&part, shape_tree, fragment)?;
    package.set_part(part_name, edited)?;
    Ok(format!("{}/shape[{position}]", slide.path))
}

pub(super) fn set_text(package: &mut NativeOfficePackage, path: &str, text: &str) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    let slide_path = path
        .split('/')
        .find(|segment| segment.starts_with("slide["))
        .map(|segment| format!("/{segment}"))
        .ok_or_else(|| {
            editor_error(
                "use.office.mutation_path_unsupported",
                "Presentation set-text requires a slide element path.",
            )
        })?;
    let slide = snapshot.get(&slide_path, 0)?;
    let part_name = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let target = locate_path(&index, &requested.path)?;
    if target.descendant("t").is_none() {
        let edited = text::insert_into_empty_target(&part, target, text)?;
        return package.set_part(part_name, edited);
    }
    let edited = replace_text_descendants(&part, target, "t", text, None)?;
    package.set_part(part_name, edited)
}

pub(super) fn set_text_format(
    package: &mut NativeOfficePackage,
    path: &str,
    format: &NativeOfficeTextFormat,
) -> UseResult<()> {
    if format.strikethrough.is_some() {
        return Err(editor_error(
            "use.office.presentation_strikethrough_unsupported",
            "Native Presentation text formatting does not support strikethrough yet.",
        ));
    }
    if format.double_strikethrough.is_some() {
        return Err(editor_error(
            "use.office.presentation_double_strikethrough_unsupported",
            "Native Presentation text formatting does not support double strikethrough.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    match requested.node_type {
        OfficeNodeType::Paragraph if format.has_character_properties() => {
            return Err(editor_error(
                "use.office.mutation_type_unsupported",
                "Presentation paragraph paths accept alignment only; address a run path for character formatting.",
            ));
        }
        OfficeNodeType::Run if format.alignment.is_some() => {
            return Err(editor_error(
                "use.office.mutation_type_unsupported",
                "Presentation run paths accept character formatting only; address the paragraph for alignment.",
            ));
        }
        OfficeNodeType::Paragraph | OfficeNodeType::Run => {}
        _ => {
            return Err(editor_error(
                "use.office.mutation_type_unsupported",
                "Native Presentation text formatting supports paragraph and run paths.",
            ));
        }
    }
    let slide_path = path
        .split('/')
        .find(|segment| segment.starts_with("slide["))
        .map(|segment| format!("/{segment}"))
        .ok_or_else(|| {
            editor_error(
                "use.office.mutation_path_unsupported",
                "Presentation text formatting requires a slide element path.",
            )
        })?;
    let slide = snapshot.get(&slide_path, 0)?;
    let part_name = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let original = package.xml_part(part_name)?;
    let mut bytes = original.raw().to_vec();
    if let Some(alignment) = format.alignment {
        bytes = set_presentation_alignment(part_name, bytes, path, alignment)?;
    }
    if format.has_character_properties() {
        bytes = ensure_character_properties(part_name, bytes, path)?;
        if let Some(bold) = format.bold {
            bytes =
                set_character_attribute(part_name, bytes, path, "b", if bold { "1" } else { "0" })?;
        }
        if let Some(italic) = format.italic {
            bytes = set_character_attribute(
                part_name,
                bytes,
                path,
                "i",
                if italic { "1" } else { "0" },
            )?;
        }
        if let Some(underline) = format.underline {
            bytes = set_character_attribute(
                part_name,
                bytes,
                path,
                "u",
                presentation_underline(underline),
            )?;
        }
        if let Some(script) = format.script {
            bytes = set_character_attribute(
                part_name,
                bytes,
                path,
                "baseline",
                presentation_baseline(script),
            )?;
        }
        if let Some(text_case) = format.text_case {
            let value = match text_case {
                NativeOfficeTextCase::None => "none",
                NativeOfficeTextCase::SmallCaps => "small",
                NativeOfficeTextCase::AllCaps => "all",
            };
            bytes = set_character_attribute(part_name, bytes, path, "cap", value)?;
        }
        if let Some(language) = &format.language {
            bytes = set_character_attribute(part_name, bytes, path, "lang", language)?;
        }
        if let Some(size) = format.font_size_centipoints {
            bytes = set_character_attribute(part_name, bytes, path, "sz", &size.to_string())?;
        }
        if let Some(family) = &format.font_family {
            for name in ["latin", "ea", "cs"] {
                bytes = set_character_font(part_name, bytes, path, name, family)?;
            }
        }
        if let Some(color) = format.text_color {
            bytes = set_character_color(part_name, bytes, path, &color.hex())?;
        }
        if let Some(highlight) = format.highlight {
            bytes = set_character_highlight(part_name, bytes, path, highlight)?;
        }
    }
    package.set_part(part_name, bytes)
}

fn presentation_underline(underline: NativeOfficeUnderline) -> &'static str {
    match underline {
        NativeOfficeUnderline::None => "none",
        NativeOfficeUnderline::Single => "sng",
        NativeOfficeUnderline::Double => "dbl",
    }
}

fn presentation_baseline(script: NativeOfficeTextScript) -> &'static str {
    match script {
        NativeOfficeTextScript::Baseline => "0",
        NativeOfficeTextScript::Superscript => "30000",
        NativeOfficeTextScript::Subscript => "-25000",
    }
}

fn set_presentation_alignment(
    part_name: &str,
    bytes: Vec<u8>,
    path: &str,
    alignment: NativeOfficeHorizontalAlignment,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let paragraph = locate_path(&index, path)?;
    let properties = if let Some(properties) = paragraph.child("pPr", 1) {
        properties
    } else {
        let tag = qualified(prefix(&paragraph.qualified_name), "pPr");
        let edited = apply_patches(
            &part,
            vec![XmlPatch::new(
                paragraph.content_range.start..paragraph.content_range.start,
                format!("<{tag}/>"),
            )],
        )?;
        return set_presentation_alignment(part_name, edited, path, alignment);
    };
    let value = match alignment {
        NativeOfficeHorizontalAlignment::Left => "l",
        NativeOfficeHorizontalAlignment::Center => "ctr",
        NativeOfficeHorizontalAlignment::Right => "r",
        NativeOfficeHorizontalAlignment::Justify => "just",
    };
    patch_start_tag_attributes(
        &part,
        properties,
        &BTreeMap::from([("algn".to_string(), Some(value.to_string()))]),
    )
}

fn ensure_character_properties(part_name: &str, bytes: Vec<u8>, path: &str) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let run = locate_path(&index, path)?;
    if run.child("rPr", 1).is_some() {
        return Ok(part.raw().to_vec());
    }
    let tag = qualified(prefix(&run.qualified_name), "rPr");
    apply_patches(
        &part,
        vec![XmlPatch::new(
            run.content_range.start..run.content_range.start,
            format!("<{tag}/>"),
        )],
    )
}

fn set_character_attribute(
    part_name: &str,
    bytes: Vec<u8>,
    path: &str,
    name: &str,
    value: &str,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let run = locate_path(&index, path)?;
    let properties = run.child("rPr", 1).ok_or_else(|| {
        editor_error(
            "use.office.presentation_run_properties_missing",
            format!("Presentation run '{path}' has no properties element."),
        )
    })?;
    patch_start_tag_attributes(
        &part,
        properties,
        &BTreeMap::from([(name.to_string(), Some(value.to_string()))]),
    )
}

fn set_character_font(
    part_name: &str,
    bytes: Vec<u8>,
    path: &str,
    name: &str,
    family: &str,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let run = locate_path(&index, path)?;
    let properties = run.child("rPr", 1).ok_or_else(|| {
        editor_error(
            "use.office.presentation_run_properties_missing",
            format!("Presentation run '{path}' has no properties element."),
        )
    })?;
    if let Some(font) = properties.child(name, 1) {
        return patch_start_tag_attributes(
            &part,
            font,
            &BTreeMap::from([("typeface".to_string(), Some(family.to_string()))]),
        );
    }
    let tag = qualified(prefix(&properties.qualified_name), name);
    insert_character_property(
        &part,
        properties,
        name,
        format!(
            "<{tag} typeface=\"{}\"/>",
            crate::xml_edit::escape_attribute(family)
        ),
    )
}

fn set_character_color(
    part_name: &str,
    bytes: Vec<u8>,
    path: &str,
    color: &str,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let run = locate_path(&index, path)?;
    let properties = run.child("rPr", 1).ok_or_else(|| {
        editor_error(
            "use.office.presentation_run_properties_missing",
            format!("Presentation run '{path}' has no properties element."),
        )
    })?;
    let prefix = prefix(&properties.qualified_name);
    let fill_tag = qualified(prefix, "solidFill");
    let color_tag = qualified(prefix, "srgbClr");
    let fragment = format!("<{fill_tag}><{color_tag} val=\"{color}\"/></{fill_tag}>");
    if let Some(fill) = properties.child("solidFill", 1) {
        if let Some(existing_color) = fill.child("srgbClr", 1) {
            return patch_start_tag_attributes(
                &part,
                existing_color,
                &BTreeMap::from([("val".to_string(), Some(color.to_string()))]),
            );
        }
        if let Some(existing_color) = fill.children.first() {
            let replacement = if existing_color.empty {
                format!("<{color_tag} val=\"{color}\"/>").into_bytes()
            } else {
                let content = part
                    .parse_bytes()
                    .get(existing_color.content_range.clone())
                    .ok_or_else(|| {
                        editor_error(
                            "use.office.presentation_color_invalid",
                            "Presentation text color content range is invalid.",
                        )
                    })?;
                let mut replacement = format!("<{color_tag} val=\"{color}\">").into_bytes();
                replacement.extend_from_slice(content);
                replacement.extend_from_slice(format!("</{color_tag}>").as_bytes());
                replacement
            };
            return apply_patches(
                &part,
                vec![XmlPatch::new(
                    existing_color.full_range.clone(),
                    replacement,
                )],
            );
        }
        return insert_child(&part, fill, format!("<{color_tag} val=\"{color}\"/>"));
    }
    if let Some(fill) = properties.children.iter().find(|child| {
        matches!(
            child.local_name.as_str(),
            "noFill" | "gradFill" | "blipFill" | "pattFill" | "grpFill"
        )
    }) {
        return apply_patches(
            &part,
            vec![XmlPatch::new(fill.full_range.clone(), fragment)],
        );
    }
    insert_character_property(&part, properties, "solidFill", fragment)
}

fn set_character_highlight(
    part_name: &str,
    bytes: Vec<u8>,
    path: &str,
    highlight: NativeOfficeHighlightColor,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let run = locate_path(&index, path)?;
    let properties = run.child("rPr", 1).ok_or_else(|| {
        editor_error(
            "use.office.presentation_run_properties_missing",
            format!("Presentation run '{path}' has no properties element."),
        )
    })?;
    let Some(color) = highlight.rgb_hex() else {
        let Some(existing) = properties.child("highlight", 1) else {
            return Ok(part.raw().to_vec());
        };
        return apply_patches(
            &part,
            vec![XmlPatch::new(existing.full_range.clone(), Vec::new())],
        );
    };
    let drawing_prefix = prefix(&properties.qualified_name);
    let highlight_tag = qualified(drawing_prefix, "highlight");
    let color_tag = qualified(drawing_prefix, "srgbClr");
    let color_fragment = format!("<{color_tag} val=\"{color}\"/>");
    let fragment = format!("<{highlight_tag}>{color_fragment}</{highlight_tag}>");
    let Some(existing) = properties.child("highlight", 1) else {
        return insert_character_property(&part, properties, "highlight", fragment);
    };
    if let Some(existing_color) = existing.child("srgbClr", 1) {
        return patch_start_tag_attributes(
            &part,
            existing_color,
            &BTreeMap::from([("val".to_string(), Some(color.to_string()))]),
        );
    }
    if let Some(existing_color) = existing.children.first() {
        return apply_patches(
            &part,
            vec![XmlPatch::new(
                existing_color.full_range.clone(),
                color_fragment,
            )],
        );
    }
    insert_child(&part, existing, color_fragment)
}

fn insert_character_property(
    part: &LosslessXmlPart,
    properties: &IndexedXmlElement,
    name: &str,
    fragment: String,
) -> UseResult<Vec<u8>> {
    let position = CHARACTER_PROPERTY_ORDER
        .iter()
        .position(|candidate| *candidate == name)
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_run_property_invalid",
                format!("Presentation run property '{name}' has no schema position."),
            )
        })?;
    insert_ordered_child(
        part,
        properties,
        fragment,
        &CHARACTER_PROPERTY_ORDER[position + 1..],
    )
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    match requested.node_type {
        OfficeNodeType::Shape => remove_object(package, &snapshot, &requested.path),
        OfficeNodeType::Table
        | OfficeNodeType::TableRow
        | OfficeNodeType::TableColumn
        | OfficeNodeType::TableCell => table::remove(package, &snapshot, &requested),
        OfficeNodeType::Slide => remove_slide(package, &requested.path),
        _ => Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native Presentation remove currently supports slides, shapes, tables, rows, columns, and grid-safe cells.",
        )),
    }
}

fn remove_object(
    package: &mut NativeOfficePackage,
    snapshot: &NativeOfficeDocument,
    path: &str,
) -> UseResult<()> {
    let slide_path = path
        .split('/')
        .find(|segment| segment.starts_with("slide["))
        .map(|segment| format!("/{segment}"))
        .ok_or_else(|| node_not_found(path))?;
    let slide = snapshot.get(&slide_path, 0)?;
    let part_name = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let part = package.xml_part(part_name)?;
    let index = index_xml(&part)?;
    let target = locate_path(&index, path)?;
    let hyperlink_relationships = super::hyperlink::owned_hyperlink_relationship_ids(target);
    let edited = crate::xml_edit::apply_patches(
        &part,
        vec![crate::xml_edit::XmlPatch::new(
            target.full_range.clone(),
            Vec::new(),
        )],
    )?;
    package.set_part(part_name, edited)?;
    super::hyperlink::remove_relationships_if_unused(package, part_name, &hyperlink_relationships)
}

fn remove_slide(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let position = path
        .strip_prefix("/slide[")
        .and_then(|position| position.strip_suffix(']'))
        .and_then(|position| position.parse::<usize>().ok())
        .filter(|position| *position > 0)
        .ok_or_else(|| node_not_found(path))?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide = snapshot.get(path, 0)?;
    let part_name = slide.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let presentation = package.xml_part("ppt/presentation.xml")?;
    let index = index_xml(&presentation)?;
    let slide_list = index
        .child("sldIdLst", 1)
        .ok_or_else(|| node_not_found(path))?;
    let slide_id = slide_list
        .child("sldId", position)
        .ok_or_else(|| node_not_found(path))?;
    let relationship_id = slide_id
        .qualified_attributes
        .iter()
        .find(|(name, _)| name.ends_with(":id"))
        .map(|(_, value)| value.clone())
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_slide_invalid",
                format!("Presentation slide '{path}' has no relationship ID."),
            )
        })?;
    super::comment::remove_presentation_slide_comments(package, &part_name)?;
    let edited = crate::xml_edit::apply_patches(
        &presentation,
        vec![crate::xml_edit::XmlPatch::new(
            slide_id.full_range.clone(),
            Vec::new(),
        )],
    )?;
    crate::opc_edit::remove_relationship(
        package,
        "ppt/_rels/presentation.xml.rels",
        &relationship_id,
    )?;
    crate::opc_edit::remove_content_type_override(package, &part_name)?;
    let (directory, file_name) = part_name.rsplit_once('/').ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            format!("Presentation slide part '{part_name}' has an invalid path."),
        )
    })?;
    let relationship_part = format!("{directory}/_rels/{file_name}.rels");
    package.remove_part(&relationship_part)?;
    if !package.remove_part(&part_name)? {
        return Err(node_not_found(path));
    }
    package.set_part("ppt/presentation.xml", edited)
}

fn slide_xml(title: &str) -> String {
    let shape = if title.is_empty() {
        String::new()
    } else {
        text_shape_xml(TextShape {
            id: 2,
            name: "Title 1",
            text: title,
            x: 914_400,
            y: 457_200,
            width: 10_363_200,
            height: 1_143_000,
            font_size: 3_200,
        })
    };
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>{shape}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )
}

struct TextShape<'a> {
    id: u32,
    name: &'a str,
    text: &'a str,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    font_size: u32,
}

fn text_shape_xml(shape: TextShape<'_>) -> String {
    let text = escape_text(shape.text);
    let space = preserve_space_attribute(&text);
    format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{}\" name=\"{}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"{}\"/><a:t{space}>{text}</a:t></a:r><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>",
        shape.id,
        escape_attribute(shape.name),
        shape.x,
        shape.y,
        shape.width,
        shape.height,
        shape.font_size
    )
}

pub(super) fn locate_path<'a>(
    root: &'a IndexedXmlElement,
    path: &str,
) -> UseResult<&'a IndexedXmlElement> {
    let segments = parse_segments(path)?;
    let mut segments = segments.into_iter();
    let slide = segments.next().ok_or_else(|| node_not_found(path))?;
    if slide.name != "slide" {
        return Err(node_not_found(path));
    }
    let mut current = root
        .descendant("spTree")
        .ok_or_else(|| node_not_found(path))?;
    for segment in segments {
        let position = segment.position.unwrap_or(1);
        current = match segment.name.as_str() {
            "shape" => current
                .child("sp", position)
                .ok_or_else(|| node_not_found(path))?,
            "picture" => current
                .child("pic", position)
                .ok_or_else(|| node_not_found(path))?,
            "group" => current
                .child("grpSp", position)
                .ok_or_else(|| node_not_found(path))?,
            "table" | "tbl" => position
                .checked_sub(1)
                .and_then(|index| {
                    current
                        .children
                        .iter()
                        .filter(|child| {
                            child.local_name == "graphicFrame" && child.descendant("tbl").is_some()
                        })
                        .nth(index)
                })
                .ok_or_else(|| node_not_found(path))?,
            "tr" | "row" => current
                .descendant("tbl")
                .and_then(|table| table.child("tr", position))
                .ok_or_else(|| node_not_found(path))?,
            "tc" | "cell" => current
                .child("tc", position)
                .ok_or_else(|| node_not_found(path))?,
            "paragraph" | "p" => current
                .descendant("txBody")
                .and_then(|body| body.child("p", position))
                .ok_or_else(|| node_not_found(path))?,
            "run" | "r" => current
                .child_any(&["r", "fld"], position)
                .ok_or_else(|| node_not_found(path))?,
            name => {
                return Err(editor_error(
                    "use.office.mutation_path_unsupported",
                    format!(
                        "Presentation path element '{name}' is not supported for native mutation."
                    ),
                ));
            }
        };
    }
    Ok(current)
}
