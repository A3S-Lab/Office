use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};
use base64::Engine as _;

use super::part::{dialect, relationship_part, relative_target, worksheet_drawing, OfficeDialect};
use super::{
    editor_error, escape_attribute, node_not_found, NativeCreatedImage, NativeOfficeImage,
    NativeOfficeImageFormat,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::CellReference;
use crate::xml_edit::{apply_patches, index_xml, insert_child, IndexedXmlElement, XmlPatch};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage, RelationshipSource};

const MAX_IMAGE_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_IMAGE_WIDTH_PX: u32 = 6 * 96;
const EMU_PER_PIXEL: u64 = 9_525;

mod format;

pub(crate) use format::inspect_image;
pub(super) use format::validate_pixel_bounds;

struct PreparedImage {
    bytes: Vec<u8>,
    format: NativeOfficeImageFormat,
    name: Option<String>,
    alt_text: String,
    width_px: u32,
    height_px: u32,
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    parent: &str,
    image: &NativeOfficeImage,
) -> UseResult<NativeCreatedImage> {
    let prepared = prepare_image(package, image)?;
    match package.kind() {
        DocumentKind::Word => add_word(package, parent, prepared),
        DocumentKind::Spreadsheet => add_spreadsheet(package, parent, prepared),
        DocumentKind::Presentation => add_presentation(package, parent, prepared),
    }
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let picture = snapshot.get(path, 0)?;
    if picture.node_type != OfficeNodeType::Picture {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            format!("Office element '{path}' is not a native picture."),
        ));
    }
    match package.kind() {
        DocumentKind::Word => remove_word(package, path),
        DocumentKind::Spreadsheet => remove_spreadsheet(package, &picture),
        DocumentKind::Presentation => remove_presentation(package, &snapshot, path),
    }
}

fn prepare_image(
    package: &NativeOfficePackage,
    image: &NativeOfficeImage,
) -> UseResult<PreparedImage> {
    let max_bytes = usize::try_from(package.limits().max_part_bytes)
        .unwrap_or(usize::MAX)
        .min(MAX_IMAGE_BYTES);
    let max_encoded = max_bytes.saturating_add(2) / 3 * 4 + 4;
    if image.data.len() > max_encoded {
        return Err(image_invalid(format!(
            "Native Office image data exceeds the {max_bytes}-byte decoded limit."
        )));
    }
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(&image.data)
        .map_err(|error| image_invalid(format!("Image data is not valid base64: {error}")))?;
    if bytes.len() > max_bytes {
        return Err(image_invalid(format!(
            "Native Office image data exceeds the {max_bytes}-byte decoded limit."
        )));
    }
    let metadata = inspect_image(&bytes, Some(image.format))?;
    let (width_px, height_px) = requested_dimensions(
        metadata.width_px,
        metadata.height_px,
        image.width_px,
        image.height_px,
    )?;
    let name = image
        .name
        .as_deref()
        .map(str::trim)
        .filter(|name| !name.is_empty())
        .map(str::to_string);
    if let Some(name) = &name {
        validate_metadata_text(name, 255, "name")?;
    }
    let alt_text = image.alt_text.clone().unwrap_or_default();
    validate_metadata_text(&alt_text, 32_767, "alternative text")?;
    Ok(PreparedImage {
        bytes,
        format: metadata.format,
        name,
        alt_text,
        width_px,
        height_px,
    })
}

fn add_word(
    package: &mut NativeOfficePackage,
    parent: &str,
    image: PreparedImage,
) -> UseResult<NativeCreatedImage> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let parent_node = snapshot.get(parent, 0)?;
    if !matches!(
        parent_node.node_type,
        OfficeNodeType::Body | OfficeNodeType::Paragraph | OfficeNodeType::TableCell
    ) || (!parent_node.path.starts_with("/body") && parent_node.path != "/body")
    {
        return Err(editor_error(
            "use.office.image_parent_unsupported",
            "Native Word pictures require /body, a body paragraph, or a body table cell.",
        ));
    }
    let owner = "word/document.xml";
    let office_dialect = dialect(package)?;
    let document = package.xml_part(owner)?;
    let index = index_xml(&document)?;
    let mut properties = Vec::new();
    index.descendants_named("docPr", &mut properties);
    let id = next_non_visual_id(&properties, "use.office.word_picture_limit")?;
    let name = image
        .name
        .clone()
        .unwrap_or_else(|| format!("Picture {id}"));
    let (media_part, relationship_id) =
        add_media_relationship(package, owner, "word/media", &image, office_dialect)?;
    let run = word_picture_xml(
        office_dialect,
        id,
        &name,
        &image.alt_text,
        &relationship_id,
        image.width_px,
        image.height_px,
    );
    let path = insert_word_picture(package, parent, &run)?;
    Ok(created_image(
        path,
        parent,
        owner,
        media_part,
        relationship_id,
        &image,
    ))
}

fn insert_word_picture(
    package: &mut NativeOfficePackage,
    parent: &str,
    run: &str,
) -> UseResult<String> {
    let owner = "word/document.xml";
    let part = package.xml_part(owner)?;
    let index = index_xml(&part)?;
    let target = super::word::locate_word_path(&index, parent)?;
    let (path, edited) = match target.local_name.as_str() {
        "p" => {
            let position = target
                .children
                .iter()
                .filter(|child| child.local_name == "r")
                .count()
                + 1;
            (
                format!("{parent}/r[{position}]"),
                insert_child(&part, target, run)?,
            )
        }
        "body" | "tc" => {
            let position = target
                .children
                .iter()
                .filter(|child| child.local_name == "p")
                .count()
                + 1;
            let paragraph = format!(
                "<w:p xmlns:w=\"{}\">{run}</w:p>",
                dialect(package)?.word_namespace()
            );
            let edited = if target.local_name == "body" {
                if let Some(section) = target
                    .children
                    .iter()
                    .find(|child| child.local_name == "sectPr")
                {
                    apply_patches(
                        &part,
                        vec![XmlPatch::new(
                            section.full_range.start..section.full_range.start,
                            paragraph,
                        )],
                    )?
                } else {
                    insert_child(&part, target, paragraph)?
                }
            } else {
                insert_child(&part, target, paragraph)?
            };
            (format!("{parent}/p[{position}]/r[1]"), edited)
        }
        _ => {
            return Err(editor_error(
                "use.office.image_parent_unsupported",
                "Native Word pictures require /body, a paragraph, or a table cell.",
            ))
        }
    };
    package.set_part(owner, edited)?;
    Ok(path)
}

fn add_spreadsheet(
    package: &mut NativeOfficePackage,
    parent: &str,
    image: PreparedImage,
) -> UseResult<NativeCreatedImage> {
    let (sheet_path, reference) = parent.rsplit_once('/').ok_or_else(|| {
        editor_error(
            "use.office.image_parent_unsupported",
            "Native Spreadsheet pictures require an anchor cell such as /Sheet1/A1.",
        )
    })?;
    if sheet_path.is_empty() || reference.contains(':') {
        return Err(editor_error(
            "use.office.image_parent_unsupported",
            "Native Spreadsheet pictures require one anchor cell such as /Sheet1/A1.",
        ));
    }
    let cell = CellReference::parse(reference).map_err(|_| {
        editor_error(
            "use.office.image_parent_unsupported",
            "Native Spreadsheet pictures require one anchor cell such as /Sheet1/A1.",
        )
    })?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path.eq_ignore_ascii_case(sheet_path)
        })
        .ok_or_else(|| node_not_found(sheet_path))?;
    let worksheet_part = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let office_dialect = dialect(package)?;
    let drawing_part = worksheet_drawing(package, &worksheet_part, office_dialect)?;
    let drawing = package.xml_part(&drawing_part)?;
    let index = index_xml(&drawing)?;
    let picture_count = direct_picture_anchors(&index).len();
    let mut properties = Vec::new();
    index.descendants_named("cNvPr", &mut properties);
    let id = next_non_visual_id(&properties, "use.office.spreadsheet_picture_limit")?;
    let name = image
        .name
        .clone()
        .unwrap_or_else(|| format!("Picture {id}"));
    let (media_part, relationship_id) =
        add_media_relationship(package, &drawing_part, "xl/media", &image, office_dialect)?;
    let fragment = spreadsheet_picture_xml(
        office_dialect,
        id,
        &name,
        &image.alt_text,
        &relationship_id,
        cell.column - 1,
        cell.row - 1,
        image.width_px,
        image.height_px,
    );
    let edited = insert_child(&drawing, &index, fragment)?;
    package.set_part(&drawing_part, edited)?;
    Ok(created_image(
        format!("{}/picture[{}]", sheet.path, picture_count + 1),
        parent,
        &drawing_part,
        media_part,
        relationship_id,
        &image,
    ))
}

fn add_presentation(
    package: &mut NativeOfficePackage,
    parent: &str,
    image: PreparedImage,
) -> UseResult<NativeCreatedImage> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide = snapshot.get(parent, 0)?;
    if slide.node_type != OfficeNodeType::Slide {
        return Err(editor_error(
            "use.office.image_parent_unsupported",
            "Native Presentation pictures require a slide parent such as /slide[1].",
        ));
    }
    let owner = slide.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let office_dialect = dialect(package)?;
    let slide_part = package.xml_part(&owner)?;
    let index = index_xml(&slide_part)?;
    let shape_tree = index
        .descendant("spTree")
        .ok_or_else(|| node_not_found(parent))?;
    let position = shape_tree
        .children
        .iter()
        .filter(|child| child.local_name == "pic")
        .count()
        + 1;
    let mut properties = Vec::new();
    shape_tree.descendants_named("cNvPr", &mut properties);
    let id = next_non_visual_id(&properties, "use.office.presentation_picture_limit")?;
    let name = image
        .name
        .clone()
        .unwrap_or_else(|| format!("Picture {id}"));
    let (media_part, relationship_id) =
        add_media_relationship(package, &owner, "ppt/media", &image, office_dialect)?;
    let fragment = presentation_picture_xml(
        office_dialect,
        id,
        &name,
        &image.alt_text,
        &relationship_id,
        image.width_px,
        image.height_px,
    );
    let edited = insert_child(&slide_part, shape_tree, fragment)?;
    package.set_part(&owner, edited)?;
    Ok(created_image(
        format!("{}/picture[{position}]", slide.path),
        parent,
        &owner,
        media_part,
        relationship_id,
        &image,
    ))
}

fn add_media_relationship(
    package: &mut NativeOfficePackage,
    owner: &str,
    directory: &str,
    image: &PreparedImage,
    office_dialect: OfficeDialect,
) -> UseResult<(String, String)> {
    let media_part = allocate_media_part(package, directory, image.format)?;
    crate::opc_edit::add_content_type_override(package, &media_part, image.format.content_type())?;
    package.set_part(&media_part, image.bytes.clone())?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        &relationship_part(owner),
        &office_dialect.relationship_type("image"),
        &relative_target(owner, &media_part),
    )?;
    Ok((media_part, relationship_id))
}

fn allocate_media_part(
    package: &NativeOfficePackage,
    directory: &str,
    format: NativeOfficeImageFormat,
) -> UseResult<String> {
    (1..=package.limits().max_entries.saturating_add(1))
        .map(|number| format!("{directory}/image{number}.{}", format.extension()))
        .find(|candidate| !package.contains_part(candidate))
        .ok_or_else(|| {
            editor_error(
                "use.office.image_name_exhausted",
                "No available native Office media part name remains.",
            )
        })
}

fn remove_word(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let owner = "word/document.xml";
    let part = package.xml_part(owner)?;
    let index = index_xml(&part)?;
    let target = super::word::locate_word_path(&index, path)?;
    remove_xml_picture(package, owner, &part, target)
}

fn remove_presentation(
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
    let owner = slide.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            "Presentation semantic slide has no source part.",
        )
    })?;
    let part = package.xml_part(&owner)?;
    let index = index_xml(&part)?;
    let target = super::presentation::locate_path(&index, path)?;
    remove_xml_picture(package, &owner, &part, target)
}

fn remove_spreadsheet(
    package: &mut NativeOfficePackage,
    picture: &crate::DocumentNode,
) -> UseResult<()> {
    let owner = picture
        .format
        .get("ownerPart")
        .map(|part| part.trim_start_matches('/').to_string())
        .ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_drawing_invalid",
                "Spreadsheet picture has no source drawing part.",
            )
        })?;
    let position = picture
        .path
        .rsplit_once("/picture[")
        .and_then(|(_, value)| value.strip_suffix(']'))
        .and_then(|value| value.parse::<usize>().ok())
        .filter(|position| *position > 0)
        .ok_or_else(|| node_not_found(&picture.path))?;
    let part = package.xml_part(&owner)?;
    let index = index_xml(&part)?;
    let anchors = direct_picture_anchors(&index);
    let target = anchors
        .get(position - 1)
        .copied()
        .ok_or_else(|| node_not_found(&picture.path))?;
    remove_xml_picture(package, &owner, &part, target)
}

fn remove_xml_picture(
    package: &mut NativeOfficePackage,
    owner: &str,
    part: &LosslessXmlPart,
    target: &IndexedXmlElement,
) -> UseResult<()> {
    let relationship_ids = embedded_relationship_ids(target);
    let edited = apply_patches(
        part,
        vec![XmlPatch::new(target.full_range.clone(), Vec::new())],
    )?;
    package.set_part(owner, edited)?;
    cleanup_removed_images(package, owner, relationship_ids)
}

fn cleanup_removed_images(
    package: &mut NativeOfficePackage,
    owner: &str,
    relationship_ids: BTreeSet<String>,
) -> UseResult<()> {
    if relationship_ids.is_empty() {
        return Ok(());
    }
    let owner_part = package.xml_part(owner)?;
    let owner_index = index_xml(&owner_part)?;
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    let model = package.opc_model()?;
    let removable = relationship_ids
        .into_iter()
        .filter(|id| !contains_image_relationship(&owner_index, id))
        .filter_map(|id| {
            let relationship = model.relationships().relationship(&source, &id)?;
            if !relationship.relationship_type.ends_with("/image") {
                return None;
            }
            let target = relationship.target.internal_part_name().map(str::to_string);
            Some((id, target))
        })
        .collect::<Vec<_>>();
    let relationship_part = relationship_part(owner);
    let targets = removable
        .iter()
        .filter_map(|(_, target)| target.clone())
        .collect::<BTreeSet<_>>();
    for (id, _) in removable {
        crate::opc_edit::remove_relationship(package, &relationship_part, &id)?;
    }
    for target in targets {
        let model = package.opc_model()?;
        let still_referenced = model
            .relationships()
            .relationships()
            .any(|(_, relationship)| {
                relationship.target.internal_part_name() == Some(target.as_str())
            });
        if still_referenced {
            continue;
        }
        if model.content_types().override_for_part(&target).is_some() {
            crate::opc_edit::remove_content_type_override(package, &target)?;
        }
        package.remove_part(&target)?;
    }
    Ok(())
}

fn embedded_relationship_ids(element: &IndexedXmlElement) -> BTreeSet<String> {
    let mut ids = BTreeSet::new();
    collect_embedded_relationship_ids(element, &mut ids);
    ids
}

fn collect_embedded_relationship_ids(element: &IndexedXmlElement, output: &mut BTreeSet<String>) {
    for (name, value) in &element.qualified_attributes {
        if is_image_relationship_attribute(name) {
            output.insert(value.clone());
        }
    }
    for child in &element.children {
        collect_embedded_relationship_ids(child, output);
    }
}

fn contains_image_relationship(element: &IndexedXmlElement, id: &str) -> bool {
    element
        .qualified_attributes
        .iter()
        .any(|(name, value)| is_image_relationship_attribute(name) && value == id)
        || element
            .children
            .iter()
            .any(|child| contains_image_relationship(child, id))
}

fn is_image_relationship_attribute(name: &str) -> bool {
    matches!(
        name.rsplit_once(':').map_or(name, |(_, local)| local),
        "embed" | "link"
    )
}

fn direct_picture_anchors(root: &IndexedXmlElement) -> Vec<&IndexedXmlElement> {
    root.children
        .iter()
        .filter(|child| {
            matches!(
                child.local_name.as_str(),
                "oneCellAnchor" | "twoCellAnchor" | "absoluteAnchor"
            ) && child.descendant("pic").is_some()
        })
        .collect()
}

fn created_image(
    path: String,
    parent: &str,
    owner: &str,
    media_part: String,
    relationship_id: String,
    image: &PreparedImage,
) -> NativeCreatedImage {
    NativeCreatedImage {
        path,
        part: format!("/{media_part}"),
        parent: parent.to_string(),
        owner_part: format!("/{owner}"),
        relationship_id,
        format: image.format,
        width_px: image.width_px,
        height_px: image.height_px,
    }
}

fn word_picture_xml(
    dialect: OfficeDialect,
    id: u32,
    name: &str,
    alt: &str,
    relationship_id: &str,
    width_px: u32,
    height_px: u32,
) -> String {
    let width = emu(width_px);
    let height = emu(height_px);
    format!(
        "<w:r xmlns:w=\"{}\"><w:drawing><wp:inline xmlns:wp=\"{}\" distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{width}\" cy=\"{height}\"/><wp:docPr id=\"{id}\" name=\"{}\" descr=\"{}\"/><a:graphic xmlns:a=\"{}\"><a:graphicData uri=\"{}\"><pic:pic xmlns:pic=\"{}\"><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"{}\" descr=\"{}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r=\"{}\" r:embed=\"{}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>",
        dialect.word_namespace(),
        dialect.word_drawing_namespace(),
        escape_attribute(name),
        escape_attribute(alt),
        dialect.drawing_namespace(),
        dialect.picture_namespace(),
        dialect.picture_namespace(),
        escape_attribute(name),
        escape_attribute(alt),
        dialect.relationship_namespace(),
        escape_attribute(relationship_id),
    )
}

#[allow(clippy::too_many_arguments)]
fn spreadsheet_picture_xml(
    dialect: OfficeDialect,
    id: u32,
    name: &str,
    alt: &str,
    relationship_id: &str,
    column: u32,
    row: u32,
    width_px: u32,
    height_px: u32,
) -> String {
    let width = emu(width_px);
    let height = emu(height_px);
    format!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\"><xdr:from><xdr:col>{column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"{width}\" cy=\"{height}\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"{id}\" name=\"{}\" descr=\"{}\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"{}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>",
        dialect.spreadsheet_drawing_namespace(),
        dialect.drawing_namespace(),
        dialect.relationship_namespace(),
        escape_attribute(name),
        escape_attribute(alt),
        escape_attribute(relationship_id),
    )
}

#[allow(clippy::too_many_arguments)]
fn presentation_picture_xml(
    dialect: OfficeDialect,
    id: u32,
    name: &str,
    alt: &str,
    relationship_id: &str,
    width_px: u32,
    height_px: u32,
) -> String {
    let width = emu(width_px);
    let height = emu(height_px);
    format!(
        "<p:pic xmlns:p=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\"><p:nvPicPr><p:cNvPr id=\"{id}\" name=\"{}\" descr=\"{}\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>",
        dialect.presentation_namespace(),
        dialect.drawing_namespace(),
        dialect.relationship_namespace(),
        escape_attribute(name),
        escape_attribute(alt),
        escape_attribute(relationship_id),
    )
}

fn next_non_visual_id(properties: &[&IndexedXmlElement], error_code: &str) -> UseResult<u32> {
    properties
        .iter()
        .filter_map(|element| element.attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| editor_error(error_code, "Native Office picture IDs are exhausted."))
}

fn emu(pixels: u32) -> u64 {
    u64::from(pixels) * EMU_PER_PIXEL
}

fn requested_dimensions(
    source_width: u32,
    source_height: u32,
    requested_width: Option<u32>,
    requested_height: Option<u32>,
) -> UseResult<(u32, u32)> {
    let dimensions = match (requested_width, requested_height) {
        (Some(width), Some(height)) => (width, height),
        (Some(width), None) => (width, scale_dimension(width, source_height, source_width)?),
        (None, Some(height)) => (
            scale_dimension(height, source_width, source_height)?,
            height,
        ),
        (None, None) => (
            DEFAULT_IMAGE_WIDTH_PX,
            scale_dimension(DEFAULT_IMAGE_WIDTH_PX, source_height, source_width)?,
        ),
    };
    validate_pixel_bounds(dimensions.0, dimensions.1)?;
    Ok(dimensions)
}

fn scale_dimension(value: u32, numerator: u32, denominator: u32) -> UseResult<u32> {
    if value == 0 || denominator == 0 {
        return Err(image_invalid("Image dimensions must be positive integers."));
    }
    let scaled = u64::from(value)
        .checked_mul(u64::from(numerator))
        .and_then(|scaled| scaled.checked_add(u64::from(denominator) / 2))
        .map(|scaled| scaled / u64::from(denominator))
        .ok_or_else(|| image_invalid("Image dimension scaling overflowed."))?
        .max(1);
    u32::try_from(scaled).map_err(|_| image_invalid("Image dimension scaling overflowed."))
}

fn validate_metadata_text(value: &str, max_chars: usize, label: &str) -> UseResult<()> {
    if value.chars().count() > max_chars || value.chars().any(char::is_control) {
        return Err(image_invalid(format!(
            "Image {label} must contain at most {max_chars} non-control characters."
        )));
    }
    Ok(())
}

fn image_invalid(message: impl Into<String>) -> UseError {
    editor_error("use.office.image_invalid", message)
}
