use std::path::{Path, PathBuf};

use a3s_use_core::UseResult;
use sha2::{Digest, Sha256};

use super::io::{source_mutated, verify_source_revision};
use crate::layout::{dimension_matches_dpi, layout_error};
use crate::xml_tree::{parse_xml_tree, XmlElement};
use crate::{
    DocumentKind, NativeOfficeDocument, NativeOfficeImage, NativeOfficeImageFormat,
    NativeOfficePackage, NativeOfficeUnit, NativeOfficeUnitInventoryOptions, PackageRevision,
    RelationshipSource, RelationshipTarget,
};

const EMU_PER_MICROMETER: u64 = 36;
const MICROMETERS_PER_INCH_MILLI: u128 = 25_400_000;
const PRESENTATION_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_PRESENTATION_RELATIONSHIPS: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(super) struct ExactPngSlide {
    pub(super) source_path: PathBuf,
    pub(super) png: Vec<u8>,
    pub(super) png_sha256: String,
    pub(super) width_px: u32,
    pub(super) height_px: u32,
    pub(super) surface_width_micrometers: u32,
    pub(super) surface_height_micrometers: u32,
}

pub(super) async fn load_candidate(
    source_path: &Path,
    source_revision: &PackageRevision,
    requested_unit: &NativeOfficeUnit,
) -> UseResult<ExactPngSlide> {
    verify_source_revision(source_path, source_revision).await?;
    let package = NativeOfficePackage::open(source_path)
        .await
        .map_err(|_| source_mutated())?;
    if package.source_revision() != source_revision {
        return Err(source_mutated());
    }
    if package.kind() != DocumentKind::Presentation {
        return Err(layout_unsupported());
    }
    let source_path = package.path().to_path_buf();
    let document =
        NativeOfficeDocument::from_package(package.clone()).map_err(|_| layout_unsupported())?;
    let inventory = document
        .inventory_units(NativeOfficeUnitInventoryOptions::default())
        .map_err(|_| layout_unsupported())?;
    if !inventory.units.iter().any(|unit| unit == requested_unit) {
        return Err(layout_error(
            "use.office.layout_unit_mismatch",
            "The requested PPTX slide identity does not match the immutable source inventory.",
        ));
    }
    let slide = document
        .root()
        .children
        .iter()
        .find(|node| node.path == requested_unit.path)
        .ok_or_else(|| {
            layout_error(
                "use.office.layout_unit_mismatch",
                "The requested PPTX slide semantic path does not exist.",
            )
        })?;
    let slide_part = slide.format.get("part").ok_or_else(layout_unsupported)?;
    let (slide_width_emu, slide_height_emu) = presentation_surface_emu(&package)?;
    let image_part =
        exact_slide_image_part(&document, slide_part, slide_width_emu, slide_height_emu)?;
    let png = package
        .part(&image_part)
        .map_err(|_| layout_unsupported())?
        .to_vec();
    let metadata = NativeOfficeImage::inspect_bytes(&png).map_err(|_| layout_unsupported())?;
    if metadata.format != NativeOfficeImageFormat::Png || !opaque_png(&png) {
        return Err(layout_unsupported());
    }
    if u128::from(metadata.width_px) * u128::from(slide_height_emu)
        != u128::from(metadata.height_px) * u128::from(slide_width_emu)
    {
        return Err(layout_unsupported());
    }
    let surface_width_micrometers = emu_to_micrometers(slide_width_emu)?;
    let surface_height_micrometers = emu_to_micrometers(slide_height_emu)?;
    let dpi_x = dpi_milli(metadata.width_px, surface_width_micrometers);
    let dpi_y = dpi_milli(metadata.height_px, surface_height_micrometers);
    if !dimension_matches_dpi(metadata.width_px, dpi_x, surface_width_micrometers)
        || !dimension_matches_dpi(metadata.height_px, dpi_y, surface_height_micrometers)
    {
        return Err(layout_unsupported());
    }
    let png_sha256 = format!("{:x}", Sha256::digest(&png));
    Ok(ExactPngSlide {
        source_path,
        png,
        png_sha256,
        width_px: metadata.width_px,
        height_px: metadata.height_px,
        surface_width_micrometers,
        surface_height_micrometers,
    })
}

fn presentation_surface_emu(package: &NativeOfficePackage) -> UseResult<(u64, u64)> {
    let part = package
        .xml_part("ppt/presentation.xml")
        .map_err(|_| layout_unsupported())?;
    let root = parse_xml_tree(&part).map_err(|_| layout_unsupported())?;
    let size = root.child("sldSz").ok_or_else(layout_unsupported)?;
    let width = positive_u64(size.attribute("cx"))?;
    let height = positive_u64(size.attribute("cy"))?;
    Ok((width, height))
}

fn exact_slide_image_part(
    document: &NativeOfficeDocument,
    slide_part: &str,
    slide_width_emu: u64,
    slide_height_emu: u64,
) -> UseResult<String> {
    let part = document
        .package()
        .xml_part(slide_part)
        .map_err(|_| layout_unsupported())?;
    let root = parse_xml_tree(&part).map_err(|_| layout_unsupported())?;
    let shape_tree = root
        .child("cSld")
        .and_then(|common| common.child("spTree"))
        .ok_or_else(layout_unsupported)?;
    let mut pictures = shape_tree
        .child_elements()
        .filter(|child| child.local_name == "pic");
    let picture = pictures.next().ok_or_else(layout_unsupported)?;
    if pictures.next().is_some()
        || shape_tree
            .child_elements()
            .any(|child| !matches!(child.local_name.as_str(), "nvGrpSpPr" | "grpSpPr" | "pic"))
    {
        return Err(layout_unsupported());
    }
    require_children(picture, &["nvPicPr", "blipFill", "spPr"])?;
    validate_non_visual_picture(picture)?;
    let properties = picture.child("spPr").ok_or_else(layout_unsupported)?;
    require_children(properties, &["xfrm", "prstGeom"])?;
    let transform = properties.child("xfrm").ok_or_else(layout_unsupported)?;
    if !transform.attributes.is_empty() {
        return Err(layout_unsupported());
    }
    require_children(transform, &["off", "ext"])?;
    let offset = transform.child("off").ok_or_else(layout_unsupported)?;
    let extent = transform.child("ext").ok_or_else(layout_unsupported)?;
    if offset.attributes.len() != 2
        || extent.attributes.len() != 2
        || signed_i64(offset.attribute("x"))? != 0
        || signed_i64(offset.attribute("y"))? != 0
        || positive_u64(extent.attribute("cx"))? != slide_width_emu
        || positive_u64(extent.attribute("cy"))? != slide_height_emu
    {
        return Err(layout_unsupported());
    }
    let geometry = properties
        .child("prstGeom")
        .ok_or_else(layout_unsupported)?;
    if geometry.attribute("prst") != Some("rect")
        || geometry.attributes.len() != 1
        || geometry
            .child_elements()
            .any(|child| child.local_name != "avLst")
    {
        return Err(layout_unsupported());
    }
    let adjustments = geometry.child("avLst").ok_or_else(layout_unsupported)?;
    if !adjustments.attributes.is_empty() || adjustments.child_elements().next().is_some() {
        return Err(layout_unsupported());
    }

    let fill = picture.child("blipFill").ok_or_else(layout_unsupported)?;
    if !fill.attributes.is_empty() {
        return Err(layout_unsupported());
    }
    require_children(fill, &["blip", "stretch"])?;
    let blip = fill.child("blip").ok_or_else(layout_unsupported)?;
    if blip.child_elements().next().is_some()
        || blip.attributes.len() != 1
        || blip.attribute("link").is_some()
    {
        return Err(layout_unsupported());
    }
    let relationship_id = blip
        .attribute_ns(PRESENTATION_RELATIONSHIPS, "embed")
        .or_else(|| blip.attribute_ns(STRICT_PRESENTATION_RELATIONSHIPS, "embed"))
        .ok_or_else(layout_unsupported)?;
    let stretch = fill.child("stretch").ok_or_else(layout_unsupported)?;
    if !stretch.attributes.is_empty() {
        return Err(layout_unsupported());
    }
    require_children(stretch, &["fillRect"])?;
    let fill_rect = stretch.child("fillRect").ok_or_else(layout_unsupported)?;
    if !fill_rect.attributes.is_empty() || fill_rect.child_elements().next().is_some() {
        return Err(layout_unsupported());
    }

    let source = RelationshipSource::Part {
        part_name: slide_part.trim_start_matches('/').to_string(),
    };
    let relationship = document
        .opc()
        .relationships()
        .relationship(&source, relationship_id)
        .ok_or_else(layout_unsupported)?;
    if !relationship.relationship_type.ends_with("/image") {
        return Err(layout_unsupported());
    }
    match &relationship.target {
        RelationshipTarget::Internal {
            part_name,
            fragment: None,
        } => Ok(part_name.clone()),
        _ => Err(layout_unsupported()),
    }
}

fn validate_non_visual_picture(picture: &XmlElement) -> UseResult<()> {
    let non_visual = picture.child("nvPicPr").ok_or_else(layout_unsupported)?;
    if !non_visual.attributes.is_empty() {
        return Err(layout_unsupported());
    }
    require_children(non_visual, &["cNvPr", "cNvPicPr", "nvPr"])?;
    let identity = non_visual.child("cNvPr").ok_or_else(layout_unsupported)?;
    if identity.attributes.iter().any(|attribute| {
        !matches!(
            attribute.local_name.as_str(),
            "id" | "name" | "descr" | "title" | "hidden"
        )
    }) || identity
        .attribute("hidden")
        .is_some_and(|value| !matches!(value, "0" | "false"))
        || identity.child_elements().next().is_some()
    {
        return Err(layout_unsupported());
    }
    for child_name in ["cNvPicPr", "nvPr"] {
        let child = non_visual
            .child(child_name)
            .ok_or_else(layout_unsupported)?;
        if !child.attributes.is_empty() || child.child_elements().next().is_some() {
            return Err(layout_unsupported());
        }
    }
    Ok(())
}

fn require_children(element: &XmlElement, expected: &[&str]) -> UseResult<()> {
    let actual = element
        .child_elements()
        .map(|child| child.local_name.as_str())
        .collect::<Vec<_>>();
    if actual.len() == expected.len()
        && expected
            .iter()
            .all(|name| actual.iter().filter(|actual| *actual == name).count() == 1)
    {
        return Ok(());
    }
    Err(layout_unsupported())
}

fn opaque_png(bytes: &[u8]) -> bool {
    if bytes.len() < 33 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" || &bytes[12..16] != b"IHDR" {
        return false;
    }
    if !matches!(bytes[25], 0 | 2) {
        return false;
    }
    let mut cursor = 8_usize;
    while let Some(header) = bytes.get(cursor..cursor.saturating_add(8)) {
        let length = u32::from_be_bytes(match header[..4].try_into() {
            Ok(length) => length,
            Err(_) => return false,
        });
        let Ok(length) = usize::try_from(length) else {
            return false;
        };
        let Some(end) = cursor
            .checked_add(12)
            .and_then(|value| value.checked_add(length))
        else {
            return false;
        };
        let Some(chunk) = bytes.get(cursor..end) else {
            return false;
        };
        if &chunk[4..8] == b"tRNS" || matches!(&chunk[4..8], b"acTL" | b"fcTL" | b"fdAT") {
            return false;
        }
        cursor = end;
        if &chunk[4..8] == b"IEND" {
            return cursor == bytes.len();
        }
    }
    false
}

fn emu_to_micrometers(value: u64) -> UseResult<u32> {
    let rounded = value
        .checked_add(EMU_PER_MICROMETER / 2)
        .ok_or_else(layout_unsupported)?
        / EMU_PER_MICROMETER;
    u32::try_from(rounded).map_err(|_| layout_unsupported())
}

pub(super) fn dpi_milli(pixels: u32, micrometers: u32) -> u32 {
    let numerator = u128::from(pixels) * MICROMETERS_PER_INCH_MILLI;
    let denominator = u128::from(micrometers);
    u32::try_from((numerator + denominator / 2) / denominator).unwrap_or(u32::MAX)
}

fn positive_u64(value: Option<&str>) -> UseResult<u64> {
    value
        .and_then(|value| value.parse::<u64>().ok())
        .filter(|value| *value > 0)
        .ok_or_else(layout_unsupported)
}

fn signed_i64(value: Option<&str>) -> UseResult<i64> {
    value
        .and_then(|value| value.parse::<i64>().ok())
        .ok_or_else(layout_unsupported)
}

fn layout_unsupported() -> a3s_use_core::UseError {
    layout_error(
        "use.office.layout_unsupported",
        "This source unit is not an opaque PNG that exactly covers one PPTX slide surface.",
    )
}
