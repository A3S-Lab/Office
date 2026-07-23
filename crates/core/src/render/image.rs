use a3s_use_core::UseResult;
use base64::Engine as _;

use super::output::BoundedOutput;
use super::render_error;
use crate::editor::inspect_image;
use crate::{
    DocumentNode, NativeOfficeDocument, NativeOfficeImageFormat, RelationshipSource,
    RelationshipTarget,
};

pub(super) struct RenderImage<'a> {
    pub(super) bytes: &'a [u8],
    pub(super) media_type: &'static str,
    pub(super) width_px: u32,
    pub(super) height_px: u32,
}

pub(super) fn resolve<'a>(
    document: &'a NativeOfficeDocument,
    node: &DocumentNode,
    owner_part: Option<&str>,
) -> UseResult<Option<RenderImage<'a>>> {
    let direct_part = node
        .format
        .get("part")
        .map(|part| part.trim_start_matches('/'));
    let part = if let Some(part) = direct_part {
        Some(part)
    } else if let (Some(owner), Some(relationship_id)) =
        (owner_part, node.format.get("relationshipId"))
    {
        let source = RelationshipSource::Part {
            part_name: owner.trim_start_matches('/').to_string(),
        };
        let Some(relationship) = document
            .opc()
            .relationships()
            .relationship(&source, relationship_id)
        else {
            return Ok(None);
        };
        if !relationship.relationship_type.ends_with("/image") {
            return Ok(None);
        }
        match &relationship.target {
            RelationshipTarget::Internal { part_name, .. } => Some(part_name.as_str()),
            RelationshipTarget::External { .. } => None,
        }
    } else {
        None
    };
    let Some(part) = part else {
        return Ok(None);
    };
    let bytes = document.package().part(part)?;
    let metadata = inspect_image(bytes, None).map_err(|error| {
        render_error(
            "use.office.render_image_invalid",
            format!("Native Office semantic preview rejected image part '/{part}': {error}"),
        )
        .with_detail("part", format!("/{part}"))
    })?;
    let media_type = match metadata.format {
        NativeOfficeImageFormat::Png => "image/png",
        NativeOfficeImageFormat::Jpeg => "image/jpeg",
        NativeOfficeImageFormat::Gif => "image/gif",
    };
    Ok(Some(RenderImage {
        bytes,
        media_type,
        width_px: metadata.width_px,
        height_px: metadata.height_px,
    }))
}

pub(super) fn write_data_url(output: &mut BoundedOutput, image: &RenderImage<'_>) -> UseResult<()> {
    let prefix = format!("data:{};base64,", image.media_type);
    let encoded_length = image
        .bytes
        .len()
        .checked_add(2)
        .map(|length| length / 3)
        .and_then(|length| length.checked_mul(4))
        .ok_or_else(|| {
            render_error(
                "use.office.render_output_too_large",
                "Native Office image encoding length overflowed.",
            )
        })?;
    output.ensure_additional(prefix.len().saturating_add(encoded_length))?;
    output.push(&prefix)?;
    let encoded = base64::engine::general_purpose::STANDARD.encode(image.bytes);
    output.push(&encoded)
}
