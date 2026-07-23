use a3s_use_core::UseResult;

use super::{container_end, local_name, parse_slide_path};
use crate::editor::presentation::{editor_error, node_not_found};
use crate::editor::{NativeOfficeInsertPosition, NativeOfficeSwapResult};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::xml_edit::{
    apply_patches, index_xml, relocate_element, swap_elements, IndexedXmlElement, XmlPatch,
};
use crate::NativeOfficePackage;

const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

pub(super) fn move_slide(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    require_presentation_root(target_parent)?;
    let source_number = parse_slide_path(path).ok_or_else(|| node_not_found(path))?;
    let part = package.xml_part("ppt/presentation.xml")?;
    let index = index_xml(&part)?;
    let slide_list = slide_list(&index)?;
    let slides = slide_elements(slide_list);
    let source_index = source_number
        .checked_sub(1)
        .filter(|index| *index < slides.len())
        .ok_or_else(|| node_not_found(path))?;
    let source = slides[source_index];
    let mut order = slides.clone();
    order.remove(source_index);
    let slot = resolve_slide_slot(&order, source_index, position, SlideOperation::Move)?;
    let insertion = order
        .get(slot)
        .map_or_else(|| container_end(slide_list), |slide| slide.full_range.start);
    let edited = relocate_element(&part, source, insertion)?;
    package.set_part("ppt/presentation.xml", edited)?;
    Ok(format!("/slide[{}]", slot + 1))
}

pub(super) fn copy_slide(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    require_presentation_root(target_parent)?;
    let source_number = parse_slide_path(path).ok_or_else(|| node_not_found(path))?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let source_node = snapshot.get(path, 0)?;
    if source_node.node_type != OfficeNodeType::Slide {
        return Err(node_not_found(path));
    }
    let source_part = source_node.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            format!("Presentation slide '{path}' has no source part."),
        )
    })?;
    require_copyable_slide(package, &source_part, path)?;

    let presentation = package.xml_part("ppt/presentation.xml")?;
    let index = index_xml(&presentation)?;
    let slide_list = slide_list(&index)?;
    let slides = slide_elements(slide_list);
    let source_index = source_number
        .checked_sub(1)
        .filter(|index| *index < slides.len())
        .ok_or_else(|| node_not_found(path))?;
    let source_id = slides[source_index];
    let slot = resolve_slide_slot(&slides, source_index, position, SlideOperation::Copy)?;
    let insertion = slides
        .get(slot)
        .map_or_else(|| container_end(slide_list), |slide| slide.full_range.start);

    let number = (1..=package.limits().max_entries.saturating_add(1))
        .find(|number| !package.contains_part(&format!("ppt/slides/slide{number}.xml")))
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_slide_limit",
                "Presentation has no available slide part number for copy.",
            )
        })?;
    let target_part = format!("ppt/slides/slide{number}.xml");
    let source_relationship_part = relationship_part(&source_part)?;
    let target_relationship_part = relationship_part(&target_part)?;
    crate::opc_edit::add_content_type_override(package, &target_part, SLIDE_CONTENT_TYPE)?;
    let source_bytes = package.part(&source_part)?.to_vec();
    let relationship_bytes = package.part(&source_relationship_part)?.to_vec();
    package.set_part(&target_part, source_bytes)?;
    package.set_part(&target_relationship_part, relationship_bytes)?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        "ppt/_rels/presentation.xml.rels",
        SLIDE_RELATIONSHIP_TYPE,
        &format!("slides/slide{number}.xml"),
    )?;
    let slide_id = slides
        .iter()
        .filter_map(|slide| slide.attributes.get("id"))
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
    let relationship_attribute = source_id
        .qualified_attributes
        .keys()
        .find(|name| name.contains(':') && local_name(name) == "id")
        .cloned()
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_slide_invalid",
                format!("Presentation slide '{path}' has no relationship attribute."),
            )
        })?;
    let tag = source_id.qualified_name.clone();
    let fragment = format!(
        "<{tag} id=\"{slide_id}\" {relationship_attribute}=\"{}\"/>",
        quick_xml::escape::escape(&relationship_id)
    );
    let edited = apply_patches(
        &presentation,
        vec![XmlPatch::new(insertion..insertion, fragment)],
    )?;
    package.set_part("ppt/presentation.xml", edited)?;
    Ok(format!("/slide[{}]", slot + 1))
}

pub(super) fn swap_slides(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    let first_number = parse_slide_path(path).ok_or_else(|| node_not_found(path))?;
    let second_number = parse_slide_path(with).ok_or_else(|| node_not_found(with))?;
    let part = package.xml_part("ppt/presentation.xml")?;
    let index = index_xml(&part)?;
    let slides = slide_elements(slide_list(&index)?);
    let first_index = first_number
        .checked_sub(1)
        .filter(|index| *index < slides.len())
        .ok_or_else(|| node_not_found(path))?;
    let second_index = second_number
        .checked_sub(1)
        .filter(|index| *index < slides.len())
        .ok_or_else(|| node_not_found(with))?;
    if first_index == second_index {
        return Ok(NativeOfficeSwapResult {
            first: path.to_string(),
            second: with.to_string(),
        });
    }
    let edited = swap_elements(&part, slides[first_index], slides[second_index])?;
    package.set_part("ppt/presentation.xml", edited)?;
    Ok(NativeOfficeSwapResult {
        first: format!("/slide[{}]", second_index + 1),
        second: format!("/slide[{}]", first_index + 1),
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SlideOperation {
    Move,
    Copy,
}

fn resolve_slide_slot(
    order: &[&IndexedXmlElement],
    source_index: usize,
    position: Option<&NativeOfficeInsertPosition>,
    operation: SlideOperation,
) -> UseResult<usize> {
    match position {
        Some(NativeOfficeInsertPosition::Index { index }) => {
            if *index > order.len() {
                return Err(editor_error(
                    "use.office.position_invalid",
                    format!(
                        "Presentation slide insertion index {index} is outside 0-{}.",
                        order.len()
                    ),
                ));
            }
            Ok(*index)
        }
        Some(NativeOfficeInsertPosition::Before { path }) => {
            slide_anchor_slot(order, source_index, path, false, operation)
        }
        Some(NativeOfficeInsertPosition::After { path }) => {
            slide_anchor_slot(order, source_index, path, true, operation)
        }
        None if operation == SlideOperation::Copy => Ok(source_index + 1),
        None => Ok(order.len()),
    }
}

fn slide_anchor_slot(
    order: &[&IndexedXmlElement],
    source_index: usize,
    anchor: &str,
    after: bool,
    operation: SlideOperation,
) -> UseResult<usize> {
    let anchor_number = parse_slide_path(anchor).ok_or_else(|| node_not_found(anchor))?;
    let anchor_index = anchor_number
        .checked_sub(1)
        .ok_or_else(|| node_not_found(anchor))?;
    if operation == SlideOperation::Move && anchor_index == source_index {
        return Ok(source_index.min(order.len()));
    }
    let slot = if operation == SlideOperation::Move {
        if anchor_index > order.len() {
            return Err(node_not_found(anchor));
        }
        anchor_index - usize::from(anchor_index > source_index)
    } else {
        if anchor_index >= order.len() {
            return Err(node_not_found(anchor));
        }
        anchor_index
    };
    Ok(slot + usize::from(after))
}

fn require_copyable_slide(
    package: &NativeOfficePackage,
    source_part: &str,
    path: &str,
) -> UseResult<()> {
    let relationship_part = relationship_part(source_part)?;
    if !package.contains_part(&relationship_part) {
        return Err(slide_copy_unsupported(path));
    }
    let relationships = package.xml_part(&relationship_part)?;
    let relationship_index = index_xml(&relationships)?;
    let entries = relationship_index
        .children
        .iter()
        .filter(|child| child.local_name == "Relationship")
        .collect::<Vec<_>>();
    if entries.len() != 1
        || !entries[0]
            .attributes
            .get("Type")
            .is_some_and(|value| value.ends_with("/slideLayout"))
        || entries[0]
            .attributes
            .get("TargetMode")
            .is_some_and(|value| value.eq_ignore_ascii_case("External"))
    {
        return Err(slide_copy_unsupported(path));
    }
    let slide = package.xml_part(source_part)?;
    let slide_index = index_xml(&slide)?;
    if contains_relationship_reference(&slide_index) {
        return Err(slide_copy_unsupported(path));
    }
    Ok(())
}

fn contains_relationship_reference(element: &IndexedXmlElement) -> bool {
    element.qualified_attributes.keys().any(|attribute| {
        attribute.contains(':') && matches!(local_name(attribute), "id" | "embed" | "link")
    }) || element.children.iter().any(contains_relationship_reference)
}

fn slide_copy_unsupported(path: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.presentation_slide_copy_unsupported",
        format!(
            "Presentation slide '{path}' owns notes, media, charts, external links, or another relationship graph that cannot yet be cloned losslessly."
        ),
    )
}

fn slide_list(root: &IndexedXmlElement) -> UseResult<&IndexedXmlElement> {
    root.child("sldIdLst", 1).ok_or_else(|| {
        editor_error(
            "use.office.presentation_slides_missing",
            "Presentation has no slide ID list.",
        )
    })
}

fn slide_elements(slide_list: &IndexedXmlElement) -> Vec<&IndexedXmlElement> {
    slide_list
        .children
        .iter()
        .filter(|child| child.local_name == "sldId")
        .collect()
}

fn require_presentation_root(target_parent: Option<&str>) -> UseResult<()> {
    if target_parent.is_none_or(|parent| matches!(parent, "/" | "/presentation")) {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Presentation slide move/copy requires the presentation root as its parent.",
        ))
    }
}

fn relationship_part(part_name: &str) -> UseResult<String> {
    let (directory, file_name) = part_name.rsplit_once('/').ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            format!("Presentation slide part '{part_name}' has an invalid path."),
        )
    })?;
    Ok(format!("{directory}/_rels/{file_name}.rels"))
}
