use a3s_use_core::UseResult;

use super::{container_end, local_name};
use crate::editor::presentation::{editor_error, node_not_found};
use crate::editor::{NativeOfficeInsertPosition, NativeOfficeSwapResult};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::xml_edit::{
    apply_patches, element_fragment, index_xml, relocate_element, swap_elements, IndexedXmlElement,
    XmlPatch,
};
use crate::{LosslessXmlPart, NativeOfficePackage};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ObjectKind {
    Shape,
    Picture,
    Table,
    Chart,
    Connector,
    Group,
}

impl ObjectKind {
    fn segment(self) -> &'static str {
        match self {
            Self::Shape => "shape",
            Self::Picture => "picture",
            Self::Table => "table",
            Self::Chart => "chart",
            Self::Connector => "connector",
            Self::Group => "group",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ObjectPath {
    slide: usize,
    kind: ObjectKind,
    position: usize,
}

impl ObjectPath {
    fn slide_path(self) -> String {
        format!("/slide[{}]", self.slide)
    }
}

#[derive(Debug, Clone, Copy)]
struct ObjectRef<'a> {
    element: &'a IndexedXmlElement,
    kind: ObjectKind,
}

#[derive(Debug, Clone, Copy)]
struct OrderedObject<'a> {
    object: ObjectRef<'a>,
    source: bool,
    copied: bool,
    first: bool,
    second: bool,
}

impl<'a> OrderedObject<'a> {
    fn new(object: ObjectRef<'a>) -> Self {
        Self {
            object,
            source: false,
            copied: false,
            first: false,
            second: false,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Operation {
    Move,
    Copy,
}

pub(super) fn move_object(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    let requested = parse_object_path(path).ok_or_else(|| node_not_found(path))?;
    require_same_slide(requested, target_parent)?;
    let slide_path = requested.slide_path();
    let part_name = resolve_slide_part(package, &slide_path)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let shape_tree = shape_tree(&index, &slide_path)?;
    let objects = top_level_objects(shape_tree);
    let source = find_object(&objects, requested).ok_or_else(|| node_not_found(path))?;
    let source_index = objects
        .iter()
        .position(|object| object.element.full_range == source.element.full_range)
        .ok_or_else(|| node_not_found(path))?;
    let mut order = objects
        .iter()
        .copied()
        .map(OrderedObject::new)
        .collect::<Vec<_>>();
    order[source_index].source = true;
    let source_token = order.remove(source_index);
    let slot = resolve_slot(
        &objects,
        &order,
        requested,
        source_index,
        position,
        Operation::Move,
    )?;
    let insertion = order.get(slot).map_or_else(
        || container_end(shape_tree),
        |token| token.object.element.full_range.start,
    );
    order.insert(slot, source_token);
    let result_path = marked_path(&order, &slide_path, |token| token.source)?;
    let edited = relocate_element(&part, source.element, insertion)?;
    package.set_part(&part_name, edited)?;
    Ok(result_path)
}

pub(super) fn copy_object(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    let requested = parse_object_path(path).ok_or_else(|| node_not_found(path))?;
    require_same_slide(requested, target_parent)?;
    let slide_path = requested.slide_path();
    let part_name = resolve_slide_part(package, &slide_path)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let shape_tree = shape_tree(&index, &slide_path)?;
    let objects = top_level_objects(shape_tree);
    let source = find_object(&objects, requested).ok_or_else(|| node_not_found(path))?;
    require_copyable_shape(source, path)?;
    let source_index = objects
        .iter()
        .position(|object| object.element.full_range == source.element.full_range)
        .ok_or_else(|| node_not_found(path))?;
    let mut order = objects
        .iter()
        .copied()
        .map(OrderedObject::new)
        .collect::<Vec<_>>();
    let slot = resolve_slot(
        &objects,
        &order,
        requested,
        source_index,
        position,
        Operation::Copy,
    )?;
    let insertion = order.get(slot).map_or_else(
        || container_end(shape_tree),
        |token| token.object.element.full_range.start,
    );
    let mut copied = OrderedObject::new(source);
    copied.copied = true;
    order.insert(slot, copied);
    let result_path = marked_path(&order, &slide_path, |token| token.copied)?;
    let id = next_non_visual_id(shape_tree)?;
    let fragment = copied_shape_fragment(&part, source.element, id)?;
    let edited = apply_patches(&part, vec![XmlPatch::new(insertion..insertion, fragment)])?;
    package.set_part(&part_name, edited)?;
    Ok(result_path)
}

pub(super) fn swap_objects(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    let first_path = parse_object_path(path).ok_or_else(|| node_not_found(path))?;
    let second_path = parse_object_path(with).ok_or_else(|| node_not_found(with))?;
    if first_path.slide != second_path.slide {
        return Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Presentation object swap requires both objects to be on the same slide.",
        ));
    }
    let slide_path = first_path.slide_path();
    let part_name = resolve_slide_part(package, &slide_path)?;
    let part = package.xml_part(&part_name)?;
    let index = index_xml(&part)?;
    let shape_tree = shape_tree(&index, &slide_path)?;
    let objects = top_level_objects(shape_tree);
    let first = find_object(&objects, first_path).ok_or_else(|| node_not_found(path))?;
    let second = find_object(&objects, second_path).ok_or_else(|| node_not_found(with))?;
    if first.element.full_range == second.element.full_range {
        return Ok(NativeOfficeSwapResult {
            first: canonical_path(&objects, &slide_path, first)?,
            second: canonical_path(&objects, &slide_path, second)?,
        });
    }

    let mut order = objects
        .iter()
        .copied()
        .map(OrderedObject::new)
        .collect::<Vec<_>>();
    let first_index = order
        .iter()
        .position(|token| token.object.element.full_range == first.element.full_range)
        .ok_or_else(|| node_not_found(path))?;
    let second_index = order
        .iter()
        .position(|token| token.object.element.full_range == second.element.full_range)
        .ok_or_else(|| node_not_found(with))?;
    order[first_index].first = true;
    order[second_index].second = true;
    order.swap(first_index, second_index);
    let first_result = marked_path(&order, &slide_path, |token| token.first)?;
    let second_result = marked_path(&order, &slide_path, |token| token.second)?;
    let edited = swap_elements(&part, first.element, second.element)?;
    package.set_part(&part_name, edited)?;
    Ok(NativeOfficeSwapResult {
        first: first_result,
        second: second_result,
    })
}

fn parse_object_path(path: &str) -> Option<ObjectPath> {
    let path = path.strip_prefix("/slide[")?;
    let (slide, object) = path.split_once("]/")?;
    let slide = slide.parse::<usize>().ok().filter(|value| *value > 0)?;
    let (kind, position) = [
        (ObjectKind::Shape, "shape["),
        (ObjectKind::Picture, "picture["),
        (ObjectKind::Table, "table["),
        (ObjectKind::Chart, "chart["),
        (ObjectKind::Connector, "connector["),
        (ObjectKind::Group, "group["),
    ]
    .into_iter()
    .find_map(|(kind, prefix)| {
        object
            .strip_prefix(prefix)
            .and_then(|position| position.strip_suffix(']'))
            .and_then(|position| position.parse::<usize>().ok())
            .filter(|position| *position > 0)
            .map(|position| (kind, position))
    })?;
    Some(ObjectPath {
        slide,
        kind,
        position,
    })
}

fn resolve_slide_part(package: &NativeOfficePackage, slide_path: &str) -> UseResult<String> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let slide = snapshot.get(slide_path, 0)?;
    if slide.node_type != OfficeNodeType::Slide {
        return Err(node_not_found(slide_path));
    }
    slide.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.presentation_slide_invalid",
            format!("Presentation slide '{slide_path}' has no source part."),
        )
    })
}

fn shape_tree<'a>(
    root: &'a IndexedXmlElement,
    slide_path: &str,
) -> UseResult<&'a IndexedXmlElement> {
    root.descendant("spTree")
        .ok_or_else(|| node_not_found(slide_path))
}

fn top_level_objects(shape_tree: &IndexedXmlElement) -> Vec<ObjectRef<'_>> {
    shape_tree
        .children
        .iter()
        .filter_map(|element| classify_object(element).map(|kind| ObjectRef { element, kind }))
        .collect()
}

fn classify_object(element: &IndexedXmlElement) -> Option<ObjectKind> {
    match element.local_name.as_str() {
        "sp" => Some(ObjectKind::Shape),
        "pic" => Some(ObjectKind::Picture),
        "cxnSp" => Some(ObjectKind::Connector),
        "grpSp" => Some(ObjectKind::Group),
        "graphicFrame" if element.descendant("tbl").is_some() => Some(ObjectKind::Table),
        "graphicFrame" if element.descendant("chart").is_some() => Some(ObjectKind::Chart),
        "graphicFrame" => Some(ObjectKind::Shape),
        _ => None,
    }
}

fn find_object<'a>(objects: &[ObjectRef<'a>], requested: ObjectPath) -> Option<ObjectRef<'a>> {
    objects
        .iter()
        .copied()
        .filter(|object| object.kind == requested.kind)
        .nth(requested.position - 1)
}

fn require_same_slide(requested: ObjectPath, target_parent: Option<&str>) -> UseResult<()> {
    let source_parent = requested.slide_path();
    if target_parent.is_none_or(|parent| parent == source_parent) {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_parent_unsupported",
            format!(
                "Native Presentation object move/copy currently requires the source slide '{source_parent}'; cross-slide relationship migration is not yet enabled."
            ),
        ))
    }
}

fn resolve_slot(
    objects: &[ObjectRef<'_>],
    order: &[OrderedObject<'_>],
    source: ObjectPath,
    source_index: usize,
    position: Option<&NativeOfficeInsertPosition>,
    operation: Operation,
) -> UseResult<usize> {
    match position {
        Some(NativeOfficeInsertPosition::Index { index }) => {
            let candidates = order
                .iter()
                .enumerate()
                .filter(|(_, token)| token.object.kind == source.kind)
                .collect::<Vec<_>>();
            if *index > candidates.len() {
                return Err(editor_error(
                    "use.office.position_invalid",
                    format!(
                        "Presentation object insertion index {index} is outside 0-{} for '{}' siblings.",
                        candidates.len(),
                        source.kind.segment()
                    ),
                ));
            }
            Ok(candidates
                .get(*index)
                .map_or(order.len(), |(slot, _)| *slot))
        }
        Some(NativeOfficeInsertPosition::Before { path }) => {
            object_anchor_slot(objects, order, source, source_index, path, false, operation)
        }
        Some(NativeOfficeInsertPosition::After { path }) => {
            object_anchor_slot(objects, order, source, source_index, path, true, operation)
        }
        None if operation == Operation::Copy => Ok(source_index + 1),
        None => Ok(order.len()),
    }
}

fn object_anchor_slot(
    objects: &[ObjectRef<'_>],
    order: &[OrderedObject<'_>],
    source: ObjectPath,
    source_index: usize,
    anchor: &str,
    after: bool,
    operation: Operation,
) -> UseResult<usize> {
    let anchor_path = parse_object_path(anchor).ok_or_else(|| node_not_found(anchor))?;
    if anchor_path.slide != source.slide {
        return Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Presentation object placement anchor must be on the source slide.",
        ));
    }
    let anchor_object = find_object(objects, anchor_path).ok_or_else(|| node_not_found(anchor))?;
    if operation == Operation::Move
        && objects[source_index].element.full_range == anchor_object.element.full_range
    {
        return Ok(source_index.min(order.len()));
    }
    let slot = order
        .iter()
        .position(|token| token.object.element.full_range == anchor_object.element.full_range)
        .ok_or_else(|| node_not_found(anchor))?;
    Ok(slot + usize::from(after))
}

fn marked_path(
    order: &[OrderedObject<'_>],
    slide_path: &str,
    marker: impl Fn(&OrderedObject<'_>) -> bool,
) -> UseResult<String> {
    let index = order
        .iter()
        .position(marker)
        .ok_or_else(|| node_not_found(slide_path))?;
    let kind = order[index].object.kind;
    let position = order[..=index]
        .iter()
        .filter(|token| token.object.kind == kind)
        .count();
    Ok(format!("{slide_path}/{}[{position}]", kind.segment()))
}

fn canonical_path(
    objects: &[ObjectRef<'_>],
    slide_path: &str,
    selected: ObjectRef<'_>,
) -> UseResult<String> {
    let index = objects
        .iter()
        .position(|object| object.element.full_range == selected.element.full_range)
        .ok_or_else(|| node_not_found(slide_path))?;
    let position = objects[..=index]
        .iter()
        .filter(|object| object.kind == selected.kind)
        .count();
    Ok(format!(
        "{slide_path}/{}[{position}]",
        selected.kind.segment()
    ))
}

fn require_copyable_shape(source: ObjectRef<'_>, path: &str) -> UseResult<()> {
    if source.kind != ObjectKind::Shape
        || source.element.local_name != "sp"
        || contains_copy_unsafe_data(source.element)
    {
        return Err(editor_error(
            "use.office.presentation_object_copy_unsupported",
            format!(
                "Presentation object '{path}' is not a plain relationship-free shape; pictures, tables, charts, placeholders, extension identities, and relationship-owning objects cannot yet be cloned losslessly."
            ),
        ));
    }
    let mut properties = Vec::new();
    source.element.descendants_named("cNvPr", &mut properties);
    if properties.len() != 1 || !properties[0].attributes.contains_key("id") {
        return Err(editor_error(
            "use.office.presentation_object_copy_unsupported",
            format!(
                "Presentation shape '{path}' does not have exactly one reusable non-visual identity record."
            ),
        ));
    }
    Ok(())
}

fn contains_copy_unsafe_data(element: &IndexedXmlElement) -> bool {
    matches!(element.local_name.as_str(), "ph" | "extLst" | "creationId")
        || element.qualified_attributes.keys().any(|attribute| {
            attribute.contains(':') && matches!(local_name(attribute), "id" | "embed" | "link")
        })
        || element.children.iter().any(contains_copy_unsafe_data)
}

fn next_non_visual_id(shape_tree: &IndexedXmlElement) -> UseResult<u32> {
    let mut properties = Vec::new();
    shape_tree.descendants_named("cNvPr", &mut properties);
    properties
        .into_iter()
        .filter_map(|element| element.attributes.get("id"))
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| {
            editor_error(
                "use.office.presentation_object_identity_exhausted",
                "Presentation non-visual object IDs are exhausted.",
            )
        })
}

fn copied_shape_fragment(
    part: &LosslessXmlPart,
    shape: &IndexedXmlElement,
    id: u32,
) -> UseResult<Vec<u8>> {
    let mut properties = Vec::new();
    shape.descendants_named("cNvPr", &mut properties);
    let properties = properties.first().copied().ok_or_else(|| {
        editor_error(
            "use.office.presentation_object_copy_unsupported",
            "Presentation shape has no non-visual identity record.",
        )
    })?;
    let mut attributes = properties.qualified_attributes.clone();
    attributes.insert("id".into(), id.to_string());
    attributes.insert("name".into(), format!("A3S Copy {id}"));
    let attributes = attributes
        .into_iter()
        .map(|(name, value)| format!(" {name}=\"{}\"", quick_xml::escape::escape(&value)))
        .collect::<String>();
    let terminator = if properties.empty { "/>" } else { ">" };
    let replacement =
        format!("<{}{attributes}{terminator}", properties.qualified_name).into_bytes();
    let relative_start = properties
        .start_tag_range
        .start
        .checked_sub(shape.full_range.start)
        .ok_or_else(fragment_error)?;
    let relative_end = properties
        .start_tag_range
        .end
        .checked_sub(shape.full_range.start)
        .ok_or_else(fragment_error)?;
    let fragment = element_fragment(part, shape)?;
    if relative_start > relative_end || relative_end > fragment.len() {
        return Err(fragment_error());
    }
    let mut copied = Vec::with_capacity(
        fragment
            .len()
            .saturating_sub(relative_end - relative_start)
            .saturating_add(replacement.len()),
    );
    copied.extend_from_slice(&fragment[..relative_start]);
    copied.extend_from_slice(&replacement);
    copied.extend_from_slice(&fragment[relative_end..]);
    Ok(copied)
}

fn fragment_error() -> a3s_use_core::UseError {
    editor_error(
        "use.office.presentation_object_copy_unsupported",
        "Presentation shape identity ranges cannot be cloned losslessly.",
    )
}
