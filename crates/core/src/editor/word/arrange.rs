use a3s_use_core::UseResult;

use super::{editor_error, locate_word_path, node_not_found, DOCUMENT_PART};
use crate::editor::{NativeOfficeInsertPosition, NativeOfficeSwapResult};
use crate::xml_edit::{
    duplicate_element, index_xml, relocate_element, swap_elements, IndexedXmlElement,
};
use crate::NativeOfficePackage;

#[derive(Debug, Clone, Copy)]
struct OrderedChild<'a> {
    element: &'a IndexedXmlElement,
    source: bool,
    copied: bool,
}

pub(in crate::editor) fn move_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    let part = package.xml_part(DOCUMENT_PART)?;
    let index = index_xml(&part)?;
    let (parent_path, parent, source) = resolve_source(&index, path)?;
    require_same_parent(&index, &parent_path, parent, target_parent)?;
    let (insertion, result_path) = resolve_placement(
        &index,
        &parent_path,
        parent,
        source,
        position,
        Operation::Move,
    )?;
    let edited = relocate_element(&part, source, insertion)?;
    package.set_part(DOCUMENT_PART, edited)?;
    Ok(result_path)
}

pub(in crate::editor) fn copy_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
    name: Option<&str>,
) -> UseResult<String> {
    if name.is_some() {
        return Err(editor_error(
            "use.office.mutation_option_unsupported",
            "Word node copy does not accept a worksheet name.",
        ));
    }
    let part = package.xml_part(DOCUMENT_PART)?;
    let index = index_xml(&part)?;
    let (parent_path, parent, source) = resolve_source(&index, path)?;
    require_same_parent(&index, &parent_path, parent, target_parent)?;
    if source.local_name == "tc" {
        return Err(editor_error(
            "use.office.word_copy_unsupported",
            "Native Word table-cell copy is not yet lossless because the table grid must be resized; copy a row or use add --type cell.",
        ));
    }
    require_identity_free(source, path)?;
    let (insertion, result_path) = resolve_placement(
        &index,
        &parent_path,
        parent,
        source,
        position,
        Operation::Copy,
    )?;
    let edited = duplicate_element(&part, source, insertion)?;
    package.set_part(DOCUMENT_PART, edited)?;
    Ok(result_path)
}

pub(in crate::editor) fn swap_nodes(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    let part = package.xml_part(DOCUMENT_PART)?;
    let index = index_xml(&part)?;
    let (first_parent_path, first_parent, first) = resolve_source(&index, path)?;
    let (second_parent_path, second_parent, second) = resolve_source(&index, with)?;
    if first_parent.full_range != second_parent.full_range {
        return Err(editor_error(
            "use.office.mutation_parent_unsupported",
            "Native Word swap requires both nodes to have the same parent.",
        ));
    }
    if first.full_range == second.full_range {
        return Ok(NativeOfficeSwapResult {
            first: path.to_string(),
            second: with.to_string(),
        });
    }

    let mut order = movable_children(first_parent)
        .into_iter()
        .map(|element| OrderedChild {
            element,
            source: element.full_range == first.full_range,
            copied: element.full_range == second.full_range,
        })
        .collect::<Vec<_>>();
    let first_index = order
        .iter()
        .position(|child| child.source)
        .ok_or_else(|| node_not_found(path))?;
    let second_index = order
        .iter()
        .position(|child| child.copied)
        .ok_or_else(|| node_not_found(with))?;
    order.swap(first_index, second_index);
    let first_path = ordered_path(&order, &first_parent_path, |child| child.source)?;
    let second_path = ordered_path(&order, &second_parent_path, |child| child.copied)?;

    let edited = swap_elements(&part, first, second)?;
    package.set_part(DOCUMENT_PART, edited)?;
    Ok(NativeOfficeSwapResult {
        first: first_path,
        second: second_path,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Operation {
    Move,
    Copy,
}

fn resolve_source<'a>(
    root: &'a IndexedXmlElement,
    path: &str,
) -> UseResult<(String, &'a IndexedXmlElement, &'a IndexedXmlElement)> {
    let (parent_path, _) = path.rsplit_once('/').ok_or_else(|| node_not_found(path))?;
    let parent_path = if parent_path.is_empty() {
        "/body"
    } else {
        parent_path
    };
    let source = locate_word_path(root, path)?;
    let parent = locate_parent(root, parent_path)?;
    if !parent
        .children
        .iter()
        .any(|child| child.full_range == source.full_range)
    {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            format!("Word node '{path}' is not a direct child of '{parent_path}'."),
        ));
    }
    require_allowed_child(parent, source)?;
    Ok((parent_path.to_string(), parent, source))
}

fn locate_parent<'a>(root: &'a IndexedXmlElement, path: &str) -> UseResult<&'a IndexedXmlElement> {
    if matches!(path, "/" | "/body") {
        root.child("body", 1).ok_or_else(|| node_not_found("/body"))
    } else {
        locate_word_path(root, path)
    }
}

fn require_same_parent(
    root: &IndexedXmlElement,
    source_parent_path: &str,
    source_parent: &IndexedXmlElement,
    target_parent: Option<&str>,
) -> UseResult<()> {
    let Some(target_parent) = target_parent else {
        return Ok(());
    };
    let target = locate_parent(root, target_parent)?;
    if target.full_range == source_parent.full_range {
        return Ok(());
    }
    Err(editor_error(
        "use.office.mutation_parent_unsupported",
        format!(
            "Native Word move/copy currently requires the source parent '{source_parent_path}'; cross-parent ownership migration is not yet enabled."
        ),
    ))
}

fn require_allowed_child(parent: &IndexedXmlElement, child: &IndexedXmlElement) -> UseResult<()> {
    let allowed = match parent.local_name.as_str() {
        "body" | "tc" => matches!(child.local_name.as_str(), "p" | "tbl"),
        "tbl" => child.local_name == "tr",
        "tr" => child.local_name == "tc",
        "p" => child.local_name == "r",
        _ => false,
    };
    if allowed {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.mutation_type_unsupported",
            format!(
                "Word element '{}' cannot be arranged inside '{}'.",
                child.local_name, parent.local_name
            ),
        ))
    }
}

fn movable_children(parent: &IndexedXmlElement) -> Vec<&IndexedXmlElement> {
    parent
        .children
        .iter()
        .filter(|child| {
            matches!(
                (parent.local_name.as_str(), child.local_name.as_str()),
                ("body" | "tc", "p" | "tbl") | ("tbl", "tr") | ("tr", "tc") | ("p", "r")
            )
        })
        .collect()
}

fn resolve_placement(
    root: &IndexedXmlElement,
    parent_path: &str,
    parent: &IndexedXmlElement,
    source: &IndexedXmlElement,
    position: Option<&NativeOfficeInsertPosition>,
    operation: Operation,
) -> UseResult<(usize, String)> {
    let mut order = movable_children(parent)
        .into_iter()
        .map(|element| OrderedChild {
            element,
            source: element.full_range == source.full_range,
            copied: false,
        })
        .collect::<Vec<_>>();
    let source_index = order
        .iter()
        .position(|child| child.source)
        .ok_or_else(|| node_not_found(parent_path))?;
    if operation == Operation::Move {
        order.remove(source_index);
    }

    let slot = match position {
        Some(NativeOfficeInsertPosition::Index { index }) => {
            let candidates = order
                .iter()
                .enumerate()
                .filter(|(_, child)| child.element.local_name == source.local_name)
                .collect::<Vec<_>>();
            if *index > candidates.len() {
                return Err(editor_error(
                    "use.office.position_invalid",
                    format!(
                        "Word insertion index {index} is outside 0-{} for '{}' siblings.",
                        candidates.len(),
                        source.local_name
                    ),
                ));
            }
            candidates
                .get(*index)
                .map_or(order.len(), |(slot, _)| *slot)
        }
        Some(NativeOfficeInsertPosition::Before { path }) => {
            resolve_anchor_slot(root, parent, source, &order, path, false, operation)?
        }
        Some(NativeOfficeInsertPosition::After { path }) => {
            resolve_anchor_slot(root, parent, source, &order, path, true, operation)?
        }
        None if operation == Operation::Copy => source_index + 1,
        None => order.len(),
    };

    let inserted = OrderedChild {
        element: source,
        source: operation == Operation::Move,
        copied: operation == Operation::Copy,
    };
    let slot = slot.min(order.len());
    let insertion = order.get(slot).map_or_else(
        || parent_end(parent),
        |child| child.element.full_range.start,
    );
    order.insert(slot, inserted);
    let result_path = ordered_path(&order, parent_path, |child| {
        if operation == Operation::Move {
            child.source
        } else {
            child.copied
        }
    })?;
    Ok((insertion, result_path))
}

fn resolve_anchor_slot(
    root: &IndexedXmlElement,
    parent: &IndexedXmlElement,
    source: &IndexedXmlElement,
    order: &[OrderedChild<'_>],
    path: &str,
    after: bool,
    operation: Operation,
) -> UseResult<usize> {
    super::validate_mutation_path(path)?;
    let anchor = locate_word_path(root, path)?;
    require_allowed_child(parent, anchor)?;
    if !parent
        .children
        .iter()
        .any(|child| child.full_range == anchor.full_range)
    {
        return Err(editor_error(
            "use.office.position_parent_mismatch",
            format!("Word anchor '{path}' is not a sibling of the source node."),
        ));
    }
    if operation == Operation::Move && anchor.full_range == source.full_range {
        return Ok(order
            .iter()
            .position(|child| child.element.full_range.start > source.full_range.start)
            .unwrap_or(order.len()));
    }
    let slot = order
        .iter()
        .position(|child| child.element.full_range == anchor.full_range)
        .ok_or_else(|| node_not_found(path))?;
    Ok(slot + usize::from(after))
}

fn parent_end(parent: &IndexedXmlElement) -> usize {
    if parent.local_name == "body" {
        if let Some(section) = parent
            .children
            .iter()
            .find(|child| child.local_name == "sectPr")
        {
            return section.full_range.start;
        }
    }
    parent.content_range.end
}

fn ordered_path(
    order: &[OrderedChild<'_>],
    parent_path: &str,
    selected: impl Fn(&OrderedChild<'_>) -> bool,
) -> UseResult<String> {
    let target = order
        .iter()
        .find(|child| selected(child))
        .ok_or_else(|| node_not_found(parent_path))?;
    let ordinal = order
        .iter()
        .take_while(|child| !selected(child))
        .filter(|child| child.element.local_name == target.element.local_name)
        .count()
        + 1;
    Ok(format!(
        "{parent_path}/{}[{ordinal}]",
        target.element.local_name
    ))
}

fn require_identity_free(element: &IndexedXmlElement, path: &str) -> UseResult<()> {
    const FORBIDDEN_ELEMENTS: &[&str] = &[
        "bookmarkStart",
        "bookmarkEnd",
        "commentRangeStart",
        "commentRangeEnd",
        "commentReference",
        "footnoteReference",
        "endnoteReference",
        "drawing",
        "object",
        "pict",
    ];
    const FORBIDDEN_ATTRIBUTES: &[&str] =
        &["id", "paraId", "textId", "anchor", "editId", "durableId"];
    if FORBIDDEN_ELEMENTS.contains(&element.local_name.as_str())
        || element.qualified_attributes.keys().any(|name| {
            let local = name
                .rsplit_once(':')
                .map_or(name.as_str(), |(_, local)| local);
            FORBIDDEN_ATTRIBUTES.contains(&local)
        })
    {
        return Err(editor_error(
            "use.office.word_copy_identity_unsupported",
            format!(
                "Word node '{path}' contains relationship or document-scoped identity data that cannot yet be cloned losslessly."
            ),
        ));
    }
    for child in &element.children {
        require_identity_free(child, path)?;
    }
    Ok(())
}
