use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{
    add_external_relationship, ensure_namespace, existing_external_relationship,
    old_relationship_id, remove_relationship_if_unused,
};
use crate::editor::part::dialect;
use crate::editor::{editor_error, node_not_found, prefix, qualified};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{
    apply_patches, index_xml, insert_child, insert_ordered_child, patch_start_tag_attributes,
    IndexedXmlElement, XmlPatch,
};
use crate::{
    LosslessXmlPart, NativeOfficeDocument, NativeOfficeHyperlink, NativeOfficeHyperlinkTarget,
    NativeOfficePackage, OfficeNodeType,
};

const HYPERLINKS_LATER_SIBLINGS: &[&str] = &[
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    hyperlink: &NativeOfficeHyperlink,
) -> UseResult<String> {
    let initial = NativeOfficeDocument::from_package(package.clone())?;
    let (requested_sheet, requested_range) = requested_scope(&initial, path)?;
    let reference = requested_range.a1();
    let requested_cell = requested_range
        .is_single_cell()
        .then(|| format!("/{requested_sheet}/{reference}"));
    if requested_cell
        .as_deref()
        .is_some_and(|cell| initial.get(cell, 0).is_err())
    {
        let display = hyperlink
            .display
            .as_deref()
            .unwrap_or_else(|| hyperlink.default_display());
        crate::editor::spreadsheet::set_text(
            package,
            requested_cell.as_deref().unwrap_or_default(),
            display,
        )?;
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = find_sheet(&snapshot, &requested_sheet).ok_or_else(|| node_not_found(path))?;
    if requested_range.is_single_cell() {
        let cell_path = format!("{}/{}", sheet.path, reference);
        let cell = snapshot.get(&cell_path, 0)?;
        if cell.node_type != OfficeNodeType::Cell {
            return Err(editor_error(
                "use.office.hyperlink_owner_unsupported",
                "Native Spreadsheet single-cell hyperlinks require a cell path.",
            ));
        }
    }
    let owner = sheet.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;

    let original = package.xml_part(owner)?;
    let relationship_namespace = dialect(package)?.relationship_namespace();
    let (bytes, relationship_prefix) = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { .. } => {
            ensure_namespace(&original, "r", relationship_namespace)?
        }
        NativeOfficeHyperlinkTarget::Internal { .. } => (original.raw().to_vec(), String::new()),
    };
    let part = LosslessXmlPart::parse(owner.clone(), bytes)?;
    let index = index_xml(&part)?;
    let hyperlinks = index.child("hyperlinks", 1);
    let existing = hyperlinks.and_then(|links| find_hyperlink(links, requested_range));
    if let Some(hyperlinks) = hyperlinks {
        reject_overlapping_hyperlinks(hyperlinks, requested_range, existing)?;
    }
    let old_id = existing.and_then(old_relationship_id);
    let relationship_id = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { uri } => Some(
            existing_external_relationship(package, owner, old_id.as_deref(), uri)?
                .map(Ok)
                .unwrap_or_else(|| add_external_relationship(package, owner, uri))?,
        ),
        NativeOfficeHyperlinkTarget::Internal { .. } => None,
    };

    let edited = if let Some(existing) = existing {
        update_existing(
            &part,
            existing,
            hyperlink,
            relationship_id.as_deref(),
            &relationship_prefix,
        )?
    } else {
        let fragment = hyperlink_fragment(
            &index,
            &reference,
            hyperlink,
            relationship_id.as_deref(),
            &relationship_prefix,
        );
        if let Some(hyperlinks) = hyperlinks {
            insert_child(&part, hyperlinks, fragment)?
        } else {
            let spreadsheet_prefix = prefix(&index.qualified_name);
            let tag = qualified(spreadsheet_prefix, "hyperlinks");
            insert_ordered_child(
                &part,
                &index,
                format!("<{tag}>{fragment}</{tag}>"),
                HYPERLINKS_LATER_SIBLINGS,
            )?
        }
    };
    package.set_part(owner, edited)?;
    remove_relationship_if_unused(package, owner, old_id.as_deref())?;
    let updated = NativeOfficeDocument::from_package(package.clone())?;
    let updated_sheet =
        find_sheet(&updated, &requested_sheet).ok_or_else(|| node_not_found(path))?;
    find_semantic_hyperlink(updated_sheet, &reference)
        .map(|node| node.path.clone())
        .ok_or_else(|| node_not_found(path))
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    if requested.node_type != OfficeNodeType::Hyperlink {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Native Spreadsheet hyperlink removal requires a hyperlink path.",
        ));
    }
    let reference = requested.format.get("ref").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_hyperlink_invalid",
            "Spreadsheet semantic hyperlink has no range reference.",
        )
    })?;
    let range = CellRange::parse(reference)?;
    let requested_sheet = path
        .trim_start_matches('/')
        .split('/')
        .next()
        .ok_or_else(|| node_not_found(path))?;
    let sheet = find_sheet(&snapshot, requested_sheet).ok_or_else(|| node_not_found(path))?;
    let owner = sheet.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let part = package.xml_part(owner)?;
    let index = index_xml(&part)?;
    let hyperlinks = index
        .child("hyperlinks", 1)
        .ok_or_else(|| node_not_found(path))?;
    let hyperlink = find_hyperlink(hyperlinks, range).ok_or_else(|| node_not_found(path))?;
    let old_id = old_relationship_id(hyperlink);
    let target = if hyperlinks
        .children
        .iter()
        .filter(|child| child.local_name == "hyperlink")
        .count()
        == 1
    {
        hyperlinks.full_range.clone()
    } else {
        hyperlink.full_range.clone()
    };
    package.set_part(
        owner,
        apply_patches(&part, vec![XmlPatch::new(target, Vec::new())])?,
    )?;
    remove_relationship_if_unused(package, owner, old_id.as_deref())
}

pub(super) fn remove_for_range(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let (requested_sheet, requested_reference) = split_cell_path(path)?;
    let removed = CellRange::parse(requested_reference)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path[1..].eq_ignore_ascii_case(requested_sheet)
        })
        .ok_or_else(|| node_not_found(path))?;
    let owner = sheet.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            "Spreadsheet semantic sheet has no source part.",
        )
    })?;
    let part = package.xml_part(owner)?;
    let index = index_xml(&part)?;
    let Some(hyperlinks) = index.child("hyperlinks", 1) else {
        return Ok(());
    };
    let all_links = hyperlinks
        .children
        .iter()
        .filter(|child| child.local_name == "hyperlink")
        .collect::<Vec<_>>();
    let mut selected = Vec::new();
    let mut relationship_ids = Vec::new();
    for hyperlink in &all_links {
        let reference = hyperlink.attributes.get("ref").ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_hyperlink_invalid",
                "Spreadsheet hyperlink has no cell reference.",
            )
        })?;
        let linked = CellRange::parse(reference)?;
        if !ranges_intersect(removed, linked) {
            continue;
        }
        if !range_contains(removed, linked) {
            return Err(editor_error(
                "use.office.hyperlink_range_conflict",
                format!(
                    "Removing Spreadsheet range '{}' would partially intersect hyperlink range '{}'.",
                    removed.a1(),
                    linked.a1()
                ),
            ));
        }
        if let Some(id) = old_relationship_id(hyperlink) {
            relationship_ids.push(id);
        }
        selected.push(*hyperlink);
    }
    if selected.is_empty() {
        return Ok(());
    }
    let patches = if selected.len() == all_links.len() {
        vec![XmlPatch::new(hyperlinks.full_range.clone(), Vec::new())]
    } else {
        selected
            .into_iter()
            .map(|hyperlink| XmlPatch::new(hyperlink.full_range.clone(), Vec::new()))
            .collect()
    };
    package.set_part(owner, apply_patches(&part, patches)?)?;
    super::remove_relationships_if_unused(package, owner, &relationship_ids)
}

fn ranges_intersect(left: CellRange, right: CellRange) -> bool {
    left.start.column <= right.end.column
        && right.start.column <= left.end.column
        && left.start.row <= right.end.row
        && right.start.row <= left.end.row
}

fn range_contains(outer: CellRange, inner: CellRange) -> bool {
    outer.start.column <= inner.start.column
        && outer.start.row <= inner.start.row
        && outer.end.column >= inner.end.column
        && outer.end.row >= inner.end.row
}

fn requested_scope(document: &NativeOfficeDocument, path: &str) -> UseResult<(String, CellRange)> {
    if let Ok(node) = document.get(path, 0) {
        if node.node_type == OfficeNodeType::Hyperlink {
            let reference = node.format.get("ref").ok_or_else(|| {
                editor_error(
                    "use.office.spreadsheet_hyperlink_invalid",
                    "Spreadsheet semantic hyperlink has no range reference.",
                )
            })?;
            let sheet = node
                .path
                .trim_start_matches('/')
                .split('/')
                .next()
                .ok_or_else(|| node_not_found(path))?;
            return Ok((sheet.to_string(), CellRange::parse(reference)?));
        }
    }
    let owner_path = path.strip_suffix("/hyperlink").unwrap_or(path);
    let (sheet, reference) = split_cell_path(owner_path)?;
    Ok((sheet.to_string(), CellRange::parse(reference)?))
}

fn find_sheet<'a>(
    document: &'a NativeOfficeDocument,
    requested: &str,
) -> Option<&'a crate::DocumentNode> {
    document.root().children.iter().find(|node| {
        node.node_type == OfficeNodeType::Worksheet
            && node.path[1..].eq_ignore_ascii_case(requested)
    })
}

fn find_semantic_hyperlink<'a>(
    node: &'a crate::DocumentNode,
    reference: &str,
) -> Option<&'a crate::DocumentNode> {
    if node.node_type == OfficeNodeType::Hyperlink
        && node
            .format
            .get("ref")
            .is_some_and(|value| value == reference)
    {
        return Some(node);
    }
    node.children
        .iter()
        .find_map(|child| find_semantic_hyperlink(child, reference))
}

fn reject_overlapping_hyperlinks(
    hyperlinks: &IndexedXmlElement,
    requested: CellRange,
    existing: Option<&IndexedXmlElement>,
) -> UseResult<()> {
    for hyperlink in hyperlinks
        .children
        .iter()
        .filter(|child| child.local_name == "hyperlink")
    {
        if existing.is_some_and(|existing| existing.full_range == hyperlink.full_range) {
            continue;
        }
        let reference = hyperlink.attributes.get("ref").ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_hyperlink_invalid",
                "Spreadsheet hyperlink has no cell or range reference.",
            )
        })?;
        let linked = CellRange::parse(reference)?;
        if ranges_intersect(requested, linked) {
            return Err(editor_error(
                "use.office.hyperlink_range_conflict",
                format!(
                    "Spreadsheet hyperlink range '{}' overlaps existing hyperlink range '{}'.",
                    requested.a1(),
                    linked.a1()
                ),
            ));
        }
    }
    Ok(())
}

fn split_cell_path(path: &str) -> UseResult<(&str, &str)> {
    let (sheet, reference) = path
        .strip_prefix('/')
        .and_then(|path| path.split_once('/'))
        .ok_or_else(|| node_not_found(path))?;
    if sheet.is_empty() || reference.is_empty() || reference.contains('/') {
        return Err(editor_error(
            "use.office.hyperlink_owner_unsupported",
            "Native Spreadsheet hyperlinks require a cell or rectangular range path such as /Sheet1/A1:C3.",
        ));
    }
    Ok((sheet, reference))
}

fn find_hyperlink(
    hyperlinks: &IndexedXmlElement,
    reference: CellRange,
) -> Option<&IndexedXmlElement> {
    hyperlinks
        .children
        .iter()
        .filter(|child| child.local_name == "hyperlink")
        .find(|child| {
            child
                .attributes
                .get("ref")
                .and_then(|value| CellRange::parse(value).ok())
                .is_some_and(|value| value == reference)
        })
}

fn update_existing(
    part: &LosslessXmlPart,
    existing: &IndexedXmlElement,
    hyperlink: &NativeOfficeHyperlink,
    relationship_id: Option<&str>,
    relationship_prefix: &str,
) -> UseResult<Vec<u8>> {
    let mut updates = removal_updates(existing, &["id", "location", "display", "tooltip"]);
    match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { .. } => {
            let id = relationship_id.ok_or_else(|| {
                editor_error(
                    "use.office.hyperlink_relationship_missing",
                    "Native Spreadsheet external hyperlink has no relationship ID.",
                )
            })?;
            updates.insert(
                qualified(Some(relationship_prefix), "id"),
                Some(id.to_string()),
            );
        }
        NativeOfficeHyperlinkTarget::Internal { location } => {
            updates.insert("location".into(), Some(location.clone()));
        }
    }
    if let Some(display) = &hyperlink.display {
        updates.insert("display".into(), Some(display.clone()));
    }
    if let Some(tooltip) = &hyperlink.tooltip {
        updates.insert("tooltip".into(), Some(tooltip.clone()));
    }
    patch_start_tag_attributes(part, existing, &updates)
}

fn hyperlink_fragment(
    root: &IndexedXmlElement,
    reference: &str,
    hyperlink: &NativeOfficeHyperlink,
    relationship_id: Option<&str>,
    relationship_prefix: &str,
) -> String {
    let tag = qualified(prefix(&root.qualified_name), "hyperlink");
    let target = match &hyperlink.target {
        NativeOfficeHyperlinkTarget::External { .. } => format!(
            " {}=\"{}\"",
            qualified(Some(relationship_prefix), "id"),
            crate::xml_edit::escape_attribute(relationship_id.unwrap_or_default())
        ),
        NativeOfficeHyperlinkTarget::Internal { location } => format!(
            " location=\"{}\"",
            crate::xml_edit::escape_attribute(location)
        ),
    };
    let display = hyperlink
        .display
        .as_ref()
        .map_or_else(String::new, |value| {
            format!(" display=\"{}\"", crate::xml_edit::escape_attribute(value))
        });
    let tooltip = hyperlink
        .tooltip
        .as_ref()
        .map_or_else(String::new, |value| {
            format!(" tooltip=\"{}\"", crate::xml_edit::escape_attribute(value))
        });
    format!(
        "<{tag} ref=\"{}\"{target}{display}{tooltip}/>",
        crate::xml_edit::escape_attribute(reference)
    )
}

fn removal_updates(
    element: &IndexedXmlElement,
    local_names: &[&str],
) -> BTreeMap<String, Option<String>> {
    element
        .qualified_attributes
        .keys()
        .filter(|name| {
            let local = name
                .rsplit_once(':')
                .map_or(name.as_str(), |(_, local)| local);
            local_names.contains(&local)
        })
        .map(|name| (name.clone(), None))
        .collect()
}
