use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use super::{editor_error, escape_attribute, index_xml, prefix, qualified, validate_mutation_path};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{first_intersecting_ranges, CellRange};
use crate::xml_edit::{
    apply_patches, insert_child, insert_ordered_child, patch_start_tag_attributes,
    IndexedXmlElement, XmlPatch,
};
use crate::{
    DocumentKind, LosslessXmlPart, NativeOfficePackage, OpcPackageModel, RelationshipSource,
    RelationshipTarget,
};

const WORKSHEET_CHILDREN_AFTER_MERGES: &[&str] = &[
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
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

struct ResolvedRange {
    sheet_path: String,
    worksheet_part: String,
    range: CellRange,
}

struct ExistingMerge<'a> {
    element: &'a IndexedXmlElement,
    range: CellRange,
    reference: String,
}

pub(super) fn merge_cells(package: &mut NativeOfficePackage, path: &str) -> UseResult<String> {
    let resolved = resolve_range(package, path)?;
    let part = package.xml_part(&resolved.worksheet_part)?;
    let index = index_xml(&part)?;
    let container = merge_container(&index)?;
    let merges = existing_merges(&index, container)?;
    validate_existing_merges(&merges, &resolved.worksheet_part)?;
    let canonical = resolved.range.a1();

    if merges
        .iter()
        .any(|existing| existing.range == resolved.range)
    {
        return Ok(format!("{}/{}", resolved.sheet_path, canonical));
    }
    if let Some(existing) = merges
        .iter()
        .find(|existing| existing.range.intersects(resolved.range))
    {
        return Err(editor_error(
            "use.office.spreadsheet_merge_overlap",
            format!(
                "Merge range '{canonical}' overlaps existing merged range '{}'.",
                existing.reference
            ),
        )
        .with_suggestion("Unmerge the existing range or choose a disjoint range.")
        .with_detail("requested", canonical)
        .with_detail("existing", existing.reference.clone()));
    }
    reject_table_overlap(package, &resolved.worksheet_part, resolved.range)?;

    let merge_cell_name = qualified(prefix(&index.qualified_name), "mergeCell");
    let merge_cell = format!(
        "<{merge_cell_name} ref=\"{}\"/>",
        escape_attribute(&canonical)
    );
    let edited = if let Some(container) = container {
        let inserted = insert_child(&part, container, merge_cell)?;
        update_count(
            &resolved.worksheet_part,
            inserted,
            merges.len().saturating_add(1),
        )?
    } else {
        let container_name = qualified(prefix(&index.qualified_name), "mergeCells");
        let fragment = format!("<{container_name} count=\"1\">{merge_cell}</{container_name}>");
        insert_ordered_child(&part, &index, fragment, WORKSHEET_CHILDREN_AFTER_MERGES)?
    };
    package.set_part(&resolved.worksheet_part, edited)?;
    Ok(format!("{}/{}", resolved.sheet_path, canonical))
}

pub(super) fn unmerge_cells(package: &mut NativeOfficePackage, path: &str) -> UseResult<String> {
    let resolved = resolve_range(package, path)?;
    let part = package.xml_part(&resolved.worksheet_part)?;
    let index = index_xml(&part)?;
    let Some(container) = merge_container(&index)? else {
        return Ok(format!("{}/{}", resolved.sheet_path, resolved.range.a1()));
    };
    let merges = existing_merges(&index, Some(container))?;
    validate_existing_merges(&merges, &resolved.worksheet_part)?;
    let matches = merges
        .iter()
        .filter(|existing| existing.range == resolved.range)
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(invalid_merge_collection(
            &resolved.worksheet_part,
            "contains duplicate merged ranges",
        ));
    }
    let Some(exact) = matches.first() else {
        let intersecting = merges
            .iter()
            .filter(|existing| resolved.range.intersects(existing.range))
            .map(|existing| existing.reference.clone())
            .collect::<Vec<_>>();
        if intersecting.is_empty() {
            return Ok(format!("{}/{}", resolved.sheet_path, resolved.range.a1()));
        }
        return Err(editor_error(
            "use.office.spreadsheet_merge_not_exact",
            format!(
                "Range '{}' does not exactly match an existing merge; it intersects {} merged range(s).",
                resolved.range.a1(),
                intersecting.len()
            ),
        )
        .with_suggestion(format!(
            "Unmerge each exact range, beginning with '{}'.",
            intersecting[0]
        ))
        .with_detail("requested", resolved.range.a1())
        .with_detail("validRanges", intersecting));
    };

    let remaining = merges.len().saturating_sub(1);
    let edited = if remaining == 0 {
        if !container_is_removable(&part, container, exact.element) {
            return Err(editor_error(
                "use.office.spreadsheet_merge_unknown_content",
                "The final merged range cannot be removed without discarding unknown mergeCells data.",
            )
            .with_suggestion(
                "Inspect the worksheet with native raw XML and preserve or relocate the unknown data first.",
            )
            .with_detail("part", resolved.worksheet_part.clone())
            .with_detail("range", exact.reference.clone()));
        }
        apply_patches(
            &part,
            vec![XmlPatch::new(container.full_range.clone(), Vec::new())],
        )?
    } else {
        let removed = apply_patches(
            &part,
            vec![XmlPatch::new(exact.element.full_range.clone(), Vec::new())],
        )?;
        update_count(&resolved.worksheet_part, removed, remaining)?
    };
    package.set_part(&resolved.worksheet_part, edited)?;
    Ok(format!("{}/{}", resolved.sheet_path, resolved.range.a1()))
}

fn resolve_range(package: &NativeOfficePackage, path: &str) -> UseResult<ResolvedRange> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Merged-cell operations are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(path)?;
    let (requested_sheet, reference) = path.rsplit_once('/').ok_or_else(|| {
        editor_error(
            "use.office.mutation_path_unsupported",
            "Merged-cell operations require a range path such as /Sheet1/A1:B2.",
        )
    })?;
    if requested_sheet.is_empty() || reference.is_empty() {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Merged-cell operations require a range path such as /Sheet1/A1:B2.",
        ));
    }
    let range = CellRange::parse(reference)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node.path.eq_ignore_ascii_case(requested_sheet)
        })
        .ok_or_else(|| {
            editor_error(
                "use.office.node_not_found",
                format!("Office semantic path '{requested_sheet}' does not exist."),
            )
            .with_detail("path", requested_sheet)
        })?;
    let worksheet_part = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{}' has no source part.", sheet.path),
        )
    })?;
    Ok(ResolvedRange {
        sheet_path: sheet.path.clone(),
        worksheet_part,
        range,
    })
}

fn merge_container(worksheet: &IndexedXmlElement) -> UseResult<Option<&IndexedXmlElement>> {
    let candidates = worksheet
        .children
        .iter()
        .filter(|child| child.local_name == "mergeCells")
        .collect::<Vec<_>>();
    if candidates.len() > 1 {
        return Err(invalid_merge_collection(
            "<worksheet>",
            "contains multiple mergeCells collections",
        ));
    }
    let Some(container) = candidates.first().copied() else {
        return Ok(None);
    };
    if container.namespace != worksheet.namespace {
        return Err(invalid_merge_collection(
            "<worksheet>",
            "uses mergeCells in an unexpected namespace",
        ));
    }
    Ok(Some(container))
}

fn existing_merges<'a>(
    worksheet: &'a IndexedXmlElement,
    container: Option<&'a IndexedXmlElement>,
) -> UseResult<Vec<ExistingMerge<'a>>> {
    let Some(container) = container else {
        return Ok(Vec::new());
    };
    container
        .children
        .iter()
        .filter(|child| child.local_name == "mergeCell" && child.namespace == worksheet.namespace)
        .map(|element| {
            let reference = element.attributes.get("ref").ok_or_else(|| {
                invalid_merge_collection("<worksheet>", "contains a mergeCell without ref")
            })?;
            let range = CellRange::parse(reference).map_err(|error| {
                invalid_merge_collection(
                    "<worksheet>",
                    format!("contains invalid merge ref '{reference}': {error}"),
                )
            })?;
            Ok(ExistingMerge {
                element,
                range,
                reference: range.a1(),
            })
        })
        .collect()
}

fn validate_existing_merges(merges: &[ExistingMerge<'_>], part_name: &str) -> UseResult<()> {
    let ranges = merges.iter().map(|merge| merge.range).collect::<Vec<_>>();
    if let Some((left, right)) = first_intersecting_ranges(&ranges) {
        return Err(invalid_merge_collection(
            part_name,
            format!(
                "contains overlapping ranges '{}' and '{}'",
                merges[left].reference, merges[right].reference
            ),
        ));
    }
    Ok(())
}

fn reject_table_overlap(
    package: &NativeOfficePackage,
    worksheet_part: &str,
    requested: CellRange,
) -> UseResult<()> {
    let model = OpcPackageModel::read(package)?;
    let source = RelationshipSource::Part {
        part_name: worksheet_part.to_string(),
    };
    for relationship in model.relationships().relationships_from(&source) {
        if !relationship.relationship_type.ends_with("/table") {
            continue;
        }
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            return Err(editor_error(
                "use.office.spreadsheet_table_invalid",
                "Spreadsheet table relationships must be internal.",
            )
            .with_detail("relationshipId", relationship.id.clone()));
        };
        let table_part = package.xml_part(part_name)?;
        let table = index_xml(&table_part)?;
        if table.local_name != "table" {
            return Err(editor_error(
                "use.office.spreadsheet_table_invalid",
                format!("Spreadsheet table part '{part_name}' has an invalid root."),
            ));
        }
        let reference = table.attributes.get("ref").ok_or_else(|| {
            editor_error(
                "use.office.spreadsheet_table_invalid",
                format!("Spreadsheet table part '{part_name}' has no range."),
            )
        })?;
        let table_range = CellRange::parse(reference).map_err(|error| {
            editor_error(
                "use.office.spreadsheet_table_invalid",
                format!(
                    "Spreadsheet table part '{part_name}' has invalid range '{reference}': {error}"
                ),
            )
        })?;
        if requested.intersects(table_range) {
            let name = table
                .attributes
                .get("displayName")
                .or_else(|| table.attributes.get("name"))
                .cloned()
                .unwrap_or_else(|| "(unnamed)".to_string());
            return Err(editor_error(
                "use.office.spreadsheet_merge_table_overlap",
                format!(
                    "Merge range '{}' overlaps Spreadsheet table '{name}' at '{}'.",
                    requested.a1(),
                    table_range.a1()
                ),
            )
            .with_suggestion("Choose a range outside the table or remove the table first.")
            .with_detail("range", requested.a1())
            .with_detail("table", name)
            .with_detail("tableRange", table_range.a1())
            .with_detail("part", part_name.clone()));
        }
    }
    Ok(())
}

fn update_count(part_name: &str, bytes: Vec<u8>, count: usize) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(part_name.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let container = merge_container(&index)?.ok_or_else(|| {
        invalid_merge_collection(part_name, "lost its mergeCells collection during mutation")
    })?;
    patch_start_tag_attributes(
        &part,
        container,
        &BTreeMap::from([("count".to_string(), Some(count.to_string()))]),
    )
}

fn container_is_removable(
    part: &LosslessXmlPart,
    container: &IndexedXmlElement,
    merge: &IndexedXmlElement,
) -> bool {
    if container
        .qualified_attributes
        .keys()
        .any(|name| name != "count")
        || container.children.len() != 1
        || container.children[0].full_range != merge.full_range
    {
        return false;
    }
    let bytes = part.parse_bytes();
    bytes
        .get(container.content_range.start..merge.full_range.start)
        .into_iter()
        .flatten()
        .chain(
            bytes
                .get(merge.full_range.end..container.content_range.end)
                .into_iter()
                .flatten(),
        )
        .all(u8::is_ascii_whitespace)
}

fn invalid_merge_collection(part_name: &str, reason: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_merge_invalid",
        format!("Spreadsheet worksheet merge collection {}.", reason.into()),
    )
    .with_detail("part", part_name)
}
