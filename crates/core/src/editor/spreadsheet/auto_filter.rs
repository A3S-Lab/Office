use a3s_use_core::UseResult;

use super::{editor_error, filter_xml, validate_mutation_path};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::{apply_patches, index_xml, insert_ordered_child, XmlPatch};
use crate::{
    DocumentKind, NativeOfficeDocument, NativeOfficePackage, NativeSpreadsheetAutoFilter,
    OfficeNodeType,
};

struct ResolvedSheet {
    path: String,
    part: String,
}

struct ResolvedFilter {
    sheet: ResolvedSheet,
    path: String,
}

pub(super) fn is_path(path: &str) -> bool {
    let normalized = path.trim_matches('/');
    normalized
        .rsplit_once('/')
        .is_some_and(|(parent, segment)| {
            !parent.contains('/') && segment.eq_ignore_ascii_case("autofilter")
        })
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    sheet: &str,
    filter: &NativeSpreadsheetAutoFilter,
) -> UseResult<String> {
    let resolved = resolve_sheet(package, sheet)?;
    let mut filter = filter.clone();
    let range = filter.validate()?;
    filter.range = range.a1();
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    if snapshot
        .get(&resolved.path, 1)?
        .children
        .iter()
        .any(|node| node.node_type == OfficeNodeType::AutoFilter)
    {
        return Err(editor_error(
            "use.office.spreadsheet_filter_exists",
            format!("Worksheet '{}' already has an AutoFilter.", resolved.path),
        ));
    }
    validate_geometry(package, &snapshot, &resolved, range)?;
    let worksheet = package.xml_part(&resolved.part)?;
    let root = index_xml(&worksheet)?;
    let prefix = root
        .qualified_name
        .rsplit_once(':')
        .map(|(prefix, _)| prefix);
    let edited = insert_ordered_child(
        &worksheet,
        &root,
        filter_xml::fragment(prefix, &filter)?,
        &[
            "sortState",
            "dataConsolidate",
            "customSheetViews",
            "mergeCells",
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
        ],
    )?;
    package.set_part(&resolved.part, edited)?;
    Ok(format!("{}/autofilter", resolved.path))
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    filter: &NativeSpreadsheetAutoFilter,
) -> UseResult<String> {
    let resolved = resolve_filter(package, path)?;
    let mut filter = filter.clone();
    let range = filter.validate()?;
    filter.range = range.a1();
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(&resolved.path, 0)?;
    require_mutable(&node)?;
    validate_geometry(package, &snapshot, &resolved.sheet, range)?;
    let worksheet = package.xml_part(&resolved.sheet.part)?;
    let root = index_xml(&worksheet)?;
    let existing = direct_filters(&root);
    if existing.len() != 1 {
        return Err(filter_part_error(
            &resolved.sheet.part,
            "does not contain exactly one worksheet AutoFilter",
        ));
    }
    filter_xml::validate_mutable(&worksheet, existing[0])?;
    let prefix = root
        .qualified_name
        .rsplit_once(':')
        .map(|(prefix, _)| prefix);
    let edited = apply_patches(
        &worksheet,
        vec![XmlPatch::new(
            existing[0].full_range.clone(),
            filter_xml::fragment(prefix, &filter)?,
        )],
    )?;
    package.set_part(&resolved.sheet.part, edited)?;
    Ok(resolved.path)
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    let resolved = resolve_filter(package, path)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    require_mutable(&snapshot.get(&resolved.path, 0)?)?;
    let worksheet = package.xml_part(&resolved.sheet.part)?;
    let root = index_xml(&worksheet)?;
    let existing = direct_filters(&root);
    if existing.len() != 1 {
        return Err(filter_part_error(
            &resolved.sheet.part,
            "does not contain exactly one worksheet AutoFilter",
        ));
    }
    filter_xml::validate_mutable(&worksheet, existing[0])?;
    let edited = apply_patches(
        &worksheet,
        vec![XmlPatch::new(existing[0].full_range.clone(), Vec::new())],
    )?;
    package.set_part(&resolved.sheet.part, edited)
}

fn resolve_sheet(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedSheet> {
    if package.kind() != DocumentKind::Spreadsheet {
        return Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Spreadsheet AutoFilter operations are available only for Spreadsheet documents.",
        ));
    }
    validate_mutation_path(requested)?;
    if requested.trim_start_matches('/').contains('/') {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Adding a Spreadsheet AutoFilter requires a worksheet path such as /Sheet1.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let sheet = snapshot
        .root()
        .children
        .iter()
        .find(|node| {
            node.node_type == OfficeNodeType::Worksheet && node.path.eq_ignore_ascii_case(requested)
        })
        .ok_or_else(|| node_not_found(requested))?;
    let part = sheet.format.get("part").cloned().ok_or_else(|| {
        editor_error(
            "use.office.spreadsheet_sheet_invalid",
            format!("Worksheet '{}' has no source part.", sheet.path),
        )
    })?;
    Ok(ResolvedSheet {
        path: sheet.path.clone(),
        part,
    })
}

fn resolve_filter(package: &NativeOfficePackage, requested: &str) -> UseResult<ResolvedFilter> {
    validate_mutation_path(requested)?;
    if !is_path(requested) {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Worksheet AutoFilter updates require a path such as /Sheet1/autofilter.",
        ));
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let node = snapshot.get(requested, 0)?;
    if node.node_type != OfficeNodeType::AutoFilter {
        return Err(editor_error(
            "use.office.mutation_path_unsupported",
            "Worksheet AutoFilter updates require a path such as /Sheet1/autofilter.",
        ));
    }
    let (sheet_path, _) = node
        .path
        .rsplit_once('/')
        .ok_or_else(|| node_not_found(requested))?;
    Ok(ResolvedFilter {
        sheet: resolve_sheet(package, sheet_path)?,
        path: node.path,
    })
}

fn validate_geometry(
    package: &NativeOfficePackage,
    snapshot: &NativeOfficeDocument,
    sheet: &ResolvedSheet,
    range: CellRange,
) -> UseResult<()> {
    let sheet_node = snapshot.get(&sheet.path, 1)?;
    for table in sheet_node
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Table)
    {
        if let Some(table_range) = table
            .format
            .get("ref")
            .and_then(|value| CellRange::parse(value).ok())
        {
            if range.intersects(table_range) {
                return Err(editor_error(
                    "use.office.spreadsheet_filter_table_overlap",
                    format!(
                        "Worksheet AutoFilter range '{}' overlaps table '{}' at '{}'.",
                        range.a1(),
                        table.path,
                        table_range.a1()
                    ),
                ));
            }
        }
    }
    let worksheet = package.xml_part(&sheet.part)?;
    let root = index_xml(&worksheet)?;
    for collection in root
        .children
        .iter()
        .filter(|child| child.local_name == "mergeCells" && child.namespace == root.namespace)
    {
        for merged in collection
            .children
            .iter()
            .filter(|child| child.local_name == "mergeCell" && child.namespace == root.namespace)
        {
            if let Some(reference) = merged.attributes.get("ref") {
                let merged_range = CellRange::parse(reference)?;
                if range.intersects(merged_range) {
                    return Err(editor_error(
                        "use.office.spreadsheet_filter_merge_overlap",
                        format!(
                            "Worksheet AutoFilter range '{}' overlaps merged range '{}'.",
                            range.a1(),
                            merged_range.a1()
                        ),
                    ));
                }
            }
        }
    }
    Ok(())
}

fn direct_filters(
    root: &crate::xml_edit::IndexedXmlElement,
) -> Vec<&crate::xml_edit::IndexedXmlElement> {
    root.children
        .iter()
        .filter(|child| child.local_name == "autoFilter" && child.namespace == root.namespace)
        .collect()
}

fn require_mutable(node: &crate::DocumentNode) -> UseResult<()> {
    if node.format.get("nativeMutable").map(String::as_str) == Some("true") {
        Ok(())
    } else {
        Err(editor_error(
            "use.office.spreadsheet_filter_unknown_content",
            format!(
                "Spreadsheet AutoFilter '{}' contains unsupported criteria, sort state, extensions, or unknown content.",
                node.path
            ),
        ))
    }
}

fn node_not_found(path: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}

fn filter_part_error(part: &str, reason: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_filter_invalid",
        format!("Spreadsheet worksheet part '{part}' {reason}."),
    )
    .with_detail("part", part)
}
