use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::{UseError, UseResult};

use super::{
    merge_error, merge_text_elements, placeholders, scan_text_elements, MergeAccumulator, MergeData,
};
use crate::xml_edit::{
    apply_patches, decoded_element_text, index_xml, IndexedXmlElement, XmlPatch,
};
use crate::{LosslessXmlPart, NativeOfficePackage};

const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";

pub(super) fn merge(
    package: &mut NativeOfficePackage,
    data: &MergeData,
    accumulator: &mut MergeAccumulator,
) -> UseResult<()> {
    let worksheet_parts = worksheet_parts(package);
    let shared_references = shared_string_references(package, &worksheet_parts)?;
    merge_shared_strings(package, &shared_references, data, accumulator)?;

    for part_name in worksheet_parts {
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        let cells = worksheet_cells(&root)?;
        let mut patches = Vec::new();
        for cell in cells {
            merge_cell(&part, cell, data, accumulator, &mut patches)?;
        }
        if !patches.is_empty() {
            package.set_part(&part_name, apply_patches(&part, patches)?)?;
            accumulator.changed(&part_name);
        }
    }
    Ok(())
}

pub(super) fn scan(
    package: &NativeOfficePackage,
    unresolved: &mut BTreeSet<String>,
) -> UseResult<()> {
    let worksheet_parts = worksheet_parts(package);
    let shared_references = shared_string_references(package, &worksheet_parts)?;
    scan_shared_strings(package, &shared_references, unresolved)?;

    for part_name in worksheet_parts {
        let part = package.xml_part(&part_name)?;
        let root = index_xml(&part)?;
        let cells = worksheet_cells(&root)?;
        for cell in cells {
            scan_cell(&part, cell, unresolved)?;
        }
    }
    Ok(())
}

fn merge_shared_strings(
    package: &mut NativeOfficePackage,
    references: &BTreeMap<usize, usize>,
    data: &MergeData,
    accumulator: &mut MergeAccumulator,
) -> UseResult<()> {
    if references.is_empty() {
        return Ok(());
    }
    if !package.contains_part(SHARED_STRINGS_PART) {
        return Err(spreadsheet_error(
            "Spreadsheet cells reference shared strings, but xl/sharedStrings.xml is missing.",
        ));
    }

    let part = package.xml_part(SHARED_STRINGS_PART)?;
    let root = index_xml(&part)?;
    let items = root
        .children
        .iter()
        .filter(|element| element.local_name == "si")
        .collect::<Vec<_>>();
    let mut patches = Vec::new();
    for (index, multiplier) in references {
        let item = items.get(*index).ok_or_else(|| {
            spreadsheet_error(format!(
                "Spreadsheet cell references missing shared-string index {index}."
            ))
            .with_detail("sharedStringIndex", *index)
        })?;
        let mut text_elements = Vec::new();
        collect_spreadsheet_text(item, &mut text_elements);
        merge_text_elements(
            &part,
            &text_elements,
            data,
            *multiplier,
            accumulator,
            &mut patches,
        )?;
    }
    if !patches.is_empty() {
        package.set_part(SHARED_STRINGS_PART, apply_patches(&part, patches)?)?;
        accumulator.changed(SHARED_STRINGS_PART);
    }
    Ok(())
}

fn scan_shared_strings(
    package: &NativeOfficePackage,
    references: &BTreeMap<usize, usize>,
    unresolved: &mut BTreeSet<String>,
) -> UseResult<()> {
    if references.is_empty() {
        return Ok(());
    }
    if !package.contains_part(SHARED_STRINGS_PART) {
        return Err(spreadsheet_error(
            "Spreadsheet cells reference shared strings, but xl/sharedStrings.xml is missing.",
        ));
    }

    let part = package.xml_part(SHARED_STRINGS_PART)?;
    let root = index_xml(&part)?;
    let items = root
        .children
        .iter()
        .filter(|element| element.local_name == "si")
        .collect::<Vec<_>>();
    for index in references.keys() {
        let item = items.get(*index).ok_or_else(|| {
            spreadsheet_error(format!(
                "Spreadsheet cell references missing shared-string index {index}."
            ))
            .with_detail("sharedStringIndex", *index)
        })?;
        let mut text_elements = Vec::new();
        collect_spreadsheet_text(item, &mut text_elements);
        scan_text_elements(&part, &text_elements, unresolved)?;
    }
    Ok(())
}

fn merge_cell(
    part: &LosslessXmlPart,
    cell: &IndexedXmlElement,
    data: &MergeData,
    accumulator: &mut MergeAccumulator,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let cell_type = cell.attributes.get("t").map_or("n", String::as_str);
    match cell_type {
        "s" => Ok(()),
        "inlineStr" => {
            let inline = single_direct_child(cell, "is")?;
            if let Some(inline) = inline {
                let mut text_elements = Vec::new();
                collect_spreadsheet_text(inline, &mut text_elements);
                merge_text_elements(part, &text_elements, data, 1, accumulator, patches)?;
            }
            Ok(())
        }
        "str" => {
            let values = direct_children(cell, "v");
            ensure_at_most_one(cell, "v", values.len())?;
            merge_text_elements(part, &values, data, 1, accumulator, patches)
        }
        _ => reject_resolved_unsupported_cell(part, cell, cell_type, data),
    }
}

fn scan_cell(
    part: &LosslessXmlPart,
    cell: &IndexedXmlElement,
    unresolved: &mut BTreeSet<String>,
) -> UseResult<()> {
    let cell_type = cell.attributes.get("t").map_or("n", String::as_str);
    match cell_type {
        "s" => Ok(()),
        "inlineStr" => {
            let inline = single_direct_child(cell, "is")?;
            if let Some(inline) = inline {
                let mut text_elements = Vec::new();
                collect_spreadsheet_text(inline, &mut text_elements);
                scan_text_elements(part, &text_elements, unresolved)?;
            }
            Ok(())
        }
        _ => {
            let values = direct_children(cell, "v");
            ensure_at_most_one(cell, "v", values.len())?;
            scan_text_elements(part, &values, unresolved)
        }
    }
}

fn reject_resolved_unsupported_cell(
    part: &LosslessXmlPart,
    cell: &IndexedXmlElement,
    cell_type: &str,
    data: &MergeData,
) -> UseResult<()> {
    let values = direct_children(cell, "v");
    ensure_at_most_one(cell, "v", values.len())?;
    if values.is_empty() {
        return Ok(());
    }
    let text = values
        .iter()
        .map(|element| decoded_element_text(part, element))
        .collect::<UseResult<Vec<_>>>()?
        .concat();
    if let Some(placeholder) = placeholders(&text)
        .into_iter()
        .find(|placeholder| data.get(&placeholder.key).is_some())
    {
        return Err(merge_error(
            "use.office.template_cell_type_unsupported",
            format!(
                "Native Office template replacement in Spreadsheet cell '{}' would mutate unsupported cell type '{cell_type}'.",
                cell_reference(cell)
            ),
        )
        .with_suggestion(
            "Store template placeholders as inline strings, shared strings, or t=\"str\" cell values.",
        )
        .with_detail("part", part.name())
        .with_detail("cell", cell_reference(cell))
        .with_detail("cellType", cell_type)
        .with_detail("placeholder", placeholder.key));
    }
    Ok(())
}

fn shared_string_references(
    package: &NativeOfficePackage,
    worksheet_parts: &[String],
) -> UseResult<BTreeMap<usize, usize>> {
    let mut references = BTreeMap::<usize, usize>::new();
    for part_name in worksheet_parts {
        let part = package.xml_part(part_name)?;
        let root = index_xml(&part)?;
        let cells = worksheet_cells(&root)?;
        for cell in cells {
            if cell.attributes.get("t").map(String::as_str) != Some("s") {
                continue;
            }
            let values = direct_children(cell, "v");
            if values.len() != 1 {
                return Err(spreadsheet_error(format!(
                    "Shared-string cell '{}' must contain exactly one direct value element.",
                    cell_reference(cell)
                ))
                .with_detail("part", part_name.as_str())
                .with_detail("cell", cell_reference(cell)));
            }
            let raw = decoded_element_text(&part, values[0])?;
            let index = raw.trim().parse::<usize>().map_err(|error| {
                spreadsheet_error(format!(
                    "Shared-string cell '{}' contains invalid index '{raw}': {error}",
                    cell_reference(cell)
                ))
                .with_detail("part", part_name.as_str())
                .with_detail("cell", cell_reference(cell))
            })?;
            let count = references.entry(index).or_default();
            *count = count.checked_add(1).ok_or_else(|| {
                spreadsheet_error("Spreadsheet shared-string reference count overflowed.")
            })?;
        }
    }
    Ok(references)
}

fn worksheet_parts(package: &NativeOfficePackage) -> Vec<String> {
    package
        .part_names()
        .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect()
}

fn worksheet_cells(root: &IndexedXmlElement) -> UseResult<Vec<&IndexedXmlElement>> {
    let sheet_data = direct_children(root, "sheetData");
    if sheet_data.len() > 1 {
        return Err(spreadsheet_error(
            "Spreadsheet worksheet contains more than one sheetData element.",
        ));
    }
    let Some(sheet_data) = sheet_data.into_iter().next() else {
        return Ok(Vec::new());
    };
    let mut cells = Vec::new();
    for row in sheet_data
        .children
        .iter()
        .filter(|child| child.local_name == "row")
    {
        cells.extend(direct_children(row, "c"));
    }
    Ok(cells)
}

fn collect_spreadsheet_text<'a>(
    element: &'a IndexedXmlElement,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        if child.local_name == "rPh" {
            continue;
        }
        if child.local_name == "t" {
            output.push(child);
        } else {
            collect_spreadsheet_text(child, output);
        }
    }
}

fn direct_children<'a>(
    element: &'a IndexedXmlElement,
    local_name: &str,
) -> Vec<&'a IndexedXmlElement> {
    element
        .children
        .iter()
        .filter(|child| child.local_name == local_name)
        .collect()
}

fn single_direct_child<'a>(
    element: &'a IndexedXmlElement,
    local_name: &str,
) -> UseResult<Option<&'a IndexedXmlElement>> {
    let children = direct_children(element, local_name);
    ensure_at_most_one(element, local_name, children.len())?;
    Ok(children.into_iter().next())
}

fn ensure_at_most_one(cell: &IndexedXmlElement, local_name: &str, count: usize) -> UseResult<()> {
    if count <= 1 {
        return Ok(());
    }
    Err(spreadsheet_error(format!(
        "Spreadsheet cell '{}' contains more than one direct '{local_name}' element.",
        cell_reference(cell)
    ))
    .with_detail("cell", cell_reference(cell)))
}

fn cell_reference(cell: &IndexedXmlElement) -> String {
    cell.attributes
        .get("r")
        .cloned()
        .unwrap_or_else(|| "<unknown>".to_string())
}

fn spreadsheet_error(message: impl Into<String>) -> UseError {
    merge_error("use.office.template_spreadsheet_invalid", message)
}
