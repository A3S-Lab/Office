use a3s_use_core::UseResult;

use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_formula::{
    LocalStructuredReferenceContext, StructuredReferenceRewritePlan,
    StructuredReferenceRewriteResult,
};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{apply_patches, escape_text, index_xml, IndexedXmlElement, XmlPatch};
use crate::{LosslessXmlPart, NativeOfficePackage};

#[derive(Debug, Clone, Copy)]
enum FormulaPartKind {
    Workbook {
        target_sheet_index: usize,
    },
    Worksheet {
        target: bool,
        table_range: CellRange,
    },
    Chart,
    Table {
        target: bool,
    },
}

#[derive(Debug, Clone, Copy)]
struct FormulaCarrier<'a> {
    element: &'a IndexedXmlElement,
    context: LocalStructuredReferenceContext,
}

pub(super) fn rewrite_table_references(
    package: &mut NativeOfficePackage,
    target_sheet_part: &str,
    target_table_part: &str,
    table_range: CellRange,
    plan: &StructuredReferenceRewritePlan,
) -> UseResult<()> {
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let worksheet_parts = snapshot
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|node| {
            node.format.get("part").cloned().ok_or_else(|| {
                formula_error(
                    "use.office.spreadsheet_sheet_invalid",
                    format!("Worksheet '{}' has no source part.", node.path),
                )
            })
        })
        .collect::<UseResult<Vec<_>>>()?;
    let target_sheet_index = worksheet_parts
        .iter()
        .position(|part| part == target_sheet_part)
        .ok_or_else(|| {
            formula_error(
                "use.office.spreadsheet_sheet_invalid",
                format!("Worksheet part '{target_sheet_part}' has no workbook sheet entry."),
            )
        })?;
    let chart_parts = package
        .part_names()
        .filter(|part| part.starts_with("xl/charts/") && part.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    let table_parts = package
        .part_names()
        .filter(|part| part.starts_with("xl/tables/") && part.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let mut matched = rewrite_part(
        package,
        "xl/workbook.xml",
        FormulaPartKind::Workbook { target_sheet_index },
        plan,
    )?;
    for part_name in &worksheet_parts {
        matched |= rewrite_part(
            package,
            part_name,
            FormulaPartKind::Worksheet {
                target: part_name == target_sheet_part,
                table_range,
            },
            plan,
        )?;
    }
    for part_name in &chart_parts {
        matched |= rewrite_part(package, part_name, FormulaPartKind::Chart, plan)?;
    }
    for part_name in &table_parts {
        matched |= rewrite_part(
            package,
            part_name,
            FormulaPartKind::Table {
                target: part_name == target_table_part,
            },
            plan,
        )?;
    }

    if matched && plan.geometry_changed() {
        clear_formula_caches(package, &worksheet_parts)?;
        clear_chart_caches(package, &chart_parts)?;
    }
    Ok(())
}

fn rewrite_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    kind: FormulaPartKind,
    plan: &StructuredReferenceRewritePlan,
) -> UseResult<bool> {
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let mut carriers = Vec::new();
    collect_formula_carriers(&root, None, kind, &mut carriers)?;
    let mut patches = Vec::new();
    let mut matched = false;
    for carrier in carriers {
        let formula = decoded_text(&part, carrier.element)?;
        let rewritten: StructuredReferenceRewriteResult =
            plan.rewrite(&formula, carrier.context)?;
        matched |= rewritten.matched;
        if rewritten.formula != formula {
            patches.push(XmlPatch::new(
                carrier.element.content_range.clone(),
                escape_text(&rewritten.formula),
            ));
        }
    }
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
    }
    Ok(matched)
}

fn collect_formula_carriers<'a>(
    element: &'a IndexedXmlElement,
    parent: Option<&'a IndexedXmlElement>,
    kind: FormulaPartKind,
    output: &mut Vec<FormulaCarrier<'a>>,
) -> UseResult<()> {
    if is_formula_element(element, kind) {
        output.push(FormulaCarrier {
            element,
            context: local_context(element, parent, kind)?,
        });
    }
    for child in &element.children {
        collect_formula_carriers(child, Some(element), kind, output)?;
    }
    Ok(())
}

fn is_formula_element(element: &IndexedXmlElement, kind: FormulaPartKind) -> bool {
    if matches!(kind, FormulaPartKind::Workbook { .. }) {
        return element.local_name == "definedName";
    }
    matches!(
        element.local_name.as_str(),
        "f" | "formula" | "formula1" | "formula2" | "calculatedColumnFormula" | "totalsRowFormula"
    )
}

fn local_context(
    element: &IndexedXmlElement,
    parent: Option<&IndexedXmlElement>,
    kind: FormulaPartKind,
) -> UseResult<LocalStructuredReferenceContext> {
    match kind {
        FormulaPartKind::Workbook { target_sheet_index } => {
            let Some(local_sheet_id) = element.attributes.get("localSheetId") else {
                return Ok(LocalStructuredReferenceContext::Unknown);
            };
            Ok(local_sheet_id
                .parse::<usize>()
                .ok()
                .filter(|index| *index == target_sheet_index)
                .map_or(LocalStructuredReferenceContext::DoesNotApply, |_| {
                    LocalStructuredReferenceContext::Unknown
                }))
        }
        FormulaPartKind::Worksheet {
            target,
            table_range,
        } => {
            if !target {
                return Ok(LocalStructuredReferenceContext::DoesNotApply);
            }
            if element.local_name != "f"
                || parent.map(|value| value.local_name.as_str()) != Some("c")
            {
                return Ok(LocalStructuredReferenceContext::Unknown);
            }
            let reference = parent
                .and_then(|cell| cell.attributes.get("r"))
                .ok_or_else(|| {
                    formula_error(
                        "use.office.spreadsheet_formula_invalid",
                        "Spreadsheet formula cell has no cell reference.",
                    )
                })
                .and_then(|reference| CellReference::parse(reference))?;
            Ok(if table_range.contains(reference) {
                LocalStructuredReferenceContext::Applies
            } else {
                LocalStructuredReferenceContext::DoesNotApply
            })
        }
        FormulaPartKind::Chart => Ok(LocalStructuredReferenceContext::Unknown),
        FormulaPartKind::Table { target } => Ok(if target {
            LocalStructuredReferenceContext::Applies
        } else {
            LocalStructuredReferenceContext::DoesNotApply
        }),
    }
}

fn clear_formula_caches(
    package: &mut NativeOfficePackage,
    worksheet_parts: &[String],
) -> UseResult<()> {
    for part_name in worksheet_parts {
        let part = package.xml_part(part_name)?;
        let root = index_xml(&part)?;
        let mut cells = Vec::new();
        root.descendants_named("c", &mut cells);
        let patches = cells
            .into_iter()
            .filter(|cell| cell.children.iter().any(|child| child.local_name == "f"))
            .filter_map(|cell| cell.children.iter().find(|child| child.local_name == "v"))
            .map(|value| XmlPatch::new(value.full_range.clone(), Vec::new()))
            .collect::<Vec<_>>();
        if !patches.is_empty() {
            package.set_part(part_name, apply_patches(&part, patches)?)?;
        }
    }
    Ok(())
}

fn clear_chart_caches(package: &mut NativeOfficePackage, chart_parts: &[String]) -> UseResult<()> {
    for part_name in chart_parts {
        let part = package.xml_part(part_name)?;
        let root = index_xml(&part)?;
        let mut patches = Vec::new();
        for cache_name in ["numCache", "strCache", "multiLvlStrCache"] {
            let mut caches = Vec::new();
            root.descendants_named(cache_name, &mut caches);
            patches.extend(
                caches
                    .into_iter()
                    .map(|cache| XmlPatch::new(cache.full_range.clone(), Vec::new())),
            );
        }
        if !patches.is_empty() {
            package.set_part(part_name, apply_patches(&part, patches)?)?;
        }
    }
    Ok(())
}

fn decoded_text(part: &LosslessXmlPart, element: &IndexedXmlElement) -> UseResult<String> {
    let bytes = part
        .parse_bytes()
        .get(element.content_range.clone())
        .ok_or_else(|| {
            formula_error(
                "use.office.spreadsheet_formula_invalid",
                format!("Formula range in '{}' is invalid.", part.name()),
            )
        })?;
    let text = std::str::from_utf8(bytes).map_err(|error| {
        formula_error(
            "use.office.spreadsheet_formula_invalid",
            format!("Formula in '{}' is not UTF-8: {error}", part.name()),
        )
    })?;
    quick_xml::escape::unescape(text)
        .map(|value| value.into_owned())
        .map_err(|error| {
            formula_error(
                "use.office.spreadsheet_formula_invalid",
                format!(
                    "Formula in '{}' contains invalid XML escapes: {error}",
                    part.name()
                ),
            )
        })
}

fn formula_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error(code, message)
}
