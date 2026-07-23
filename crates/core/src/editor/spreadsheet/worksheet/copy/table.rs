use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::copy_error;
use super::graph::ClonePlan;
use crate::spreadsheet_formula::rewrite_formula_sheet_name;
use crate::xml_edit::{apply_patches, escape_text, index_xml, IndexedXmlElement, XmlPatch};
use crate::{LosslessXmlPart, NativeOfficePackage};

#[derive(Debug)]
pub(super) struct TableIdentityPlan {
    pub(super) by_part: BTreeMap<String, TableIdentity>,
    pub(super) names: Vec<(String, String)>,
}

#[derive(Debug)]
pub(super) struct TableIdentity {
    id: u32,
    name: String,
}

impl TableIdentityPlan {
    pub(super) fn build(package: &NativeOfficePackage, plan: &ClonePlan) -> UseResult<Self> {
        let mut maximum_id = 0_u32;
        let mut names = BTreeSet::new();
        for part_name in package
            .part_names()
            .filter(|part| part.starts_with("xl/tables/") && part.ends_with(".xml"))
        {
            let part = package.xml_part(part_name)?;
            let root = index_xml(&part)?;
            if root.local_name != "table" {
                continue;
            }
            if let Some(id) = root
                .attributes
                .get("id")
                .and_then(|value| value.parse::<u32>().ok())
            {
                maximum_id = maximum_id.max(id);
            }
            for attribute in ["name", "displayName"] {
                if let Some(name) = root.attributes.get(attribute) {
                    names.insert(name.to_ascii_lowercase());
                }
            }
        }

        let mut by_part = BTreeMap::new();
        let mut replacements = Vec::new();
        for source in plan
            .parts
            .keys()
            .filter(|part| part.starts_with("xl/tables/") && part.ends_with(".xml"))
        {
            let part = package.xml_part(source)?;
            let root = index_xml(&part)?;
            let old_name = root
                .attributes
                .get("displayName")
                .or_else(|| root.attributes.get("name"))
                .cloned()
                .ok_or_else(|| copy_error(format!("Spreadsheet table '{source}' has no name.")))?;
            maximum_id = maximum_id
                .checked_add(1)
                .ok_or_else(|| copy_error("Spreadsheet table IDs are exhausted."))?;
            let name = unique_table_name(&old_name, maximum_id, &mut names)?;
            replacements.push((old_name, name.clone()));
            by_part.insert(
                source.clone(),
                TableIdentity {
                    id: maximum_id,
                    name,
                },
            );
        }
        Ok(Self {
            by_part,
            names: replacements,
        })
    }
}

fn unique_table_name(
    source: &str,
    suggested_number: u32,
    names: &mut BTreeSet<String>,
) -> UseResult<String> {
    let table_number = source
        .strip_prefix("Table")
        .filter(|suffix| !suffix.is_empty() && suffix.chars().all(|value| value.is_ascii_digit()));
    let base = if table_number.is_some() {
        "Table".to_string()
    } else {
        format!("{source}_Copy")
    };
    for number in suggested_number..=u32::MAX {
        let candidate = if table_number.is_some() {
            format!("{base}{number}")
        } else if number == suggested_number {
            base.clone()
        } else {
            format!("{base}{number}")
        };
        if candidate.chars().count() <= 255 && names.insert(candidate.to_ascii_lowercase()) {
            return Ok(candidate);
        }
    }
    Err(copy_error("Spreadsheet table names are exhausted."))
}

pub(super) fn rewrite_table_identity(
    target: &str,
    bytes: Vec<u8>,
    identity: &TableIdentity,
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(target.to_string(), bytes)?;
    let root = index_xml(&part)?;
    if root.local_name != "table" {
        return Err(copy_error(format!(
            "Spreadsheet table part '{target}' has an invalid root."
        )));
    }
    let updates = BTreeMap::from([
        ("id".to_string(), identity.id.to_string()),
        ("name".to_string(), identity.name.clone()),
        ("displayName".to_string(), identity.name.clone()),
    ]);
    apply_patches(
        &part,
        vec![XmlPatch::new(
            root.start_tag_range.clone(),
            super::super::updated_start_tag(&root, &updates),
        )],
    )
}

pub(super) fn should_rewrite_formulas(part_name: &str) -> bool {
    part_name.starts_with("xl/worksheets/")
        || part_name.starts_with("xl/charts/")
        || part_name.starts_with("xl/tables/")
}

pub(super) fn rewrite_cloned_formulas(
    target: &str,
    bytes: Vec<u8>,
    source_sheet: &str,
    target_sheet: &str,
    table_names: &[(String, String)],
) -> UseResult<Vec<u8>> {
    let part = LosslessXmlPart::parse(target.to_string(), bytes)?;
    let index = index_xml(&part)?;
    let mut formulas = Vec::new();
    collect_formula_elements(&index, &mut formulas);
    let mut patches = Vec::new();
    for formula in formulas {
        let text = super::super::decoded_text(&part, formula)?;
        let rewritten = rewrite_formula_sheet_name(&text, source_sheet, target_sheet)?;
        let rewritten = rewrite_table_names(&rewritten, table_names);
        if rewritten != text {
            patches.push(XmlPatch::new(
                formula.content_range.clone(),
                escape_text(&rewritten),
            ));
        }
    }
    if patches.is_empty() {
        Ok(part.raw().to_vec())
    } else {
        apply_patches(&part, patches)
    }
}

fn collect_formula_elements<'a>(
    element: &'a IndexedXmlElement,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        if matches!(
            child.local_name.as_str(),
            "f" | "formula" | "calculatedColumnFormula" | "totalsRowFormula"
        ) {
            output.push(child);
        }
        collect_formula_elements(child, output);
    }
}

fn rewrite_table_names(formula: &str, replacements: &[(String, String)]) -> String {
    let mut output = String::with_capacity(formula.len());
    let mut cursor = 0_usize;
    while cursor < formula.len() {
        if formula.as_bytes()[cursor] == b'"' {
            let end = quoted_string_end(formula, cursor);
            output.push_str(&formula[cursor..end]);
            cursor = end;
            continue;
        }
        let character = formula[cursor..].chars().next().unwrap_or_default();
        if character.is_alphabetic() || character == '_' {
            let start = cursor;
            cursor += character.len_utf8();
            while cursor < formula.len() {
                let candidate = formula[cursor..].chars().next().unwrap_or_default();
                if !(candidate.is_alphanumeric() || matches!(candidate, '_' | '.')) {
                    break;
                }
                cursor += candidate.len_utf8();
            }
            let token = &formula[start..cursor];
            if formula[cursor..].starts_with('[') {
                if let Some((_, replacement)) = replacements
                    .iter()
                    .find(|(source, _)| source.eq_ignore_ascii_case(token))
                {
                    output.push_str(replacement);
                    continue;
                }
            }
            output.push_str(token);
            continue;
        }
        output.push(character);
        cursor += character.len_utf8();
    }
    output
}

fn quoted_string_end(value: &str, start: usize) -> usize {
    let bytes = value.as_bytes();
    let mut cursor = start + 1;
    while cursor < bytes.len() {
        if bytes[cursor] == b'"' {
            if bytes.get(cursor + 1) == Some(&b'"') {
                cursor += 2;
                continue;
            }
            return cursor + 1;
        }
        let character = value[cursor..].chars().next().unwrap_or_default();
        cursor += character.len_utf8();
    }
    value.len()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn structured_table_names_are_rewritten_outside_strings() {
        assert_eq!(
            rewrite_table_names(
                r#"SUM(Table1[Value])+\"Table1[Value]\"+Table10[Value]"#,
                &[("Table1".into(), "Table2".into())]
            ),
            r#"SUM(Table2[Value])+\"Table1[Value]\"+Table10[Value]"#
        );
    }
}
