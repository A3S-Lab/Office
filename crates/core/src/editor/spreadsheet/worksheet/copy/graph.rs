use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::UseResult;

use super::copy_error;
use super::table::{
    rewrite_cloned_formulas, rewrite_table_identity, should_rewrite_formulas, TableIdentityPlan,
};
use crate::xml_edit::{apply_patches, index_xml, IndexedXmlElement, XmlPatch};
use crate::{NativeOfficePackage, OpcPackageModel, RelationshipSource, RelationshipTarget};

#[derive(Debug)]
pub(super) struct ClonePlan {
    pub(super) parts: BTreeMap<String, String>,
}

impl ClonePlan {
    pub(super) fn build(
        package: &NativeOfficePackage,
        model: &OpcPackageModel,
        source_part: &str,
    ) -> UseResult<Self> {
        let mut allocator = PartNameAllocator::new(package);
        let mut parts = BTreeMap::new();
        clone_part_graph(
            package,
            model,
            source_part,
            true,
            &mut allocator,
            &mut parts,
        )?;
        let relationship_parts = parts
            .keys()
            .filter(|source| package.contains_part(&relationship_part(source)))
            .count();
        let added = parts
            .len()
            .checked_add(relationship_parts)
            .ok_or_else(|| copy_error("Worksheet copy part count overflowed."))?;
        let resulting_entries = package
            .part_names()
            .count()
            .checked_add(added)
            .ok_or_else(|| copy_error("Worksheet copy package entry count overflowed."))?;
        if resulting_entries > package.limits().max_entries {
            return Err(super::super::super::editor_error(
                "use.office.package_too_many_parts",
                format!(
                    "Worksheet copy would create {resulting_entries} package entries; the limit is {}.",
                    package.limits().max_entries
                ),
            ));
        }
        Ok(Self { parts })
    }

    pub(super) fn target(&self, source: &str) -> Option<&str> {
        self.parts.get(source).map(String::as_str)
    }

    pub(super) fn apply(
        &self,
        package: &mut NativeOfficePackage,
        model: &OpcPackageModel,
        source_sheet: &str,
        target_sheet: &str,
        table_identities: &TableIdentityPlan,
    ) -> UseResult<()> {
        for (source, target) in &self.parts {
            let mut bytes = package.part(source)?.to_vec();
            if should_rewrite_formulas(source) {
                bytes = rewrite_cloned_formulas(
                    target,
                    bytes,
                    source_sheet,
                    target_sheet,
                    &table_identities.names,
                )?;
            }
            if let Some(identity) = table_identities.by_part.get(source) {
                bytes = rewrite_table_identity(target, bytes, identity)?;
            }
            package.set_part(target, bytes)?;
            if let Some(content_type) = model.content_types().override_for_part(source) {
                crate::opc_edit::add_content_type_override(package, target, content_type)?;
            }
        }

        for (source, target) in &self.parts {
            let source_relationships = relationship_part(source);
            if !package.contains_part(&source_relationships) {
                continue;
            }
            let target_relationships = relationship_part(target);
            let bytes = rewrite_relationship_targets(
                package,
                model,
                source,
                target,
                &source_relationships,
                &self.parts,
            )?;
            package.set_part(&target_relationships, bytes)?;
            if let Some(content_type) = model
                .content_types()
                .override_for_part(&source_relationships)
            {
                crate::opc_edit::add_content_type_override(
                    package,
                    &target_relationships,
                    content_type,
                )?;
            }
        }
        Ok(())
    }
}

fn clone_part_graph(
    package: &NativeOfficePackage,
    model: &OpcPackageModel,
    source_part: &str,
    force_clone: bool,
    allocator: &mut PartNameAllocator,
    parts: &mut BTreeMap<String, String>,
) -> UseResult<String> {
    if let Some(target) = parts.get(source_part) {
        return Ok(target.clone());
    }
    if !force_clone && is_shared_workbook_part(source_part) {
        return Ok(source_part.to_string());
    }
    if !package.contains_part(source_part) {
        return Err(copy_error(format!(
            "Worksheet relationship target '{source_part}' does not exist."
        )));
    }
    let target = allocator.allocate(source_part)?;
    parts.insert(source_part.to_string(), target.clone());

    let source = RelationshipSource::Part {
        part_name: source_part.to_string(),
    };
    for relationship in model.relationships().relationships_from(&source) {
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            continue;
        };
        if relationship.relationship_type.ends_with("/pivotTable") {
            return Err(super::super::super::editor_error(
                "use.office.spreadsheet_copy_pivot_unsupported",
                "Worksheet copy cannot yet clone pivot tables safely; remove the pivot table or keep using the compatibility provider.",
            ));
        }
        if is_shared_relationship(&relationship.relationship_type, part_name) {
            continue;
        }
        clone_part_graph(package, model, part_name, false, allocator, parts)?;
    }
    Ok(target)
}

pub(super) fn is_shared_relationship(relationship_type: &str, part_name: &str) -> bool {
    is_shared_workbook_part(part_name)
        || [
            "/styles",
            "/theme",
            "/sharedStrings",
            "/calcChain",
            "/connections",
            "/externalLink",
            "/pivotCacheDefinition",
            "/person",
        ]
        .iter()
        .any(|suffix| relationship_type.ends_with(suffix))
}

fn is_shared_workbook_part(part_name: &str) -> bool {
    part_name == "xl/workbook.xml"
        || part_name == "xl/styles.xml"
        || part_name == "xl/sharedStrings.xml"
        || part_name == "xl/calcChain.xml"
        || part_name == "xl/connections.xml"
        || part_name.starts_with("xl/theme/")
        || part_name.starts_with("xl/pivotCache/")
        || part_name.starts_with("xl/externalLinks/")
        || part_name.starts_with("xl/persons/")
        || part_name.starts_with("docProps/")
}

#[derive(Debug)]
struct PartNameAllocator {
    reserved: BTreeSet<String>,
    limit: usize,
}

impl PartNameAllocator {
    fn new(package: &NativeOfficePackage) -> Self {
        Self {
            reserved: package.part_names().map(str::to_string).collect(),
            limit: package.limits().max_entries.saturating_add(1),
        }
    }

    fn allocate(&mut self, source: &str) -> UseResult<String> {
        let (directory, file_name) = source
            .rsplit_once('/')
            .map_or(("", source), |(directory, file_name)| {
                (directory, file_name)
            });
        let (stem, extension) = file_name
            .rsplit_once('.')
            .map_or((file_name, ""), |(stem, extension)| (stem, extension));
        let prefix = stem.trim_end_matches(|character: char| character.is_ascii_digit());
        let prefix = if prefix.is_empty() { stem } else { prefix };
        for number in 1..=self.limit {
            let file_name = if extension.is_empty() {
                format!("{prefix}{number}")
            } else {
                format!("{prefix}{number}.{extension}")
            };
            let candidate = if directory.is_empty() {
                file_name
            } else {
                format!("{directory}/{file_name}")
            };
            if self.reserved.insert(candidate.clone()) {
                return Ok(candidate);
            }
        }
        Err(copy_error(format!(
            "Worksheet copy cannot allocate a package part derived from '{source}'."
        )))
    }
}

fn rewrite_relationship_targets(
    package: &NativeOfficePackage,
    model: &OpcPackageModel,
    source_part: &str,
    target_part: &str,
    relationship_part_name: &str,
    mappings: &BTreeMap<String, String>,
) -> UseResult<Vec<u8>> {
    let part = package.xml_part(relationship_part_name)?;
    let index = index_xml(&part)?;
    let source = RelationshipSource::Part {
        part_name: source_part.to_string(),
    };
    let mut patches = Vec::new();
    for element in index
        .children
        .iter()
        .filter(|element| element.local_name == "Relationship")
    {
        let Some(id) = element.attributes.get("Id") else {
            continue;
        };
        let Some(relationship) = model.relationships().relationship(&source, id) else {
            return Err(copy_error(format!(
                "Relationship '{id}' from '{source_part}' is missing from the OPC graph."
            )));
        };
        let RelationshipTarget::Internal {
            part_name,
            fragment,
        } = &relationship.target
        else {
            continue;
        };
        let Some(target) = mappings.get(part_name) else {
            continue;
        };
        let updates = BTreeMap::from([(
            qualified_attribute_name(element, "Target"),
            relative_target(target_part, target, fragment.as_deref()),
        )]);
        patches.push(XmlPatch::new(
            element.start_tag_range.clone(),
            super::super::updated_start_tag(element, &updates),
        ));
    }
    if patches.is_empty() {
        Ok(part.raw().to_vec())
    } else {
        apply_patches(&part, patches)
    }
}

pub(super) fn relationship_part(part_name: &str) -> String {
    part_name.rsplit_once('/').map_or_else(
        || format!("_rels/{part_name}.rels"),
        |(directory, file_name)| format!("{directory}/_rels/{file_name}.rels"),
    )
}

pub(super) fn relative_target(source: &str, target: &str, fragment: Option<&str>) -> String {
    let source_directory = source
        .rsplit_once('/')
        .map_or_else(Vec::new, |(directory, _)| directory.split('/').collect());
    let target_segments = target.split('/').collect::<Vec<_>>();
    let common = source_directory
        .iter()
        .zip(&target_segments)
        .take_while(|(source, target)| source == target)
        .count();
    let mut segments = Vec::new();
    segments.extend(std::iter::repeat_n("..", source_directory.len() - common));
    segments.extend(target_segments[common..].iter().copied());
    let mut value = segments.join("/");
    if let Some(fragment) = fragment {
        value.push('#');
        value.push_str(fragment);
    }
    value
}

fn qualified_attribute_name(element: &IndexedXmlElement, local_name: &str) -> String {
    element
        .qualified_attributes
        .keys()
        .find(|name| {
            name.rsplit_once(':')
                .map_or(name.as_str(), |(_, local)| local)
                == local_name
        })
        .cloned()
        .unwrap_or_else(|| local_name.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relative_targets_follow_cloned_part_locations() {
        assert_eq!(
            relative_target("xl/worksheets/sheet2.xml", "xl/drawings/drawing2.xml", None),
            "../drawings/drawing2.xml"
        );
        assert_eq!(
            relative_target(
                "xl/drawings/drawing2.xml",
                "xl/media/image2.png",
                Some("preview")
            ),
            "../media/image2.png#preview"
        );
    }
}
