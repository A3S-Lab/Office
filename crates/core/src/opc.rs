use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::{UseError, UseResult};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use serde::{Deserialize, Serialize};

use crate::discovery::office_error;
use crate::xml::{attribute, decode_attributes, LosslessXmlPart, XmlLimits};
use crate::{DocumentKind, NativeOfficePackage};

const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
const CONTENT_TYPES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const RELATIONSHIPS_CONTENT_TYPE: &str = "application/vnd.openxmlformats-package.relationships+xml";
const OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STRICT_OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

/// Content type declarations from `[Content_Types].xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContentTypes {
    defaults: BTreeMap<String, String>,
    overrides: BTreeMap<String, String>,
}

impl ContentTypes {
    pub fn default_for_extension(&self, extension: &str) -> Option<&str> {
        self.defaults
            .get(&extension.trim_start_matches('.').to_ascii_lowercase())
            .map(String::as_str)
    }

    pub fn override_for_part(&self, part_name: &str) -> Option<&str> {
        self.overrides
            .get(part_name.trim_start_matches('/'))
            .map(String::as_str)
    }

    pub fn content_type(&self, part_name: &str) -> Option<&str> {
        let part_name = part_name.trim_start_matches('/');
        self.override_for_part(part_name).or_else(|| {
            part_name
                .rsplit_once('.')
                .and_then(|(_, extension)| self.default_for_extension(extension))
        })
    }

    pub fn defaults(&self) -> impl Iterator<Item = (&str, &str)> {
        self.defaults
            .iter()
            .map(|(extension, content_type)| (extension.as_str(), content_type.as_str()))
    }

    pub fn overrides(&self) -> impl Iterator<Item = (&str, &str)> {
        self.overrides
            .iter()
            .map(|(part, content_type)| (part.as_str(), content_type.as_str()))
    }

    fn read(package: &NativeOfficePackage, limits: XmlLimits) -> UseResult<Self> {
        let part = package.xml_part_with_limits(CONTENT_TYPES_PART, limits)?;
        require_root(
            &part,
            "Types",
            CONTENT_TYPES_NAMESPACE,
            "use.office.content_types_invalid",
        )?;
        let mut defaults = BTreeMap::new();
        let mut overrides = BTreeMap::new();
        let mut folded_overrides = BTreeSet::new();
        let mut reader = part.reader();
        let mut depth = 0_usize;

        loop {
            let (resolution, event) = reader.read_resolved_event().map_err(|error| {
                opc_error(
                    "use.office.content_types_invalid",
                    format!("Failed to parse content types: {error}"),
                    CONTENT_TYPES_PART,
                )
            })?;
            match event {
                Event::Start(start) => {
                    let namespace = namespace(resolution)?;
                    if depth == 1 && namespace.as_deref() == Some(CONTENT_TYPES_NAMESPACE) {
                        read_content_type_declaration(
                            &part,
                            &start,
                            &reader,
                            &mut defaults,
                            &mut overrides,
                            &mut folded_overrides,
                        )?;
                    }
                    depth += 1;
                }
                Event::Empty(start) => {
                    let namespace = namespace(resolution)?;
                    if depth == 1 && namespace.as_deref() == Some(CONTENT_TYPES_NAMESPACE) {
                        read_content_type_declaration(
                            &part,
                            &start,
                            &reader,
                            &mut defaults,
                            &mut overrides,
                            &mut folded_overrides,
                        )?;
                    }
                }
                Event::End(_) => depth = depth.saturating_sub(1),
                Event::Eof => break,
                _ => {}
            }
        }

        Ok(Self {
            defaults,
            overrides,
        })
    }
}

/// The package or part that owns a set of OPC relationships.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case")]
pub enum RelationshipSource {
    Package,
    Part { part_name: String },
}

impl RelationshipSource {
    pub fn part_name(&self) -> Option<&str> {
        match self {
            Self::Package => None,
            Self::Part { part_name } => Some(part_name),
        }
    }
}

/// A safe, resolved relationship target. External targets are inert data.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "mode", rename_all = "kebab-case")]
pub enum RelationshipTarget {
    Internal {
        part_name: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        fragment: Option<String>,
    },
    External {
        uri: String,
    },
}

impl RelationshipTarget {
    pub fn internal_part_name(&self) -> Option<&str> {
        match self {
            Self::Internal { part_name, .. } => Some(part_name),
            Self::External { .. } => None,
        }
    }
}

/// One relationship from an OPC `.rels` part.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Relationship {
    pub id: String,
    pub relationship_type: String,
    pub target: RelationshipTarget,
}

/// Namespace-aware relationship graph for an OOXML package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RelationshipGraph {
    by_source: BTreeMap<RelationshipSource, Vec<Relationship>>,
}

impl RelationshipGraph {
    pub fn relationships_from(&self, source: &RelationshipSource) -> &[Relationship] {
        self.by_source.get(source).map_or(&[], Vec::as_slice)
    }

    pub fn relationship(&self, source: &RelationshipSource, id: &str) -> Option<&Relationship> {
        self.relationships_from(source)
            .iter()
            .find(|relationship| relationship.id == id)
    }

    pub fn sources(&self) -> impl Iterator<Item = &RelationshipSource> {
        self.by_source.keys()
    }

    pub fn relationships(&self) -> impl Iterator<Item = (&RelationshipSource, &Relationship)> {
        self.by_source.iter().flat_map(|(source, relationships)| {
            relationships
                .iter()
                .map(move |relationship| (source, relationship))
        })
    }

    fn read(package: &NativeOfficePackage, limits: XmlLimits) -> UseResult<Self> {
        let relationship_parts = package
            .part_names()
            .filter_map(relationship_source_from_part)
            .collect::<Vec<_>>();
        let mut by_source = BTreeMap::new();

        for (relationship_part, source) in relationship_parts {
            if let RelationshipSource::Part { part_name } = &source {
                if !package.contains_part(part_name) {
                    return Err(opc_error(
                        "use.office.relationship_source_missing",
                        format!(
                            "Relationship part '{relationship_part}' refers to missing source part '{part_name}'."
                        ),
                        &relationship_part,
                    ));
                }
            }
            let relationships = read_relationships(package, &relationship_part, &source, limits)?;
            if by_source.insert(source, relationships).is_some() {
                return Err(opc_error(
                    "use.office.relationship_source_duplicate",
                    "OOXML package contains duplicate relationship sources.",
                    &relationship_part,
                ));
            }
        }

        if !by_source.contains_key(&RelationshipSource::Package) {
            return Err(opc_error(
                "use.office.relationship_root_missing",
                "OOXML package has no root relationship part.",
                "_rels/.rels",
            ));
        }
        Ok(Self { by_source })
    }
}

/// Parsed OPC metadata shared by all native Office format engines.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpcPackageModel {
    content_types: ContentTypes,
    relationships: RelationshipGraph,
}

impl OpcPackageModel {
    pub fn read(package: &NativeOfficePackage) -> UseResult<Self> {
        Self::read_with_xml_limits(package, XmlLimits::default())
    }

    pub fn read_with_xml_limits(
        package: &NativeOfficePackage,
        limits: XmlLimits,
    ) -> UseResult<Self> {
        let content_types = ContentTypes::read(package, limits)?;
        validate_content_types(package, &content_types)?;
        let relationships = RelationshipGraph::read(package, limits)?;
        validate_relationship_targets(package, &relationships)?;
        validate_main_relationship(package.kind(), &relationships)?;
        Ok(Self {
            content_types,
            relationships,
        })
    }

    pub fn content_types(&self) -> &ContentTypes {
        &self.content_types
    }

    pub fn relationships(&self) -> &RelationshipGraph {
        &self.relationships
    }
}

impl NativeOfficePackage {
    pub fn opc_model(&self) -> UseResult<OpcPackageModel> {
        OpcPackageModel::read(self)
    }
}

fn read_content_type_declaration(
    part: &LosslessXmlPart,
    start: &quick_xml::events::BytesStart<'_>,
    reader: &quick_xml::reader::NsReader<&[u8]>,
    defaults: &mut BTreeMap<String, String>,
    overrides: &mut BTreeMap<String, String>,
    folded_overrides: &mut BTreeSet<String>,
) -> UseResult<()> {
    let local_name = local_name(part.name(), start.local_name().as_ref())?;
    let attributes = decode_attributes(part.name(), start, reader)?;
    match local_name.as_str() {
        "Default" => {
            let extension = required_attribute(part.name(), &attributes, "Extension")?;
            validate_extension(part.name(), extension)?;
            let content_type = required_attribute(part.name(), &attributes, "ContentType")?;
            validate_content_type(part.name(), content_type)?;
            let extension = extension.to_ascii_lowercase();
            if defaults
                .insert(extension.clone(), content_type.to_string())
                .is_some()
            {
                return Err(opc_error(
                    "use.office.content_types_duplicate",
                    format!("Duplicate default content type for extension '{extension}'."),
                    part.name(),
                ));
            }
        }
        "Override" => {
            let part_name = required_attribute(part.name(), &attributes, "PartName")?;
            let part_name = normalize_override_name(part.name(), part_name)?;
            let content_type = required_attribute(part.name(), &attributes, "ContentType")?;
            validate_content_type(part.name(), content_type)?;
            if !folded_overrides.insert(part_name.to_ascii_lowercase())
                || overrides
                    .insert(part_name.clone(), content_type.to_string())
                    .is_some()
            {
                return Err(opc_error(
                    "use.office.content_types_duplicate",
                    format!("Duplicate content type override for part '{part_name}'."),
                    part.name(),
                ));
            }
        }
        _ => {}
    }
    Ok(())
}

fn read_relationships(
    package: &NativeOfficePackage,
    relationship_part: &str,
    source: &RelationshipSource,
    limits: XmlLimits,
) -> UseResult<Vec<Relationship>> {
    let part = package.xml_part_with_limits(relationship_part, limits)?;
    require_root(
        &part,
        "Relationships",
        RELATIONSHIPS_NAMESPACE,
        "use.office.relationships_invalid",
    )?;
    let mut reader = part.reader();
    let mut depth = 0_usize;
    let mut relationships = Vec::new();
    let mut ids = BTreeSet::new();

    loop {
        let (resolution, event) = reader.read_resolved_event().map_err(|error| {
            opc_error(
                "use.office.relationships_invalid",
                format!("Failed to parse relationships: {error}"),
                relationship_part,
            )
        })?;
        match event {
            Event::Start(start) => {
                let namespace = namespace(resolution)?;
                if depth == 1
                    && namespace.as_deref() == Some(RELATIONSHIPS_NAMESPACE)
                    && start.local_name().as_ref() == b"Relationship"
                {
                    relationships.push(read_relationship(
                        relationship_part,
                        source,
                        &start,
                        &reader,
                        &mut ids,
                    )?);
                }
                depth += 1;
            }
            Event::Empty(start) => {
                let namespace = namespace(resolution)?;
                if depth == 1
                    && namespace.as_deref() == Some(RELATIONSHIPS_NAMESPACE)
                    && start.local_name().as_ref() == b"Relationship"
                {
                    relationships.push(read_relationship(
                        relationship_part,
                        source,
                        &start,
                        &reader,
                        &mut ids,
                    )?);
                }
            }
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(relationships)
}

fn read_relationship(
    relationship_part: &str,
    source: &RelationshipSource,
    start: &quick_xml::events::BytesStart<'_>,
    reader: &quick_xml::reader::NsReader<&[u8]>,
    ids: &mut BTreeSet<String>,
) -> UseResult<Relationship> {
    let attributes = decode_attributes(relationship_part, start, reader)?;
    let id = required_attribute(relationship_part, &attributes, "Id")?;
    validate_token(relationship_part, "relationship ID", id)?;
    if !ids.insert(id.to_string()) {
        return Err(opc_error(
            "use.office.relationship_id_duplicate",
            format!("Relationship ID '{id}' is duplicated."),
            relationship_part,
        ));
    }
    let relationship_type = required_attribute(relationship_part, &attributes, "Type")?;
    validate_uri_data(relationship_part, "relationship type", relationship_type)?;
    let raw_target = required_attribute(relationship_part, &attributes, "Target")?;
    let target = match attribute(&attributes, "TargetMode") {
        None | Some("Internal") => resolve_internal_target(relationship_part, source, raw_target)?,
        Some("External") => {
            validate_uri_data(
                relationship_part,
                "external relationship target",
                raw_target,
            )?;
            RelationshipTarget::External {
                uri: raw_target.to_string(),
            }
        }
        Some(mode) => {
            return Err(opc_error(
                "use.office.relationship_mode_invalid",
                format!("Relationship target mode '{mode}' is not Internal or External."),
                relationship_part,
            ));
        }
    };
    Ok(Relationship {
        id: id.to_string(),
        relationship_type: relationship_type.to_string(),
        target,
    })
}

fn resolve_internal_target(
    relationship_part: &str,
    source: &RelationshipSource,
    target: &str,
) -> UseResult<RelationshipTarget> {
    validate_uri_data(relationship_part, "internal relationship target", target)?;
    let (path, fragment) = target
        .split_once('#')
        .map_or((target, None), |(path, fragment)| {
            (path, (!fragment.is_empty()).then(|| fragment.to_string()))
        });
    if path.is_empty() || path.contains('?') || path.starts_with("//") || has_uri_scheme(path) {
        return Err(opc_error(
            "use.office.relationship_target_invalid",
            format!("Internal relationship target '{target}' is not a package part URI."),
            relationship_part,
        ));
    }

    let mut segments = if path.starts_with('/') {
        Vec::new()
    } else {
        source
            .part_name()
            .and_then(|part| part.rsplit_once('/').map(|(directory, _)| directory))
            .map_or_else(Vec::new, |directory| {
                directory.split('/').map(str::to_string).collect()
            })
    };
    for segment in path.trim_start_matches('/').split('/') {
        match segment {
            "" => {
                return Err(opc_error(
                    "use.office.relationship_target_invalid",
                    format!("Internal relationship target '{target}' contains an empty segment."),
                    relationship_part,
                ));
            }
            "." => {}
            ".." => {
                if segments.pop().is_none() {
                    return Err(opc_error(
                        "use.office.relationship_target_escape",
                        format!(
                            "Internal relationship target '{target}' escapes the package root."
                        ),
                        relationship_part,
                    ));
                }
            }
            value => segments.push(value.to_string()),
        }
    }
    if segments.is_empty() {
        return Err(opc_error(
            "use.office.relationship_target_invalid",
            format!("Internal relationship target '{target}' resolves to the package root."),
            relationship_part,
        ));
    }
    Ok(RelationshipTarget::Internal {
        part_name: segments.join("/"),
        fragment,
    })
}

fn validate_content_types(
    package: &NativeOfficePackage,
    content_types: &ContentTypes,
) -> UseResult<()> {
    for part_name in package.part_names() {
        if part_name == CONTENT_TYPES_PART {
            continue;
        }
        let Some(content_type) = content_types.content_type(part_name) else {
            return Err(opc_error(
                "use.office.content_type_missing",
                format!("OOXML package part '{part_name}' has no declared content type."),
                CONTENT_TYPES_PART,
            ));
        };
        if relationship_source_from_part(part_name).is_some()
            && content_type != RELATIONSHIPS_CONTENT_TYPE
        {
            return Err(opc_error(
                "use.office.content_type_mismatch",
                format!(
                    "Relationship part '{part_name}' must use content type '{RELATIONSHIPS_CONTENT_TYPE}'."
                ),
                CONTENT_TYPES_PART,
            ));
        }
    }

    let main_part = package.kind().main_part();
    let actual = content_types.content_type(main_part).ok_or_else(|| {
        opc_error(
            "use.office.content_type_missing",
            format!("Main document part '{main_part}' has no content type."),
            CONTENT_TYPES_PART,
        )
    })?;
    let expected = package.kind().main_content_type();
    if actual != expected {
        return Err(opc_error(
            "use.office.content_type_mismatch",
            format!(
                "Main document part '{main_part}' has content type '{actual}', expected '{expected}'."
            ),
            CONTENT_TYPES_PART,
        ));
    }
    for (part_name, _) in content_types.overrides() {
        if !package.contains_part(part_name) {
            return Err(opc_error(
                "use.office.content_type_part_missing",
                format!("Content type override references missing part '{part_name}'."),
                CONTENT_TYPES_PART,
            ));
        }
    }
    Ok(())
}

fn validate_relationship_targets(
    package: &NativeOfficePackage,
    relationships: &RelationshipGraph,
) -> UseResult<()> {
    for (source, relationship) in relationships.relationships() {
        let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
            continue;
        };
        if !package.contains_part(part_name) {
            return Err(opc_error(
                "use.office.relationship_target_missing",
                format!(
                    "Relationship '{}' from {:?} targets missing part '{part_name}'.",
                    relationship.id, source
                ),
                &relationship_part_for_source(source),
            ));
        }
    }
    Ok(())
}

fn validate_main_relationship(
    kind: DocumentKind,
    relationships: &RelationshipGraph,
) -> UseResult<()> {
    let office_documents = relationships
        .relationships_from(&RelationshipSource::Package)
        .iter()
        .filter(|relationship| {
            matches!(
                relationship.relationship_type.as_str(),
                OFFICE_DOCUMENT_RELATIONSHIP | STRICT_OFFICE_DOCUMENT_RELATIONSHIP
            )
        })
        .collect::<Vec<_>>();
    let [relationship] = office_documents.as_slice() else {
        return Err(opc_error(
            "use.office.main_relationship_invalid",
            format!(
                "OOXML package must contain exactly one root officeDocument relationship; found {}.",
                office_documents.len()
            ),
            "_rels/.rels",
        ));
    };
    if relationship.target.internal_part_name() != Some(kind.main_part()) {
        return Err(opc_error(
            "use.office.main_relationship_invalid",
            format!(
                "Root officeDocument relationship must target '{}'.",
                kind.main_part()
            ),
            "_rels/.rels",
        ));
    }
    Ok(())
}

fn relationship_source_from_part(part_name: &str) -> Option<(String, RelationshipSource)> {
    if part_name == "_rels/.rels" {
        return Some((part_name.to_string(), RelationshipSource::Package));
    }
    let (directory, relationship_name) = part_name.rsplit_once("/_rels/")?;
    let source_name = relationship_name.strip_suffix(".rels")?;
    if directory.is_empty() || source_name.is_empty() || source_name.contains('/') {
        return None;
    }
    Some((
        part_name.to_string(),
        RelationshipSource::Part {
            part_name: format!("{directory}/{source_name}"),
        },
    ))
}

fn relationship_part_for_source(source: &RelationshipSource) -> String {
    match source {
        RelationshipSource::Package => "_rels/.rels".to_string(),
        RelationshipSource::Part { part_name } => part_name.rsplit_once('/').map_or_else(
            || format!("_rels/{part_name}.rels"),
            |(directory, file_name)| format!("{directory}/_rels/{file_name}.rels"),
        ),
    }
}

fn normalize_override_name(metadata_part: &str, part_name: &str) -> UseResult<String> {
    let Some(part_name) = part_name.strip_prefix('/') else {
        return Err(opc_error(
            "use.office.content_types_invalid",
            format!("Content type override '{part_name}' must be an absolute part name."),
            metadata_part,
        ));
    };
    validate_part_path(metadata_part, part_name)?;
    Ok(part_name.to_string())
}

fn validate_part_path(metadata_part: &str, part_name: &str) -> UseResult<()> {
    if part_name.is_empty()
        || part_name.contains('\\')
        || part_name.chars().any(char::is_control)
        || part_name
            .split('/')
            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
    {
        return Err(opc_error(
            "use.office.relationship_target_invalid",
            format!("Package part name '{part_name}' is invalid."),
            metadata_part,
        ));
    }
    Ok(())
}

fn require_root(
    part: &LosslessXmlPart,
    local_name: &str,
    namespace: &str,
    code: &str,
) -> UseResult<()> {
    if part.root().local_name == local_name && part.root().namespace.as_deref() == Some(namespace) {
        return Ok(());
    }
    Err(opc_error(
        code,
        format!(
            "XML part '{}' must have root element '{{{namespace}}}{local_name}'.",
            part.name()
        ),
        part.name(),
    ))
}

fn namespace(resolution: ResolveResult<'_>) -> UseResult<Option<String>> {
    match resolution {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .map(|value| Some(value.to_string()))
            .map_err(|error| {
                office_error(
                    "use.office.xml_encoding_invalid",
                    format!("XML namespace is not valid UTF-8: {error}"),
                )
            }),
        ResolveResult::Unknown(prefix) => Err(office_error(
            "use.office.xml_namespace_invalid",
            format!(
                "XML uses unbound namespace prefix '{}'.",
                String::from_utf8_lossy(&prefix)
            ),
        )),
    }
}

fn local_name(part_name: &str, bytes: &[u8]) -> UseResult<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|error| {
            opc_error(
                "use.office.xml_encoding_invalid",
                format!("XML element name is not valid UTF-8: {error}"),
                part_name,
            )
        })
}

fn required_attribute<'a>(
    part_name: &str,
    attributes: &'a [(String, String)],
    name: &str,
) -> UseResult<&'a str> {
    attribute(attributes, name).ok_or_else(|| {
        opc_error(
            "use.office.xml_attribute_missing",
            format!("XML element is missing required attribute '{name}'."),
            part_name,
        )
    })
}

fn validate_extension(part_name: &str, extension: &str) -> UseResult<()> {
    if extension.is_empty()
        || extension.starts_with('.')
        || extension.contains(['/', '\\'])
        || extension.chars().any(char::is_control)
    {
        return Err(opc_error(
            "use.office.content_types_invalid",
            format!("Content type extension '{extension}' is invalid."),
            part_name,
        ));
    }
    Ok(())
}

fn validate_content_type(part_name: &str, content_type: &str) -> UseResult<()> {
    if content_type.is_empty()
        || !content_type.contains('/')
        || content_type.chars().any(char::is_control)
    {
        return Err(opc_error(
            "use.office.content_types_invalid",
            format!("Content type '{content_type}' is invalid."),
            part_name,
        ));
    }
    Ok(())
}

fn validate_token(part_name: &str, description: &str, value: &str) -> UseResult<()> {
    if value.is_empty() || value.chars().any(char::is_whitespace) {
        return Err(opc_error(
            "use.office.relationships_invalid",
            format!("The {description} '{value}' is invalid."),
            part_name,
        ));
    }
    Ok(())
}

fn validate_uri_data(part_name: &str, description: &str, value: &str) -> UseResult<()> {
    if value.is_empty() || value.chars().any(char::is_control) || value.contains('\\') {
        return Err(opc_error(
            "use.office.relationship_target_invalid",
            format!("The {description} '{value}' is invalid."),
            part_name,
        ));
    }
    Ok(())
}

fn has_uri_scheme(value: &str) -> bool {
    let Some((scheme, _)) = value.split_once(':') else {
        return false;
    };
    !scheme.is_empty()
        && scheme.chars().enumerate().all(|(index, character)| {
            if index == 0 {
                character.is_ascii_alphabetic()
            } else {
                character.is_ascii_alphanumeric() || matches!(character, '+' | '-' | '.')
            }
        })
}

fn opc_error(code: &str, message: impl Into<String>, part: &str) -> UseError {
    office_error(code, message).with_detail("part", part)
}

impl DocumentKind {
    pub(crate) fn main_content_type(self) -> &'static str {
        match self {
            Self::Word => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            }
            Self::Spreadsheet => {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            }
            Self::Presentation => {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            }
        }
    }
}
