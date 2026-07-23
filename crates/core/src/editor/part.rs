use std::collections::BTreeMap;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use super::editor_error;
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::xml_edit::{index_xml, patch_start_tag_attributes, IndexedXmlElement};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage, RelationshipSource};

mod spreadsheet;

pub(super) use spreadsheet::worksheet_drawing;

const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEET: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_PRESENTATION: &str =
    "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATION: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const TRANSITIONAL_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
const TRANSITIONAL_SPREADSHEET_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_SPREADSHEET_DRAWING: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const TRANSITIONAL_DRAWING: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const TRANSITIONAL_PICTURE: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const STRICT_PICTURE: &str = "http://purl.oclc.org/ooxml/drawingml/picture";
const TRANSITIONAL_WORD_DRAWING: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const STRICT_WORD_DRAWING: &str = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";

const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const DRAWING_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const HEADER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

/// A known OOXML part type that the native engine can create safely.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficePartType {
    Chart,
    Header,
    Footer,
}

/// Receipt for a typed OOXML part created with its content type and relationship.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeCreatedPart {
    pub path: String,
    pub part: String,
    pub parent: String,
    pub owner_part: String,
    pub relationship_id: String,
    #[serde(rename = "type")]
    pub part_type: NativeOfficePartType,
}

pub(super) fn add(
    package: &mut NativeOfficePackage,
    parent: &str,
    part_type: NativeOfficePartType,
) -> UseResult<NativeCreatedPart> {
    match (package.kind(), part_type) {
        (DocumentKind::Word, NativeOfficePartType::Chart) => {
            add_word_part(package, parent, part_type, WordPart::Chart)
        }
        (DocumentKind::Word, NativeOfficePartType::Header) => {
            add_word_part(package, parent, part_type, WordPart::Header)
        }
        (DocumentKind::Word, NativeOfficePartType::Footer) => {
            add_word_part(package, parent, part_type, WordPart::Footer)
        }
        (DocumentKind::Spreadsheet, NativeOfficePartType::Chart) => {
            spreadsheet::add_chart(package, parent)
        }
        (DocumentKind::Presentation, NativeOfficePartType::Chart) => {
            add_presentation_chart(package, parent)
        }
        (kind, unsupported) => Err(part_error(
            "use.office.part_type_unsupported",
            format!(
                "Native {kind:?} documents do not support typed {unsupported:?} part creation."
            ),
        )),
    }
}

#[derive(Debug, Clone, Copy)]
enum WordPart {
    Chart,
    Header,
    Footer,
}

fn add_word_part(
    package: &mut NativeOfficePackage,
    parent: &str,
    part_type: NativeOfficePartType,
    word_part: WordPart,
) -> UseResult<NativeCreatedPart> {
    if parent != "/" {
        return Err(part_error(
            "use.office.part_parent_unsupported",
            "Native Word chart, header, and footer parts require the document root parent '/'.",
        ));
    }
    let dialect = dialect(package)?;
    let owner = "word/document.xml";
    let (directory, stem, content_type, relationship_name, root_xml) = match word_part {
        WordPart::Chart => (
            "word/charts",
            "chart",
            CHART_CONTENT_TYPE,
            "chart",
            chart_xml(dialect),
        ),
        WordPart::Header => (
            "word",
            "header",
            HEADER_CONTENT_TYPE,
            "header",
            word_container_xml("hdr", dialect),
        ),
        WordPart::Footer => (
            "word",
            "footer",
            FOOTER_CONTENT_TYPE,
            "footer",
            word_container_xml("ftr", dialect),
        ),
    };
    let position = relationship_count(package, owner, relationship_name)? + 1;
    create_owned_part(CreatePart {
        package,
        parent: "/",
        owner,
        directory,
        stem,
        content_type,
        relationship_type: &dialect.relationship_type(relationship_name),
        xml: &root_xml,
        path: &format!("/{relationship_name}[{position}]"),
        part_type,
    })
}

fn add_presentation_chart(
    package: &mut NativeOfficePackage,
    parent: &str,
) -> UseResult<NativeCreatedPart> {
    let slide = semantic_parent(package, parent, OfficeNodeType::Slide, "slide")?;
    let owner = source_part(&slide, "Presentation slide")?;
    let dialect = dialect(package)?;
    let position = relationship_count(package, &owner, "chart")? + 1;
    create_owned_part(CreatePart {
        package,
        parent: &slide.path,
        owner: &owner,
        directory: "ppt/charts",
        stem: "chart",
        content_type: CHART_CONTENT_TYPE,
        relationship_type: &dialect.relationship_type("chart"),
        xml: &chart_xml(dialect),
        path: &format!("{}/chart[{position}]", slide.path),
        part_type: NativeOfficePartType::Chart,
    })
}

struct CreatePart<'a> {
    package: &'a mut NativeOfficePackage,
    parent: &'a str,
    owner: &'a str,
    directory: &'a str,
    stem: &'a str,
    content_type: &'a str,
    relationship_type: &'a str,
    xml: &'a str,
    path: &'a str,
    part_type: NativeOfficePartType,
}

fn create_owned_part(request: CreatePart<'_>) -> UseResult<NativeCreatedPart> {
    let part_name = allocate_part(request.package, request.directory, request.stem)?;
    LosslessXmlPart::parse(part_name.clone(), request.xml.as_bytes().to_vec())?;
    crate::opc_edit::add_content_type_override(request.package, &part_name, request.content_type)?;
    request
        .package
        .set_part(&part_name, request.xml.as_bytes().to_vec())?;
    let relationship_id = crate::opc_edit::add_relationship(
        request.package,
        &relationship_part(request.owner),
        request.relationship_type,
        &relative_target(request.owner, &part_name),
    )?;
    Ok(NativeCreatedPart {
        path: request.path.to_string(),
        part: part_uri(&part_name),
        parent: request.parent.to_string(),
        owner_part: part_uri(request.owner),
        relationship_id,
        part_type: request.part_type,
    })
}

fn semantic_parent(
    package: &NativeOfficePackage,
    parent: &str,
    expected: OfficeNodeType,
    label: &str,
) -> UseResult<crate::DocumentNode> {
    let node = NativeOfficeDocument::from_package(package.clone())?.get(parent, 0)?;
    if node.node_type != expected {
        return Err(part_error(
            "use.office.part_parent_unsupported",
            format!("Native part creation requires a {label} parent."),
        ));
    }
    Ok(node)
}

fn source_part(node: &crate::DocumentNode, label: &str) -> UseResult<String> {
    node.format.get("part").cloned().ok_or_else(|| {
        part_error(
            "use.office.part_parent_invalid",
            format!("{label} has no source OOXML part."),
        )
    })
}

fn relationship_count(
    package: &NativeOfficePackage,
    owner: &str,
    relationship_name: &str,
) -> UseResult<usize> {
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    Ok(package
        .opc_model()?
        .relationships()
        .relationships_from(&source)
        .iter()
        .filter(|relationship| {
            relationship
                .relationship_type
                .ends_with(&format!("/{relationship_name}"))
        })
        .count())
}

fn allocate_part(package: &NativeOfficePackage, directory: &str, stem: &str) -> UseResult<String> {
    (1..=package.limits().max_entries.saturating_add(1))
        .map(|number| format!("{directory}/{stem}{number}.xml"))
        .find(|candidate| !package.contains_part(candidate))
        .ok_or_else(|| {
            part_error(
                "use.office.part_name_exhausted",
                format!("No available '{stem}' OOXML part name remains."),
            )
        })
}

pub(super) fn relationship_part(part_name: &str) -> String {
    part_name.rsplit_once('/').map_or_else(
        || format!("_rels/{part_name}.rels"),
        |(directory, file_name)| format!("{directory}/_rels/{file_name}.rels"),
    )
}

pub(super) fn relative_target(source: &str, target: &str) -> String {
    let source_directory = source
        .rsplit_once('/')
        .map(|(directory, _)| directory.split('/').collect::<Vec<_>>())
        .unwrap_or_default();
    let target_segments = target.split('/').collect::<Vec<_>>();
    let common = source_directory
        .iter()
        .zip(&target_segments)
        .take_while(|(left, right)| left == right)
        .count();
    std::iter::repeat_n("..", source_directory.len().saturating_sub(common))
        .chain(target_segments[common..].iter().copied())
        .collect::<Vec<_>>()
        .join("/")
}

pub(super) fn ensure_namespace(
    part: &LosslessXmlPart,
    preferred: &str,
    namespace: &str,
    error_code: &str,
) -> UseResult<(Vec<u8>, String)> {
    let root = index_xml(part)?;
    if let Some(prefix) = bound_prefix(&root, namespace) {
        return Ok((part.raw().to_vec(), prefix));
    }
    let prefix = (0..=64)
        .map(|offset| {
            if offset == 0 {
                preferred.to_string()
            } else {
                format!("{preferred}{offset}")
            }
        })
        .find(|candidate| {
            !root
                .qualified_attributes
                .contains_key(&format!("xmlns:{candidate}"))
        })
        .ok_or_else(|| {
            editor_error(
                error_code,
                format!("OOXML part '{}' has no free namespace prefix.", part.name()),
            )
        })?;
    let updates = BTreeMap::from([(format!("xmlns:{prefix}"), Some(namespace.to_string()))]);
    Ok((patch_start_tag_attributes(part, &root, &updates)?, prefix))
}

fn bound_prefix(root: &IndexedXmlElement, namespace: &str) -> Option<String> {
    root.qualified_attributes.iter().find_map(|(name, value)| {
        name.strip_prefix("xmlns:")
            .filter(|_| value == namespace)
            .map(str::to_string)
    })
}

#[derive(Debug, Clone, Copy)]
pub(super) enum OfficeDialect {
    Transitional,
    Strict,
}

impl OfficeDialect {
    pub(super) fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_RELATIONSHIPS,
            Self::Strict => STRICT_RELATIONSHIPS,
        }
    }

    pub(super) fn relationship_type(self, name: &str) -> String {
        format!("{}/{name}", self.relationship_namespace())
    }

    pub(super) fn word_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_WORD,
            Self::Strict => STRICT_WORD,
        }
    }

    pub(super) fn spreadsheet_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_SPREADSHEET,
            Self::Strict => STRICT_SPREADSHEET,
        }
    }

    fn chart_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_CHART,
            Self::Strict => STRICT_CHART,
        }
    }

    pub(super) fn spreadsheet_drawing_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_SPREADSHEET_DRAWING,
            Self::Strict => STRICT_SPREADSHEET_DRAWING,
        }
    }

    pub(super) fn drawing_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_DRAWING,
            Self::Strict => STRICT_DRAWING,
        }
    }

    pub(super) fn picture_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_PICTURE,
            Self::Strict => STRICT_PICTURE,
        }
    }

    pub(super) fn word_drawing_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_WORD_DRAWING,
            Self::Strict => STRICT_WORD_DRAWING,
        }
    }

    pub(super) fn presentation_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_PRESENTATION,
            Self::Strict => STRICT_PRESENTATION,
        }
    }
}

pub(super) fn dialect(package: &NativeOfficePackage) -> UseResult<OfficeDialect> {
    let namespace = package
        .xml_part(package.kind().main_part())?
        .root()
        .namespace
        .clone();
    match (package.kind(), namespace.as_deref()) {
        (DocumentKind::Word, Some(TRANSITIONAL_WORD))
        | (DocumentKind::Spreadsheet, Some(TRANSITIONAL_SPREADSHEET))
        | (DocumentKind::Presentation, Some(TRANSITIONAL_PRESENTATION)) => {
            Ok(OfficeDialect::Transitional)
        }
        (DocumentKind::Word, Some(STRICT_WORD))
        | (DocumentKind::Spreadsheet, Some(STRICT_SPREADSHEET))
        | (DocumentKind::Presentation, Some(STRICT_PRESENTATION)) => Ok(OfficeDialect::Strict),
        _ => Err(part_error(
            "use.office.part_dialect_unsupported",
            "Native part creation requires a recognized transitional or strict OOXML namespace.",
        )),
    }
}

fn chart_xml(dialect: OfficeDialect) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><c:chartSpace xmlns:c=\"{}\"><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart></c:chartSpace>",
        dialect.chart_namespace()
    )
}

fn word_container_xml(root: &str, dialect: OfficeDialect) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:{root} xmlns:w=\"{}\"><w:p/></w:{root}>",
        dialect.word_namespace()
    )
}

fn drawing_xml(dialect: OfficeDialect) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{}\"/>",
        dialect.spreadsheet_drawing_namespace()
    )
}

fn part_uri(part_name: &str) -> String {
    format!("/{part_name}")
}

fn part_error(code: &str, message: impl Into<String>) -> UseError {
    editor_error(code, message)
}
