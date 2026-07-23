use std::sync::Arc;

use a3s_use_core::{UseError, UseResult};
use quick_xml::events::Event;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use crate::discovery::office_error;
use crate::{LosslessXmlPart, NativeOfficePackage, XmlEncoding, XmlRootName};

/// A safely parsed XML package part exposed by the native raw read API.
///
/// `xml` is normalized to UTF-8 for callers. `byte_length` and `sha256` describe
/// the original package bytes, which remain available through
/// [`NativeOfficePackage::part`].
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeRawXmlPart {
    pub part: String,
    pub byte_length: u64,
    pub sha256: String,
    pub encoding: XmlEncoding,
    pub root: XmlRootName,
    pub xml: String,
}

pub(super) fn inspect(
    package: &NativeOfficePackage,
    requested_part: &str,
) -> UseResult<NativeRawXmlPart> {
    let part = package.xml_part(requested_part)?;
    let byte_length = u64::try_from(part.raw().len()).map_err(|_| {
        raw_error(
            "use.office.raw_part_too_large",
            part.name(),
            "OOXML XML part length cannot be represented on this platform.",
        )
    })?;
    let xml = std::str::from_utf8(part.parse_bytes())
        .map_err(|error| {
            raw_error(
                "use.office.raw_xml_invalid",
                part.name(),
                format!("OOXML XML part is not valid normalized UTF-8: {error}"),
            )
        })?
        .to_string();

    Ok(NativeRawXmlPart {
        part: part_uri(part.name()),
        byte_length,
        sha256: format!("{:x}", Sha256::digest(part.raw())),
        encoding: part.encoding(),
        root: part.root().clone(),
        xml,
    })
}

pub(super) fn replace(
    package: &mut NativeOfficePackage,
    requested_part: &str,
    xml: &str,
) -> UseResult<String> {
    let current = package.xml_part(requested_part)?;
    let part_name = current.name().to_string();
    require_mutable_part(&part_name)?;

    let input_length = u64::try_from(xml.len()).map_err(|_| {
        raw_error(
            "use.office.raw_input_too_large",
            &part_name,
            "Replacement XML length cannot be represented on this platform.",
        )
    })?;
    if input_length > package.limits().max_part_bytes {
        return Err(raw_error(
            "use.office.raw_input_too_large",
            &part_name,
            format!(
                "Replacement XML exceeds the {}-byte package part limit.",
                package.limits().max_part_bytes
            ),
        ));
    }

    let replacement = LosslessXmlPart::parse(
        part_name.clone(),
        Arc::<[u8]>::from(xml.as_bytes().to_vec()),
    )?;
    require_utf8_declaration(&replacement)?;
    if replacement.root() != current.root() {
        return Err(raw_error(
            "use.office.raw_root_mismatch",
            &part_name,
            format!(
                "Replacement root '{{{}}}{}' does not match existing root '{{{}}}{}'.",
                namespace_label(replacement.root()),
                replacement.root().local_name,
                namespace_label(current.root()),
                current.root().local_name
            ),
        )
        .with_detail("expectedRoot", serde_json::json!(current.root()))
        .with_detail("actualRoot", serde_json::json!(replacement.root())));
    }

    package.set_part(&part_name, replacement.raw().to_vec())?;
    Ok(part_uri(&part_name))
}

fn require_mutable_part(part_name: &str) -> UseResult<()> {
    if part_name.eq_ignore_ascii_case("[Content_Types].xml") || is_relationship_part(part_name) {
        return Err(raw_error(
            "use.office.raw_part_protected",
            part_name,
            "Content-type and relationship parts cannot be replaced through raw XML access.",
        )
        .with_suggestion(
            "Use a typed native Office operation that updates OPC metadata transactionally.",
        ));
    }
    Ok(())
}

fn is_relationship_part(part_name: &str) -> bool {
    let mut segments = part_name.rsplit('/');
    let Some(file_name) = segments.next() else {
        return false;
    };
    file_name.to_ascii_lowercase().ends_with(".rels")
        && segments
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("_rels"))
}

fn require_utf8_declaration(part: &LosslessXmlPart) -> UseResult<()> {
    let mut reader = part.reader();
    loop {
        let (_, event) = reader.read_resolved_event().map_err(|error| {
            raw_error(
                "use.office.raw_xml_invalid",
                part.name(),
                format!("Failed to inspect the replacement XML declaration: {error}"),
            )
        })?;
        match event {
            Event::Decl(declaration) => {
                let Some(encoding) = declaration.encoding() else {
                    return Ok(());
                };
                let encoding = encoding.map_err(|error| {
                    raw_error(
                        "use.office.raw_xml_invalid",
                        part.name(),
                        format!("Replacement XML has an invalid encoding declaration: {error}"),
                    )
                })?;
                if encoding.eq_ignore_ascii_case(b"utf-8") {
                    return Ok(());
                }
                let encoding = String::from_utf8_lossy(&encoding);
                return Err(raw_error(
                    "use.office.raw_encoding_unsupported",
                    part.name(),
                    format!(
                        "Replacement XML declares '{encoding}'; raw replacement accepts UTF-8 XML only."
                    ),
                ));
            }
            Event::Start(_) | Event::Empty(_) | Event::Eof => return Ok(()),
            _ => {}
        }
    }
}

fn namespace_label(root: &XmlRootName) -> &str {
    root.namespace.as_deref().unwrap_or("")
}

fn part_uri(part_name: &str) -> String {
    format!("/{part_name}")
}

fn raw_error(code: &str, part_name: &str, message: impl Into<String>) -> UseError {
    office_error(code, message).with_detail("part", part_uri(part_name))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{NativeOfficeEditor, NativeOfficeMutation};

    const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn word_document(text: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
        )
    }

    #[tokio::test]
    async fn inspects_and_replaces_an_existing_xml_part() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("raw.docx");
        let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

        let before = editor.raw_xml_part("/word/document.xml").unwrap();
        assert_eq!(before.part, "/word/document.xml");
        assert_eq!(before.encoding, XmlEncoding::Utf8);
        assert_eq!(before.root.local_name, "document");
        assert_eq!(before.root.namespace.as_deref(), Some(WORD_NAMESPACE));
        assert_eq!(before.sha256.len(), 64);
        assert_eq!(before.byte_length as usize, before.xml.len());

        let result = editor
            .replace_xml_part("word/document.xml", word_document("Native raw"))
            .unwrap();
        assert_eq!(result, "/word/document.xml");
        assert_eq!(editor.snapshot().unwrap().text_view().text, "Native raw");
    }

    #[tokio::test]
    async fn rejects_protected_parts_unsafe_xml_and_root_changes_without_dirtying() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("safe.docx");
        let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
        let original = editor.package().part("word/document.xml").unwrap().to_vec();

        let protected = editor
            .replace_xml_part("/[Content_Types].xml", "<Types/>")
            .unwrap_err();
        assert_eq!(protected.code, "use.office.raw_part_protected");
        let relationships = editor
            .replace_xml_part("/word/_rels/document.xml.rels", "<Relationships/>")
            .unwrap_err();
        assert_eq!(relationships.code, "use.office.raw_part_protected");

        let mismatch = editor
            .replace_xml_part("word/document.xml", "<document/>")
            .unwrap_err();
        assert_eq!(mismatch.code, "use.office.raw_root_mismatch");

        let doctype = editor
            .replace_xml_part(
                "word/document.xml",
                format!(
                    r#"<!DOCTYPE w:document [<!ENTITY x SYSTEM "file:///etc/passwd">]><w:document xmlns:w="{WORD_NAMESPACE}"><w:body/></w:document>"#
                ),
            )
            .unwrap_err();
        assert_eq!(doctype.code, "use.office.xml_doctype_forbidden");

        let encoding = editor
            .replace_xml_part(
                "word/document.xml",
                word_document("encoding").replace("UTF-8", "UTF-16"),
            )
            .unwrap_err();
        assert_eq!(encoding.code, "use.office.raw_encoding_unsupported");
        assert_eq!(
            editor.package().part("word/document.xml").unwrap(),
            original
        );
        assert!(!editor.is_dirty());
    }

    #[tokio::test]
    async fn replacement_participates_in_semantic_validation_and_batch_rollback() {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join("rollback.docx");
        let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
        let original = editor.package().part("word/document.xml").unwrap().to_vec();

        let invalid_semantics =
            format!(r#"<w:document xmlns:w="{WORD_NAMESPACE}"><w:notBody/></w:document>"#);
        let error = editor
            .replace_xml_part("word/document.xml", invalid_semantics)
            .unwrap_err();
        assert_eq!(error.code, "use.office.batch_validation_failed");
        assert_eq!(
            editor.package().part("word/document.xml").unwrap(),
            original
        );

        let mutations = [
            NativeOfficeMutation::ReplaceXmlPart {
                part: "/word/document.xml".to_string(),
                xml: word_document("must roll back"),
            },
            NativeOfficeMutation::SetText {
                path: "/body/p[999]".to_string(),
                text: "missing".to_string(),
            },
        ];
        let error = editor.apply_batch(&mutations).unwrap_err();
        assert_eq!(error.code, "use.office.node_not_found");
        assert_eq!(
            editor.package().part("word/document.xml").unwrap(),
            original
        );
        assert!(!editor.is_dirty());

        let json = serde_json::to_value(&mutations[0]).unwrap();
        assert_eq!(json["operation"], "replace-xml-part");
        assert_eq!(json["part"], "/word/document.xml");
    }
}
