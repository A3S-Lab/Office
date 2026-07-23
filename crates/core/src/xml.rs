use std::collections::BTreeSet;
use std::sync::Arc;

use a3s_use_core::{UseError, UseResult};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use serde::{Deserialize, Serialize};

use crate::discovery::office_error;

/// Resource limits applied while parsing one XML package part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct XmlLimits {
    pub max_depth: usize,
    pub max_events: usize,
    pub max_attributes_per_element: usize,
}

impl Default for XmlLimits {
    fn default() -> Self {
        Self {
            max_depth: 256,
            max_events: 5_000_000,
            max_attributes_per_element: 4_096,
        }
    }
}

/// The source encoding of a loss-preserved OOXML XML part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum XmlEncoding {
    Utf8,
    Utf16LittleEndian,
    Utf16BigEndian,
}

/// Namespace-aware name of the root element in an XML part.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct XmlRootName {
    pub namespace: Option<String>,
    pub local_name: String,
}

/// An immutable XML part that retains its original bytes exactly.
///
/// XML is validated into a safe UTF-8 parsing buffer, but `raw()` always returns
/// the original package bytes. This lets semantic readers support UTF-16 while
/// untouched parts remain byte-for-byte stable.
#[derive(Debug, Clone)]
pub struct LosslessXmlPart {
    name: String,
    raw: Arc<[u8]>,
    parse_bytes: Arc<[u8]>,
    encoding: XmlEncoding,
    root: XmlRootName,
}

impl LosslessXmlPart {
    pub fn parse(name: impl Into<String>, bytes: impl Into<Arc<[u8]>>) -> UseResult<Self> {
        Self::parse_with_limits(name, bytes, XmlLimits::default())
    }

    pub fn parse_with_limits(
        name: impl Into<String>,
        bytes: impl Into<Arc<[u8]>>,
        limits: XmlLimits,
    ) -> UseResult<Self> {
        validate_limits(limits)?;
        let name = name.into();
        let raw = bytes.into();
        let (parse_bytes, encoding) = prepare_xml_bytes(&name, raw.clone())?;
        let root = validate_xml(&name, &parse_bytes, limits)?;
        Ok(Self {
            name,
            raw,
            parse_bytes,
            encoding,
            root,
        })
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn raw(&self) -> &[u8] {
        &self.raw
    }

    pub fn encoding(&self) -> XmlEncoding {
        self.encoding
    }

    pub fn root(&self) -> &XmlRootName {
        &self.root
    }

    pub(crate) fn reader(&self) -> NsReader<&[u8]> {
        configured_reader(&self.parse_bytes)
    }

    pub(crate) fn parse_bytes(&self) -> &[u8] {
        &self.parse_bytes
    }

    pub(crate) fn raw_prefix(&self) -> &[u8] {
        if self.encoding == XmlEncoding::Utf8 {
            &self.raw[..self.raw.len().saturating_sub(self.parse_bytes.len())]
        } else {
            &[]
        }
    }
}

pub(crate) fn decode_attributes(
    part_name: &str,
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> UseResult<Vec<(String, String)>> {
    let mut values = Vec::new();
    let mut seen = BTreeSet::new();
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| {
            xml_error(
                part_name,
                "use.office.xml_invalid",
                format!("Invalid XML attribute: {error}"),
            )
        })?;
        let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
            xml_error(
                part_name,
                "use.office.xml_encoding_invalid",
                format!("XML attribute name is not valid UTF-8: {error}"),
            )
        })?;
        if !seen.insert(key.to_string()) {
            return Err(xml_error(
                part_name,
                "use.office.xml_invalid",
                format!("XML element contains duplicate attribute '{key}'."),
            ));
        }
        if !is_namespace_declaration(key)
            && matches!(
                reader.resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            )
        {
            return Err(xml_error(
                part_name,
                "use.office.xml_namespace_invalid",
                format!("XML attribute '{key}' uses an unbound namespace prefix."),
            ));
        }
        let value = attribute
            .decode_and_unescape_value(reader.decoder())
            .map_err(|error| {
                xml_error(
                    part_name,
                    "use.office.xml_invalid",
                    format!("Invalid value for XML attribute '{key}': {error}"),
                )
            })?
            .into_owned();
        values.push((key.to_string(), value));
    }
    Ok(values)
}

pub(crate) fn attribute<'a>(attributes: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|(key, _)| key == name)
        .map(|(_, value)| value.as_str())
}

fn configured_reader(bytes: &[u8]) -> NsReader<&[u8]> {
    let mut reader = NsReader::from_reader(bytes);
    let config = reader.config_mut();
    config.allow_dangling_amp = false;
    config.allow_unmatched_ends = false;
    config.check_comments = true;
    config.check_end_names = true;
    config.expand_empty_elements = false;
    config.trim_text(false);
    reader
}

fn validate_xml(part_name: &str, bytes: &[u8], limits: XmlLimits) -> UseResult<XmlRootName> {
    let mut reader = configured_reader(bytes);
    let mut depth = 0_usize;
    let mut event_count = 0_usize;
    let mut root = None;
    let mut declaration_seen = false;

    loop {
        let next = reader.read_resolved_event();
        let (resolution, event) = match next {
            Ok(event) => event,
            Err(error) => {
                let position = reader.error_position();
                return Err(xml_error(
                    part_name,
                    "use.office.xml_invalid",
                    format!("Malformed XML at byte {position}: {error}"),
                ));
            }
        };
        if matches!(event, Event::Eof) {
            break;
        }
        event_count = event_count.checked_add(1).ok_or_else(|| {
            xml_limit_error(
                part_name,
                "XML event count overflowed while parsing the part.",
            )
        })?;
        if event_count > limits.max_events {
            return Err(xml_limit_error(
                part_name,
                format!("XML part exceeds the {}-event limit.", limits.max_events),
            ));
        }

        match event {
            Event::Start(start) => {
                let namespace = resolved_namespace(part_name, resolution)?;
                validate_element(part_name, &start, &reader, limits)?;
                if depth == 0 {
                    if root.is_some() {
                        return Err(xml_error(
                            part_name,
                            "use.office.xml_invalid",
                            "XML part contains more than one root element.",
                        ));
                    }
                    root = Some(root_name(part_name, &start, namespace)?);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| xml_limit_error(part_name, "XML nesting depth overflowed."))?;
                if depth > limits.max_depth {
                    return Err(xml_limit_error(
                        part_name,
                        format!(
                            "XML part exceeds the {}-level depth limit.",
                            limits.max_depth
                        ),
                    ));
                }
            }
            Event::Empty(start) => {
                let namespace = resolved_namespace(part_name, resolution)?;
                validate_element(part_name, &start, &reader, limits)?;
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| xml_limit_error(part_name, "XML nesting depth overflowed."))?;
                if element_depth > limits.max_depth {
                    return Err(xml_limit_error(
                        part_name,
                        format!(
                            "XML part exceeds the {}-level depth limit.",
                            limits.max_depth
                        ),
                    ));
                }
                if depth == 0 {
                    if root.is_some() {
                        return Err(xml_error(
                            part_name,
                            "use.office.xml_invalid",
                            "XML part contains more than one root element.",
                        ));
                    }
                    root = Some(root_name(part_name, &start, namespace)?);
                }
            }
            Event::End(end) => {
                resolved_namespace(part_name, resolution)?;
                std::str::from_utf8(end.name().as_ref()).map_err(|error| {
                    xml_error(
                        part_name,
                        "use.office.xml_encoding_invalid",
                        format!("XML element name is not valid UTF-8: {error}"),
                    )
                })?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    xml_error(
                        part_name,
                        "use.office.xml_invalid",
                        "XML part contains an unmatched closing element.",
                    )
                })?;
            }
            Event::Text(text) => {
                let text = text.xml_content().map_err(|error| {
                    xml_error(
                        part_name,
                        "use.office.xml_encoding_invalid",
                        format!("XML text cannot be decoded: {error}"),
                    )
                })?;
                if depth == 0 && !text.trim().is_empty() {
                    return Err(xml_error(
                        part_name,
                        "use.office.xml_invalid",
                        "XML character data cannot appear outside the root element.",
                    ));
                }
            }
            Event::CData(text) => {
                text.decode().map_err(|error| {
                    xml_error(
                        part_name,
                        "use.office.xml_encoding_invalid",
                        format!("XML CDATA cannot be decoded: {error}"),
                    )
                })?;
                if depth == 0 {
                    return Err(xml_error(
                        part_name,
                        "use.office.xml_invalid",
                        "XML CDATA cannot appear outside the root element.",
                    ));
                }
            }
            Event::GeneralRef(reference) => {
                if depth == 0 {
                    return Err(xml_error(
                        part_name,
                        "use.office.xml_invalid",
                        "XML entity references cannot appear outside the root element.",
                    ));
                }
                validate_reference(
                    part_name,
                    &reference.decode().map_err(|error| {
                        xml_error(
                            part_name,
                            "use.office.xml_encoding_invalid",
                            format!("XML entity reference cannot be decoded: {error}"),
                        )
                    })?,
                )?;
            }
            Event::DocType(_) => {
                return Err(xml_error(
                    part_name,
                    "use.office.xml_doctype_forbidden",
                    "DTD declarations and external XML entities are forbidden.",
                ));
            }
            Event::Decl(declaration) => {
                if declaration_seen || root.is_some() || event_count != 1 {
                    return Err(xml_error(
                        part_name,
                        "use.office.xml_invalid",
                        "The XML declaration must appear at most once at the start of the part.",
                    ));
                }
                declaration_seen = true;
                declaration.version().map_err(|error| {
                    xml_error(
                        part_name,
                        "use.office.xml_invalid",
                        format!("Invalid XML declaration: {error}"),
                    )
                })?;
            }
            Event::PI(instruction) => {
                reader
                    .decoder()
                    .decode(instruction.as_ref())
                    .map_err(|error| {
                        xml_error(
                            part_name,
                            "use.office.xml_encoding_invalid",
                            format!("XML processing instruction cannot be decoded: {error}"),
                        )
                    })?;
            }
            Event::Comment(comment) => {
                comment.decode().map_err(|error| {
                    xml_error(
                        part_name,
                        "use.office.xml_encoding_invalid",
                        format!("XML comment cannot be decoded: {error}"),
                    )
                })?;
            }
            Event::Eof => unreachable!(),
        }
    }

    if depth != 0 {
        return Err(xml_error(
            part_name,
            "use.office.xml_invalid",
            "XML part ended before all elements were closed.",
        ));
    }
    root.ok_or_else(|| {
        xml_error(
            part_name,
            "use.office.xml_invalid",
            "XML part does not contain a root element.",
        )
    })
}

fn validate_element(
    part_name: &str,
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    limits: XmlLimits,
) -> UseResult<()> {
    std::str::from_utf8(start.name().as_ref()).map_err(|error| {
        xml_error(
            part_name,
            "use.office.xml_encoding_invalid",
            format!("XML element name is not valid UTF-8: {error}"),
        )
    })?;
    let attributes = decode_attributes(part_name, start, reader)?;
    if attributes.len() > limits.max_attributes_per_element {
        return Err(xml_limit_error(
            part_name,
            format!(
                "XML element exceeds the {}-attribute limit.",
                limits.max_attributes_per_element
            ),
        ));
    }
    Ok(())
}

fn root_name(
    part_name: &str,
    start: &BytesStart<'_>,
    namespace: Option<String>,
) -> UseResult<XmlRootName> {
    let local_name = std::str::from_utf8(start.local_name().as_ref())
        .map_err(|error| {
            xml_error(
                part_name,
                "use.office.xml_encoding_invalid",
                format!("XML root name is not valid UTF-8: {error}"),
            )
        })?
        .to_string();
    Ok(XmlRootName {
        namespace,
        local_name,
    })
}

fn resolved_namespace(part_name: &str, resolution: ResolveResult<'_>) -> UseResult<Option<String>> {
    match resolution {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .map(|value| Some(value.to_string()))
            .map_err(|error| {
                xml_error(
                    part_name,
                    "use.office.xml_encoding_invalid",
                    format!("XML namespace is not valid UTF-8: {error}"),
                )
            }),
        ResolveResult::Unknown(prefix) => Err(xml_error(
            part_name,
            "use.office.xml_namespace_invalid",
            format!(
                "XML element uses unbound namespace prefix '{}'.",
                String::from_utf8_lossy(&prefix)
            ),
        )),
    }
}

fn validate_reference(part_name: &str, reference: &str) -> UseResult<()> {
    if matches!(reference, "amp" | "lt" | "gt" | "apos" | "quot") {
        return Ok(());
    }
    let scalar = if let Some(hex) = reference.strip_prefix("#x") {
        u32::from_str_radix(hex, 16).ok()
    } else if let Some(decimal) = reference.strip_prefix('#') {
        decimal.parse::<u32>().ok()
    } else {
        None
    };
    if scalar.is_some_and(is_xml_character) {
        return Ok(());
    }
    Err(xml_error(
        part_name,
        "use.office.xml_entity_forbidden",
        format!("XML entity reference '&{reference};' is not a safe built-in reference."),
    ))
}

fn is_xml_character(value: u32) -> bool {
    matches!(value, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

fn prepare_xml_bytes(part_name: &str, raw: Arc<[u8]>) -> UseResult<(Arc<[u8]>, XmlEncoding)> {
    if let Some(bytes) = raw.strip_prefix(&[0xFF, 0xFE]) {
        return decode_utf16(part_name, bytes, true)
            .map(|bytes| (Arc::from(bytes), XmlEncoding::Utf16LittleEndian));
    }
    if let Some(bytes) = raw.strip_prefix(&[0xFE, 0xFF]) {
        return decode_utf16(part_name, bytes, false)
            .map(|bytes| (Arc::from(bytes), XmlEncoding::Utf16BigEndian));
    }
    if raw.starts_with(&[b'<', 0]) {
        return decode_utf16(part_name, &raw, true)
            .map(|bytes| (Arc::from(bytes), XmlEncoding::Utf16LittleEndian));
    }
    if raw.starts_with(&[0, b'<']) {
        return decode_utf16(part_name, &raw, false)
            .map(|bytes| (Arc::from(bytes), XmlEncoding::Utf16BigEndian));
    }
    let parse_bytes = raw
        .strip_prefix(&[0xEF, 0xBB, 0xBF])
        .map_or_else(|| raw.clone(), |bytes| Arc::from(bytes.to_vec()));
    Ok((parse_bytes, XmlEncoding::Utf8))
}

fn decode_utf16(part_name: &str, bytes: &[u8], little_endian: bool) -> UseResult<Vec<u8>> {
    let mut chunks = bytes.chunks_exact(2);
    let units = chunks
        .by_ref()
        .map(|chunk| {
            let pair = [chunk[0], chunk[1]];
            if little_endian {
                u16::from_le_bytes(pair)
            } else {
                u16::from_be_bytes(pair)
            }
        })
        .collect::<Vec<_>>();
    if !chunks.remainder().is_empty() {
        return Err(xml_error(
            part_name,
            "use.office.xml_encoding_invalid",
            "UTF-16 XML contains an incomplete code unit.",
        ));
    }
    String::from_utf16(&units)
        .map(|text| text.into_bytes())
        .map_err(|error| {
            xml_error(
                part_name,
                "use.office.xml_encoding_invalid",
                format!("UTF-16 XML contains an invalid surrogate pair: {error}"),
            )
        })
}

fn validate_limits(limits: XmlLimits) -> UseResult<()> {
    if limits.max_depth == 0 || limits.max_events == 0 || limits.max_attributes_per_element == 0 {
        return Err(office_error(
            "use.office.xml_limits_invalid",
            "XML limits must be positive.",
        ));
    }
    Ok(())
}

fn is_namespace_declaration(name: &str) -> bool {
    name == "xmlns" || name.starts_with("xmlns:")
}

fn xml_limit_error(part_name: &str, message: impl Into<String>) -> UseError {
    xml_error(part_name, "use.office.xml_limit_exceeded", message)
}

fn xml_error(part_name: &str, code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message).with_detail("part", part_name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preserves_raw_bytes_and_resolves_root_namespace() {
        let bytes = br#"<?xml version="1.0"?><w:document xmlns:w="urn:word"><w:p>Hi &amp; bye</w:p></w:document>"#;
        let part = LosslessXmlPart::parse("word/document.xml", bytes.as_slice()).unwrap();

        assert_eq!(part.raw(), bytes);
        assert_eq!(part.encoding(), XmlEncoding::Utf8);
        assert_eq!(part.root().local_name, "document");
        assert_eq!(part.root().namespace.as_deref(), Some("urn:word"));
    }

    #[test]
    fn supports_utf16_without_rewriting_original_bytes() {
        let text = "<?xml version=\"1.0\" encoding=\"UTF-16\"?><root>你好</root>";
        let mut bytes = vec![0xFF, 0xFE];
        bytes.extend(text.encode_utf16().flat_map(u16::to_le_bytes));

        let part = LosslessXmlPart::parse("customXml/item.xml", bytes.clone()).unwrap();

        assert_eq!(part.raw(), bytes);
        assert_eq!(part.encoding(), XmlEncoding::Utf16LittleEndian);
        assert_eq!(part.root().local_name, "root");

        let text = "<root>plain</root>";
        let little_endian = text
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        let part = LosslessXmlPart::parse("customXml/no-bom.xml", little_endian).unwrap();
        assert_eq!(part.encoding(), XmlEncoding::Utf16LittleEndian);

        let big_endian = text
            .encode_utf16()
            .flat_map(u16::to_be_bytes)
            .collect::<Vec<_>>();
        let part = LosslessXmlPart::parse("customXml/no-bom-be.xml", big_endian).unwrap();
        assert_eq!(part.encoding(), XmlEncoding::Utf16BigEndian);
    }

    #[test]
    fn rejects_doctype_custom_entities_and_unbound_prefixes() {
        let doctype =
            br#"<!DOCTYPE root [<!ENTITY x SYSTEM "file:///etc/passwd">]><root>&x;</root>"#;
        assert_eq!(
            LosslessXmlPart::parse("unsafe.xml", doctype.as_slice())
                .unwrap_err()
                .code,
            "use.office.xml_doctype_forbidden"
        );
        assert_eq!(
            LosslessXmlPart::parse("unsafe.xml", b"<root>&custom;</root>".as_slice())
                .unwrap_err()
                .code,
            "use.office.xml_entity_forbidden"
        );
        assert_eq!(
            LosslessXmlPart::parse("unsafe.xml", b"<x:root/>".as_slice())
                .unwrap_err()
                .code,
            "use.office.xml_namespace_invalid"
        );
    }

    #[test]
    fn applies_depth_event_and_attribute_limits() {
        let depth = XmlLimits {
            max_depth: 2,
            ..XmlLimits::default()
        };
        assert_eq!(
            LosslessXmlPart::parse_with_limits("deep.xml", b"<a><b><c/></b></a>".as_slice(), depth)
                .unwrap_err()
                .code,
            "use.office.xml_limit_exceeded"
        );

        let events = XmlLimits {
            max_events: 2,
            ..XmlLimits::default()
        };
        assert_eq!(
            LosslessXmlPart::parse_with_limits("events.xml", b"<a><b/></a>".as_slice(), events)
                .unwrap_err()
                .code,
            "use.office.xml_limit_exceeded"
        );

        let attributes = XmlLimits {
            max_attributes_per_element: 1,
            ..XmlLimits::default()
        };
        assert_eq!(
            LosslessXmlPart::parse_with_limits(
                "attrs.xml",
                b"<a first=\"1\" second=\"2\"/>".as_slice(),
                attributes,
            )
            .unwrap_err()
            .code,
            "use.office.xml_limit_exceeded"
        );
    }

    #[test]
    fn lossless_xml_part_is_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<LosslessXmlPart>();
    }
}
