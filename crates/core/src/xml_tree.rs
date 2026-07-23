use a3s_use_core::{UseError, UseResult};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use crate::discovery::office_error;
use crate::xml::LosslessXmlPart;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct XmlAttribute {
    pub namespace: Option<String>,
    pub local_name: String,
    pub value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum XmlNode {
    Element(XmlElement),
    Text(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct XmlElement {
    pub namespace: Option<String>,
    pub local_name: String,
    pub attributes: Vec<XmlAttribute>,
    pub children: Vec<XmlNode>,
}

impl XmlElement {
    pub fn attribute(&self, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| attribute.local_name == local_name)
            .map(|attribute| attribute.value.as_str())
    }

    pub fn attribute_ns(&self, namespace: &str, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace.as_deref() == Some(namespace)
                    && attribute.local_name == local_name
            })
            .map(|attribute| attribute.value.as_str())
    }

    pub fn child(&self, local_name: &str) -> Option<&XmlElement> {
        self.child_elements()
            .find(|element| element.local_name == local_name)
    }

    pub fn children_named<'a>(
        &'a self,
        local_name: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> + 'a {
        self.child_elements()
            .filter(move |element| element.local_name == local_name)
    }

    pub fn child_elements(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|child| match child {
            XmlNode::Element(element) => Some(element),
            XmlNode::Text(_) => None,
        })
    }

    pub fn descendants<'a>(&'a self, output: &mut Vec<&'a XmlElement>) {
        for child in self.child_elements() {
            output.push(child);
            child.descendants(output);
        }
    }
}

pub(crate) fn parse_xml_tree(part: &LosslessXmlPart) -> UseResult<XmlElement> {
    let mut reader = part.reader();
    let mut stack = Vec::new();
    let mut root = None;

    loop {
        let (resolution, event) = reader.read_resolved_event().map_err(|error| {
            tree_error(
                part.name(),
                format!("Failed to build semantic XML tree: {error}"),
            )
        })?;
        match event {
            Event::Start(start) => {
                let namespace = namespace(part.name(), resolution)?;
                let element = element(part.name(), namespace, &start, &reader)?;
                stack.push(element);
            }
            Event::Empty(start) => {
                let namespace = namespace(part.name(), resolution)?;
                let element = element(part.name(), namespace, &start, &reader)?;
                append_element(part.name(), &mut stack, &mut root, element)?;
            }
            Event::End(_) => {
                let element = stack.pop().ok_or_else(|| {
                    tree_error(
                        part.name(),
                        "XML semantic tree encountered an unmatched closing element.",
                    )
                })?;
                append_element(part.name(), &mut stack, &mut root, element)?;
            }
            Event::Text(text) => {
                let text = text.xml_content().map_err(|error| {
                    tree_error(part.name(), format!("Failed to decode XML text: {error}"))
                })?;
                append_text(&mut stack, text.into_owned());
            }
            Event::CData(text) => {
                let text = text.decode().map_err(|error| {
                    tree_error(part.name(), format!("Failed to decode XML CDATA: {error}"))
                })?;
                append_text(&mut stack, text.into_owned());
            }
            Event::GeneralRef(reference) => {
                let reference = reference.decode().map_err(|error| {
                    tree_error(
                        part.name(),
                        format!("Failed to decode XML entity reference: {error}"),
                    )
                })?;
                append_text(&mut stack, decode_reference(part.name(), &reference)?);
            }
            Event::Eof => break,
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {}
            Event::DocType(_) => {
                return Err(tree_error(
                    part.name(),
                    "DTD declarations are forbidden in semantic XML trees.",
                ));
            }
        }
    }
    if !stack.is_empty() {
        return Err(tree_error(
            part.name(),
            "XML semantic tree ended with unclosed elements.",
        ));
    }
    root.ok_or_else(|| tree_error(part.name(), "XML semantic tree has no root element."))
}

fn element(
    part_name: &str,
    element_namespace: Option<String>,
    start: &BytesStart<'_>,
    reader: &quick_xml::reader::NsReader<&[u8]>,
) -> UseResult<XmlElement> {
    let local_name = std::str::from_utf8(start.local_name().as_ref())
        .map(str::to_string)
        .map_err(|error| {
            tree_error(
                part_name,
                format!("XML element name is not valid UTF-8: {error}"),
            )
        })?;
    let mut attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute
            .map_err(|error| tree_error(part_name, format!("Invalid XML attribute: {error}")))?;
        let raw_name = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
            tree_error(
                part_name,
                format!("XML attribute name is not valid UTF-8: {error}"),
            )
        })?;
        if raw_name == "xmlns" || raw_name.starts_with("xmlns:") {
            continue;
        }
        let namespace = namespace(part_name, reader.resolve_attribute(attribute.key).0)?;
        let local_name = std::str::from_utf8(attribute.key.local_name().as_ref())
            .map(str::to_string)
            .map_err(|error| {
                tree_error(
                    part_name,
                    format!("XML attribute name is not valid UTF-8: {error}"),
                )
            })?;
        let value = attribute
            .decode_and_unescape_value(reader.decoder())
            .map_err(|error| {
                tree_error(
                    part_name,
                    format!("Failed to decode XML attribute '{raw_name}': {error}"),
                )
            })?
            .into_owned();
        attributes.push(XmlAttribute {
            namespace,
            local_name,
            value,
        });
    }
    Ok(XmlElement {
        namespace: element_namespace,
        local_name,
        attributes,
        children: Vec::new(),
    })
}

fn append_element(
    part_name: &str,
    stack: &mut [XmlElement],
    root: &mut Option<XmlElement>,
    element: XmlElement,
) -> UseResult<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Element(element));
    } else if root.replace(element).is_some() {
        return Err(tree_error(
            part_name,
            "XML semantic tree contains more than one root element.",
        ));
    }
    Ok(())
}

fn append_text(stack: &mut [XmlElement], text: String) {
    if let Some(parent) = stack.last_mut() {
        match parent.children.last_mut() {
            Some(XmlNode::Text(previous)) => previous.push_str(&text),
            _ => parent.children.push(XmlNode::Text(text)),
        }
    }
}

fn namespace(part_name: &str, resolution: ResolveResult<'_>) -> UseResult<Option<String>> {
    match resolution {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .map(|value| Some(value.to_string()))
            .map_err(|error| {
                tree_error(
                    part_name,
                    format!("XML namespace is not valid UTF-8: {error}"),
                )
            }),
        ResolveResult::Unknown(prefix) => Err(tree_error(
            part_name,
            format!(
                "XML uses unbound namespace prefix '{}'.",
                String::from_utf8_lossy(&prefix)
            ),
        )),
    }
}

fn decode_reference(part_name: &str, reference: &str) -> UseResult<String> {
    let value = match reference {
        "amp" => '&',
        "lt" => '<',
        "gt" => '>',
        "apos" => '\'',
        "quot" => '"',
        _ if reference.starts_with("#x") => u32::from_str_radix(&reference[2..], 16)
            .ok()
            .and_then(char::from_u32)
            .ok_or_else(|| invalid_reference(part_name, reference))?,
        _ if reference.starts_with('#') => reference[1..]
            .parse::<u32>()
            .ok()
            .and_then(char::from_u32)
            .ok_or_else(|| invalid_reference(part_name, reference))?,
        _ => return Err(invalid_reference(part_name, reference)),
    };
    Ok(value.to_string())
}

fn invalid_reference(part_name: &str, reference: &str) -> UseError {
    tree_error(
        part_name,
        format!("XML entity reference '&{reference};' cannot be decoded safely."),
    )
}

fn tree_error(part_name: &str, message: impl Into<String>) -> UseError {
    office_error("use.office.xml_tree_invalid", message).with_detail("part", part_name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builds_namespace_aware_tree_and_decodes_text() {
        let part = LosslessXmlPart::parse(
            "word/document.xml",
            br#"<w:document xmlns:w="urn:w" xmlns:r="urn:r"><w:p r:id="rId1"><w:t>A&amp;B</w:t></w:p></w:document>"#
                .as_slice(),
        )
        .unwrap();

        let root = parse_xml_tree(&part).unwrap();

        assert_eq!(root.namespace.as_deref(), Some("urn:w"));
        assert_eq!(root.local_name, "document");
        let paragraph = root.child("p").unwrap();
        assert_eq!(paragraph.attribute_ns("urn:r", "id"), Some("rId1"));
        assert_eq!(
            paragraph.child("t").unwrap().children[0],
            XmlNode::Text("A&B".into())
        );
    }
}
