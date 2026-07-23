use a3s_use_core::UseResult;

use super::{editor_error, prefix, preserve_space_attribute, qualified};
use crate::xml_edit::{apply_patches, escape_text, insert_child, IndexedXmlElement, XmlPatch};
use crate::LosslessXmlPart;

pub(super) fn insert_into_empty_target(
    part: &LosslessXmlPart,
    target: &IndexedXmlElement,
    text: &str,
) -> UseResult<Vec<u8>> {
    let container = match target.local_name.as_str() {
        "tc" => target
            .descendant("txBody")
            .and_then(|body| body.child("p", 1))
            .ok_or_else(|| {
                editor_error(
                    "use.office.presentation_table_cell_invalid",
                    "Presentation table cell has no text paragraph.",
                )
            })?,
        "p" | "r" | "fld" => target,
        _ => {
            return Err(editor_error(
                "use.office.mutation_type_unsupported",
                format!(
                    "Presentation element '{}' has no editable text run.",
                    target.local_name
                ),
            ))
        }
    };
    let namespace_prefix = prefix(&container.qualified_name);
    let text_tag = qualified(namespace_prefix, "t");
    let escaped = escape_text(text);
    let space = preserve_space_attribute(text);
    if matches!(container.local_name.as_str(), "r" | "fld") {
        return insert_child(
            part,
            container,
            format!("<{text_tag}{space}>{escaped}</{text_tag}>"),
        );
    }

    let run_tag = qualified(namespace_prefix, "r");
    let run_properties_tag = qualified(namespace_prefix, "rPr");
    let fragment = format!(
        "<{run_tag}><{run_properties_tag} lang=\"en-US\"/><{text_tag}{space}>{escaped}</{text_tag}></{run_tag}>"
    );
    if let Some(end_properties) = container
        .children
        .iter()
        .find(|child| child.local_name == "endParaRPr")
    {
        apply_patches(
            part,
            vec![XmlPatch::new(
                end_properties.full_range.start..end_properties.full_range.start,
                fragment,
            )],
        )
    } else {
        insert_child(part, container, fragment)
    }
}
