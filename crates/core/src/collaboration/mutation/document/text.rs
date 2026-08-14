use a3s_use_core::UseResult;
use yrs::types::Attrs;
use yrs::{Any, Out, ReadTxn, Text, Transact, Xml, XmlElementRef, XmlFragment, XmlOut, XmlTextRef};

use super::super::super::{collaboration_error, NativeOfficeCollaborationManifest};
use super::super::utf16_len;
use super::identity::paragraph_text_id_rotations;

const MAX_DOCUMENT_TEXT_REPLACEMENTS: u32 = 4_096;

pub(super) fn validate_text_replacement(search: &str, expected_matches: u32) -> UseResult<()> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Document text replacement requires a non-empty search string.",
        ));
    }
    if !(1..=MAX_DOCUMENT_TEXT_REPLACEMENTS).contains(&expected_matches) {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            format!(
                "Document text replacement expects between 1 and {MAX_DOCUMENT_TEXT_REPLACEMENTS} matches."
            ),
        )
        .with_detail("expectedMatches", expected_matches as u64)
        .with_detail(
            "maxExpectedMatches",
            MAX_DOCUMENT_TEXT_REPLACEMENTS as u64,
        ));
    }
    Ok(())
}

pub(super) fn replace_document_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    replacement: &str,
    expected_matches: u32,
) -> UseResult<()> {
    let root = format!("{}.document.content", manifest.namespace);
    let fragment = doc.get_or_insert_xml_fragment(root);
    let transaction = doc.transact();
    let delete_utf16 = utf16_len(search)?;
    let mut replacements = Vec::new();
    for node in fragment.successors(&transaction) {
        let XmlOut::Text(text) = node else {
            continue;
        };
        collect_text_replacements(&text, &transaction, search, delete_utf16, &mut replacements)?;
        if replacements.len() > MAX_DOCUMENT_TEXT_REPLACEMENTS as usize {
            break;
        }
    }

    let actual_matches = u32::try_from(replacements.len()).unwrap_or(u32::MAX);
    if actual_matches != expected_matches {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            format!(
                "Document text replacement found {actual_matches} matches, not the expected {expected_matches}."
            ),
        )
        .with_suggestion(
            "Read the latest collaborative Document state and retry with an exact search value and match count.",
        )
        .with_detail("actualMatches", actual_matches as u64)
        .with_detail("expectedMatches", expected_matches as u64));
    }
    if search == replacement {
        return Ok(());
    }

    let mut paragraphs = Vec::<XmlElementRef>::new();
    for replacement_target in &replacements {
        let Some(XmlOut::Element(parent)) = replacement_target.text.parent() else {
            continue;
        };
        if !paragraphs.contains(&parent) {
            paragraphs.push(parent);
        }
    }
    let text_id_rotations = paragraph_text_id_rotations(&paragraphs, &transaction)?;
    drop(transaction);

    // Replacing from the end keeps every previously computed UTF-16 offset
    // valid within its Y.XmlText. The store lock prevents a local concurrent
    // writer from changing this in-memory replica between scan and commit.
    let mut transaction = doc.transact_mut();
    for replacement_target in replacements.into_iter().rev() {
        replacement_target.text.remove_range(
            &mut transaction,
            replacement_target.index_utf16,
            replacement_target.delete_utf16,
        );
        if replacement.is_empty() {
            continue;
        }
        match replacement_target.attributes {
            Some(attributes) => replacement_target.text.insert_with_attributes(
                &mut transaction,
                replacement_target.index_utf16,
                replacement,
                attributes,
            ),
            None => replacement_target.text.insert(
                &mut transaction,
                replacement_target.index_utf16,
                replacement,
            ),
        }
    }
    for (paragraph, next_text_id) in text_id_rotations {
        paragraph.insert_attribute(&mut transaction, "textId", next_text_id);
    }
    Ok(())
}

fn collect_text_replacements<T: ReadTxn>(
    text: &XmlTextRef,
    transaction: &T,
    search: &str,
    delete_utf16: u32,
    output: &mut Vec<DocumentTextReplacement>,
) -> UseResult<()> {
    for run in plain_text_runs(text, transaction)? {
        for (byte_index, _) in run.text.match_indices(search) {
            let relative_utf16 = utf16_len(&run.text[..byte_index])?;
            let index_utf16 = run.start_utf16.checked_add(relative_utf16).ok_or_else(|| {
                collaboration_error(
                    "office.collaboration.mutation_too_large",
                    "The Document text offset exceeds the supported UTF-16 range.",
                )
            })?;
            let attributes = run
                .format_spans
                .iter()
                .find(|span| span.start_utf16 <= index_utf16 && index_utf16 < span.end_utf16)
                .and_then(|span| span.attributes.clone());
            output.push(DocumentTextReplacement {
                text: text.clone(),
                index_utf16,
                delete_utf16,
                attributes,
            });
            if output.len() > MAX_DOCUMENT_TEXT_REPLACEMENTS as usize {
                return Ok(());
            }
        }
    }
    Ok(())
}

fn plain_text_runs<T: ReadTxn>(text: &XmlTextRef, transaction: &T) -> UseResult<Vec<PlainTextRun>> {
    let mut runs = Vec::new();
    let mut current = PlainTextRun::default();
    let mut cursor_utf16 = 0_u32;

    for chunk in text.diff(transaction, |_| ()) {
        match chunk.insert {
            Out::Any(Any::String(value)) => {
                if value.is_empty() {
                    continue;
                }
                if current.text.is_empty() {
                    current.start_utf16 = cursor_utf16;
                }
                let chunk_len = utf16_len(&value)?;
                let end_utf16 = cursor_utf16.checked_add(chunk_len).ok_or_else(|| {
                    collaboration_error(
                        "office.collaboration.mutation_too_large",
                        "The Document text exceeds the supported UTF-16 offset range.",
                    )
                })?;
                current.text.push_str(&value);
                current.format_spans.push(FormatSpan {
                    start_utf16: cursor_utf16,
                    end_utf16,
                    attributes: chunk.attributes.map(|attributes| *attributes),
                });
                cursor_utf16 = end_utf16;
            }
            _ => {
                if !current.text.is_empty() {
                    runs.push(std::mem::take(&mut current));
                }
                // Y.Text embeds occupy one logical position. Treat them as a
                // hard search boundary so a replacement never deletes a
                // ProseMirror inline atom by accident.
                cursor_utf16 = cursor_utf16.checked_add(1).ok_or_else(|| {
                    collaboration_error(
                        "office.collaboration.mutation_too_large",
                        "The Document text exceeds the supported UTF-16 offset range.",
                    )
                })?;
            }
        }
    }
    if !current.text.is_empty() {
        runs.push(current);
    }
    Ok(runs)
}

#[derive(Debug)]
struct DocumentTextReplacement {
    text: XmlTextRef,
    index_utf16: u32,
    delete_utf16: u32,
    attributes: Option<Attrs>,
}

#[derive(Debug, Default)]
struct PlainTextRun {
    start_utf16: u32,
    text: String,
    format_spans: Vec<FormatSpan>,
}

#[derive(Debug)]
struct FormatSpan {
    start_utf16: u32,
    end_utf16: u32,
    attributes: Option<Attrs>,
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::sync::Arc;

    use yrs::{Any, Text, Transact, Xml, XmlElementPrelim, XmlFragment, XmlTextPrelim};

    use super::*;
    use crate::collaboration::document::new_replica_document;
    use crate::collaboration::{
        NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
        NativeOfficeCollaborationMode, NativeOfficeCollaborationMutation,
    };

    #[test]
    fn replacement_contract_rejects_empty_or_unbounded_match_counts() {
        for mutation in [
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: String::new(),
                replacement: "value".to_owned(),
                expected_matches: 1,
            },
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: "value".to_owned(),
                replacement: "next".to_owned(),
                expected_matches: 0,
            },
            NativeOfficeCollaborationMutation::DocumentReplaceText {
                search: "value".to_owned(),
                replacement: "next".to_owned(),
                expected_matches: MAX_DOCUMENT_TEXT_REPLACEMENTS + 1,
            },
        ] {
            let error = super::super::validate_document_mutation(&mutation).unwrap_err();
            assert_eq!(error.code, "office.collaboration.mutation_invalid");
        }
    }

    #[test]
    fn replacement_preserves_format_and_rotates_the_word_text_identity_once() {
        let doc = new_replica_document(
            7,
            "a3s.office",
            NativeOfficeCollaborationArtifactKind::Document,
        );
        let fragment = doc.get_or_insert_xml_fragment("a3s.office.document.content");
        let (paragraph, text) = {
            let mut transaction = doc.transact_mut();
            let paragraph =
                fragment.push_back(&mut transaction, XmlElementPrelim::empty("paragraph"));
            paragraph.insert_attribute(&mut transaction, "paragraphId", "00000001");
            paragraph.insert_attribute(&mut transaction, "textId", "00000002");
            let text =
                paragraph.push_back(&mut transaction, XmlTextPrelim::new("plain bold bold end"));
            let mut attributes = Attrs::new();
            attributes.insert("bold".into(), Any::Map(Arc::new(HashMap::new())));
            text.format(&mut transaction, 6, 9, attributes);
            (paragraph, text)
        };
        let manifest = manifest();

        replace_document_text(&doc, &manifest, "bold", "strong", 2).unwrap();

        let transaction = doc.transact();
        let chunks = text.diff(&transaction, |_| ());
        assert!(chunks.iter().any(|chunk| {
            matches!(&chunk.insert, Out::Any(Any::String(value)) if value.as_ref().contains("strong"))
                && chunk
                    .attributes
                    .as_ref()
                    .is_some_and(|attributes| attributes.contains_key("bold"))
        }));
        assert!(matches!(
            paragraph.get_attribute(&transaction, "textId"),
            Some(Out::Any(Any::String(value))) if value.as_ref() == "00000003"
        ));
    }

    fn manifest() -> NativeOfficeCollaborationManifest {
        NativeOfficeCollaborationManifest {
            format: String::new(),
            schema_version: 1,
            protocol: String::new(),
            protocol_version: 1,
            namespace: "a3s.office".to_owned(),
            artifact_id: "document".to_owned(),
            kind: NativeOfficeCollaborationArtifactKind::Document,
            actor_id: "agent".to_owned(),
            actor_kind: NativeOfficeCollaborationActorKind::Agent,
            mode: NativeOfficeCollaborationMode::Edit,
            client_id: 7,
        }
    }
}
