use a3s_use_core::UseResult;
use yrs::{GetString, Text, Transact};

use super::super::{
    collaboration_error, NativeOfficeCollaborationDocumentTextFindResult,
    NativeOfficeCollaborationDocumentTextMatch, NativeOfficeCollaborationManifest,
};
use super::{is_utf16_boundary, utf16_len};

const MAX_MARKDOWN_TEXT_REPLACEMENTS: u32 = 4_096;

pub(super) fn validate_text_replacement(
    search: &str,
    expected_matches: u32,
    occurrence: Option<u32>,
) -> UseResult<()> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Markdown text replacement requires a non-empty search string.",
        ));
    }
    if !(1..=MAX_MARKDOWN_TEXT_REPLACEMENTS).contains(&expected_matches) {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            format!(
                "Markdown text replacement expects between 1 and {MAX_MARKDOWN_TEXT_REPLACEMENTS} matches."
            ),
        )
        .with_detail("expectedMatches", expected_matches as u64)
        .with_detail(
            "maxExpectedMatches",
            MAX_MARKDOWN_TEXT_REPLACEMENTS as u64,
        ));
    }
    if let Some(occurrence) = occurrence {
        if !(1..=expected_matches).contains(&occurrence) {
            return Err(collaboration_error(
                "office.collaboration.mutation_invalid",
                format!(
                    "Markdown text replacement occurrence must be from 1 through the expected match count {expected_matches}."
                ),
            )
            .with_detail("occurrence", u64::from(occurrence))
            .with_detail("expectedMatches", expected_matches as u64));
        }
    }
    Ok(())
}

/// List Markdown source matches in the same non-overlapping walk replace uses.
pub(in crate::collaboration) fn find_markdown_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    limit: usize,
) -> UseResult<NativeOfficeCollaborationDocumentTextFindResult> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Markdown text find requires a non-empty search string.",
        ));
    }
    if limit == 0 {
        return Err(collaboration_error(
            "office.collaboration.find_limit_invalid",
            "Markdown text find limit must be at least 1.",
        ));
    }
    let source = markdown_source(doc, manifest);
    let offsets = collect_match_offsets(&source, search)?;
    let match_count = if offsets.len() > MAX_MARKDOWN_TEXT_REPLACEMENTS as usize {
        MAX_MARKDOWN_TEXT_REPLACEMENTS as usize
    } else {
        offsets.len()
    };
    let truncated = match_count > limit;
    let mut matches = Vec::new();
    for (index, index_utf16) in offsets.into_iter().take(limit).enumerate() {
        let occurrence = u32::try_from(index + 1).map_err(|_| {
            collaboration_error(
                "office.collaboration.mutation_too_large",
                "Markdown text find exceeded the supported occurrence range.",
            )
        })?;
        matches.push(NativeOfficeCollaborationDocumentTextMatch {
            occurrence,
            text: search.to_owned(),
            paragraph_id: None,
            text_id: None,
            sheet_id: None,
            row: None,
            column: None,
            container_kind: None,
            container_id: None,
            element_id: None,
            field_id: None,
            annotation_id: None,
            page_index: None,
            annotation_type: None,
            index_utf16,
        });
    }
    Ok(NativeOfficeCollaborationDocumentTextFindResult {
        search: search.to_owned(),
        match_count,
        truncated,
        matches,
    })
}

pub(super) fn replace_markdown_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    replacement: &str,
    expected_matches: u32,
    occurrence: Option<u32>,
) -> UseResult<()> {
    let root = format!("{}.markdown.source", manifest.namespace);
    let text = doc.get_or_insert_text(root);
    let source = text.get_string(&doc.transact());
    let delete_utf16 = utf16_len(search)?;
    let offsets = collect_match_offsets(&source, search)?;

    let actual_matches = u32::try_from(offsets.len()).unwrap_or(u32::MAX);
    if actual_matches != expected_matches {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            format!(
                "Markdown text replacement found {actual_matches} matches, not the expected {expected_matches}."
            ),
        )
        .with_suggestion(
            "Find Markdown text matches first, then retry with an exact search value and match count.",
        )
        .with_detail("actualMatches", actual_matches as u64)
        .with_detail("expectedMatches", expected_matches as u64));
    }
    let offsets = if let Some(occurrence) = occurrence {
        let Some(previous) = occurrence.checked_sub(1) else {
            return Err(collaboration_error(
                "office.collaboration.mutation_invalid",
                "Markdown text replacement occurrence is 1-based.",
            ));
        };
        let index = usize::try_from(previous).map_err(|_| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                "Markdown text replacement occurrence is outside the supported range.",
            )
        })?;
        vec![offsets.into_iter().nth(index).ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                "Markdown text replacement occurrence is outside the current matches.",
            )
        })?]
    } else {
        offsets
    };
    if search == replacement {
        return Ok(());
    }

    // Replace from the end so earlier UTF-16 offsets stay valid.
    let mut transaction = doc.transact_mut();
    for index_utf16 in offsets.into_iter().rev() {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
        if !replacement.is_empty() {
            text.insert(&mut transaction, index_utf16, replacement);
        }
    }
    Ok(())
}

pub(super) fn apply_markdown_replace(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    markdown: &str,
) -> UseResult<()> {
    let root = format!("{}.markdown.source", manifest.namespace);
    let text = doc.get_or_insert_text(root);
    let current = text.get_string(&doc.transact());
    let (index_utf16, delete_utf16, insert) = minimal_text_replacement(&current, markdown);
    if delete_utf16 == 0 && insert.is_empty() {
        return Ok(());
    }
    let mut transaction = doc.transact_mut();
    if delete_utf16 > 0 {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
    }
    if !insert.is_empty() {
        text.insert(&mut transaction, index_utf16, &insert);
    }
    Ok(())
}

pub(super) fn apply_markdown_splice(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    index_utf16: u32,
    delete_utf16: u32,
    insert: &str,
) -> UseResult<()> {
    let root = format!("{}.markdown.source", manifest.namespace);
    let text = doc.get_or_insert_text(root);
    let current = text.get_string(&doc.transact());
    let current_len = utf16_len(&current)?;
    let end_utf16 = index_utf16
        .checked_add(delete_utf16)
        .ok_or_else(|| invalid_markdown_range(index_utf16, delete_utf16, current_len))?;
    if index_utf16 > current_len
        || end_utf16 > current_len
        || !is_utf16_boundary(&current, index_utf16)
        || !is_utf16_boundary(&current, end_utf16)
    {
        return Err(invalid_markdown_range(
            index_utf16,
            delete_utf16,
            current_len,
        ));
    }
    if delete_utf16 == 0 && insert.is_empty() {
        return Ok(());
    }
    let mut transaction = doc.transact_mut();
    if delete_utf16 > 0 {
        text.remove_range(&mut transaction, index_utf16, delete_utf16);
    }
    if !insert.is_empty() {
        text.insert(&mut transaction, index_utf16, insert);
    }
    Ok(())
}

fn markdown_source(doc: &yrs::Doc, manifest: &NativeOfficeCollaborationManifest) -> String {
    let root = format!("{}.markdown.source", manifest.namespace);
    let text = doc.get_or_insert_text(root);
    text.get_string(&doc.transact())
}

fn collect_match_offsets(source: &str, search: &str) -> UseResult<Vec<u32>> {
    let mut offsets = Vec::new();
    let mut start = 0_usize;
    while let Some(relative) = source[start..].find(search) {
        let byte_index = start + relative;
        offsets.push(utf16_len(&source[..byte_index])?);
        if offsets.len() > MAX_MARKDOWN_TEXT_REPLACEMENTS as usize {
            break;
        }
        start = byte_index + search.len();
    }
    Ok(offsets)
}

fn minimal_text_replacement(current: &str, replacement: &str) -> (u32, u32, String) {
    let mut current_prefix_bytes = 0;
    let mut replacement_prefix_bytes = 0;
    let mut index_utf16 = 0;
    for (current_character, replacement_character) in current.chars().zip(replacement.chars()) {
        if current_character != replacement_character {
            break;
        }
        current_prefix_bytes += current_character.len_utf8();
        replacement_prefix_bytes += replacement_character.len_utf8();
        index_utf16 += current_character.len_utf16() as u32;
    }

    let current_tail = &current[current_prefix_bytes..];
    let replacement_tail = &replacement[replacement_prefix_bytes..];
    let mut current_suffix_bytes = 0;
    let mut replacement_suffix_bytes = 0;
    for (current_character, replacement_character) in current_tail
        .chars()
        .rev()
        .zip(replacement_tail.chars().rev())
    {
        if current_character != replacement_character {
            break;
        }
        current_suffix_bytes += current_character.len_utf8();
        replacement_suffix_bytes += replacement_character.len_utf8();
    }

    let current_middle = &current_tail[..current_tail.len() - current_suffix_bytes];
    let replacement_middle = &replacement_tail[..replacement_tail.len() - replacement_suffix_bytes];
    let delete_utf16 = current_middle.encode_utf16().count() as u32;
    let insert = replacement_middle.to_owned();
    (index_utf16, delete_utf16, insert)
}

fn invalid_markdown_range(
    index_utf16: u32,
    delete_utf16: u32,
    length_utf16: u32,
) -> a3s_use_core::UseError {
    collaboration_error(
        "office.collaboration.mutation_range_invalid",
        "The Markdown splice range is outside the current source or splits a UTF-16 surrogate pair.",
    )
    .with_suggestion("Inspect the latest state vector and retry with UTF-16 code-unit offsets.")
    .with_detail("indexUtf16", index_utf16 as u64)
    .with_detail("deleteUtf16", delete_utf16 as u64)
    .with_detail("lengthUtf16", length_utf16 as u64)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::collaboration::document::new_replica_document;
    use crate::collaboration::{
        NativeOfficeCollaborationActorKind, NativeOfficeCollaborationArtifactKind,
        NativeOfficeCollaborationMode,
    };

    fn manifest() -> NativeOfficeCollaborationManifest {
        NativeOfficeCollaborationManifest {
            format: String::new(),
            schema_version: 1,
            protocol: String::new(),
            protocol_version: 1,
            namespace: "a3s.office".to_owned(),
            artifact_id: "markdown".to_owned(),
            kind: NativeOfficeCollaborationArtifactKind::Markdown,
            actor_id: "agent".to_owned(),
            actor_kind: NativeOfficeCollaborationActorKind::Agent,
            mode: NativeOfficeCollaborationMode::Edit,
            client_id: 7,
        }
    }

    #[test]
    fn replacement_retains_unicode_scalar_boundaries() {
        assert_eq!(
            minimal_text_replacement("A😀B", "A🦀B"),
            (1, 2, "🦀".to_owned())
        );
    }

    #[test]
    fn find_and_replace_share_non_overlapping_utf16_walk() {
        let doc = new_replica_document(
            7,
            "a3s.office",
            NativeOfficeCollaborationArtifactKind::Markdown,
        );
        let manifest = manifest();
        {
            let root = format!("{}.markdown.source", manifest.namespace);
            let text = doc.get_or_insert_text(root);
            let mut tx = doc.transact_mut();
            text.insert(&mut tx, 0, "one Draft two Draft three Draft");
        }

        let found = find_markdown_text(&doc, &manifest, "Draft", 50).unwrap();
        assert_eq!(found.match_count, 3);
        assert_eq!(
            found
                .matches
                .iter()
                .map(|hit| hit.occurrence)
                .collect::<Vec<_>>(),
            vec![1, 2, 3]
        );
        assert_eq!(
            found
                .matches
                .iter()
                .map(|hit| hit.index_utf16)
                .collect::<Vec<_>>(),
            vec![4, 14, 26]
        );

        replace_markdown_text(&doc, &manifest, "Draft", "Final", 3, Some(2)).unwrap();
        assert_eq!(
            markdown_source(&doc, &manifest),
            "one Draft two Final three Draft"
        );

        let limited = find_markdown_text(&doc, &manifest, "Draft", 1).unwrap();
        assert_eq!(limited.match_count, 2);
        assert!(limited.truncated);
        assert_eq!(limited.matches.len(), 1);
        assert_eq!(limited.matches[0].occurrence, 1);
    }

    #[test]
    fn replace_rejects_match_count_conflict() {
        let doc = new_replica_document(
            8,
            "a3s.office",
            NativeOfficeCollaborationArtifactKind::Markdown,
        );
        let manifest = manifest();
        {
            let root = format!("{}.markdown.source", manifest.namespace);
            let text = doc.get_or_insert_text(root);
            let mut tx = doc.transact_mut();
            text.insert(&mut tx, 0, "Draft Draft");
        }
        let error = replace_markdown_text(&doc, &manifest, "Draft", "Final", 1, None).unwrap_err();
        assert_eq!(error.code, "office.collaboration.mutation_match_conflict");
    }
}
