use a3s_use_core::UseResult;

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationDocumentTextFindResult,
    NativeOfficeCollaborationDocumentTextMatch, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationPresentationContainerKind,
};
use super::super::utf16_len;
use super::element::write_element_text;
use super::state::{ordered_container_ids, read_element_state};

const MAX_PRESENTATION_TEXT_MATCHES: usize = 4_096;

const CONTAINER_KINDS: [NativeOfficeCollaborationPresentationContainerKind; 3] = [
    NativeOfficeCollaborationPresentationContainerKind::Slide,
    NativeOfficeCollaborationPresentationContainerKind::Master,
    NativeOfficeCollaborationPresentationContainerKind::Layout,
];

pub(super) struct PresentationTextHit {
    kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: String,
    element_id: String,
    byte_index: usize,
    index_utf16: u32,
}

/// List active scene-element text matches in the same walk replace uses.
///
/// Every non-overlapping match is one hit. Each hit still names the scene
/// element so a structural edit can use `presentation-update-element`.
pub(in crate::collaboration) fn find_presentation_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    limit: usize,
) -> UseResult<NativeOfficeCollaborationDocumentTextFindResult> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Presentation text find requires a non-empty search string.",
        ));
    }
    if limit == 0 {
        return Err(collaboration_error(
            "office.collaboration.find_limit_invalid",
            "Presentation text find limit must be at least 1.",
        ));
    }
    let hits = collect_presentation_text_hits(doc, manifest, search)?;
    let overflow = hits.len() > MAX_PRESENTATION_TEXT_MATCHES;
    let match_count = hits.len().min(MAX_PRESENTATION_TEXT_MATCHES);
    let truncated = overflow || match_count > limit;
    let mut matches = Vec::new();
    for (index, hit) in hits.into_iter().take(match_count.min(limit)).enumerate() {
        let occurrence = u32::try_from(index + 1).map_err(|_| {
            collaboration_error(
                "office.collaboration.mutation_too_large",
                "Presentation text find exceeded the supported occurrence range.",
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
            container_kind: Some(hit.kind.as_str().to_owned()),
            container_id: Some(hit.container_id),
            element_id: Some(hit.element_id),
            index_utf16: hit.index_utf16,
        });
    }
    Ok(NativeOfficeCollaborationDocumentTextFindResult {
        search: search.to_owned(),
        match_count,
        truncated,
        matches,
    })
}

pub(super) fn validate_presentation_text_replacement(
    search: &str,
    expected_matches: u32,
    occurrence: Option<u32>,
) -> UseResult<()> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "Presentation text replacement requires a non-empty search string.",
        ));
    }
    let max = u32::try_from(MAX_PRESENTATION_TEXT_MATCHES).unwrap_or(u32::MAX);
    if !(1..=max).contains(&expected_matches) {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            format!("Presentation text replacement expects between 1 and {max} matches."),
        )
        .with_detail("expectedMatches", expected_matches as u64)
        .with_detail("maxExpectedMatches", max as u64));
    }
    if let Some(occurrence) = occurrence {
        if !(1..=expected_matches).contains(&occurrence) {
            return Err(collaboration_error(
                "office.collaboration.mutation_invalid",
                format!(
                    "Presentation text replacement occurrence must be from 1 through the expected match count {expected_matches}."
                ),
            )
            .with_detail("occurrence", u64::from(occurrence))
            .with_detail("expectedMatches", expected_matches as u64));
        }
    }
    Ok(())
}

pub(super) fn replace_presentation_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    replacement: &str,
    expected_matches: u32,
    occurrence: Option<u32>,
) -> UseResult<()> {
    let hits = collect_presentation_text_hits(doc, manifest, search)?;
    if hits.len() > MAX_PRESENTATION_TEXT_MATCHES {
        return Err(collaboration_error(
            "office.collaboration.mutation_too_large",
            format!(
                "Presentation text replacement exceeds {MAX_PRESENTATION_TEXT_MATCHES} matches."
            ),
        ));
    }
    let actual_matches = u32::try_from(hits.len()).unwrap_or(u32::MAX);
    if actual_matches != expected_matches {
        return Err(collaboration_error(
            "office.collaboration.mutation_match_conflict",
            format!(
                "Presentation text replacement found {actual_matches} matches, not the expected {expected_matches}."
            ),
        )
        .with_suggestion(
            "Find Presentation text matches first, then retry with an exact search value and match count.",
        )
        .with_detail("actualMatches", actual_matches as u64)
        .with_detail("expectedMatches", expected_matches as u64));
    }
    let hits = if let Some(occurrence) = occurrence {
        let index = usize::try_from(occurrence - 1).map_err(|_| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                "Presentation text replacement occurrence is outside the supported range.",
            )
        })?;
        vec![hits.into_iter().nth(index).ok_or_else(|| {
            collaboration_error(
                "office.collaboration.mutation_invalid",
                "Presentation text replacement occurrence is outside the current matches.",
            )
        })?]
    } else {
        hits
    };
    if search == replacement {
        return Ok(());
    }

    let mut groups: Vec<PresentationTextHitGroup> = Vec::new();
    for hit in hits {
        if let Some(group) = groups.iter_mut().find(|group| {
            group.kind == hit.kind
                && group.container_id == hit.container_id
                && group.element_id == hit.element_id
        }) {
            group.byte_indexes.push(hit.byte_index);
        } else {
            groups.push(PresentationTextHitGroup {
                kind: hit.kind,
                container_id: hit.container_id,
                element_id: hit.element_id,
                byte_indexes: vec![hit.byte_index],
            });
        }
    }
    for group in groups {
        let state = read_element_state(doc, manifest, group.kind, &group.container_id)?;
        let text = state
            .records
            .get(&group.element_id)
            .and_then(|record| record.value.get("text"))
            .and_then(|value| value.as_str())
            .ok_or_else(|| {
                collaboration_error(
                    "office.collaboration.mutation_match_conflict",
                    format!(
                        "Presentation scene element '{}' no longer has text to replace.",
                        group.element_id
                    ),
                )
            })?
            .to_owned();
        let next = replace_byte_spans(&text, search, replacement, &group.byte_indexes);
        write_element_text(
            doc,
            manifest,
            group.kind,
            &group.container_id,
            &group.element_id,
            &next,
        )?;
    }
    Ok(())
}

struct PresentationTextHitGroup {
    kind: NativeOfficeCollaborationPresentationContainerKind,
    container_id: String,
    element_id: String,
    byte_indexes: Vec<usize>,
}

fn collect_presentation_text_hits(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
) -> UseResult<Vec<PresentationTextHit>> {
    let mut hits = Vec::new();
    for kind in CONTAINER_KINDS {
        for container_id in ordered_container_ids(doc, manifest, kind)? {
            let state = read_element_state(doc, manifest, kind, &container_id)?;
            for element_id in &state.active_order {
                let Some(record) = state.records.get(element_id) else {
                    continue;
                };
                if record.tombstoned {
                    continue;
                }
                let Some(text) = record.value.get("text").and_then(|value| value.as_str()) else {
                    continue;
                };
                let mut start = 0_usize;
                while let Some(relative) = text[start..].find(search) {
                    let byte_index = start + relative;
                    hits.push(PresentationTextHit {
                        kind,
                        container_id: container_id.clone(),
                        element_id: element_id.clone(),
                        byte_index,
                        index_utf16: utf16_len(&text[..byte_index])?,
                    });
                    if hits.len() > MAX_PRESENTATION_TEXT_MATCHES {
                        return Ok(hits);
                    }
                    start = byte_index + search.len();
                }
            }
        }
    }
    Ok(hits)
}

fn replace_byte_spans(
    text: &str,
    search: &str,
    replacement: &str,
    byte_indexes: &[usize],
) -> String {
    let mut next = text.to_owned();
    for &start in byte_indexes.iter().rev() {
        let end = start + search.len();
        if end <= next.len() {
            next.replace_range(start..end, replacement);
        }
    }
    next
}
