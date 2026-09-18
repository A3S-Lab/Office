use a3s_use_core::UseResult;

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationDocumentTextFindResult,
    NativeOfficeCollaborationDocumentTextMatch, NativeOfficeCollaborationManifest,
    NativeOfficeCollaborationPresentationContainerKind,
};
use super::super::utf16_len;
use super::state::{ordered_container_ids, read_element_state};

const MAX_PRESENTATION_TEXT_FIND_MATCHES: usize = 4_096;

const CONTAINER_KINDS: [NativeOfficeCollaborationPresentationContainerKind; 3] = [
    NativeOfficeCollaborationPresentationContainerKind::Slide,
    NativeOfficeCollaborationPresentationContainerKind::Master,
    NativeOfficeCollaborationPresentationContainerKind::Layout,
];

/// List active scene elements whose `text` contains `search`.
///
/// One hit per element, in slide, master, then layout order. The edit that
/// follows is still `presentation-update-element`; this walk does not invent a
/// text-replace mutation.
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

    let mut matches = Vec::new();
    let mut match_count = 0_usize;
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
                let Some(byte_index) = text.find(search) else {
                    continue;
                };
                match_count += 1;
                if match_count > MAX_PRESENTATION_TEXT_FIND_MATCHES {
                    break;
                }
                if matches.len() >= limit {
                    continue;
                }
                let occurrence = u32::try_from(matches.len() + 1).map_err(|_| {
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
                    container_kind: Some(kind.as_str().to_owned()),
                    container_id: Some(container_id.clone()),
                    element_id: Some(element_id.clone()),
                    index_utf16: utf16_len(&text[..byte_index])?,
                });
            }
            if match_count > MAX_PRESENTATION_TEXT_FIND_MATCHES {
                break;
            }
        }
        if match_count > MAX_PRESENTATION_TEXT_FIND_MATCHES {
            break;
        }
    }
    if match_count > MAX_PRESENTATION_TEXT_FIND_MATCHES {
        match_count = MAX_PRESENTATION_TEXT_FIND_MATCHES;
    }
    Ok(NativeOfficeCollaborationDocumentTextFindResult {
        search: search.to_owned(),
        match_count,
        truncated: match_count > limit,
        matches,
    })
}
