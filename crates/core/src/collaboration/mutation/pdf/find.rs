use a3s_use_core::UseResult;

use super::super::super::{
    collaboration_error, NativeOfficeCollaborationDocumentTextFindResult,
    NativeOfficeCollaborationDocumentTextMatch, NativeOfficeCollaborationManifest,
};
use super::super::utf16_len;
use super::annotation::ordered_freetext_contents;
use super::{ordered_pdf_form_values, PdfRecordCollectionRoots};

const MAX_PDF_TEXT_FIND_MATCHES: usize = 4_096;
const FREETEXT_ANNOTATION_TYPE: u32 = 3;

/// List PDF form values and FreeText annotation contents that contain `search`.
///
/// Walk order is form-value collection order, then FreeText annotation order.
/// One hit per matching form field or FreeText annotation (first substring).
/// Following edits stay `pdf-set-form-value` / `pdf-update-annotation`; this
/// walk does not invent PDF body replace-text.
pub(in crate::collaboration) fn find_pdf_text(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    search: &str,
    limit: usize,
) -> UseResult<NativeOfficeCollaborationDocumentTextFindResult> {
    if search.is_empty() {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "PDF text find requires a non-empty search string.",
        ));
    }
    if limit == 0 {
        return Err(collaboration_error(
            "office.collaboration.find_limit_invalid",
            "PDF text find limit must be at least 1.",
        ));
    }

    let mut matches = Vec::new();
    let mut match_count = 0_usize;

    let form_roots = PdfRecordCollectionRoots::new(doc, manifest, "form-values");
    for (field_id, value) in ordered_pdf_form_values(doc, &form_roots)? {
        let Some(byte_index) = value.find(search) else {
            continue;
        };
        push_match(
            &mut matches,
            &mut match_count,
            limit,
            search,
            PdfFindHit {
                field_id: Some(field_id),
                annotation_id: None,
                page_index: None,
                annotation_type: None,
                index_utf16: utf16_len(&value[..byte_index])?,
            },
        )?;
        if match_count > MAX_PDF_TEXT_FIND_MATCHES {
            break;
        }
    }

    if match_count <= MAX_PDF_TEXT_FIND_MATCHES {
        for (annotation_id, page_index, contents) in ordered_freetext_contents(doc, manifest)? {
            let Some(byte_index) = contents.find(search) else {
                continue;
            };
            push_match(
                &mut matches,
                &mut match_count,
                limit,
                search,
                PdfFindHit {
                    field_id: None,
                    annotation_id: Some(annotation_id),
                    page_index: Some(page_index),
                    annotation_type: Some(FREETEXT_ANNOTATION_TYPE),
                    index_utf16: utf16_len(&contents[..byte_index])?,
                },
            )?;
            if match_count > MAX_PDF_TEXT_FIND_MATCHES {
                break;
            }
        }
    }

    if match_count > MAX_PDF_TEXT_FIND_MATCHES {
        match_count = MAX_PDF_TEXT_FIND_MATCHES;
    }
    Ok(NativeOfficeCollaborationDocumentTextFindResult {
        search: search.to_owned(),
        match_count,
        truncated: match_count > limit,
        matches,
    })
}

struct PdfFindHit {
    field_id: Option<String>,
    annotation_id: Option<String>,
    page_index: Option<u32>,
    annotation_type: Option<u32>,
    index_utf16: u32,
}

fn push_match(
    matches: &mut Vec<NativeOfficeCollaborationDocumentTextMatch>,
    match_count: &mut usize,
    limit: usize,
    search: &str,
    hit: PdfFindHit,
) -> UseResult<()> {
    *match_count += 1;
    if *match_count > MAX_PDF_TEXT_FIND_MATCHES {
        return Ok(());
    }
    if matches.len() >= limit {
        return Ok(());
    }
    let occurrence = u32::try_from(matches.len() + 1).map_err(|_| {
        collaboration_error(
            "office.collaboration.mutation_too_large",
            "PDF text find exceeded the supported occurrence range.",
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
        field_id: hit.field_id,
        annotation_id: hit.annotation_id,
        page_index: hit.page_index,
        annotation_type: hit.annotation_type,
        index_utf16: hit.index_utf16,
    });
    Ok(())
}
