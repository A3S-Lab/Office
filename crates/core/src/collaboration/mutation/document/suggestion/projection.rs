use a3s_use_core::UseResult;
use yrs::ReadTxn;

use super::collect_suggestions;
use crate::collaboration::{
    NativeOfficeCollaborationDocumentSuggestion,
    NativeOfficeCollaborationDocumentSuggestionPlacement,
};

pub(in crate::collaboration) fn project_document_suggestions<T: ReadTxn>(
    transaction: &T,
    fragment: &yrs::XmlFragmentRef,
) -> UseResult<Vec<NativeOfficeCollaborationDocumentSuggestion>> {
    let mut suggestions = collect_suggestions(transaction, fragment)?
        .into_values()
        .collect::<Vec<_>>();
    suggestions.sort_by(|left, right| {
        let left_position = left
            .placements
            .first()
            .map(|placement| (placement.text_node, placement.start_utf16));
        let right_position = right
            .placements
            .first()
            .map(|placement| (placement.text_node, placement.start_utf16));
        left_position
            .cmp(&right_position)
            .then_with(|| {
                left.identity
                    .kind
                    .as_str()
                    .cmp(right.identity.kind.as_str())
            })
            .then_with(|| left.identity.id.cmp(&right.identity.id))
    });

    Ok(suggestions
        .into_iter()
        .map(|suggestion| NativeOfficeCollaborationDocumentSuggestion {
            id: suggestion.identity.id,
            kind: suggestion.identity.kind,
            actor_id: suggestion.identity.actor_id,
            author: suggestion.identity.author,
            created_at: suggestion.identity.date,
            text: suggestion.text,
            placements: suggestion
                .placements
                .into_iter()
                .map(
                    |placement| NativeOfficeCollaborationDocumentSuggestionPlacement {
                        paragraph_id: placement.paragraph_id,
                        text_id: placement.text_id,
                        start_utf16: placement.start_utf16,
                        end_utf16: placement.end_utf16,
                        text: placement.text,
                    },
                )
                .collect(),
        })
        .collect())
}
