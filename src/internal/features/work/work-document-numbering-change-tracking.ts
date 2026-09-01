import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import { ReplaceAroundStep } from '@tiptap/pm/transform';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import { serializeDocumentNumberingChange } from './work-document-numbering-changes';

interface DocumentNumberingChangeTrackingOptions {
  createChange: () => WorkDocumentChangeIdentity;
}

export function trackDocumentNumberingChangeTransaction(
  transaction: Transaction,
  state: EditorState,
  options: DocumentNumberingChangeTrackingOptions,
): boolean {
  if (!transaction.steps.some((step) => step instanceof ReplaceAroundStep)) {
    return false;
  }
  let identity: WorkDocumentChangeIdentity | null = null;
  let tracked = false;
  state.doc.descendants((before, position) => {
    if (before.type.name !== 'orderedList') return;
    const after = transaction.doc.nodeAt(position);
    if (
      !after ||
      after.type !== before.type ||
      !before.content.eq(after.content)
    ) {
      return;
    }
    const previous = serializeDocumentNumberingChange(before.attrs);
    const current = serializeDocumentNumberingChange(after.attrs);
    if (previous === current || hasNumberingChange(after)) return;
    identity ??= options.createChange();
    transaction.setNodeMarkup(position, undefined, {
      ...after.attrs,
      numberingChangeKind: 'numbering',
      numberingChangeId: identity.id,
      numberingChangeActorId: identity.actorId ?? '',
      numberingChangeAuthor: identity.author || 'A3S Work',
      numberingChangeDate: identity.date || new Date().toISOString(),
      numberingChangeBefore: previous,
    });
    tracked = true;
  });
  return tracked;
}

function hasNumberingChange(node: ProseMirrorNode): boolean {
  return node.attrs.numberingChangeKind === 'numbering';
}
