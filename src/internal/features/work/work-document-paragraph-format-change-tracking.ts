import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import { ReplaceAroundStep } from '@tiptap/pm/transform';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import { serializeDocumentParagraphFormatting } from './work-document-paragraph-format-changes';

interface DocumentParagraphFormattingTrackingOptions {
  createChange: () => WorkDocumentChangeIdentity;
}

export function trackDocumentParagraphFormattingTransaction(
  transaction: Transaction,
  state: EditorState,
  options: DocumentParagraphFormattingTrackingOptions,
): boolean {
  if (!transaction.steps.some((step) => step instanceof ReplaceAroundStep)) {
    return false;
  }
  let identity: WorkDocumentChangeIdentity | null = null;
  let tracked = false;
  state.doc.descendants((before, position) => {
    if (!isParagraph(before)) return;
    const after = transaction.doc.nodeAt(position);
    if (!after || after.type !== before.type || !isParagraph(after)) return;
    const previous = serializeDocumentParagraphFormatting(before.attrs);
    const current = serializeDocumentParagraphFormatting(after.attrs);
    if (previous === current || !sameContentIgnoringMarks(before, after)) {
      return;
    }
    if (after.attrs.paragraphChangeKind === 'paragraph-formatting') return;
    identity ??= options.createChange();
    transaction.setNodeMarkup(position, undefined, {
      ...after.attrs,
      paragraphChangeKind: 'paragraph-formatting',
      paragraphChangeId: identity.id,
      paragraphChangeActorId: identity.actorId ?? '',
      paragraphChangeAuthor: identity.author || 'A3S Work',
      paragraphChangeDate: identity.date || new Date().toISOString(),
      paragraphChangeBefore: previous,
    });
    tracked = true;
  });
  return tracked;
}

function isParagraph(node: ProseMirrorNode): boolean {
  return node.type.name === 'paragraph' || node.type.name === 'heading';
}

function sameContentIgnoringMarks(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  return (
    JSON.stringify(childContentIgnoringMarks(before)) ===
    JSON.stringify(childContentIgnoringMarks(after))
  );
}

function childContentIgnoringMarks(node: ProseMirrorNode): unknown[] {
  const content: unknown[] = [];
  node.forEach((child) => {
    content.push(contentIgnoringMarks(child));
  });
  return content;
}

function contentIgnoringMarks(node: ProseMirrorNode): unknown {
  if (node.isText) return { type: node.type.name, text: node.text ?? '' };
  const content = childContentIgnoringMarks(node);
  return {
    type: node.type.name,
    ...(Object.keys(node.attrs).length ? { attrs: node.attrs } : {}),
    ...(content.length ? { content } : {}),
  };
}
