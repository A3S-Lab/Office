import { createNodeFromContent, type Content, type Editor } from '@tiptap/core';
import { type Fragment, Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { Transaction } from '@tiptap/pm/state';

export type ExternalDocumentContentResult =
  | 'applied'
  | 'unchanged'
  | 'unsupported';

/**
 * Returns whether an editor transaction represents a user-authored document
 * change that the controlled host must persist.
 */
export function shouldPublishDocumentUpdate(transaction: Transaction): boolean {
  return (
    !transaction.getMeta('preventUpdate') &&
    transaction.getMeta('addToHistory') !== false
  );
}

/**
 * Applies controlled host content as the smallest ProseMirror transaction.
 * This keeps selections, editor instances, and local history stable while an
 * agent or another realtime producer extends the document.
 */
export function applyExternalDocumentContent(
  editor: Editor,
  content: Content,
): ExternalDocumentContentResult {
  let nextDocument: ProseMirrorNode | Fragment;
  try {
    nextDocument = createNodeFromContent(content, editor.schema, {
      errorOnInvalidContent: true,
      slice: false,
    });
  } catch {
    return 'unsupported';
  }
  if (!(nextDocument instanceof ProseMirrorNode)) return 'unsupported';

  const currentDocument = editor.state.doc;
  const start = currentDocument.content.findDiffStart(nextDocument.content);
  if (start === null) return 'unchanged';
  const end = currentDocument.content.findDiffEnd(nextDocument.content);
  let currentEnd = end?.a ?? currentDocument.content.size;
  let nextEnd = end?.b ?? nextDocument.content.size;

  // Tree diff boundaries may cross for a pure insertion or deletion at a
  // block edge. Node.slice() requires an ordered range, so advance both end
  // positions by the overlap while preserving the exact replacement delta.
  const overlap = start - Math.min(currentEnd, nextEnd);
  if (overlap > 0) {
    currentEnd += overlap;
    nextEnd += overlap;
  }

  try {
    const transaction = editor.state.tr
      // Controlled content is an exact host snapshot. `replaceRange` may
      // broaden a cross-block deletion to preserve an editing heuristic,
      // which can become pathologically expensive when a later snapshot is
      // shorter. Apply the already schema-valid diff range exactly instead.
      .replace(start, currentEnd, nextDocument.slice(start, nextEnd))
      .setMeta('addToHistory', false)
      .setMeta('preventUpdate', true);
    editor.view.dispatch(transaction);
    return 'applied';
  } catch {
    return 'unsupported';
  }
}
