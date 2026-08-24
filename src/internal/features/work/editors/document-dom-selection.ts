import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';

/**
 * Reconciles a visible browser text selection before a capture-phase command
 * snapshots editor state. Selection-change delivery can lag behind a keyboard
 * shortcut under load, leaving ProseMirror with an older collapsed selection.
 */
export function synchronizeDocumentEditorSelectionFromDom(
  editor: Editor,
): boolean {
  if (editor.isDestroyed) return false;
  const { view } = editor;
  const domSelection = view.dom.ownerDocument.getSelection();
  if (
    !domSelection ||
    domSelection.isCollapsed ||
    domSelection.rangeCount !== 1
  ) {
    return false;
  }
  const range = domSelection.getRangeAt(0);
  if (
    !view.dom.contains(range.startContainer) ||
    !view.dom.contains(range.endContainer)
  ) {
    return false;
  }

  let from: number;
  let to: number;
  try {
    from = view.posAtDOM(range.startContainer, range.startOffset, 1);
    to = view.posAtDOM(range.endContainer, range.endOffset, -1);
  } catch {
    return false;
  }
  const maximum = editor.state.doc.content.size;
  if (from < 0 || to <= from || to > maximum) return false;

  const nextSelection = TextSelection.between(
    editor.state.doc.resolve(from),
    editor.state.doc.resolve(to),
    1,
  );
  if (nextSelection.empty || nextSelection.eq(editor.state.selection)) {
    return false;
  }
  view.dispatch(
    editor.state.tr.setSelection(nextSelection).setMeta('addToHistory', false),
  );
  return true;
}
