import type { Editor } from '@tiptap/core';
import { Fragment, type Node as ProseMirrorNode } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';

const MOVE_CONTAINER_TYPES = new Set([
  'doc',
  'documentSection',
  'blockquote',
  'listItem',
]);

/**
 * WPS/Word Alt+Shift+↑ / Alt+Shift+↓: swap the current block with the
 * adjacent sibling inside the nearest movable container.
 */
export function moveDocumentBlock(
  editor: Editor,
  direction: -1 | 1,
): boolean {
  if (editor.isDestroyed || !editor.view) return false;

  const { state } = editor;
  const { $from } = state.selection;

  let depth = $from.depth;
  while (depth > 0) {
    const node = $from.node(depth);
    const parent = $from.node(depth - 1);
    if (node.isBlock && MOVE_CONTAINER_TYPES.has(parent.type.name)) break;
    depth -= 1;
  }
  if (depth <= 0) return false;

  const parentDepth = depth - 1;
  const index = $from.index(parentDepth);
  const parent = $from.node(parentDepth);
  const targetIndex = index + direction;
  if (targetIndex < 0 || targetIndex >= parent.childCount) return false;

  const children: ProseMirrorNode[] = [];
  for (let childIndex = 0; childIndex < parent.childCount; childIndex += 1) {
    children.push(parent.child(childIndex));
  }
  const [moved] = children.splice(index, 1);
  if (!moved) return false;
  children.splice(targetIndex, 0, moved);

  const contentStart = $from.start(parentDepth);
  const contentEnd = $from.end(parentDepth);
  let tr = state.tr.replaceWith(
    contentStart,
    contentEnd,
    Fragment.from(children),
  );

  let offset = contentStart;
  for (let childIndex = 0; childIndex < targetIndex; childIndex += 1) {
    offset += children[childIndex]!.nodeSize;
  }
  const caret = Math.min(offset + 1, offset + moved.nodeSize - 1);
  tr = tr
    .setSelection(TextSelection.near(tr.doc.resolve(caret)))
    .scrollIntoView();
  editor.view.dispatch(tr);
  return true;
}
