import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { cycleDocumentSelectionCaseText } from '../work-document-text-case';

export function cycleDocumentSelectionCase(editor: Editor): boolean {
  if (editor.isDestroyed) return false;
  const range = selectionRangeForCaseCycle(editor);
  if (!range) return false;
  const text = editor.state.doc.textBetween(range.from, range.to, '\n');
  if (!text) return false;
  const next = cycleDocumentSelectionCaseText(text);
  if (next === text) return false;
  const { tr } = editor.state;
  tr.insertText(next, range.from, range.to);
  tr.setSelection(
    TextSelection.create(tr.doc, range.from, range.from + next.length),
  );
  editor.view.dispatch(tr.scrollIntoView());
  return true;
}

function selectionRangeForCaseCycle(
  editor: Editor,
): { from: number; to: number } | null {
  const { from, to, empty } = editor.state.selection;
  if (!empty) return { from, to };

  const $pos = editor.state.selection.$from;
  if (!$pos.parent.isTextblock) return null;
  const blockStart = $pos.start();
  const text = $pos.parent.textBetween(
    0,
    $pos.parent.content.size,
    undefined,
    '\0',
  );
  if (!text) return null;
  const offset = $pos.parentOffset;
  let left = offset;
  let right = offset;
  while (left > 0 && isWordCharacter(text.charAt(left - 1))) left -= 1;
  while (right < text.length && isWordCharacter(text.charAt(right))) right += 1;
  if (left === right) return null;
  return { from: blockStart + left, to: blockStart + right };
}

function isWordCharacter(character: string): boolean {
  return /\p{L}|\p{N}/u.test(character);
}
