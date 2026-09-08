import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { moveDocumentBlock } from '../src/internal/features/work/editors/document-move-block';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('moves the active paragraph among siblings with Alt+Shift arrows', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>First</p><p>Second</p><p>Third</p>',
  });

  let secondPos = 1;
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text === 'Second') {
      secondPos = pos;
      return false;
    }
    return undefined;
  });
  editor.commands.setTextSelection(secondPos);
  expect(moveDocumentBlock(editor, -1)).toBe(true);
  expect(paragraphTexts(editor)).toEqual(['Second', 'First', 'Third']);

  expect(moveDocumentBlock(editor, 1)).toBe(true);
  expect(paragraphTexts(editor)).toEqual(['First', 'Second', 'Third']);

  expect(moveDocumentBlock(editor, 1)).toBe(true);
  expect(paragraphTexts(editor)).toEqual(['First', 'Third', 'Second']);

  expect(moveDocumentBlock(editor, 1)).toBe(false);
  expect(paragraphTexts(editor)).toEqual(['First', 'Third', 'Second']);
});

function paragraphTexts(current: Editor): string[] {
  const texts: string[] = [];
  current.state.doc.descendants((node) => {
    if (node.type.name === 'paragraph') texts.push(node.textContent);
  });
  return texts;
}
