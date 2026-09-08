import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { documentFieldUnlinkRanges } from '../src/internal/features/work/work-document-field-node';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('unlinks a selected WPS field to its result text with one undo step', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-document-field="true" data-field-id="field-page" data-field-kind="page" data-field-instruction="PAGE" data-field-display="7" data-field-code="{ PAGE }">7</span></p>',
  });

  const fieldPos = fieldPosition(editor);
  expect(String(editor.state.doc.nodeAt(fieldPos)?.attrs.display ?? '')).toBe(
    '7',
  );

  editor.view.dispatch(
    editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, fieldPos)),
  );
  expect(documentFieldUnlinkRanges(editor.state)).toEqual([
    { from: fieldPos, to: fieldPos + 1, text: '7' },
  ]);

  expect(editor.commands.unlinkDocumentFields()).toBe(true);
  expect(editor.getHTML()).toContain('7');
  expect(editor.getHTML()).not.toContain('data-document-field');
  expect(editor.commands.undo()).toBe(true);
  expect(editor.getHTML()).toContain('data-document-field');
  expect(editor.getHTML()).toContain('data-field-display="7"');
});

test('unlinks every field covered by a text selection', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-document-field="true" data-field-id="field-a" data-field-kind="page" data-field-instruction="PAGE" data-field-display="1">1</span> and <span data-document-field="true" data-field-id="field-b" data-field-kind="numPages" data-field-instruction="NUMPAGES" data-field-display="3">3</span></p>',
  });

  editor.commands.selectAll();
  expect(documentFieldUnlinkRanges(editor.state).length).toBe(2);
  expect(editor.commands.unlinkDocumentFields()).toBe(true);
  expect(editor.getText()).toContain('1 and 3');
  expect(editor.getHTML()).not.toContain('data-document-field');
});

test('refuses to unlink when the selection does not cover a field', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Plain text</p>',
  });
  editor.commands.setTextSelection({ from: 1, to: 5 });
  expect(documentFieldUnlinkRanges(editor.state)).toEqual([]);
  expect(editor.commands.unlinkDocumentFields()).toBe(false);
});

function fieldPosition(current: Editor): number {
  let position = -1;
  current.state.doc.descendants((node, pos) => {
    if (node.type.name === 'documentField') {
      position = pos;
      return false;
    }
    return true;
  });
  if (position < 0) throw new Error('Expected a document field node.');
  return position;
}
