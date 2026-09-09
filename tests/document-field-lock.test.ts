import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import { documentFieldLockTargets } from '../src/internal/features/work/work-document-field-node';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('locks selected WPS fields against F9 refresh', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-document-field="true" data-field-id="field-page" data-field-kind="page" data-field-instruction="PAGE" data-field-display="1">1</span></p>',
  });

  const fieldPos = fieldPosition(editor);
  editor.view.dispatch(
    editor.state.tr.setSelection(
      NodeSelection.create(editor.state.doc, fieldPos),
    ),
  );
  expect(documentFieldLockTargets(editor.state)).toEqual([
    { from: fieldPos, to: fieldPos + 1, locked: false },
  ]);

  expect(editor.commands.setDocumentFieldsLocked(true)).toBe(true);
  expect(editor.getHTML()).toContain('data-field-locked="true"');
  const lockedPos = fieldPosition(editor);
  expect(editor.state.doc.nodeAt(lockedPos)?.attrs.locked).toBe(true);

  expect(
    editor.commands.refreshDocumentFields(documentContent(editor), {
      resolveContext: () => ({
        pageNumber: 9,
        totalPages: 9,
        sectionNumber: 1,
        sectionPages: 9,
      }),
    }),
  ).toBe(false);
  expect(String(editor.state.doc.nodeAt(lockedPos)?.attrs.display)).toBe('1');

  editor.view.dispatch(
    editor.state.tr.setSelection(
      NodeSelection.create(editor.state.doc, lockedPos),
    ),
  );
  expect(editor.commands.setDocumentFieldsLocked(false)).toBe(true);
  expect(editor.getHTML()).not.toContain('data-field-locked');
  expect(
    editor.commands.refreshDocumentFields(documentContent(editor), {
      resolveContext: () => ({
        pageNumber: 9,
        totalPages: 9,
        sectionNumber: 1,
        sectionPages: 9,
      }),
    }),
  ).toBe(true);
  expect(String(editor.state.doc.nodeAt(lockedPos)?.attrs.display)).toBe('9');
});

test('refuses to lock when the selection does not cover a field', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Plain text</p>',
  });
  editor.commands.setTextSelection({ from: 1, to: 5 });
  expect(documentFieldLockTargets(editor.state)).toEqual([]);
  expect(editor.commands.setDocumentFieldsLocked(true)).toBe(false);
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

function documentContent(current: Editor): WorkDocumentContent {
  return {
    type: 'document',
    title: 'Lock fields',
    html: current.getHTML(),
    sections: [],
  };
}
