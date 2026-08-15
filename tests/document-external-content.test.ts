import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { afterEach, expect, test } from '@rstest/core';
import {
  applyExternalDocumentContent,
  shouldPublishDocumentUpdate,
} from '../src/internal/features/work/editors/document-external-content';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('applies a controlled document extension without replacing the editor state', () => {
  let updateCount = 0;
  editor = new Editor({
    content: '<p>Alpha</p><p>Stable paragraph</p>',
    extensions: [StarterKit],
    onUpdate: () => {
      updateCount += 1;
    },
  });
  editor.commands.setTextSelection(2);
  const beforeDocument = editor.state.doc;

  expect(
    applyExternalDocumentContent(
      editor,
      '<p>Alpha beta</p><p>Stable paragraph</p>',
    ),
  ).toBe('applied');

  expect(editor.getText()).toBe('Alpha beta\n\nStable paragraph');
  expect(editor.state.doc).not.toBe(beforeDocument);
  expect(editor.state.selection.from).toBe(2);
  expect(updateCount).toBe(0);
});

test('does not dispatch a transaction for equivalent controlled content', () => {
  editor = new Editor({
    content: '<p>Already current</p>',
    extensions: [StarterKit],
  });
  const beforeDocument = editor.state.doc;

  expect(applyExternalDocumentContent(editor, '<p>Already current</p>')).toBe(
    'unchanged',
  );
  expect(editor.state.doc).toBe(beforeDocument);
});

test('replaces a longer controlled document with an exact shorter snapshot', () => {
  let updateCount = 0;
  editor = new Editor({
    content:
      '<h1>Software Development Contract</h1><h2>Parties</h2><p>The Client appoints the Developer.</p><h2>Scope</h2><p>The Developer will deliver the agreed software milestones.</p><h2>Acceptance</h2><p>Acceptance requires written confirmation within five business days.</p>',
    extensions: [StarterKit],
    onUpdate: () => {
      updateCount += 1;
    },
  });

  expect(
    applyExternalDocumentContent(
      editor,
      '<h1>Software Development Contract</h1><h2>Parties</h2><p>The Client appoints the Developer.</p>',
    ),
  ).toBe('applied');

  expect(editor.getHTML()).toBe(
    '<h1>Software Development Contract</h1><h2>Parties</h2><p>The Client appoints the Developer.</p>',
  );
  expect(updateCount).toBe(0);
});

test('publishes only undoable user document transactions', () => {
  editor = new Editor({
    content: '<p>Current document</p>',
    extensions: [StarterKit],
  });

  expect(
    shouldPublishDocumentUpdate(editor.state.tr.setMeta('preventUpdate', true)),
  ).toBe(false);
  expect(
    shouldPublishDocumentUpdate(editor.state.tr.setMeta('addToHistory', false)),
  ).toBe(false);
  expect(shouldPublishDocumentUpdate(editor.state.tr.insertText('A'))).toBe(
    true,
  );
});
