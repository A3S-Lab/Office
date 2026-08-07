import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { runDocumentWpsShortcut } from '../src/internal/features/work/editors/document-wps-shortcuts';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('executes WPS Writer font, paragraph, and style shortcuts', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'e' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('center');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: '>', code: 'Period', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('textStyle').fontSize).toBe('12pt');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: '[' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('textStyle').fontSize).toBe('10.5pt');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: '1' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').lineHeight).toBe('1');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: '2', code: 'Digit2', altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('heading', { level: 2 })).toBe(true);
});

test('executes scoped WPS Writer review shortcuts and ignores unrelated keys', () => {
  editor = createEditor();
  const calls: string[] = [];
  const callbacks = createCallbacks(calls);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'e', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'm', altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F7', ctrlKey: false }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F9', ctrlKey: false }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'p' }), callbacks),
  ).toBe(false);
  expect(calls).toEqual(['track', 'comment', 'spellcheck', 'refresh']);
});

function createEditor(): Editor {
  const currentEditor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Shortcut text</p>',
  });
  currentEditor.commands.setTextSelection({ from: 1, to: 14 });
  return currentEditor;
}

function shortcut(
  overrides: Partial<
    Pick<
      KeyboardEvent,
      'altKey' | 'code' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'
    >
  >,
) {
  return {
    altKey: false,
    code: '',
    ctrlKey: true,
    key: '',
    metaKey: false,
    shiftKey: false,
    ...overrides,
  };
}

function createCallbacks(calls: string[] = []) {
  return {
    canInsertComment: true,
    canRefreshFields: true,
    onInsertComment: () => calls.push('comment'),
    onRefreshFields: () => calls.push('refresh'),
    onToggleSpellcheck: () => calls.push('spellcheck'),
    onToggleTrackChanges: () => calls.push('track'),
  };
}
