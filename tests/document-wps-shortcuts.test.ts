import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { runDocumentWpsShortcut } from '../src/internal/features/work/editors/document-wps-shortcuts';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

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

  editor.commands.setTextSelection({
    from: 1,
    to: editor.state.doc.content.size,
  });
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'a', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('textStyle').textCase).toBe('all-caps');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'k', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('textStyle').textCase).toBe('small-caps');
  expect(editor.getHTML()).toContain('data-office-text-case="small-caps"');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'h', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('textStyle').hiddenText).toBe(true);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        ctrlKey: false,
        key: 'h',
        metaKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('textStyle').hiddenText).toBe(false);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'd', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('underline').underlineStyle).toBe('double');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'w', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('underline').underlineStyle).toBe('words');
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
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'g', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'p' }), callbacks),
  ).toBe(false);
  expect(calls).toEqual([
    'track',
    'comment',
    'spellcheck',
    'refresh',
    'wordCount',
  ]);
});

test('routes the unshifted font-dialog shortcut without colliding with double underline', () => {
  editor = createEditor();
  const calls: string[] = [];
  const callbacks = createCallbacks(calls);

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'd' }), callbacks),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'd', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(calls).toEqual(['fontDialog']);
  expect(editor.getAttributes('underline').underlineStyle).toBe('double');
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
    onOpenFontDialog: () => calls.push('fontDialog'),
    onOpenWordCount: () => calls.push('wordCount'),
    onRefreshFields: () => calls.push('refresh'),
    onToggleSpellcheck: () => calls.push('spellcheck'),
    onToggleTrackChanges: () => calls.push('track'),
  };
}
