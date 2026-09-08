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
      shortcut({ key: 'F9', ctrlKey: false, altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F9', ctrlKey: false, shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F9', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'F11' }), callbacks),
  ).toBe(true);
  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F11', shiftKey: true }),
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
    'fieldCodes',
    'selectedFieldCodes',
    'unlink',
    'lock',
    'unlock',
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

test('clears selection formatting with WPS Ctrl+Space', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  editor
    .chain()
    .focus()
    .setTextSelection({ from: 1, to: 14 })
    .toggleBold()
    .toggleItalic()
    .setTextAlign('center')
    .run();
  expect(editor.isActive('bold')).toBe(true);
  expect(editor.isActive('italic')).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('center');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: ' ', code: 'Space' }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('bold')).toBe(false);
  expect(editor.isActive('italic')).toBe(false);
  expect(editor.getAttributes('paragraph').textAlign ?? null).toBeNull();
});

test('steps paragraph indent with WPS Ctrl+M / Ctrl+Shift+M', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'm' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').indentLevel).toBe(1);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'm', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').indentLevel ?? 0).toBe(0);
});

test('steps paragraph indent with WPS Shift+Alt+. / Shift+Alt+,', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        key: '.',
        code: 'Period',
        ctrlKey: false,
        altKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').indentLevel).toBe(1);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        key: ',',
        code: 'Comma',
        ctrlKey: false,
        altKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').indentLevel ?? 0).toBe(0);
});

test('steps paragraph indent with WPS Alt+Shift+ArrowRight / ArrowLeft', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        key: 'ArrowRight',
        code: 'ArrowRight',
        ctrlKey: false,
        altKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').indentLevel).toBe(1);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        key: 'ArrowLeft',
        code: 'ArrowLeft',
        ctrlKey: false,
        altKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').indentLevel ?? 0).toBe(0);
});

test('applies WPS Writer paragraph align and line-spacing shortcuts', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'l' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('left');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'e' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('center');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'r' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('right');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: 'j' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('justify');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'j', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').textAlign).toBe('distribute');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: '5' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').lineHeight).toBe('1.5');

  expect(
    runDocumentWpsShortcut(editor, shortcut({ key: '2' }), callbacks),
  ).toBe(true);
  expect(editor.getAttributes('paragraph').lineHeight).toBe('2');
});

test('applies WPS Writer heading shortcuts Ctrl+Alt+1 / 2 / 3', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: '1', code: 'Digit1', altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('heading', { level: 1 })).toBe(true);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: '2', code: 'Digit2', altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('heading', { level: 2 })).toBe(true);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: '3', code: 'Digit3', altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('heading', { level: 3 })).toBe(true);
});

test('applies the Normal / 正文 style with WPS Ctrl+Shift+N', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: '2', code: 'Digit2', altKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('heading', { level: 2 })).toBe(true);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'n', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('paragraph')).toBe(true);
  expect(editor.isActive('heading')).toBe(false);
});

test('toggles the default bullet list with WPS Ctrl+Shift+L', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'l', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('bulletList')).toBe(true);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'l', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.isActive('bulletList')).toBe(false);
});

test('moves paragraphs with WPS Alt+Shift+ArrowUp / ArrowDown', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Alpha</p><p>Beta</p><p>Gamma</p>',
  });
  const callbacks = createCallbacks();
  let betaPos = 1;
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text === 'Beta') {
      betaPos = pos;
      return false;
    }
    return undefined;
  });
  editor.commands.setTextSelection(betaPos);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        key: 'ArrowUp',
        code: 'ArrowUp',
        ctrlKey: false,
        altKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(paragraphTexts(editor)).toEqual(['Beta', 'Alpha', 'Gamma']);

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({
        key: 'ArrowDown',
        code: 'ArrowDown',
        ctrlKey: false,
        altKey: true,
        shiftKey: true,
      }),
      callbacks,
    ),
  ).toBe(true);
  expect(paragraphTexts(editor)).toEqual(['Alpha', 'Beta', 'Gamma']);
});

function paragraphTexts(current: Editor): string[] {
  const texts: string[] = [];
  current.state.doc.descendants((node) => {
    if (node.type.name === 'paragraph') texts.push(node.textContent);
  });
  return texts;
}

test('toggles double strikethrough with WPS Ctrl+Shift+X', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'x', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('strike').strikeStyle).toBe('double');
  expect(editor.getHTML()).toContain('data-office-strike-style="double"');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'x', shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getAttributes('strike').strikeStyle).toBe('none');
});

test('cycles selected text case with WPS Shift+F3', () => {
  editor = createEditor();
  const callbacks = createCallbacks();

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F3', ctrlKey: false, shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getText()).toBe('SHORTCUT TEXT');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F3', ctrlKey: false, shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getText()).toBe('Shortcut Text');

  expect(
    runDocumentWpsShortcut(
      editor,
      shortcut({ key: 'F3', ctrlKey: false, shiftKey: true }),
      callbacks,
    ),
  ).toBe(true);
  expect(editor.getText()).toBe('shortcut text');
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
    onToggleFieldCodes: () => calls.push('fieldCodes'),
    onToggleSelectedFieldCodes: () => {
      calls.push('selectedFieldCodes');
      return true;
    },
    onUnlinkFields: () => {
      calls.push('unlink');
      return true;
    },
    onLockFields: () => {
      calls.push('lock');
      return true;
    },
    onUnlockFields: () => {
      calls.push('unlock');
      return true;
    },
    onToggleSpellcheck: () => calls.push('spellcheck'),
    onToggleTrackChanges: () => calls.push('track'),
  };
}
