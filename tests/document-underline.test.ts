import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentUnderlineStyle,
  workDocumentUnderlineStyles,
} from '../src/internal/features/work/work-document-underline';

describe('Writer underline formatting', () => {
  test('defines every native Word underline style without aliases', () => {
    expect(workDocumentUnderlineStyles).toEqual([
      'none',
      'single',
      'words',
      'double',
      'thick',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashedHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dashDotHeavy',
      'dotDotDash',
      'dashDotDotHeavy',
      'wave',
      'wavyHeavy',
      'wavyDouble',
    ]);
  });

  test('keeps native styles, color, explicit none, shortcuts, and Undo in one mark', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Underline fidelity</p>',
    });
    editor.commands.selectAll();

    expect(
      editor.commands.setDocumentUnderline('wave', { color: '#c00000' }),
    ).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('wave');
    expect(editor.getAttributes('underline')).toMatchObject({
      underlineColor: '#c00000',
      underlineStyle: 'wave',
    });
    expect(editor.getHTML()).toContain('data-office-underline-style="wave"');
    expect(editor.getHTML()).toContain('data-office-underline-color="#c00000"');
    expect(editor.getHTML()).toContain('text-decoration-style: wavy');

    expect(editor.commands.keyboardShortcut('Mod-Shift-d')).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('double');
    expect(editor.getHTML()).toContain('text-decoration-style: double');

    expect(editor.commands.keyboardShortcut('Mod-Shift-w')).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('words');
    expect(editor.commands.keyboardShortcut('Mod-Shift-w')).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('none');
    expect(editor.getHTML()).toContain('data-office-underline-style="none"');
    expect(editor.getHTML()).toContain('text-decoration-line: none');

    expect(editor.commands.undo()).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('words');
    expect(editor.commands.keyboardShortcut('Mod-u')).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('none');
    expect(editor.commands.keyboardShortcut('Mod-u')).toBe(true);
    expect(documentUnderlineStyle(editor)).toBe('single');
    editor.destroy();
  });

  test('projects native-only variants to bounded CSS while retaining metadata', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<p><u data-office-underline-style="dashDotDotHeavy" data-office-underline-color="#4472c4">Native metadata</u></p>',
    });

    expect(documentUnderlineStyle(editor)).toBe('dashDotDotHeavy');
    const html = editor.getHTML();
    expect(html).toContain('data-office-underline-style="dashDotDotHeavy"');
    expect(html).toContain('text-decoration-style: dashed');
    expect(html).toContain('text-decoration-thickness: 0.14em');
    editor.destroy();
  });
});
