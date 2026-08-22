import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentStrikeStyle,
  workDocumentStrikeStyles,
} from '../src/internal/features/work/work-document-strike';
import { resolvedDocxStrikeFormatting } from '../src/internal/features/work/work-docx-strike';

describe('Writer strikethrough formatting', () => {
  test('models the complete native single and double strike family', () => {
    expect(workDocumentStrikeStyles).toEqual(['none', 'single', 'double']);
    expect(resolvedDocxStrikeFormatting(true, undefined)).toEqual({
      style: 'single',
    });
    expect(resolvedDocxStrikeFormatting(true, true)).toEqual({
      style: 'double',
    });
    expect(resolvedDocxStrikeFormatting(true, false)).toEqual({
      style: 'single',
    });
    expect(resolvedDocxStrikeFormatting(false, true)).toEqual({
      style: 'double',
    });
    expect(resolvedDocxStrikeFormatting(false, false)).toEqual({
      style: 'none',
    });
    expect(resolvedDocxStrikeFormatting(undefined, undefined)).toBeUndefined();
    expect(resolvedDocxStrikeFormatting(undefined, undefined, true)).toEqual({
      style: 'none',
    });
  });

  test('keeps native style, explicit none, Undo, and shortcut ownership in one mark', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Strike fidelity</p>',
    });
    editor.commands.selectAll();

    expect(editor.commands.setDocumentStrike('double')).toBe(true);
    expect(documentStrikeStyle(editor)).toBe('double');
    expect(editor.getAttributes('strike')).toMatchObject({
      strikeStyle: 'double',
    });
    expect(editor.getHTML()).toContain('data-office-strike-style="double"');
    expect(editor.getHTML()).toContain('text-decoration-style: double');

    expect(editor.commands.toggleStrike()).toBe(true);
    expect(documentStrikeStyle(editor)).toBe('none');
    expect(editor.getHTML()).toContain('data-office-strike-style="none"');
    expect(editor.getHTML()).toContain('text-decoration-line: none');

    expect(editor.commands.undo()).toBe(true);
    expect(documentStrikeStyle(editor)).toBe('double');
    editor.commands.keyboardShortcut('Mod-Shift-s');
    expect(documentStrikeStyle(editor)).toBe('double');

    expect(editor.commands.toggleDocumentStrikeStyle('single')).toBe(true);
    expect(documentStrikeStyle(editor)).toBe('single');
    expect(editor.commands.toggleDocumentStrikeStyle('single')).toBe(true);
    expect(documentStrikeStyle(editor)).toBe('none');
    editor.destroy();
  });

  test('projects native double strike to CSS without flattening metadata', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<p><s data-office-strike-style="double">Native double strike</s></p>',
    });

    expect(documentStrikeStyle(editor)).toBe('double');
    expect(editor.getHTML()).toContain('data-office-strike-style="double"');
    expect(editor.getHTML()).toContain('text-decoration-style: double');
    editor.destroy();
  });
});
