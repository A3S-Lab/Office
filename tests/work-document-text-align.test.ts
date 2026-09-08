import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { DocumentTextAlign } from '../src/internal/features/work/work-document-text-align';

describe('DocumentTextAlign distribute', () => {
  test('persists distribute via data attribute and valid justify CSS', () => {
    const editor = new Editor({
      extensions: [
        StarterKit,
        DocumentTextAlign.configure({
          types: ['paragraph'],
          alignments: ['left', 'center', 'right', 'justify', 'distribute'],
        }),
      ],
      content: '<p>Distribute me</p>',
    });

    expect(editor.chain().focus().setTextAlign('distribute').run()).toBe(true);
    expect(editor.getAttributes('paragraph').textAlign).toBe('distribute');
    expect(editor.getHTML()).toContain('data-office-text-align="distribute"');
    expect(editor.getHTML()).toContain('text-align: justify');
    expect(editor.getHTML()).toContain('text-align-last: justify');
    expect(editor.isActive({ textAlign: 'distribute' })).toBe(true);

    editor.destroy();
  });
});
