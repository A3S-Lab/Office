import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { activeDocumentSection } from '../src/internal/features/work/work-document-section-editor';

describe('document section commands', () => {
  test('inserts, updates, and merges sections through TipTap commands', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>First</p><p>Second</p>',
        '</section>',
      ].join(''),
    });
    editor.commands.setTextSelection(9);

    expect(editor.commands.insertDocumentSection('continuous')).toBe(true);
    expect(editor.state.doc.childCount).toBe(2);
    expect(activeDocumentSection(editor)?.index).toBe(1);

    const active = activeDocumentSection(editor);
    expect(active).not.toBeNull();
    if (!active) throw new Error('Expected an active document section.');
    expect(
      editor.commands.updateActiveDocumentSection({
        ...active.layout,
        orientation: 'landscape',
      }),
    ).toBe(true);
    expect(activeDocumentSection(editor)?.layout.orientation).toBe('landscape');

    expect(editor.commands.mergeDocumentSectionWithPrevious()).toBe(true);
    expect(editor.state.doc.childCount).toBe(1);
    expect(editor.getText()).toContain('First');
    expect(editor.getText()).toContain('Second');

    editor.destroy();
  });
});
