import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { describe, expect, test } from '@rstest/core';
import {
  DocumentParagraphIdentity,
  normalizeDocumentParagraphId,
  normalizeDocumentParagraphIdentity,
  type WorkDocumentParagraphIdentity,
} from '../src/internal/features/work/work-document-paragraph-identity';

describe('document paragraph identity', () => {
  test('accepts only the native positive 31-bit identifier range', () => {
    expect(normalizeDocumentParagraphId('00000001')).toBe('00000001');
    expect(normalizeDocumentParagraphId('7fffffff')).toBe('7FFFFFFF');
    expect(normalizeDocumentParagraphId('00000000')).toBeNull();
    expect(normalizeDocumentParagraphId('80000000')).toBeNull();
  });

  test('keeps the paragraph ID, preserves format-only versions, and rotates text versions', () => {
    const editor = paragraphEditor(
      '<p data-office-paragraph-id="1a2b3c4d" data-office-paragraph-text-id="1a2b3c4e">Alpha</p>',
    );
    const initial = paragraphIdentities(editor)[0];
    expect(initial).toEqual({
      paragraphId: '1A2B3C4D',
      textId: '1A2B3C4E',
    });

    editor.commands.setTextSelection({ from: 1, to: 6 });
    expect(editor.commands.toggleBold()).toBe(true);
    expect(paragraphIdentities(editor)[0]).toEqual(initial);

    expect(editor.commands.insertContentAt(3, '!')).toBe(true);
    const edited = paragraphIdentities(editor)[0];
    expect(edited?.paragraphId).toBe(initial?.paragraphId);
    expect(edited?.textId).not.toBe(initial?.textId);
    editor.destroy();
  });

  test('gives a split paragraph a fresh identity', () => {
    const editor = paragraphEditor(
      '<p data-office-paragraph-id="2A2B3C4D" data-office-paragraph-text-id="2A2B3C4E">Alpha</p>',
    );
    editor.commands.setTextSelection(3);
    expect(editor.commands.splitBlock()).toBe(true);
    const identities = paragraphIdentities(editor);
    expect(identities).toHaveLength(2);
    expect(new Set(identities.map((item) => item.paragraphId)).size).toBe(2);
    expect(identities.some((item) => item.paragraphId === '2A2B3C4D')).toBe(
      true,
    );
    editor.destroy();
  });

  test('repairs copied identities without changing ordinary paragraphs', () => {
    const identity =
      'data-office-paragraph-id="3A2B3C4D" data-office-paragraph-text-id="3A2B3C4E"';
    const editor = paragraphEditor(
      `<p ${identity}>First</p><p ${identity}>Copied</p><p data-office-paragraph-id="invalid">Malformed</p>`,
    );
    const identities = paragraphIdentities(editor);
    expect(identities).toHaveLength(2);
    expect(new Set(identities.map((item) => item.paragraphId)).size).toBe(2);
    expect(identities[0]).toEqual({
      paragraphId: '3A2B3C4D',
      textId: '3A2B3C4E',
    });
    expect(editor.getHTML()).toContain('<p>Malformed</p>');
    editor.destroy();
  });
});

function paragraphEditor(content: string): Editor {
  return new Editor({
    extensions: [StarterKit, DocumentParagraphIdentity],
    content,
  });
}

function paragraphIdentities(editor: Editor): WorkDocumentParagraphIdentity[] {
  const identities: WorkDocumentParagraphIdentity[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return;
    const identity = normalizeDocumentParagraphIdentity({
      paragraphId: node.attrs.paragraphId,
      textId: node.attrs.textId,
    });
    if (identity) identities.push(identity);
  });
  return identities;
}
