import { Editor } from '@tiptap/core';
import FontFamily from '@tiptap/extension-text-style/font-family';
import FontSize from '@tiptap/extension-text-style/font-size';
import { TextStyle } from '@tiptap/extension-text-style';
import TextAlign from '@tiptap/extension-text-align';
import StarterKit from '@tiptap/starter-kit';
import { describe, expect, test } from '@rstest/core';
import { DocumentSection } from '../src/internal/features/work/work-document-section-node';
import {
  changeDocumentIndent,
  clearDocumentFormatting,
  DocumentParagraphFormatting,
  setDocumentLineHeight,
} from '../src/internal/features/work/work-document-paragraph-formatting';

function createEditor(content = '<p>A3S Office</p>'): Editor {
  return new Editor({
    extensions: [
      StarterKit,
      DocumentSection,
      TextStyle,
      FontFamily,
      FontSize,
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      DocumentParagraphFormatting,
    ],
    content,
  });
}

describe('document formatting', () => {
  test('keeps character and paragraph formatting in the TipTap document', () => {
    const editor = createEditor();

    editor
      .chain()
      .setTextSelection({ from: 1, to: 11 })
      .setFontFamily('Arial')
      .setFontSize('12pt')
      .setTextAlign('justify')
      .run();
    setDocumentLineHeight(editor, '1.5');
    changeDocumentIndent(editor, 1);

    expect(editor.getHTML()).toContain(
      'style="text-align: justify; line-height: 1.5; margin-left: 24px;"',
    );
    expect(editor.getHTML()).toContain(
      'style="font-family: Arial; font-size: 12pt;"',
    );

    editor.destroy();
  });

  test('clears direct formatting without leaving empty style spans', () => {
    const editor = createEditor();

    editor
      .chain()
      .setTextSelection({ from: 1, to: 11 })
      .setFontFamily('Arial')
      .setFontSize('18pt')
      .setTextAlign('right')
      .run();
    setDocumentLineHeight(editor, '2');
    changeDocumentIndent(editor, 2);
    clearDocumentFormatting(editor);

    expect(editor.getHTML()).toBe('<p>A3S Office</p>');

    editor.destroy();
  });

  test('preserves document sections when direct formatting is cleared', () => {
    const editor = createEditor(
      [
        '<section data-document-section="true" data-section-id="section-alpha">',
        '<h2 style="text-align: right; line-height: 2; margin-left: 48px;"',
        ' data-office-indent-level="2">',
        '<span style="font-family: Arial; font-size: 18pt;">A3S Office</span>',
        '</h2>',
        '</section>',
      ].join(''),
    );

    editor.commands.selectAll();
    clearDocumentFormatting(editor);

    expect(editor.getHTML()).toContain(
      '<section data-document-section="true" data-section-id="section-alpha"',
    );
    expect(editor.getHTML()).toContain('<p>A3S Office</p>');

    editor.destroy();
  });
});
