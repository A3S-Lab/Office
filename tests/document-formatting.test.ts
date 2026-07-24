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
  documentIndentLevel,
  documentParagraphIndent,
  documentParagraphPagination,
  documentParagraphDirection,
  documentParagraphSpacing,
  DocumentParagraphFormatting,
  setDocumentIndentLevel,
  setDocumentLineHeight,
  setDocumentParagraphIndent,
  setDocumentParagraphPagination,
  setDocumentParagraphDirection,
  setDocumentParagraphSpacing,
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
    setDocumentParagraphIndent(editor, {
      left: 24,
      right: 18,
      firstLine: -12,
    });

    expect(editor.getHTML()).toContain('data-office-indent-level="1"');
    expect(editor.getHTML()).toContain('data-office-indent-right="18"');
    expect(editor.getHTML()).toContain('data-office-indent-first-line="-12"');
    expect(editor.getHTML()).toContain('text-align: justify');
    expect(editor.getHTML()).toContain('line-height: 1.5');
    expect(editor.getHTML()).toContain('margin-left: 24px');
    expect(editor.getHTML()).toContain('margin-right: 18px');
    expect(editor.getHTML()).toContain('text-indent: -12px');
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
    setDocumentParagraphIndent(editor, {
      left: 48,
      right: 24,
      firstLine: 12,
    });
    clearDocumentFormatting(editor);

    expect(editor.getHTML()).toBe('<p>A3S Office</p>');

    editor.destroy();
  });

  test('reads and directly updates all active paragraph indents', () => {
    const editor = createEditor('<p>First</p><p>Second paragraph</p>');
    editor.commands.setTextSelection(10);
    const selection = editor.state.selection.toJSON();

    expect(documentIndentLevel(editor)).toBe(0);
    expect(
      setDocumentParagraphIndent(
        editor,
        { left: 72, right: 18, firstLine: -12 },
        { restoreFocus: false },
      ),
    ).toBe(true);

    expect(documentIndentLevel(editor)).toBe(3);
    expect(documentParagraphIndent(editor)).toEqual({
      left: 72,
      right: 18,
      firstLine: -12,
    });
    expect(editor.state.selection.toJSON()).toEqual(selection);
    expect(editor.getHTML()).toContain('data-office-indent-level="3"');
    expect(editor.getHTML()).toContain('data-office-indent-right="18"');
    expect(editor.getHTML()).toContain('data-office-indent-first-line="-12"');
    expect(editor.getHTML()).toContain('Second paragraph</p>');

    setDocumentIndentLevel(editor, 2);
    changeDocumentIndent(editor, 1);
    expect(documentParagraphIndent(editor)).toEqual({
      left: 72,
      right: 18,
      firstLine: -12,
    });

    editor.destroy();
  });

  test('reads legacy CSS indents and clamps hanging indents to the page edge', () => {
    const editor = createEditor(
      '<p style="margin-left: 48px; margin-right: 12px; text-indent: -80px;">Legacy</p>',
    );

    expect(documentParagraphIndent(editor)).toEqual({
      left: 48,
      right: 12,
      firstLine: -48,
    });

    editor.destroy();
  });

  test('keeps typed paragraph pagination properties in the document model', () => {
    const editor = createEditor();

    expect(documentParagraphPagination(editor)).toEqual({
      keepLines: false,
      keepWithNext: false,
      pageBreakBefore: false,
      widowControl: true,
    });

    expect(
      setDocumentParagraphPagination(editor, {
        keepLines: true,
        keepWithNext: true,
        pageBreakBefore: true,
        widowControl: false,
      }),
    ).toBe(true);
    expect(documentParagraphPagination(editor)).toEqual({
      keepLines: true,
      keepWithNext: true,
      pageBreakBefore: true,
      widowControl: false,
    });
    expect(editor.getHTML()).toContain('data-office-keep-lines="true"');
    expect(editor.getHTML()).toContain('data-office-keep-with-next="true"');
    expect(editor.getHTML()).toContain('data-office-page-break-before="true"');
    expect(editor.getHTML()).toContain('data-office-widow-control="false"');

    clearDocumentFormatting(editor);
    expect(editor.getHTML()).toBe('<p>A3S Office</p>');

    editor.destroy();
  });

  test('keeps explicit paragraph direction in the structured document', () => {
    const editor = createEditor('<p>English</p><p dir="rtl">שלום</p>');
    editor.commands.setTextSelection(12);

    expect(documentParagraphDirection(editor)).toBe('rtl');
    expect(setDocumentParagraphDirection(editor, 'ltr')).toBe(true);
    expect(documentParagraphDirection(editor)).toBe('ltr');
    expect(editor.getHTML()).toContain('<p dir="ltr">שלום</p>');

    setDocumentParagraphDirection(editor, 'rtl');
    expect(editor.getHTML()).toContain('<p dir="rtl">שלום</p>');

    clearDocumentFormatting(editor);
    expect(editor.getHTML()).toContain('<p>שלום</p>');

    editor.destroy();
  });

  test('parses legacy CSS direction into the paragraph model', () => {
    const editor = createEditor(
      '<h2 style="direction: rtl">عنوان المستند</h2>',
    );

    expect(documentParagraphDirection(editor)).toBe('rtl');
    expect(editor.getHTML()).toContain('<h2 dir="rtl">عنوان المستند</h2>');

    editor.destroy();
  });

  test('stores list direction on the semantic list item', () => {
    const editor = createEditor('<ul><li><p>مرحبا A3S שלום</p></li></ul>');
    editor.commands.setTextSelection(4);

    expect(setDocumentParagraphDirection(editor, 'rtl')).toBe(true);
    expect(documentParagraphDirection(editor)).toBe('rtl');
    expect(editor.getHTML()).toContain(
      '<li dir="rtl"><p>مرحبا A3S שלום</p></li>',
    );
    expect(editor.getHTML()).not.toContain('<p dir="rtl">');

    clearDocumentFormatting(editor);
    expect(editor.getHTML()).toContain('<li><p>مرحبا A3S שלום</p></li>');

    editor.destroy();
  });

  test('keeps paragraph spacing and the OOXML line rule in the document model', () => {
    const editor = createEditor();

    setDocumentParagraphSpacing(editor, {
      before: 12,
      after: 6,
      lineHeight: '1.5',
      lineRule: 'auto',
    });

    expect(documentParagraphSpacing(editor)).toEqual({
      before: 12,
      after: 6,
      lineHeight: '1.5',
      lineRule: 'auto',
    });
    expect(editor.getHTML()).toContain('data-office-space-before="12"');
    expect(editor.getHTML()).toContain('data-office-space-after="6"');
    expect(editor.getHTML()).toContain('data-office-line-rule="auto"');
    expect(editor.getHTML()).toContain('margin-top: 12pt');
    expect(editor.getHTML()).toContain('margin-bottom: 6pt');
    expect(editor.getHTML()).toContain('line-height: 1.5');

    editor.destroy();
  });

  test('uses the heading style as the keep-with-next default', () => {
    const editor = createEditor('<h2>Heading</h2><p>Body</p>');
    editor.commands.setTextSelection(2);

    expect(documentParagraphPagination(editor).keepWithNext).toBe(true);
    setDocumentParagraphPagination(editor, {
      ...documentParagraphPagination(editor),
      keepWithNext: false,
    });
    expect(documentParagraphPagination(editor).keepWithNext).toBe(false);
    expect(editor.getHTML()).toContain('data-office-keep-with-next="false"');

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
