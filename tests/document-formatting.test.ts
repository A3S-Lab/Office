import { Editor } from '@tiptap/core';
import FontSize from '@tiptap/extension-text-style/font-size';
import TextAlign from '@tiptap/extension-text-align';
import StarterKit from '@tiptap/starter-kit';
import { describe, expect, test } from '@rstest/core';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { DocumentSection } from '../src/internal/features/work/work-document-section-node';
import { toggleDocumentSuperscript } from '../src/internal/features/work/work-document-character-formatting';
import {
  changeDocumentIndent,
  clearDocumentParagraphPagination,
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
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DocumentFontFamily,
  DocumentTextStyle,
} from '../src/internal/features/work/work-document-word-line-metrics';

function createEditor(content = '<p>A3S Office</p>'): Editor {
  return new Editor({
    extensions: [
      StarterKit,
      DocumentSection,
      DocumentTextStyle,
      DocumentFontFamily,
      FontSize,
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      DocumentParagraphFormatting,
    ],
    content,
  });
}

describe('document formatting', () => {
  test('uses keyboard shortcuts without stacking vertical-position marks', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>x<sub>2</sub></p>',
    });
    editor.commands.setTextSelection(textRange(editor, '2'));

    expect(editor.commands.keyboardShortcut('Mod-.')).toBe(true);
    expect(editor.getHTML()).toContain('<sup>2</sup>');
    expect(editor.getHTML()).not.toContain('<sub>2</sub>');

    expect(editor.commands.keyboardShortcut('Mod-,')).toBe(true);
    expect(editor.getHTML()).toContain('<sub>2</sub>');
    expect(editor.getHTML()).not.toContain('<sup>2</sup>');
    editor.destroy();
  });

  test('keeps superscript and subscript editable across a DOCX round trip', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = '<p>H<sub>water</sub>O and x<sup>power</sup></p>';
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: '<p>H<sub>2</sub>O</p>',
        footerHtml: '<p>x<sup>2</sup></p>',
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createArtifactBlob(artifact);
    const imported = await importOfficeFile(
      new File([blob], 'vertical-position.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });

    expect(editor.getHTML()).toContain('<sub>water</sub>');
    expect(editor.getHTML()).toContain('<sup>power</sup>');
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      '<sub>2</sub>',
    );
    expect(imported.content.pageChrome?.default.footerHtml).toContain(
      '<sup>2</sup>',
    );

    editor.commands.setTextSelection(textRange(editor, 'water'));
    expect(toggleDocumentSuperscript(editor)).toBe(true);
    expect(editor.getHTML()).toContain('<sup>water</sup>');
    expect(editor.getHTML()).not.toContain('<sub>water</sub>');

    editor.commands.selectAll();
    clearDocumentFormatting(editor);
    expect(documentHasMark(editor, 'subscript')).toBe(false);
    expect(documentHasMark(editor, 'superscript')).toBe(false);

    editor.destroy();
  });

  test('round-trips paragraph alignment through DOCX import', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<p style="text-align: center;">Centered paragraph</p>',
      '<p style="text-align: right;">Right-aligned paragraph</p>',
      '<p style="text-align: justify;">Justified paragraph</p>',
      '<ol><li dir="rtl"><p style="text-align: right;">Aligned list item</p></li></ol>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const imported = await importOfficeFile(
      new File([blob], 'paragraph-alignment.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );

    expect(paragraphByText(html, 'Centered paragraph')?.style.textAlign).toBe(
      'center',
    );
    expect(
      paragraphByText(html, 'Right-aligned paragraph')?.style.textAlign,
    ).toBe('right');
    expect(paragraphByText(html, 'Justified paragraph')?.style.textAlign).toBe(
      'justify',
    );
    expect(paragraphByText(html, 'Aligned list item')?.style.textAlign).toBe(
      'right',
    );
    expect(
      paragraphByText(html, 'Aligned list item')?.closest('li'),
    ).toHaveAttribute('dir', 'rtl');
  });

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
    expect(editor.getHTML()).toContain('font-family: Arial');
    expect(editor.getHTML()).toContain('font-size: 12pt');
    expect(editor.getHTML()).toContain(
      'data-office-word-line-height-factor="1.15"',
    );
    expect(editor.getHTML()).toContain(
      '--work-document-word-line-height-factor: 1.15',
    );

    editor.destroy();
  });

  test('updates and clears the WPS line metric with the font family', () => {
    const editor = createEditor();

    editor
      .chain()
      .setTextSelection({ from: 1, to: 11 })
      .setFontFamily('Microsoft YaHei')
      .run();

    expect(editor.getHTML()).toContain(
      'data-office-word-line-height-factor="1.7143"',
    );
    editor
      .chain()
      .setTextSelection({ from: 1, to: 11 })
      .unsetFontFamily()
      .run();
    expect(editor.getHTML()).toBe('<p>A3S Office</p>');

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

    expect(clearDocumentParagraphPagination(editor)).toBe(true);
    expect(documentParagraphPagination(editor)).toEqual({
      keepLines: false,
      keepWithNext: false,
      pageBreakBefore: false,
      widowControl: true,
    });
    expect(editor.getHTML()).toBe('<p>A3S Office</p>');

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
    expect(editor.getHTML()).toContain('data-office-auto-line-height="1.725"');
    expect(editor.getHTML()).toContain(
      '--work-document-office-auto-line-height: 1.725',
    );
    expect(editor.getHTML()).toContain(
      '--work-document-office-auto-line-shift: -0.2875em',
    );
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

    clearDocumentParagraphPagination(editor);
    expect(documentParagraphPagination(editor).keepWithNext).toBe(true);
    expect(editor.getHTML()).toBe('<h2>Heading</h2><p>Body</p>');

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

function paragraphByText(
  document: Document,
  text: string,
): HTMLElement | undefined {
  return Array.from(
    document.body.querySelectorAll<HTMLElement>('p, h1, h2, h3'),
  ).find((element) => element.textContent === text);
}

function textRange(editor: Editor, text: string): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  editor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Text "${text}" was not found.`);
  return range;
}

function documentHasMark(editor: Editor, markName: string): boolean {
  let found = false;
  editor.state.doc.descendants((node) => {
    if (node.marks.some((mark) => mark.type.name === markName)) found = true;
  });
  return found;
}
