import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import {
  copyDocumentFormatting,
  pasteDocumentFormatting,
} from '../src/internal/features/work/editors/document-format-clipboard';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DOCUMENT_RUN_BORDER_ATTRIBUTE,
  documentRunBorderDomAttributes,
  normalizeDocumentRunBorder,
  parseDocumentRunBorder,
  serializeDocumentRunBorder,
} from '../src/internal/features/work/work-document-run-border';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document run borders', () => {
  test('normalizes the complete line-border model and projects an editable inline box', () => {
    const border = {
      style: 'double' as const,
      color: {
        value: '#4472c4' as const,
        theme: {
          theme: 'accent1',
          resolved: '#4472c4' as const,
          tint: '33',
        },
      },
      size: 12,
      space: 2,
      shadow: true,
      frame: true,
    };
    expect(normalizeDocumentRunBorder(border)).toEqual(border);
    const serialized = serializeDocumentRunBorder(border);
    expect(parseDocumentRunBorder(serialized)).toEqual(border);
    expect(documentRunBorderDomAttributes(border)).toEqual({
      [DOCUMENT_RUN_BORDER_ATTRIBUTE]: serialized,
      style: expect.stringContaining('box-decoration-break: clone'),
    });
    expect(documentRunBorderDomAttributes(border).style).toContain(
      'border: 2px double #4472c4',
    );
    expect(documentRunBorderDomAttributes(border).style).toContain(
      'padding: 2.667px',
    );

    for (const invalid of [
      { ...border, style: 'apples' },
      { ...border, size: 1 },
      { ...border, size: 97 },
      { ...border, space: 32 },
      { ...border, shadow: 'true' },
      { ...border, extra: true },
      [],
      '',
    ]) {
      expect(normalizeDocumentRunBorder(invalid)).toBeNull();
    }
    expect(normalizeDocumentRunBorder({ style: 'nil' })).toEqual({
      style: 'nil',
    });
    expect(normalizeDocumentRunBorder({ style: 'none' })).toEqual({
      style: 'none',
    });
  });

  test('authors, toggles, undoes, and paints a native character border without inventing a shortcut', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Border source and border target</p>',
    });
    const source = textRange(editor, 'Border source');
    editor.commands.setTextSelection(source);
    expect(
      editor.commands.setDocumentRunBorder({
        style: 'single',
        color: { value: '#c00000' },
        size: 4,
        space: 1,
      }),
    ).toBe(true);
    const authored = editor.getAttributes('textStyle').runBorder;
    expect(parseDocumentRunBorder(authored)).toEqual({
      style: 'single',
      color: { value: '#c00000' },
      size: 4,
      space: 1,
    });
    expect(editor.getHTML()).toContain(DOCUMENT_RUN_BORDER_ATTRIBUTE);
    expect(copyDocumentFormatting(editor)).toBe(true);

    const target = textRange(editor, 'border target');
    editor.commands.setTextSelection(target);
    expect(pasteDocumentFormatting(editor)).toBe(true);
    expect(
      parseDocumentRunBorder(editor.getAttributes('textStyle').runBorder),
    ).toEqual({
      style: 'single',
      color: { value: '#c00000' },
      size: 4,
      space: 1,
    });
    expect(editor.commands.undo()).toBe(true);
    editor.commands.setTextSelection(target);
    expect(editor.getAttributes('textStyle').runBorder).toBeUndefined();

    editor.commands.setTextSelection(source);
    expect(editor.commands.toggleDocumentRunBorder()).toBe(true);
    expect(
      parseDocumentRunBorder(editor.getAttributes('textStyle').runBorder),
    ).toEqual({ style: 'nil' });
    expect(editor.commands.undo()).toBe(true);
    expect(
      parseDocumentRunBorder(editor.getAttributes('textStyle').runBorder),
    ).toEqual({
      style: 'single',
      color: { value: '#c00000' },
      size: 4,
      space: 1,
    });
    expect(editor.commands.unsetDocumentRunBorder()).toBe(true);
    expect(editor.getAttributes('textStyle').runBorder).toBeUndefined();
  });

  test('resolves defaults and styles, accepts strict OOXML, and fails closed on untrusted leaves', () => {
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr>
          <w:bdr w:val="single" w:color="auto" w:sz="4" w:space="1"/>
        </w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Framed">
          <w:rPr><w:bdr w:val="double" w:color="4472C4" w:sz="12" w:space="2" w:shadow="1" w:frame="true"/></w:rPr>
        </w:style>
      </w:styles>
    `);
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Framed"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:bdr w:val="nil"/></w:rPr><w:t>reset</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>default</w:t></w:r></w:p>
    `);
    expect(
      markDocxRunFormatting(document, styles).runs.map(
        ({ formatting }) => formatting.runBorder,
      ),
    ).toEqual([
      {
        style: 'double',
        color: { value: '#4472c4' },
        size: 12,
        space: 2,
        shadow: true,
        frame: true,
      },
      { style: 'nil' },
      {
        style: 'single',
        color: { value: 'auto' },
        size: 4,
        space: 1,
      },
    ]);

    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r><s:rPr>
        <s:bdr s:val="dashDotStroked" s:color="AABBCC" s:sz="8" s:space="3" s:shadow="on" s:frame="off"/>
      </s:rPr><s:t>strict</s:t></s:r></s:p></s:body></s:document>
    `);
    expect(markDocxRunFormatting(strict).runs[0]?.formatting.runBorder).toEqual(
      {
        style: 'dashDotStroked',
        color: { value: '#aabbcc' },
        size: 8,
        space: 3,
        shadow: true,
        frame: false,
      },
    );

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:bdr evil:val="single"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:bdr evil:val="single"/></w:rPr><w:t>spoofed attribute</w:t></w:r>
        <w:r><w:rPr><w:bdr w:val="single"/><w:bdr w:val="double"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:bdr w:val="apples"/></w:rPr><w:t>art border</w:t></w:r>
        <w:r><w:rPr><w:bdr w:val="single" w:extra="1"/></w:rPr><w:t>extra</w:t></w:r>
        <w:r><w:rPr><w:bdr w:val="single"><w:b/></w:bdr></w:rPr><w:t>child</w:t></w:r>
        <w:r><w:rPr><w:bdr w:val="single">text</w:bdr></w:rPr><w:t>text</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.runBorder,
      ),
    ).toEqual(Array.from({ length: 7 }, () => ({ style: 'nil' })));
  });

  test('exports and reopens exact borders across editable stories and formatting revisions', async () => {
    const border = {
      style: 'thinThickMediumGap' as const,
      color: { value: '#4472c4' as const },
      size: 16,
      space: 3,
      shadow: true,
      frame: true,
    };
    const serialized = serializeDocumentRunBorder(border);
    if (!serialized) throw new Error('Expected a serialized run border.');
    const before = JSON.stringify([
      { type: 'textStyle', attrs: { runBorder: serialized } },
    ]);
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.trackChanges = true;
    artifact.content.html = [
      `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="run-border" data-change-author="Ada Reviewer" data-change-date="2026-08-24T10:00:00.000Z">${borderSpan(border, 'body')}</span>`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="border-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="border-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="border-footnote">',
      `<p>${borderSpan({ style: 'double', size: 8 }, 'footnote')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="border-endnote">',
      `<p>${borderSpan({ style: 'dotted', size: 6 }, 'endnote')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${borderSpan({ style: 'single', size: 4 }, 'header')}</p>`,
        footerHtml: `<p>${borderSpan({ style: 'dashed', size: 8 }, 'footer')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const storyXml = (
      await Promise.all(
        Object.keys(archive.files)
          .filter((path) =>
            /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i.test(
              path,
            ),
          )
          .map(async (path) => (await archive.file(path)?.async('text')) ?? ''),
      )
    ).join('\n');
    expect(storyXml).not.toContain('A3SOfficeRunBorder');
    expect(storyXml).toContain('<w:bdr');
    expect(storyXml).toContain('w:val="thinThickMediumGap"');
    expect(storyXml).toContain('w:color="4472C4"');
    expect(storyXml).toContain('w:shadow="1"');
    expect(storyXml).toContain('w:frame="1"');
    expect(storyXml).toMatch(
      /<w:rPrChange\b[^>]*>[\s\S]*?<w:bdr\b[^>]*w:val="thinThickMediumGap"/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'run-border.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(reopened.content.html).toContain(DOCUMENT_RUN_BORDER_ATTRIBUTE);
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      DOCUMENT_RUN_BORDER_ATTRIBUTE,
    );
    const html = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    expect(
      parseDocumentRunBorder(
        html
          .querySelector(`[${DOCUMENT_RUN_BORDER_ATTRIBUTE}]`)
          ?.getAttribute(DOCUMENT_RUN_BORDER_ATTRIBUTE),
      ),
    ).toEqual(border);
  });

  test('emits semantic markup when DOCX formatting markers are applied', () => {
    const document = wordXml(`
      <w:p><w:r><w:rPr><w:bdr w:val="wave" w:color="C00000" w:sz="8" w:space="1"/></w:rPr><w:t>bordered</w:t></w:r></w:p>
    `);
    const markers = markDocxRunFormatting(document);
    const html = new DOMParser().parseFromString(
      `<p>${markers.runs[0]?.startMarker}bordered${markers.runs[0]?.endMarker}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    const span = html.querySelector<HTMLElement>(
      `[${DOCUMENT_RUN_BORDER_ATTRIBUTE}]`,
    );
    expect(span).not.toBeNull();
    expect(span?.style.borderStyle).toBe('solid');
    expect(parseDocumentRunBorder(span?.dataset.officeRunBorder)).toEqual({
      style: 'wave',
      color: { value: '#c00000' },
      size: 8,
      space: 1,
    });
  });
});

function wordXml(body: string): Document {
  return parseXml(`
    <w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>
  `);
}

function borderSpan(
  border: NonNullable<ReturnType<typeof normalizeDocumentRunBorder>>,
  text: string,
): string {
  const attributes = documentRunBorderDomAttributes(border);
  return `<span ${DOCUMENT_RUN_BORDER_ATTRIBUTE}='${attributes[DOCUMENT_RUN_BORDER_ATTRIBUTE]}' style="${attributes.style}">${text}</span>`;
}

function textRange(
  currentEditor: Editor,
  text: string,
): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  currentEditor.state.doc.descendants((node, position) => {
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
