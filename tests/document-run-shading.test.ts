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
  DOCUMENT_RUN_SHADING_ATTRIBUTE,
  DOCUMENT_RUN_SHADING_PATTERNS,
  documentRunShadingDomAttributes,
  normalizeDocumentRunShading,
  parseDocumentRunShading,
  serializeDocumentRunShading,
} from '../src/internal/features/work/work-document-run-shading';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import {
  attribute,
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document run shading', () => {
  test('normalizes every native pattern and projects bounded inline paint', () => {
    const shading = {
      pattern: 'diagCross' as const,
      color: {
        value: '#4472c4' as const,
        theme: {
          theme: 'accent1',
          resolved: '#4472c4',
          tint: '33',
        },
      },
      fill: {
        value: '#fff2cc' as const,
        theme: {
          theme: 'background1',
          resolved: '#fff2cc',
          shade: 'BF',
        },
      },
    };
    expect(normalizeDocumentRunShading(shading)).toEqual(shading);
    const serialized = serializeDocumentRunShading(shading);
    expect(parseDocumentRunShading(serialized)).toEqual(shading);
    expect(documentRunShadingDomAttributes(shading)).toEqual({
      [DOCUMENT_RUN_SHADING_ATTRIBUTE]: serialized,
      style: expect.stringContaining('box-decoration-break: clone'),
    });
    expect(documentRunShadingDomAttributes(shading).style).toContain(
      'background-image:',
    );

    for (const pattern of DOCUMENT_RUN_SHADING_PATTERNS) {
      expect(normalizeDocumentRunShading({ pattern })).toEqual({ pattern });
    }
    for (const invalid of [
      { ...shading, pattern: 'unknown' },
      { ...shading, color: { value: '#12345' } },
      { ...shading, fill: { value: 'transparent' } },
      { ...shading, extra: true },
      [],
      '',
    ]) {
      expect(normalizeDocumentRunShading(invalid)).toBeNull();
    }
    expect(parseDocumentRunShading('x'.repeat(4_097))).toBeNull();
  });

  test('authors, resets, copies, pastes, and undoes semantic run shading', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Shading source and shading target</p>',
    });
    const source = textRange(editor, 'Shading source');
    editor.commands.setTextSelection(source);
    expect(
      editor.commands.setDocumentRunShading({
        pattern: 'pct25',
        color: { value: '#4472c4' },
        fill: { value: '#fff2cc' },
      }),
    ).toBe(true);
    expect(
      parseDocumentRunShading(editor.getAttributes('textStyle').runShading),
    ).toEqual({
      pattern: 'pct25',
      color: { value: '#4472c4' },
      fill: { value: '#fff2cc' },
    });
    expect(editor.getHTML()).toContain(DOCUMENT_RUN_SHADING_ATTRIBUTE);
    expect(copyDocumentFormatting(editor)).toBe(true);

    const target = textRange(editor, 'shading target');
    editor.commands.setTextSelection(target);
    expect(pasteDocumentFormatting(editor)).toBe(true);
    expect(
      parseDocumentRunShading(editor.getAttributes('textStyle').runShading),
    ).toMatchObject({ pattern: 'pct25' });
    expect(editor.commands.undo()).toBe(true);
    editor.commands.setTextSelection(target);
    expect(editor.getAttributes('textStyle').runShading).toBeUndefined();

    editor.commands.setTextSelection(source);
    expect(editor.commands.setDocumentRunShading({ pattern: 'nil' })).toBe(
      true,
    );
    expect(
      parseDocumentRunShading(editor.getAttributes('textStyle').runShading),
    ).toEqual({ pattern: 'nil' });
    expect(editor.commands.undo()).toBe(true);
    expect(editor.commands.unsetDocumentRunShading()).toBe(true);
    expect(editor.getAttributes('textStyle').runShading).toBeUndefined();
  });

  test('resolves defaults, table and character styles, direct resets, and strict markup', () => {
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr>
          <w:shd w:val="clear" w:fill="FFF2CC"/>
        </w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="character" w:styleId="PatternBase">
          <w:rPr><w:shd w:val="pct20" w:color="4472C4" w:fill="DDEBF7"/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="Pattern">
          <w:basedOn w:val="PatternBase"/>
          <w:rPr><w:shd w:val="diagCross" w:color="C00000" w:fill="FCE4D6"/></w:rPr>
        </w:style>
        <w:style w:type="table" w:styleId="Grid">
          <w:tblStylePr w:type="wholeTable"><w:rPr>
            <w:shd w:val="horzStripe" w:color="70AD47" w:fill="E2F0D9"/>
          </w:rPr></w:tblStylePr>
        </w:style>
      </w:styles>
    `);
    const document = wordXml(`
      <w:p><w:r><w:rPr><w:rStyle w:val="Pattern"/></w:rPr><w:t>styled</w:t></w:r></w:p>
      <w:tbl><w:tblPr><w:tblStyle w:val="Grid"/></w:tblPr><w:tr><w:tc><w:p>
        <w:r><w:t>table</w:t></w:r>
        <w:r><w:rPr><w:shd w:val="nil"/></w:rPr><w:t>reset</w:t></w:r>
      </w:p></w:tc></w:tr></w:tbl>
      <w:p><w:r><w:t>default</w:t></w:r></w:p>
    `);
    expect(
      markDocxRunFormatting(document, styles, undefined, styles).runs.map(
        ({ formatting }) => formatting.runShading,
      ),
    ).toEqual([
      {
        pattern: 'diagCross',
        color: { value: '#c00000' },
        fill: { value: '#fce4d6' },
      },
      {
        pattern: 'horzStripe',
        color: { value: '#70ad47' },
        fill: { value: '#e2f0d9' },
      },
      { pattern: 'nil' },
      { pattern: 'clear', fill: { value: '#fff2cc' } },
    ]);

    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r><s:rPr>
        <s:shd s:val="pct37" s:color="112233" s:fill="AABBCC"/>
      </s:rPr><s:t>strict</s:t></s:r></s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strict).runs[0]?.formatting.runShading,
    ).toEqual({
      pattern: 'pct37',
      color: { value: '#112233' },
      fill: { value: '#aabbcc' },
    });
  });

  test('fails closed on malformed leaves and preserves highlight independently', () => {
    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:shd evil:val="clear"/></w:rPr><w:t>spoofed</w:t></w:r>
        <w:r><w:rPr><w:shd evil:val="clear"/></w:rPr><w:t>spoofed attr</w:t></w:r>
        <w:r><w:rPr><w:shd w:val="clear"/><w:shd w:val="solid"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:shd w:val="unknown"/></w:rPr><w:t>unknown</w:t></w:r>
        <w:r><w:rPr><w:shd w:val="clear" w:extra="1"/></w:rPr><w:t>extra</w:t></w:r>
        <w:r><w:rPr><w:shd w:val="clear"><w:b/></w:shd></w:rPr><w:t>child</w:t></w:r>
        <w:r><w:rPr><w:shd w:val="clear">text</w:shd></w:rPr><w:t>text</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.runShading,
      ),
    ).toEqual(Array.from({ length: 7 }, () => ({ pattern: 'nil' })));

    const coexist = wordXml(`
      <w:p><w:r><w:rPr>
        <w:highlight w:val="darkCyan"/>
        <w:shd w:val="pct25" w:color="112233" w:fill="DDEEFF"/>
      </w:rPr><w:t>Both</w:t></w:r></w:p>
    `);
    const markers = markDocxRunFormatting(coexist);
    expect(markers.runs[0]?.formatting).toMatchObject({
      highlight: 'darkCyan',
      backgroundColor: '#008080',
      runShading: {
        pattern: 'pct25',
        color: { value: '#112233' },
        fill: { value: '#ddeeff' },
      },
    });
    const html = new DOMParser().parseFromString(
      `<p>${markers.runs[0]?.startMarker}Both${markers.runs[0]?.endMarker}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    expect(html.body.innerHTML).toContain(DOCUMENT_RUN_SHADING_ATTRIBUTE);
    expect(html.body.innerHTML).toContain('data-office-highlight="darkCyan"');

    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html.body.innerHTML,
    });
    const serialized = editor.getHTML();
    expect(serialized).toContain(DOCUMENT_RUN_SHADING_ATTRIBUTE);
    expect(serialized).toContain('data-office-highlight="darkCyan"');
  });

  test('exports and reopens exact shading across editable stories and revisions', async () => {
    const shading = {
      pattern: 'pct62' as const,
      color: {
        value: '#548235' as const,
        theme: {
          theme: 'accent6',
          resolved: '#548235',
          shade: 'BF',
        },
      },
      fill: {
        value: '#d2d2d2' as const,
        theme: {
          theme: 'accent3',
          resolved: '#d2d2d2',
          tint: '80',
          shade: '40',
        },
      },
    };
    const serialized = serializeDocumentRunShading(shading);
    if (!serialized) throw new Error('Expected serialized run shading.');
    const before = JSON.stringify([
      { type: 'textStyle', attrs: { runShading: serialized } },
    ]);
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.trackChanges = true;
    artifact.content.html = [
      `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="run-shading" data-change-author="Ada Reviewer" data-change-date="2026-08-24T10:00:00.000Z">${shadingSpan(shading, 'body')}</span>`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="shading-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="shading-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="shading-footnote">',
      `<p>${shadingSpan({ pattern: 'diagCross', color: { value: '#112233' }, fill: { value: '#ddeeff' } }, 'footnote')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="shading-endnote">',
      `<p>${shadingSpan({ pattern: 'nil' }, 'endnote')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${shadingSpan({ pattern: 'horzStripe', color: { value: '#4472c4' }, fill: { value: '#fff2cc' } }, 'header')}</p>`,
        footerHtml: `<p>${shadingSpan({ pattern: 'clear', fill: { value: '#e2f0d9' } }, 'footer')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const stories = await Promise.all(
      Object.keys(archive.files)
        .filter((path) =>
          /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i.test(
            path,
          ),
        )
        .map(async (path) => (await archive.file(path)?.async('text')) ?? ''),
    );
    const storyXml = stories.join('\n');
    expect(storyXml).not.toContain('A3SOfficeRunShading');
    expect(storyXml).toMatch(
      /<w:shd\b[^>]*w:val="pct62"[^>]*w:color="548235"[^>]*w:themeColor="accent6"[^>]*w:themeShade="BF"[^>]*w:fill="D2D2D2"[^>]*w:themeFill="accent3"[^>]*w:themeFillTint="80"[^>]*w:themeFillShade="40"/,
    );
    expect(storyXml).toMatch(
      /<w:rPrChange\b[^>]*>[\s\S]*?<w:shd\b[^>]*w:val="pct62"/,
    );
    expect(storyXml).toContain('w:val="diagCross"');
    expect(storyXml).toContain('w:val="nil"');

    const reopened = await importOfficeFile(
      new File([blob], 'run-shading.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(reopened.content.html).toContain(DOCUMENT_RUN_SHADING_ATTRIBUTE);
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      DOCUMENT_RUN_SHADING_ATTRIBUTE,
    );
    const html = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    expect(
      parseDocumentRunShading(
        html
          .querySelector(`[${DOCUMENT_RUN_SHADING_ATTRIBUTE}]`)
          ?.getAttribute(DOCUMENT_RUN_SHADING_ATTRIBUTE),
      ),
    ).toEqual(shading);
  });

  test('round-trips independent themed channels and native highlight together', async () => {
    const document = wordXml(`
      <w:p><w:r><w:rPr>
        <w:highlight w:val="yellow"/>
        <w:shd w:val="pct20" w:color="548235" w:themeColor="accent6" w:themeShade="BF" w:fill="D2D2D2" w:themeFill="accent3" w:themeFillTint="80"/>
      </w:rPr><w:t>Themed pair</w:t></w:r></w:p>
    `);
    const theme = parseXml(`
      <a:theme xmlns:a="${DRAWING_NAMESPACE}"><a:themeElements><a:clrScheme name="Office">
        <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      </a:clrScheme></a:themeElements></a:theme>
    `);
    const markers = markDocxRunFormatting(document, undefined, theme);
    const marker = markers.runs[0];
    if (!marker) throw new Error('Expected run-formatting markers.');
    const html = new DOMParser().parseFromString(
      `<p>${marker.startMarker}Themed pair${marker.endMarker}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const exported = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const properties = descendants(exported, 'rPr').find(
      (element) => element.parentElement?.localName === 'r',
    );
    if (!properties) throw new Error('Expected exported run properties.');
    const highlight = descendants(properties, 'highlight')[0];
    const shading = descendants(properties, 'shd')[0];
    expect(attribute(highlight, 'val')).toBe('yellow');
    expect(attribute(shading, 'val')).toBe('pct20');
    expect(attribute(shading, 'themeColor')).toBe('accent6');
    expect(attribute(shading, 'themeShade')).toBe('BF');
    expect(attribute(shading, 'themeFill')).toBe('accent3');
    expect(attribute(shading, 'themeFillTint')).toBe('80');
  });
});

function wordXml(body: string): Document {
  return parseXml(`
    <w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>
  `);
}

function shadingSpan(
  shading: NonNullable<ReturnType<typeof normalizeDocumentRunShading>>,
  text: string,
): string {
  const attributes = documentRunShadingDomAttributes(shading);
  return `<span ${DOCUMENT_RUN_SHADING_ATTRIBUTE}='${attributes[DOCUMENT_RUN_SHADING_ATTRIBUTE]}' style="${attributes.style}">${text}</span>`;
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
