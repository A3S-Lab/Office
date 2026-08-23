import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import {
  DOCUMENT_CHARACTER_SCALE_ATTRIBUTE,
  DOCUMENT_CHARACTER_SCALE_MAX_PERCENT,
  documentCharacterScaleDomAttributes,
  documentCharacterScalePercentFromCss,
  normalizeDocumentCharacterScalePercent,
} from '../src/internal/features/work/work-document-character-scale';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document character scale', () => {
  test('normalizes the native range and preserves explicit 100 percent', () => {
    expect(normalizeDocumentCharacterScalePercent(undefined)).toBeNull();
    expect(normalizeDocumentCharacterScalePercent('')).toBeNull();
    expect(normalizeDocumentCharacterScalePercent('+1')).toBe(1);
    expect(normalizeDocumentCharacterScalePercent('100')).toBe(100);
    expect(
      normalizeDocumentCharacterScalePercent(
        DOCUMENT_CHARACTER_SCALE_MAX_PERCENT,
      ),
    ).toBe(DOCUMENT_CHARACTER_SCALE_MAX_PERCENT);
    expect(normalizeDocumentCharacterScalePercent(0)).toBeNull();
    expect(
      normalizeDocumentCharacterScalePercent(
        DOCUMENT_CHARACTER_SCALE_MAX_PERCENT + 1,
      ),
    ).toBeNull();
    expect(normalizeDocumentCharacterScalePercent(100.5)).toBeNull();

    expect(documentCharacterScaleDomAttributes(80)).toEqual({
      [DOCUMENT_CHARACTER_SCALE_ATTRIBUTE]: '80',
      style: 'font-stretch: 80%',
    });
    expect(documentCharacterScaleDomAttributes(100)).toEqual({
      [DOCUMENT_CHARACTER_SCALE_ATTRIBUTE]: '100',
      style: 'font-stretch: 100%',
    });
    expect(documentCharacterScalePercentFromCss('125%')).toBe(125);
    expect(documentCharacterScalePercentFromCss('normal')).toBeNull();
    expect(documentCharacterScalePercentFromCss('100.5%')).toBeNull();
  });

  test('authors exact scale as one undoable mark change', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Character scale</p>',
    });
    editor.commands.setTextSelection(textRange(editor, 'Character'));

    expect(editor.commands.setDocumentCharacterScale(80)).toBe(true);
    expect(editor.getAttributes('textStyle').characterScalePercent).toBe(80);
    expect(editor.getHTML()).toContain(
      'data-office-character-scale-percent="80"',
    );
    expect(editor.getHTML()).toContain('font-stretch: 80%');

    expect(editor.commands.setDocumentCharacterScale(100)).toBe(true);
    expect(editor.getAttributes('textStyle').characterScalePercent).toBe(100);
    expect(editor.getHTML()).toContain(
      'data-office-character-scale-percent="100"',
    );

    expect(editor.commands.setDocumentCharacterScale(601)).toBe(false);
    expect(editor.commands.setDocumentCharacterScale(0)).toBe(false);
    expect(editor.commands.setDocumentCharacterScale(100.5)).toBe(false);
    expect(editor.getAttributes('textStyle').characterScalePercent).toBe(100);

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').characterScalePercent).toBe(80);
  });

  test('resolves style inheritance and a direct default reset without flattening scale', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Scaled"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:w/></w:rPr><w:t>default</w:t></w:r>
        <w:r><w:rPr><w:w w:val="80"/></w:rPr><w:t>condensed</w:t></w:r>
      </w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:w w:val="90"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Scaled">
          <w:rPr><w:w w:val="125"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(
      markers.runs.map(({ formatting }) => formatting.characterScalePercent),
    ).toEqual([125, 100, 80]);

    const html = new DOMParser().parseFromString(
      `<p>${markers.runs
        .map(
          ({ startMarker, endMarker }, index) =>
            `${startMarker}run-${index}${endMarker}`,
        )
        .join(' ')}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    expect(
      Array.from(html.querySelectorAll<HTMLElement>('span')).map((span) => ({
        scale: span.dataset.officeCharacterScalePercent,
        style: span.style.fontStretch,
      })),
    ).toEqual([
      { scale: '125', style: '125%' },
      { scale: '100', style: '100%' },
      { scale: '80', style: '80%' },
    ]);
  });

  test('accepts strict integers and rejects malformed or spoofed scale', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr><s:w s:val="125"/></s:rPr><s:t>strict</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strict).runs[0]?.formatting.characterScalePercent,
    ).toBe(125);

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:w w:val="80"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:w evil:val="80"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:w w:val="0"/></w:rPr><w:t>too small</w:t></w:r>
        <w:r><w:rPr><w:w w:val="601"/></w:rPr><w:t>too large</w:t></w:r>
        <w:r><w:rPr><w:w w:val="100.5"/></w:rPr><w:t>fractional</w:t></w:r>
        <w:r><w:rPr><w:w w:val="80"/><w:w w:val="90"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:w w:val="80"><w:b/></w:w></w:rPr><w:t>nested</w:t></w:r>
        <w:r><w:rPr><w:w w:val="80">text</w:w></w:rPr><w:t>text child</w:t></w:r>
        <w:r><w:rPr><w:w w:val="80" w:extra="1"/></w:rPr><w:t>extra attribute</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.characterScalePercent,
      ),
    ).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
  });

  test('imports strict native character-scale revisions', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr>
          <s:w s:val="80"/>
          <s:rPrChange s:id="8" s:author="Ada Reviewer" s:date="2026-08-23T09:15:00.000Z">
            <s:rPr><s:w/></s:rPr>
          </s:rPrChange>
        </s:rPr>
        <s:t>strict revision</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    const marker = markDocxRunFormatting(strict).runs[0];
    expect(marker?.formatting.characterScalePercent).toBe(80);
    expect(marker?.change).toMatchObject({
      author: 'Ada Reviewer',
      date: '2026-08-23T09:15:00.000Z',
      id: 'docx-format-change-8',
    });
    expect(marker?.change?.before).toContain('"characterScalePercent":100');
  });

  test('exports and reopens exact scale in body, header, and footer', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${scaleSpan(80, 'condensed')} ${scaleSpan(
      100,
      'explicit-default',
    )} ${scaleSpan(125, 'expanded')}</p>`;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${scaleSpan(75, 'header-scale')}</p>`,
        footerHtml: `<p>${scaleSpan(150, 'footer-scale')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const parts = await Promise.all(
      Object.keys(archive.files)
        .filter((path) =>
          /^word\/(?:document|header\d*|footer\d*)\.xml$/i.test(path),
        )
        .map(
          async (path) =>
            [path, (await archive.file(path)?.async('text')) ?? ''] as const,
        ),
    );
    const xml = parts.map(([, source]) => source).join('\n');
    for (const scale of [80, 100, 125, 75, 150]) {
      expect(xml).toMatch(new RegExp(`<w:w\\b[^>]*w:val="${scale}"`));
    }

    const imported = await importOfficeFile(
      new File([blob], 'character-scale.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    for (const scale of [80, 100, 125]) {
      expect(imported.content.html).toContain(
        `data-office-character-scale-percent="${scale}"`,
      );
    }
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      'data-office-character-scale-percent="75"',
    );
    expect(imported.content.pageChrome?.default.footerHtml).toContain(
      'data-office-character-scale-percent="150"',
    );
  });

  test('exports and reopens native character-scale revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before =
      '[{"type":"textStyle","attrs":{"characterScalePercent":100}}]';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-scale" data-change-author="Ada Reviewer" data-change-date="2026-08-23T09:30:00.000Z">${scaleSpan(80, 'Changed scale')}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed scale',
    );
    if (!run) throw new Error('Expected changed character-scale run.');
    const properties = directChild(run, 'rPr');
    const currentScale = directChild(properties ?? run, 'w');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    const priorScale = directChild(prior ?? run, 'w');
    if (!currentScale || !change || !prior || !priorScale) {
      throw new Error('Expected current and prior character scale.');
    }
    expect(attribute(currentScale, 'val')).toBe('80');
    expect(attribute(priorScale, 'val')).toBe('100');

    const reopened = await importOfficeFile(
      new File([blob], 'character-scale-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      'data-office-character-scale-percent="80"',
    );
    expect(reopened.content.html).toContain(
      '&quot;characterScalePercent&quot;:100',
    );
  });

  test('tracks and restores exact native character scale', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'formatting-scale',
          author: 'Reviewer',
          date: '2026-08-23T09:45:00.000Z',
        }),
      }),
      content: `<section data-document-section="true"><p>${scaleSpan(
        125,
        'Tracked scale',
      )}</p></section>`,
    });
    editor.commands.setTextSelection(textRange(editor, 'Tracked scale'));

    expect(editor.commands.setDocumentCharacterScale(80)).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'formatting-scale',
        kind: 'formatting',
        text: 'Tracked scale',
      }),
    ]);
    expect(editor.getHTML()).toContain('&quot;characterScalePercent&quot;:125');

    expect(editor.commands.rejectDocumentChange('formatting-scale')).toBe(true);
    expect(editor.getAttributes('textStyle').characterScalePercent).toBe(125);
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function scaleSpan(scale: number, content: string): string {
  const attributes = documentCharacterScaleDomAttributes(scale);
  return `<span ${DOCUMENT_CHARACTER_SCALE_ATTRIBUTE}="${attributes[DOCUMENT_CHARACTER_SCALE_ATTRIBUTE]}" style="${attributes.style}">${content}</span>`;
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
