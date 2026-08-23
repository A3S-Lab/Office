import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE,
  DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS,
  documentKerningDomAttributes,
  documentKerningIsEffective,
  normalizeDocumentKerningThresholdHalfPoints,
} from '../src/internal/features/work/work-document-kerning';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { DOCX_EXPLICIT_ZERO_KERNING_THRESHOLD_SENTINEL } from '../src/internal/features/work/work-docx-kerning';
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

describe('document kerning threshold', () => {
  test('normalizes native half-points and projects the effective state', () => {
    expect(normalizeDocumentKerningThresholdHalfPoints(undefined)).toBeNull();
    expect(normalizeDocumentKerningThresholdHalfPoints('')).toBeNull();
    expect(normalizeDocumentKerningThresholdHalfPoints('+0')).toBe(0);
    expect(normalizeDocumentKerningThresholdHalfPoints('24')).toBe(24);
    expect(normalizeDocumentKerningThresholdHalfPoints('2.4e1')).toBeNull();
    expect(normalizeDocumentKerningThresholdHalfPoints(false)).toBeNull();
    expect(normalizeDocumentKerningThresholdHalfPoints([])).toBeNull();
    expect(
      normalizeDocumentKerningThresholdHalfPoints(
        DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS,
      ),
    ).toBe(DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS);
    expect(normalizeDocumentKerningThresholdHalfPoints(-1)).toBeNull();
    expect(
      normalizeDocumentKerningThresholdHalfPoints(
        DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS + 1,
      ),
    ).toBeNull();
    expect(normalizeDocumentKerningThresholdHalfPoints(0.5)).toBeNull();

    expect(documentKerningDomAttributes(24, '12pt')).toEqual({
      [DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE]: '24',
      style: 'font-kerning: normal',
    });
    expect(documentKerningDomAttributes(25, '12pt')).toEqual({
      [DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE]: '25',
      style: 'font-kerning: none',
    });
    expect(documentKerningDomAttributes(0, '1pt')).toEqual({
      [DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE]: '0',
      style: 'font-kerning: normal',
    });
    expect(documentKerningIsEffective(null, '48pt')).toBe(false);
    expect(documentKerningIsEffective(24, '16px')).toBe(true);
    expect(documentKerningIsEffective(25, '16px')).toBe(false);
  });

  test('recomputes rendered kerning when the effective font size changes', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p><span style="font-size: 10pt">Kerning text</span></p>',
    });
    editor.commands.setTextSelection(textRange(editor, 'Kerning text'));

    expect(editor.commands.setDocumentKerningThreshold(24)).toBe(true);
    expect(editor.getAttributes('textStyle').kerningThresholdHalfPoints).toBe(
      24,
    );
    expect(editor.getHTML()).toContain(
      'data-office-kerning-threshold-half-points="24"',
    );
    expect(editor.getHTML()).toContain('font-kerning: none');

    expect(editor.chain().setFontSize('12pt').run()).toBe(true);
    expect(editor.getHTML()).toContain('font-kerning: normal');
    expect(editor.chain().setFontSize('11pt').run()).toBe(true);
    expect(editor.getHTML()).toContain('font-kerning: none');

    expect(editor.commands.setDocumentKerningThreshold(0)).toBe(true);
    expect(editor.getHTML()).toContain(
      'data-office-kerning-threshold-half-points="0"',
    );
    expect(editor.getHTML()).toContain('font-kerning: normal');
    expect(editor.commands.unsetDocumentKerningThreshold()).toBe(true);
    expect(editor.getHTML()).not.toContain(
      'data-office-kerning-threshold-half-points',
    );

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').kerningThresholdHalfPoints).toBe(
      0,
    );
  });

  test('resolves inherited thresholds and explicit zero without flattening them', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Kerned"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="0"/></w:rPr><w:t>all-sizes</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="25"/></w:rPr><w:t>below-threshold</w:t></w:r>
      </w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="30"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Kerned">
          <w:rPr><w:kern w:val="24"/><w:sz w:val="24"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(
      markers.runs.map(
        ({ formatting }) => formatting.kerningThresholdHalfPoints,
      ),
    ).toEqual([24, 0, 25]);

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
      Array.from(
        html.querySelectorAll<HTMLElement>(
          `span[${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}]`,
        ),
      ).map((span) => ({
        threshold: span.getAttribute(DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE),
        kerning: span.style.fontKerning,
        fontSize: span.style.fontSize,
      })),
    ).toEqual([
      { threshold: '24', kerning: 'normal', fontSize: '12pt' },
      { threshold: '0', kerning: 'normal', fontSize: '12pt' },
      { threshold: '25', kerning: 'none', fontSize: '12pt' },
    ]);
  });

  test('resolves inherited font size and kerning in page chrome', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: '<p>Styled header kerning</p>',
        footerHtml: '',
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const archive = await JSZip.loadAsync(
      await (await createDocxBlob(artifact.content)).arrayBuffer(),
    );
    const headerPath = Object.keys(archive.files).find((path) =>
      /^word\/header\d+\.xml$/i.test(path),
    );
    const headerSource = headerPath
      ? await archive.file(headerPath)?.async('text')
      : null;
    const stylesSource = await archive.file('word/styles.xml')?.async('text');
    if (!headerPath || !headerSource || !stylesSource) {
      throw new Error('Expected generated header and styles parts.');
    }

    const headerDocument = parseXml(headerSource);
    const paragraph = descendants(headerDocument, 'p').find((element) =>
      element.textContent?.includes('Styled header kerning'),
    );
    if (!paragraph) throw new Error('Expected the generated header paragraph.');
    let paragraphProperties = directChild(paragraph, 'pPr');
    if (!paragraphProperties) {
      paragraphProperties = headerDocument.createElementNS(
        WORD_NAMESPACE,
        'w:pPr',
      );
      paragraph.insertBefore(paragraphProperties, paragraph.firstChild);
    }
    directChild(paragraphProperties, 'pStyle')?.remove();
    const styleReference = headerDocument.createElementNS(
      WORD_NAMESPACE,
      'w:pStyle',
    );
    styleReference.setAttributeNS(WORD_NAMESPACE, 'w:val', 'KernedHeader');
    paragraphProperties.prepend(styleReference);
    archive.file(
      headerPath,
      new XMLSerializer().serializeToString(headerDocument),
    );

    const stylesDocument = parseXml(stylesSource);
    const style = stylesDocument.createElementNS(WORD_NAMESPACE, 'w:style');
    style.setAttributeNS(WORD_NAMESPACE, 'w:type', 'paragraph');
    style.setAttributeNS(WORD_NAMESPACE, 'w:styleId', 'KernedHeader');
    const runProperties = stylesDocument.createElementNS(
      WORD_NAMESPACE,
      'w:rPr',
    );
    for (const [name, value] of [
      ['sz', '24'],
      ['kern', '24'],
    ] as const) {
      const property = stylesDocument.createElementNS(
        WORD_NAMESPACE,
        `w:${name}`,
      );
      property.setAttributeNS(WORD_NAMESPACE, 'w:val', value);
      runProperties.append(property);
    }
    style.append(runProperties);
    stylesDocument.documentElement.append(style);
    archive.file(
      'word/styles.xml',
      new XMLSerializer().serializeToString(stylesDocument),
    );

    const imported = await importOfficeFile(
      new File(
        [await archive.generateAsync({ type: 'uint8array' })],
        'styled-header-kerning.docx',
        {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
      ),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const headerHtml = imported.content.pageChrome?.default.headerHtml ?? '';
    expect(headerHtml).toContain(
      `${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}="24"`,
    );
    expect(headerHtml).toContain('font-size: 12pt');
    expect(headerHtml).toContain('font-kerning: normal');
  });

  test('accepts strict integers and rejects malformed or spoofed kerning', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr><s:kern s:val="+0024"/></s:rPr><s:t>strict</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strict).runs[0]?.formatting
        .kerningThresholdHalfPoints,
    ).toBe(24);

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:kern w:val="24"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:kern evil:val="24"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:kern/></w:rPr><w:t>missing value</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="-1"/></w:rPr><w:t>negative</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="3278"/></w:rPr><w:t>too large</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="24.5"/></w:rPr><w:t>fractional</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="24"/><w:kern w:val="25"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="24"><w:b/></w:kern></w:rPr><w:t>nested</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="24">text</w:kern></w:rPr><w:t>text child</w:t></w:r>
        <w:r><w:rPr><w:kern w:val="24" w:extra="1"/></w:rPr><w:t>extra attribute</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.kerningThresholdHalfPoints,
      ),
    ).toEqual(Array.from({ length: 10 }, () => undefined));

    const strictMeasure = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr><s:kern s:val="12pt"/></s:rPr><s:t>measure</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strictMeasure).runs[0]?.formatting
        .kerningThresholdHalfPoints,
    ).toBeUndefined();
  });

  test('imports strict native kerning revisions', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr>
          <s:kern s:val="24"/>
          <s:rPrChange s:id="9" s:author="Ada Reviewer" s:date="2026-08-23T11:15:00.000Z">
            <s:rPr><s:kern s:val="0"/></s:rPr>
          </s:rPrChange>
        </s:rPr>
        <s:t>strict revision</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    const marker = markDocxRunFormatting(strict).runs[0];
    expect(marker?.formatting.kerningThresholdHalfPoints).toBe(24);
    expect(marker?.change).toMatchObject({
      author: 'Ada Reviewer',
      date: '2026-08-23T11:15:00.000Z',
      id: 'docx-format-change-9',
    });
    expect(marker?.change?.before).toContain('"kerningThresholdHalfPoints":0');
  });

  test('exports and reopens exact thresholds across every editable Word story', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      `<p>${kerningSpan(24, 12, 'threshold')} ${kerningSpan(
        0,
        8,
        'all-sizes',
      )} ${kerningSpan(25, 12, 'below-threshold')}`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="kerning-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="kerning-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="kerning-footnote">',
      `<p>${kerningSpan(0, 8, 'footnote-kerning')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="kerning-endnote">',
      `<p>${kerningSpan(25, 12, 'endnote-kerning')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${kerningSpan(0, 10, 'header-kerning')}</p>`,
        footerHtml: `<p>${kerningSpan(24, 12, 'footer-kerning')}</p>`,
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
    for (const threshold of [0, 24, 25]) {
      expect(xml).toMatch(new RegExp(`<w:kern\\b[^>]*w:val="${threshold}"`));
    }
    expect(xml).not.toContain(
      `w:val="${DOCX_EXPLICIT_ZERO_KERNING_THRESHOLD_SENTINEL}"`,
    );
    expect(await archive.file('word/footnotes.xml')?.async('text')).toMatch(
      /<w:kern\b[^>]*w:val="0"/,
    );
    expect(await archive.file('word/endnotes.xml')?.async('text')).toMatch(
      /<w:kern\b[^>]*w:val="25"/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'kerning.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    for (const threshold of [0, 24, 25]) {
      expect(reopened.content.html).toContain(
        `${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}="${threshold}"`,
      );
    }
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      `${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}="0"`,
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      `${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}="24"`,
    );
    const reopenedDocument = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    expect(
      reopenedDocument.querySelector(
        'aside[data-note-kind="footnote"] span[data-office-kerning-threshold-half-points="0"]',
      ),
    ).not.toBeNull();
    expect(
      reopenedDocument.querySelector(
        'aside[data-note-kind="endnote"] span[data-office-kerning-threshold-half-points="25"]',
      ),
    ).not.toBeNull();
  });

  test('exports and reopens native kerning revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before =
      '[{"type":"textStyle","attrs":{"kerningThresholdHalfPoints":0}}]';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-kerning" data-change-author="Ada Reviewer" data-change-date="2026-08-23T11:30:00.000Z">${kerningSpan(24, 12, 'Changed kerning')}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed kerning',
    );
    if (!run) throw new Error('Expected changed kerning run.');
    const properties = directChild(run, 'rPr');
    const currentKerning = directChild(properties ?? run, 'kern');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    const priorKerning = directChild(prior ?? run, 'kern');
    if (!currentKerning || !change || !prior || !priorKerning) {
      throw new Error('Expected current and prior kerning thresholds.');
    }
    expect(attribute(currentKerning, 'val')).toBe('24');
    expect(attribute(priorKerning, 'val')).toBe('0');

    const reopened = await importOfficeFile(
      new File([blob], 'kerning-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      `${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}="24"`,
    );
    expect(reopened.content.html).toContain(
      '&quot;kerningThresholdHalfPoints&quot;:0',
    );
  });

  test('tracks and rejects an exact native kerning threshold', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'formatting-kerning',
          author: 'Reviewer',
          date: '2026-08-23T11:45:00.000Z',
        }),
      }),
      content: `<section data-document-section="true"><p>${kerningSpan(
        24,
        12,
        'Tracked kerning',
      )}</p></section>`,
    });
    editor.commands.setTextSelection(textRange(editor, 'Tracked kerning'));

    expect(editor.commands.setDocumentKerningThreshold(0)).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'formatting-kerning',
        kind: 'formatting',
        text: 'Tracked kerning',
      }),
    ]);
    expect(editor.getHTML()).toContain(
      '&quot;kerningThresholdHalfPoints&quot;:24',
    );

    expect(editor.commands.rejectDocumentChange('formatting-kerning')).toBe(
      true,
    );
    expect(editor.getAttributes('textStyle').kerningThresholdHalfPoints).toBe(
      24,
    );
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function kerningSpan(
  threshold: number,
  fontSizePoints: number,
  content: string,
): string {
  const attributes = documentKerningDomAttributes(
    threshold,
    `${fontSizePoints}pt`,
  );
  return `<span ${DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE}="${attributes[DOCUMENT_KERNING_THRESHOLD_ATTRIBUTE]}" style="font-size: ${fontSizePoints}pt; ${attributes.style}">${content}</span>`;
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
