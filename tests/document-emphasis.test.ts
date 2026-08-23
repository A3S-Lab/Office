import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import {
  DOCUMENT_EMPHASIS_MARK_ATTRIBUTE,
  documentEmphasisMarkDomAttributes,
  normalizeDocumentEmphasisMark,
  type WorkDocumentEmphasisMark,
} from '../src/internal/features/work/work-document-emphasis';
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

describe('document emphasis marks', () => {
  test('normalizes all five native values and exposes reversible commands', () => {
    expect(normalizeDocumentEmphasisMark('none')).toBe('none');
    expect(normalizeDocumentEmphasisMark('dot')).toBe('dot');
    expect(normalizeDocumentEmphasisMark('comma')).toBe('comma');
    expect(normalizeDocumentEmphasisMark('circle')).toBe('circle');
    expect(normalizeDocumentEmphasisMark('underDot')).toBe('underDot');
    for (const invalid of [
      undefined,
      null,
      '',
      'Dot',
      'under-dot',
      'underDot ',
      1,
      false,
      [],
    ]) {
      expect(normalizeDocumentEmphasisMark(invalid)).toBeNull();
    }

    expect(documentEmphasisMarkDomAttributes('dot')).toEqual({
      [DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]: 'dot',
      style:
        'text-emphasis-style:filled dot;text-emphasis-position:over right;-webkit-text-emphasis-style:filled dot;-webkit-text-emphasis-position:over right',
    });
    expect(documentEmphasisMarkDomAttributes('comma')).toEqual({
      [DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]: 'comma',
      style:
        'text-emphasis-style:",";text-emphasis-position:over right;-webkit-text-emphasis-style:",";-webkit-text-emphasis-position:over right',
    });
    expect(documentEmphasisMarkDomAttributes('circle')).toEqual({
      [DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]: 'circle',
      style:
        'text-emphasis-style:open circle;text-emphasis-position:over right;-webkit-text-emphasis-style:open circle;-webkit-text-emphasis-position:over right',
    });
    expect(documentEmphasisMarkDomAttributes('underDot')).toEqual({
      [DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]: 'underDot',
      style:
        'text-emphasis-style:filled dot;text-emphasis-position:under right;-webkit-text-emphasis-style:filled dot;-webkit-text-emphasis-position:under right',
    });
    expect(documentEmphasisMarkDomAttributes('none')).toEqual({
      [DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]: 'none',
      style: 'text-emphasis-style:none;-webkit-text-emphasis-style:none',
    });

    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>原生着重号</p>',
    });
    editor.commands.selectAll();
    expect(editor.commands.setDocumentEmphasisMark('circle')).toBe(true);
    expect(editor.getAttributes('textStyle').emphasisMark).toBe('circle');
    expect(editor.getHTML()).toContain(
      `${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="circle"`,
    );
    expect(editor.commands.setDocumentEmphasisMark('invalid')).toBe(false);
    expect(editor.commands.unsetDocumentEmphasisMark()).toBe(true);
    expect(editor.getHTML()).not.toContain(DOCUMENT_EMPHASIS_MARK_ATTRIBUTE);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').emphasisMark).toBe('circle');
  });

  test('resolves inherited marks and preserves explicit none', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Emphasized"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:em w:val="none"/></w:rPr><w:t>reset</w:t></w:r>
        <w:r><w:rPr><w:em w:val="underDot"/></w:rPr><w:t>below</w:t></w:r>
      </w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:em w:val="dot"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Emphasized">
          <w:rPr><w:em w:val="circle"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(
      markers.runs.map(({ formatting }) => formatting.emphasisMark),
    ).toEqual(['circle', 'none', 'underDot']);

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
          `span[${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}]`,
        ),
      ).map((span) => ({
        mark: span.getAttribute(DOCUMENT_EMPHASIS_MARK_ATTRIBUTE),
        position: span.style.getPropertyValue('text-emphasis-position'),
        style: span.style.getPropertyValue('text-emphasis-style'),
      })),
    ).toEqual([
      { mark: 'circle', position: 'over right', style: 'open circle' },
      { mark: 'none', position: '', style: 'none' },
      { mark: 'underDot', position: 'under right', style: 'filled dot' },
    ]);
  });

  test('accepts strict native values and rejects malformed or spoofed properties', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr><s:em s:val="comma"/></s:rPr><s:t>strict</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    expect(markDocxRunFormatting(strict).runs[0]?.formatting.emphasisMark).toBe(
      'comma',
    );

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:em w:val="dot"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:em evil:val="dot"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:em/></w:rPr><w:t>missing value</w:t></w:r>
        <w:r><w:rPr><w:em w:val="Dot"/></w:rPr><w:t>wrong case</w:t></w:r>
        <w:r><w:rPr><w:em w:val="under-dot"/></w:rPr><w:t>wrong token</w:t></w:r>
        <w:r><w:rPr><w:em w:val="dot"/><w:em w:val="circle"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:em w:val="dot"><w:b/></w:em></w:rPr><w:t>nested</w:t></w:r>
        <w:r><w:rPr><w:em w:val="dot">text</w:em></w:rPr><w:t>text child</w:t></w:r>
        <w:r><w:rPr><w:em w:val="dot" w:extra="1"/></w:rPr><w:t>extra attribute</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.emphasisMark,
      ),
    ).toEqual(Array.from({ length: 9 }, () => undefined));
  });

  test('resolves inherited emphasis in page chrome', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: '<p>Styled header emphasis</p>',
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
      element.textContent?.includes('Styled header emphasis'),
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
    styleReference.setAttributeNS(WORD_NAMESPACE, 'w:val', 'EmphasizedHeader');
    paragraphProperties.prepend(styleReference);
    archive.file(
      headerPath,
      new XMLSerializer().serializeToString(headerDocument),
    );

    const stylesDocument = parseXml(stylesSource);
    const style = stylesDocument.createElementNS(WORD_NAMESPACE, 'w:style');
    style.setAttributeNS(WORD_NAMESPACE, 'w:type', 'paragraph');
    style.setAttributeNS(WORD_NAMESPACE, 'w:styleId', 'EmphasizedHeader');
    const runProperties = stylesDocument.createElementNS(
      WORD_NAMESPACE,
      'w:rPr',
    );
    const emphasis = stylesDocument.createElementNS(WORD_NAMESPACE, 'w:em');
    emphasis.setAttributeNS(WORD_NAMESPACE, 'w:val', 'circle');
    runProperties.append(emphasis);
    style.append(runProperties);
    stylesDocument.documentElement.append(style);
    archive.file(
      'word/styles.xml',
      new XMLSerializer().serializeToString(stylesDocument),
    );

    const imported = await importOfficeFile(
      new File(
        [await archive.generateAsync({ type: 'uint8array' })],
        'styled-header-emphasis.docx',
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
      `${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="circle"`,
    );
    expect(headerHtml).toContain('text-emphasis-style:open circle');
  });

  test('exports and reopens all native values across every editable Word story', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      `<p>${emphasisSpan('dot', '上点')} ${emphasisSpan(
        'comma',
        '逗号',
      )} ${emphasisSpan('circle', '圆圈')} ${emphasisSpan('none', '显式无')}`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="emphasis-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="emphasis-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="emphasis-footnote">',
      `<p>${emphasisSpan('underDot', '脚注下点')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="emphasis-endnote">',
      `<p>${emphasisSpan('comma', '尾注逗号')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${emphasisSpan('circle', '页眉圆圈')}</p>`,
        footerHtml: `<p>${emphasisSpan('none', '页脚显式无')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const sources = await Promise.all(
      Object.keys(archive.files)
        .filter((path) =>
          /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i.test(
            path,
          ),
        )
        .map(async (path) => (await archive.file(path)?.async('text')) ?? ''),
    );
    const xml = sources.join('\n');
    for (const mark of ['none', 'dot', 'comma', 'circle', 'underDot']) {
      expect(xml).toMatch(new RegExp(`<w:em\\b[^>]*w:val="${mark}"`));
    }

    const reopened = await importOfficeFile(
      new File([blob], 'emphasis.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    for (const mark of ['none', 'dot', 'comma', 'circle'] as const) {
      expect(reopened.content.html).toContain(
        `${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="${mark}"`,
      );
    }
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      `${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="circle"`,
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      `${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="none"`,
    );
    const reopenedDocument = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    expect(
      reopenedDocument.querySelector(
        `aside[data-note-kind="footnote"] span[${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="underDot"]`,
      ),
    ).not.toBeNull();
    expect(
      reopenedDocument.querySelector(
        `aside[data-note-kind="endnote"] span[${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="comma"]`,
      ),
    ).not.toBeNull();
  });

  test('imports, exports, accepts, and rejects native emphasis formatting revisions', async () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr>
          <s:em s:val="circle"/>
          <s:rPrChange s:id="9" s:author="Ada Reviewer" s:date="2026-08-23T13:00:00.000Z">
            <s:rPr><s:em s:val="none"/></s:rPr>
          </s:rPrChange>
        </s:rPr>
        <s:t>strict revision</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    const importedMarker = markDocxRunFormatting(strict).runs[0];
    expect(importedMarker?.formatting.emphasisMark).toBe('circle');
    expect(importedMarker?.change?.before).toContain('"emphasisMark":"none"');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = '[{"type":"textStyle","attrs":{"emphasisMark":"none"}}]';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-emphasis" data-change-author="Ada Reviewer" data-change-date="2026-08-23T13:15:00.000Z">${emphasisSpan('circle', 'Changed emphasis')}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed emphasis',
    );
    if (!run) throw new Error('Expected changed emphasis run.');
    const properties = directChild(run, 'rPr');
    const current = directChild(properties ?? run, 'em');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    const previous = directChild(prior ?? run, 'em');
    if (!current || !change || !prior || !previous) {
      throw new Error('Expected current and prior emphasis marks.');
    }
    expect(attribute(current, 'val')).toBe('circle');
    expect(attribute(previous, 'val')).toBe('none');

    const reopened = await importOfficeFile(
      new File([blob], 'emphasis-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      `${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="circle"`,
    );
    expect(reopened.content.html).toContain(
      '&quot;emphasisMark&quot;:&quot;none&quot;',
    );

    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'formatting-emphasis',
          author: 'Reviewer',
          date: '2026-08-23T13:30:00.000Z',
        }),
      }),
      content: `<section data-document-section="true"><p>${emphasisSpan(
        'dot',
        'Tracked emphasis',
      )}</p></section>`,
    });
    editor.commands.setTextSelection(textRange(editor, 'Tracked emphasis'));
    expect(editor.commands.setDocumentEmphasisMark('underDot')).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'formatting-emphasis',
        kind: 'formatting',
        text: 'Tracked emphasis',
      }),
    ]);
    expect(editor.getHTML()).toContain(
      '&quot;emphasisMark&quot;:&quot;dot&quot;',
    );
    expect(editor.commands.rejectDocumentChange('formatting-emphasis')).toBe(
      true,
    );
    expect(editor.getAttributes('textStyle').emphasisMark).toBe('dot');

    editor.commands.setTextSelection(textRange(editor, 'Tracked emphasis'));
    expect(editor.commands.setDocumentEmphasisMark('underDot')).toBe(true);
    expect(editor.commands.acceptDocumentChange('formatting-emphasis')).toBe(
      true,
    );
    expect(editor.getAttributes('textStyle').emphasisMark).toBe('underDot');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function emphasisSpan(mark: WorkDocumentEmphasisMark, content: string): string {
  const attributes = documentEmphasisMarkDomAttributes(mark);
  return `<span ${DOCUMENT_EMPHASIS_MARK_ATTRIBUTE}="${attributes[DOCUMENT_EMPHASIS_MARK_ATTRIBUTE]}" style="${attributes.style}">${content}</span>`;
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
