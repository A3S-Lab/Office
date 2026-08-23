import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import {
  DOCUMENT_CHARACTER_POSITION_ATTRIBUTE,
  DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY,
  DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS,
  documentCharacterPositionDomAttributes,
  normalizeDocumentCharacterPositionHalfPoints,
} from '../src/internal/features/work/work-document-character-position';
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

describe('document character position', () => {
  test('normalizes signed half-points and projects exact native metadata', () => {
    expect(normalizeDocumentCharacterPositionHalfPoints(undefined)).toBeNull();
    expect(normalizeDocumentCharacterPositionHalfPoints('+2')).toBe(2);
    expect(normalizeDocumentCharacterPositionHalfPoints('-3')).toBe(-3);
    expect(normalizeDocumentCharacterPositionHalfPoints(0)).toBe(0);
    expect(
      normalizeDocumentCharacterPositionHalfPoints(
        DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS,
      ),
    ).toBe(DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS);
    expect(
      normalizeDocumentCharacterPositionHalfPoints(
        DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS + 1,
      ),
    ).toBeNull();
    expect(normalizeDocumentCharacterPositionHalfPoints(1.5)).toBeNull();

    expect(documentCharacterPositionDomAttributes(-3)).toEqual({
      [DOCUMENT_CHARACTER_POSITION_ATTRIBUTE]: '-3',
      style: `${DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY}: -1.5pt`,
    });
    expect(documentCharacterPositionDomAttributes(0)).toEqual({
      [DOCUMENT_CHARACTER_POSITION_ATTRIBUTE]: '0',
      style: `${DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY}: 0pt`,
    });
  });

  test('authors raised, lowered, and explicit normal positions as one undoable mark', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Character position</p>',
    });
    editor.commands.setTextSelection(textRange(editor, 'Character'));

    expect(editor.commands.setDocumentCharacterPosition(4)).toBe(true);
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      4,
    );
    expect(editor.getHTML()).toContain(
      'data-office-character-position-half-points="4"',
    );
    expect(editor.getHTML()).toContain(
      '--work-document-character-position: 2pt',
    );

    expect(editor.commands.setDocumentCharacterPosition(-3)).toBe(true);
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      -3,
    );
    expect(editor.getHTML()).toContain(
      '--work-document-character-position: -1.5pt',
    );

    expect(editor.commands.setDocumentCharacterPosition(0)).toBe(true);
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      0,
    );
    expect(editor.getHTML()).toContain(
      'data-office-character-position-half-points="0"',
    );

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      -3,
    );
  });

  test('keeps native position while subscript and superscript take visual precedence', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Position</p>',
    });
    editor.commands.setTextSelection(textRange(editor, 'Position'));
    expect(editor.commands.setDocumentCharacterPosition(4)).toBe(true);
    expect(editor.commands.toggleDocumentSuperscript()).toBe(true);
    expect(editor.isActive('superscript')).toBe(true);
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      4,
    );

    expect(editor.commands.toggleDocumentSuperscript()).toBe(true);
    expect(editor.isActive('superscript')).toBe(false);
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      4,
    );
  });

  test('resolves style inheritance and direct zero without flattening position', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Raised"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:position w:val="0"/></w:rPr><w:t>normal</w:t></w:r>
        <w:r><w:rPr><w:position w:val="-3"/></w:rPr><w:t>lowered</w:t></w:r>
      </w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:position w:val="2"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Raised">
          <w:rPr><w:position w:val="4"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(
      markers.runs.map(
        ({ formatting }) => formatting.characterPositionHalfPoints,
      ),
    ).toEqual([4, 0, -3]);

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
        position: span.dataset.officeCharacterPositionHalfPoints,
        style: span.style.getPropertyValue(
          DOCUMENT_CHARACTER_POSITION_CSS_PROPERTY,
        ),
      })),
    ).toEqual([
      { position: '4', style: '2pt' },
      { position: '0', style: '0pt' },
      { position: '-3', style: '-1.5pt' },
    ]);
  });

  test('accepts exact strict measures and rejects spoofed or unsafe position', () => {
    for (const [source, expected] of [
      ['6.35mm', 36],
      ['1.27cm', 72],
      ['0.125in', 18],
      ['9pt', 18],
      ['0.75pc', 18],
      ['-0.5pt', -1],
      ['0pt', 0],
    ] as const) {
      const strict = parseXml(`
        <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
          <s:rPr><s:position s:val="${source}"/></s:rPr><s:t>${source}</s:t>
        </s:r></s:p></s:body></s:document>
      `);
      expect(
        markDocxRunFormatting(strict).runs[0]?.formatting
          .characterPositionHalfPoints,
      ).toBe(expected);
    }

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:position w:val="2"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:position evil:val="2"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:position w:val="3169"/></w:rPr><w:t>too high</w:t></w:r>
        <w:r><w:rPr><w:position w:val="-3169"/></w:rPr><w:t>too low</w:t></w:r>
        <w:r><w:rPr><w:position w:val="1.5"/></w:rPr><w:t>fractional</w:t></w:r>
        <w:r><w:rPr><w:position w:val="2"/><w:position w:val="4"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:position w:val="2"><w:b/></w:position></w:rPr><w:t>nested</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.characterPositionHalfPoints,
      ),
    ).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);

    const inexactStrict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr><s:position s:val="0.1pt"/></s:rPr><s:t>inexact</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(inexactStrict).runs[0]?.formatting
        .characterPositionHalfPoints,
    ).toBeUndefined();
  });

  test('imports strict native character-position revisions', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr>
          <s:position s:val="-2"/>
          <s:rPrChange s:id="7" s:author="Ada Reviewer" s:date="2026-08-23T08:15:00.000Z">
            <s:rPr><s:position s:val="2pt"/></s:rPr>
          </s:rPrChange>
        </s:rPr>
        <s:t>strict revision</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    const marker = markDocxRunFormatting(strict).runs[0];
    expect(marker?.formatting.characterPositionHalfPoints).toBe(-2);
    expect(marker?.change).toMatchObject({
      author: 'Ada Reviewer',
      date: '2026-08-23T08:15:00.000Z',
      id: 'docx-format-change-7',
    });
    expect(marker?.change?.before).toContain('"characterPositionHalfPoints":4');
  });

  test('exports and reopens exact position in body, header, and footer', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${positionSpan(4, 'raised')} ${positionSpan(
      -3,
      'lowered',
    )} ${positionSpan(0, 'normal')} ${positionSpan(
      2,
      '<sup>raised-super</sup>',
    )}</p>`;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${positionSpan(3, 'header-position')}</p>`,
        footerHtml: `<p>${positionSpan(-2, 'footer-position')}</p>`,
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
    for (const position of [4, -3, 0, 2, 3, -2]) {
      expect(xml).toMatch(new RegExp(`<w:position\\b[^>]*w:val="${position}"`));
    }
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const combinedRun = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'raised-super',
    );
    if (!combinedRun) throw new Error('Expected combined position run.');
    const combinedProperties = directChild(combinedRun, 'rPr');
    const combinedPosition = directChild(
      combinedProperties ?? combinedRun,
      'position',
    );
    const combinedVerticalAlign = directChild(
      combinedProperties ?? combinedRun,
      'vertAlign',
    );
    if (!combinedPosition || !combinedVerticalAlign) {
      throw new Error('Expected native position and superscript properties.');
    }
    expect(attribute(combinedPosition, 'val')).toBe('2');
    expect(attribute(combinedVerticalAlign, 'val')).toBe('superscript');

    const imported = await importOfficeFile(
      new File([blob], 'character-position.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    for (const position of [4, -3, 0, 2]) {
      expect(imported.content.html).toContain(
        `data-office-character-position-half-points="${position}"`,
      );
    }
    expect(imported.content.html).toContain('<sup>raised-super</sup>');
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      'data-office-character-position-half-points="3"',
    );
    expect(imported.content.pageChrome?.default.footerHtml).toContain(
      'data-office-character-position-half-points="-2"',
    );
  });

  test('exports and reopens native character-position revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before =
      '[{"type":"textStyle","attrs":{"characterPositionHalfPoints":0}}]';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-position" data-change-author="Ada Reviewer" data-change-date="2026-08-23T08:30:00.000Z">${positionSpan(4, 'Changed position')}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed position',
    );
    if (!run) throw new Error('Expected changed character-position run.');
    const properties = directChild(run, 'rPr');
    const currentPosition = directChild(properties ?? run, 'position');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    const priorPosition = directChild(prior ?? run, 'position');
    if (!currentPosition || !change || !prior || !priorPosition) {
      throw new Error('Expected current and prior character position.');
    }
    expect(attribute(currentPosition, 'val')).toBe('4');
    expect(attribute(priorPosition, 'val')).toBe('0');

    const reopened = await importOfficeFile(
      new File([blob], 'character-position-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      'data-office-character-position-half-points="4"',
    );
    expect(reopened.content.html).toContain(
      '&quot;characterPositionHalfPoints&quot;:0',
    );
  });

  test('tracks and restores exact native character position', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'formatting-position',
          author: 'Reviewer',
          date: '2026-08-23T09:30:00.000Z',
        }),
      }),
      content: `<section data-document-section="true"><p>${positionSpan(
        3,
        'Tracked position',
      )}</p></section>`,
    });
    editor.commands.setTextSelection(textRange(editor, 'Tracked position'));

    expect(editor.commands.setDocumentCharacterPosition(-4)).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'formatting-position',
        kind: 'formatting',
        text: 'Tracked position',
      }),
    ]);
    expect(editor.getHTML()).toContain(
      '&quot;characterPositionHalfPoints&quot;:3',
    );

    expect(editor.commands.rejectDocumentChange('formatting-position')).toBe(
      true,
    );
    expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
      3,
    );
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function positionSpan(position: number, content: string): string {
  const attributes = documentCharacterPositionDomAttributes(position);
  return `<span ${DOCUMENT_CHARACTER_POSITION_ATTRIBUTE}="${attributes[DOCUMENT_CHARACTER_POSITION_ATTRIBUTE]}" style="${attributes.style}">${content}</span>`;
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
