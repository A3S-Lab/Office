import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import {
  DOCUMENT_CHARACTER_SPACING_ATTRIBUTE,
  DOCUMENT_CHARACTER_SPACING_MAX_TWIPS,
  documentCharacterSpacingDomAttributes,
  normalizeDocumentCharacterSpacingTwips,
} from '../src/internal/features/work/work-document-character-spacing';
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

describe('document character spacing', () => {
  test('normalizes signed twips and projects exact native metadata to CSS', () => {
    expect(normalizeDocumentCharacterSpacingTwips(undefined)).toBeNull();
    expect(normalizeDocumentCharacterSpacingTwips('+20')).toBe(20);
    expect(normalizeDocumentCharacterSpacingTwips('-30')).toBe(-30);
    expect(normalizeDocumentCharacterSpacingTwips(0)).toBe(0);
    expect(
      normalizeDocumentCharacterSpacingTwips(
        DOCUMENT_CHARACTER_SPACING_MAX_TWIPS,
      ),
    ).toBe(DOCUMENT_CHARACTER_SPACING_MAX_TWIPS);
    expect(
      normalizeDocumentCharacterSpacingTwips(
        DOCUMENT_CHARACTER_SPACING_MAX_TWIPS + 1,
      ),
    ).toBeNull();
    expect(normalizeDocumentCharacterSpacingTwips(1.5)).toBeNull();

    expect(documentCharacterSpacingDomAttributes(-30)).toEqual({
      [DOCUMENT_CHARACTER_SPACING_ATTRIBUTE]: '-30',
      style: 'letter-spacing: -1.5pt',
    });
    expect(documentCharacterSpacingDomAttributes(0)).toEqual({
      [DOCUMENT_CHARACTER_SPACING_ATTRIBUTE]: '0',
      style: 'letter-spacing: 0pt',
    });
  });

  test('authors expanded, condensed, and explicit normal spacing as one undoable mark change', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Character spacing</p>',
    });
    editor.commands.setTextSelection(textRange(editor, 'Character'));

    expect(editor.commands.setDocumentCharacterSpacing(40)).toBe(true);
    expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(40);
    expect(editor.getHTML()).toContain(
      'data-office-character-spacing-twips="40"',
    );
    expect(editor.getHTML()).toContain('letter-spacing: 2pt');

    expect(editor.commands.setDocumentCharacterSpacing(-20)).toBe(true);
    expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(-20);
    expect(editor.getHTML()).toContain('letter-spacing: -1pt');

    expect(editor.commands.setDocumentCharacterSpacing(0)).toBe(true);
    expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(0);
    expect(editor.getHTML()).toContain(
      'data-office-character-spacing-twips="0"',
    );
    expect(editor.getHTML()).toContain('letter-spacing: 0pt');

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(-20);
  });

  test('resolves style inheritance and direct zero without flattening signed spacing', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Expanded"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:spacing w:val="0"/></w:rPr><w:t>normal</w:t></w:r>
        <w:r><w:rPr><w:spacing w:val="-30"/></w:rPr><w:t>condensed</w:t></w:r>
      </w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:spacing w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="Expanded">
          <w:rPr><w:spacing w:val="40"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(
      markers.runs.map(({ formatting }) => formatting.characterSpacingTwips),
    ).toEqual([40, 0, -30]);

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
        spacing: span.dataset.officeCharacterSpacingTwips,
        style: span.style.letterSpacing,
      })),
    ).toEqual([
      { spacing: '40', style: '2pt' },
      { spacing: '0', style: '0pt' },
      { spacing: '-30', style: '-1.5pt' },
    ]);
  });

  test('accepts strict WordprocessingML and ignores spoofed or out-of-range spacing', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr><s:spacing s:val="-45"/></s:rPr><s:t>strict</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strict).runs[0]?.formatting.characterSpacingTwips,
    ).toBe(-45);

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:spacing w:val="60"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:spacing evil:val="70"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:spacing w:val="31681"/></w:rPr><w:t>too wide</w:t></w:r>
        <w:r><w:rPr><w:spacing w:val="-31681"/></w:rPr><w:t>too tight</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.characterSpacingTwips,
      ),
    ).toEqual([undefined, undefined, undefined, undefined]);
  });

  test('exports and reopens exact spacing in body, header, and footer', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${spacingSpan(40, 'expanded')} ${spacingSpan(
      -30,
      'condensed',
    )} ${spacingSpan(0, 'normal')}</p>`;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${spacingSpan(25, 'header-spacing')}</p>`,
        footerHtml: `<p>${spacingSpan(-15, 'footer-spacing')}</p>`,
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
    for (const spacing of [40, -30, 0, 25, -15]) {
      expect(xml).toMatch(new RegExp(`<w:spacing\\b[^>]*w:val="${spacing}"`));
    }
    expect(xml).not.toContain('31681');

    const imported = await importOfficeFile(
      new File([blob], 'character-spacing.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    for (const spacing of [40, -30, 0]) {
      expect(imported.content.html).toContain(
        `data-office-character-spacing-twips="${spacing}"`,
      );
    }
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      'data-office-character-spacing-twips="25"',
    );
    expect(imported.content.pageChrome?.default.footerHtml).toContain(
      'data-office-character-spacing-twips="-15"',
    );
  });

  test('exports and reopens native character-spacing revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = '[{"type":"textStyle","attrs":{"characterSpacingTwips":0}}]';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-spacing" data-change-author="Ada Reviewer" data-change-date="2026-08-23T08:00:00.000Z">${spacingSpan(40, 'Changed spacing')}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed spacing',
    );
    if (!run) throw new Error('Expected changed character-spacing run.');
    const properties = directChild(run, 'rPr');
    const currentSpacing = directChild(properties ?? run, 'spacing');
    if (!currentSpacing) throw new Error('Expected current character spacing.');
    expect(attribute(currentSpacing, 'val')).toBe('40');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    const priorSpacing = directChild(prior ?? run, 'spacing');
    if (!priorSpacing) throw new Error('Expected prior character spacing.');
    expect(attribute(priorSpacing, 'val')).toBe('0');

    const reopened = await importOfficeFile(
      new File([blob], 'character-spacing-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      'data-office-character-spacing-twips="40"',
    );
    expect(reopened.content.html).toContain(
      '&quot;characterSpacingTwips&quot;:0',
    );
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function spacingSpan(spacing: number, text: string): string {
  const attributes = documentCharacterSpacingDomAttributes(spacing);
  return `<span ${DOCUMENT_CHARACTER_SPACING_ATTRIBUTE}="${attributes[DOCUMENT_CHARACTER_SPACING_ATTRIBUTE]}" style="${attributes.style}">${text}</span>`;
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
