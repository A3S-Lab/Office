import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { applyDocumentFontDialogPatch } from '../src/internal/features/work/editors/document-font-dialog-model';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE,
  DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE,
  DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE,
  DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE,
  documentLegacyTextEffectsConflict,
  documentLegacyTextEffectsDomAttributes,
  documentLegacyTextEffectsFromTextStyleAttributes,
  normalizeDocumentLegacyTextEffect,
} from '../src/internal/features/work/work-document-legacy-text-effects';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  DocxLegacyTextEffectsPatchCollector,
  patchDocxLegacyTextEffects,
} from '../src/internal/features/work/work-docx-legacy-text-effects-export';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
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

describe('document legacy text effects', () => {
  test('models four nullable effects while retaining the valid outline-shadow pair', () => {
    expect(normalizeDocumentLegacyTextEffect(true)).toBe(true);
    expect(normalizeDocumentLegacyTextEffect(false)).toBe(false);
    expect(normalizeDocumentLegacyTextEffect('true')).toBe(true);
    expect(normalizeDocumentLegacyTextEffect('false')).toBe(false);
    expect(normalizeDocumentLegacyTextEffect('1')).toBe(true);
    expect(normalizeDocumentLegacyTextEffect('0')).toBe(false);
    for (const invalid of [undefined, null, '', 'TRUE', 'on', 1, 0, [], {}]) {
      expect(normalizeDocumentLegacyTextEffect(invalid)).toBeNull();
    }

    expect(
      documentLegacyTextEffectsDomAttributes({
        outline: true,
        shadow: false,
        emboss: false,
        imprint: false,
      }),
    ).toEqual({
      [DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE]: 'true',
      [DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE]: 'false',
      [DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE]: 'false',
      [DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE]: 'false',
    });
    expect(
      documentLegacyTextEffectsConflict({ outline: true, shadow: true }),
    ).toBe(false);
    expect(
      documentLegacyTextEffectsFromTextStyleAttributes({
        legacyTextOutline: 'invalid',
      }),
    ).toBeNull();
    for (const conflict of [
      { outline: true, emboss: true },
      { outline: true, imprint: true },
      { shadow: true, emboss: true },
      { shadow: true, imprint: true },
      { emboss: true, imprint: true },
    ]) {
      expect(documentLegacyTextEffectsConflict(conflict)).toBe(true);
    }

    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Legacy effects</p>',
    });
    editor.commands.selectAll();
    expect(
      editor.commands.setMark('textStyle', {
        legacyTextOutline: true,
        legacyTextShadow: true,
        legacyTextEmboss: false,
        legacyTextImprint: false,
      }),
    ).toBe(true);
    expect(editor.getAttributes('textStyle')).toMatchObject({
      legacyTextOutline: true,
      legacyTextShadow: true,
      legacyTextEmboss: false,
      legacyTextImprint: false,
    });
    expect(editor.getHTML()).toContain(
      `${DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE}="true"`,
    );
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).not.toContain(
      DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE,
    );
  });

  test('resolves defaults and styles while preserving direct false resets', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="ReliefStyle"/></w:pPr>
        <w:r><w:t>styled relief</w:t></w:r>
        <w:r><w:rPr><w:emboss w:val="0"/><w:imprint/></w:rPr><w:t>direct imprint</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>default outline shadow</w:t></w:r></w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:outline/><w:shadow/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="ReliefStyle">
          <w:rPr><w:outline w:val="0"/><w:shadow w:val="0"/><w:emboss/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(markers.runs.map(({ formatting }) => formatting)).toEqual([
      expect.objectContaining({
        legacyTextOutline: false,
        legacyTextShadow: false,
        legacyTextEmboss: true,
      }),
      expect.objectContaining({
        legacyTextOutline: false,
        legacyTextShadow: false,
        legacyTextEmboss: false,
        legacyTextImprint: true,
      }),
      expect.objectContaining({
        legacyTextOutline: true,
        legacyTextShadow: true,
      }),
    ]);

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
          `span[${DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE}]`,
        ),
      ).map((span) => span.getAttribute(DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE)),
    ).toEqual(['true', 'false']);
  });

  test('accepts strict on-off leaves and ignores malformed or conflicting properties', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p>
        <s:r><s:rPr><s:outline/><s:shadow s:val="on"/><s:emboss s:val="off"/><s:imprint s:val="0"/></s:rPr><s:t>strict effects</s:t></s:r>
      </s:p></s:body></s:document>
    `);
    expect(markDocxRunFormatting(strict).runs[0]?.formatting).toMatchObject({
      legacyTextOutline: true,
      legacyTextShadow: true,
      legacyTextEmboss: false,
      legacyTextImprint: false,
    });

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:outline/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:shadow evil:val="1"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:emboss w:val="yes"/></w:rPr><w:t>invalid token</w:t></w:r>
        <w:r><w:rPr><w:imprint/><w:imprint w:val="0"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:outline><w:b/></w:outline></w:rPr><w:t>nested</w:t></w:r>
        <w:r><w:rPr><w:shadow>text</w:shadow></w:rPr><w:t>text child</w:t></w:r>
        <w:r><w:rPr><w:emboss w:extra="1"/></w:rPr><w:t>extra attribute</w:t></w:r>
        <w:r><w:rPr><w:outline/><w:emboss/></w:rPr><w:t>conflict</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(({ formatting }) => [
        formatting.legacyTextOutline,
        formatting.legacyTextShadow,
        formatting.legacyTextEmboss,
        formatting.legacyTextImprint,
      ]),
    ).toEqual(
      Array.from({ length: 8 }, () => [
        undefined,
        undefined,
        undefined,
        undefined,
      ]),
    );
  });

  test('restores nested style markers and writes native properties in schema order', async () => {
    const collector = new DocxLegacyTextEffectsPatchCollector('existing');
    const first = collector.marker(
      { outline: true, shadow: true, emboss: false, imprint: false },
      'Emphasis',
    );
    const second = collector.marker(
      { outline: false, shadow: false, emboss: true, imprint: false },
      first,
    );
    expect(collector.patches[1]).toMatchObject({
      marker: second,
      style: 'Emphasis',
      effects: { outline: false, shadow: false, emboss: true, imprint: false },
    });

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(
        wordXml(`
          <w:p><w:r><w:rPr><w:rStyle w:val="${second}"/></w:rPr><w:t>Embossed</w:t></w:r></w:p>
        `),
      ),
    );
    const patched = await patchDocxLegacyTextEffects(
      await archive.generateAsync({ type: 'arraybuffer' }),
      collector.patches.slice(1),
    );
    const reopened = await JSZip.loadAsync(patched);
    const document = parseXml(
      (await reopened.file('word/document.xml')?.async('text')) ?? '',
    );
    const properties = directChild(descendants(document, 'r')[0], 'rPr');
    if (!properties) throw new Error('Expected run properties.');
    expect(
      directChildren(properties).map(({ localName }) => localName),
    ).toEqual(['rStyle', 'outline', 'shadow', 'emboss', 'imprint']);
    expect(
      attribute(directChild(properties, 'rStyle') ?? properties, 'val'),
    ).toBe('Emphasis');
    expect(
      attribute(directChild(properties, 'outline') ?? properties, 'val'),
    ).toBe('0');
    expect(
      attribute(directChild(properties, 'emboss') ?? properties, 'val'),
    ).toBeNull();
  });

  test('exports and reopens explicit values across editable Word stories and formatting revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before =
      '[{"type":"textStyle","attrs":{"legacyTextEmboss":false,"legacyTextImprint":false,"legacyTextOutline":true,"legacyTextShadow":true}}]';
    artifact.content.html = [
      `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="legacy-effects" data-change-author="Ada Reviewer" data-change-date="2026-08-24T09:00:00.000Z">${effectSpan(
        { outline: false, shadow: false, emboss: true, imprint: false },
        'body emboss',
      )}</span>`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="legacy-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="legacy-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="legacy-footnote">',
      `<p>${effectSpan({ outline: true, shadow: true }, 'footnote outline shadow')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="legacy-endnote">',
      `<p>${effectSpan({ imprint: true }, 'endnote imprint')}</p></aside>`,
    ].join('');
    artifact.content.trackChanges = true;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${effectSpan({ outline: true }, 'header outline')}</p>`,
        footerHtml: `<p>${effectSpan({ shadow: true }, 'footer shadow')}</p>`,
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
    for (const property of ['outline', 'shadow', 'emboss', 'imprint']) {
      expect(storyXml).toContain(`<w:${property}`);
    }
    expect(storyXml).not.toContain('A3SOfficeLegacyTextEffects');
    expect(storyXml).toMatch(
      /<w:rPrChange\b[^>]*>[\s\S]*?<w:outline\/>[\s\S]*?<w:shadow\/>[\s\S]*?<w:emboss\b[^>]*w:val="0"[\s\S]*?<w:imprint\b[^>]*w:val="0"/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'legacy-text-effects.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(reopened.content.html).toContain(
      `${DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE}="true"`,
    );
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      `${DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE}="true"`,
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      `${DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE}="true"`,
    );
    expect(reopened.content.html).toContain(
      `${DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE}="true"`,
    );
  });

  test('tracks, rejects, accepts, and undoes one conflict-safe effect transaction', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'legacy-text-effects-change',
          author: 'Reviewer',
          date: '2026-08-24T09:30:00.000Z',
        }),
      }),
      content: `<section data-document-section="true"><p>${effectSpan(
        { outline: true, shadow: true, emboss: false, imprint: false },
        'Tracked effects',
      )}</p></section>`,
    });
    const selection = textRange(editor, 'Tracked effects');
    editor.commands.setTextSelection(selection);
    expect(
      applyDocumentFontDialogPatch(editor, selection, {
        legacyTextEmboss: true,
      }),
    ).toBe(false);
    expect(
      applyDocumentFontDialogPatch(editor, selection, {
        legacyTextOutline: false,
        legacyTextShadow: false,
        legacyTextEmboss: true,
        legacyTextImprint: false,
      }),
    ).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'legacy-text-effects-change',
        kind: 'formatting',
        text: 'Tracked effects',
      }),
    ]);
    expect(editor.getHTML()).toContain('&quot;legacyTextOutline&quot;:true');
    expect(
      editor.commands.rejectDocumentChange('legacy-text-effects-change'),
    ).toBe(true);
    editor.commands.setTextSelection(selection);
    expect(editor.getAttributes('textStyle')).toMatchObject({
      legacyTextOutline: true,
      legacyTextShadow: true,
      legacyTextEmboss: false,
      legacyTextImprint: false,
    });

    expect(
      applyDocumentFontDialogPatch(editor, selection, {
        legacyTextOutline: false,
        legacyTextShadow: false,
        legacyTextEmboss: true,
        legacyTextImprint: false,
      }),
    ).toBe(true);
    expect(
      editor.commands.acceptDocumentChange('legacy-text-effects-change'),
    ).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    editor.commands.setTextSelection(selection);
    expect(editor.getAttributes('textStyle')).toMatchObject({
      legacyTextOutline: false,
      legacyTextShadow: false,
      legacyTextEmboss: true,
      legacyTextImprint: false,
    });
    expect(editor.commands.undo()).toBe(true);
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function effectSpan(
  effects: Partial<
    Record<'outline' | 'shadow' | 'emboss' | 'imprint', boolean>
  >,
  content: string,
): string {
  const attributes = documentLegacyTextEffectsDomAttributes(effects);
  return `<span ${Object.entries(attributes)
    .map(([name, value]) => `${name}="${value}"`)
    .join(' ')}>${content}</span>`;
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
