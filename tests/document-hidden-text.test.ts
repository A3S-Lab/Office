import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DOCUMENT_HIDDEN_TEXT_ATTRIBUTE,
  documentHiddenTextDomAttributes,
  normalizeDocumentHiddenText,
} from '../src/internal/features/work/work-document-hidden-text';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  DocxHiddenTextPatchCollector,
  patchDocxHiddenText,
} from '../src/internal/features/work/work-docx-hidden-text-export';
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

describe('document hidden text', () => {
  test('models inherit, hidden, and explicit-visible states with the native shortcut and one-step Undo', () => {
    expect(normalizeDocumentHiddenText(true)).toBe(true);
    expect(normalizeDocumentHiddenText(false)).toBe(false);
    expect(normalizeDocumentHiddenText('true')).toBe(true);
    expect(normalizeDocumentHiddenText('false')).toBe(false);
    expect(normalizeDocumentHiddenText('1')).toBe(true);
    expect(normalizeDocumentHiddenText('0')).toBe(false);
    for (const invalid of [undefined, null, '', 'TRUE', 'yes', 1, 0, [], {}]) {
      expect(normalizeDocumentHiddenText(invalid)).toBeNull();
    }
    expect(documentHiddenTextDomAttributes(true)).toEqual({
      [DOCUMENT_HIDDEN_TEXT_ATTRIBUTE]: 'true',
    });
    expect(documentHiddenTextDomAttributes(false)).toEqual({
      [DOCUMENT_HIDDEN_TEXT_ATTRIBUTE]: 'false',
    });
    expect(documentHiddenTextDomAttributes(null)).toEqual({});

    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Hidden text fidelity</p>',
    });
    editor.commands.selectAll();

    expect(editor.commands.setDocumentHiddenText(true)).toBe(true);
    expect(editor.getAttributes('textStyle').hiddenText).toBe(true);
    expect(editor.getHTML()).toContain(
      `${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="true"`,
    );

    expect(editor.commands.keyboardShortcut('Mod-Shift-h')).toBe(true);
    expect(editor.getAttributes('textStyle').hiddenText).toBe(false);
    expect(editor.getHTML()).toContain(
      `${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="false"`,
    );

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').hiddenText).toBe(true);
    expect(editor.commands.unsetDocumentHiddenText()).toBe(true);
    expect(editor.getHTML()).not.toContain(DOCUMENT_HIDDEN_TEXT_ATTRIBUTE);
  });

  test('resolves document defaults and paragraph styles while preserving direct resets', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="HiddenStyle"/></w:pPr>
        <w:r><w:t>styled hidden</w:t></w:r>
        <w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>direct visible</w:t></w:r>
      </w:p>
      <w:p><w:r><w:t>default hidden</w:t></w:r></w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:vanish/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="HiddenStyle">
          <w:rPr><w:vanish w:val="false"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(markers.runs.map(({ formatting }) => formatting.hiddenText)).toEqual(
      [false, false, true],
    );

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
          `span[${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}]`,
        ),
      ).map((span) => span.getAttribute(DOCUMENT_HIDDEN_TEXT_ATTRIBUTE)),
    ).toEqual(['false', 'false', 'true']);
  });

  test('accepts strict on/off values and ignores malformed, duplicate, or namespace-spoofed properties', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p>
        <s:r><s:rPr><s:vanish/></s:rPr><s:t>empty true</s:t></s:r>
        <s:r><s:rPr><s:vanish s:val="on"/></s:rPr><s:t>on true</s:t></s:r>
        <s:r><s:rPr><s:vanish s:val="off"/></s:rPr><s:t>off false</s:t></s:r>
      </s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strict).runs.map(
        ({ formatting }) => formatting.hiddenText,
      ),
    ).toEqual([true, true, false]);

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:vanish/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:vanish evil:val="1"/></w:rPr><w:t>spoofed value</w:t></w:r>
        <w:r><w:rPr><w:vanish w:val="yes"/></w:rPr><w:t>invalid token</w:t></w:r>
        <w:r><w:rPr><w:vanish/><w:vanish w:val="0"/></w:rPr><w:t>duplicate</w:t></w:r>
        <w:r><w:rPr><w:vanish><w:b/></w:vanish></w:rPr><w:t>nested</w:t></w:r>
        <w:r><w:rPr><w:vanish>text</w:vanish></w:rPr><w:t>text child</w:t></w:r>
        <w:r><w:rPr><w:vanish w:extra="1"/></w:rPr><w:t>extra attribute</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(
        ({ formatting }) => formatting.hiddenText,
      ),
    ).toEqual(Array.from({ length: 7 }, () => undefined));
  });

  test('restores an existing character style after applying nested export markers', async () => {
    const collector = new DocxHiddenTextPatchCollector('existing content');
    const visibleMarker = collector.marker(false, 'Emphasis');
    const hiddenMarker = collector.marker(true, visibleMarker);
    expect(collector.patches).toEqual([
      { marker: visibleMarker, style: 'Emphasis', value: false },
      { marker: hiddenMarker, style: 'Emphasis', value: true },
    ]);

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(
        wordXml(`
          <w:p>
            <w:r><w:rPr><w:rStyle w:val="${visibleMarker}"/></w:rPr><w:t>Visible reset</w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="${hiddenMarker}"/></w:rPr><w:t>Hidden</w:t></w:r>
          </w:p>
        `),
      ),
    );
    const patched = await patchDocxHiddenText(
      await archive.generateAsync({ type: 'arraybuffer' }),
      collector.patches,
    );
    const reopened = await JSZip.loadAsync(patched);
    const document = parseXml(
      (await reopened.file('word/document.xml')?.async('text')) ?? '',
    );
    const runs = descendants(document, 'r');
    expect(
      runs.map((run) =>
        attribute(
          directChild(directChild(run, 'rPr') ?? run, 'rStyle') ?? run,
          'val',
        ),
      ),
    ).toEqual(['Emphasis', 'Emphasis']);
    expect(
      runs.map((run) =>
        attribute(
          directChild(directChild(run, 'rPr') ?? run, 'vanish') ?? run,
          'val',
        ),
      ),
    ).toEqual(['0', null]);
    expect(
      runs.every(
        (run) =>
          directChildren(directChild(run, 'rPr') ?? run).findIndex(
            ({ localName }) => localName === 'rStyle',
          ) <
          directChildren(directChild(run, 'rPr') ?? run).findIndex(
            ({ localName }) => localName === 'vanish',
          ),
      ),
    ).toBe(true);
  });

  test('exports and reopens explicit true and false across every editable Word story', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      `<p>${hiddenSpan(true, 'body hidden')} ${hiddenSpan(
        false,
        'body visible reset',
      )}`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="hidden-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="hidden-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="hidden-footnote">',
      `<p>${hiddenSpan(true, 'footnote hidden')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="hidden-endnote">',
      `<p>${hiddenSpan(false, 'endnote visible reset')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${hiddenSpan(true, 'header hidden')}</p>`,
        footerHtml: `<p>${hiddenSpan(false, 'footer visible reset')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const storySources = await Promise.all(
      Object.keys(archive.files)
        .filter((path) =>
          /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i.test(
            path,
          ),
        )
        .map(async (path) => (await archive.file(path)?.async('text')) ?? ''),
    );
    const xml = storySources.join('\n');
    expect(xml).toMatch(/<w:vanish\/?\s*>/);
    expect(xml).toMatch(/<w:vanish\b[^>]*w:val="0"\s*\/?\s*>/);
    expect(xml).not.toContain('A3SOfficeHiddenText');

    const reopened = await importOfficeFile(
      new File([blob], 'hidden-text.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(reopened.content.html).toContain(
      `${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="true"`,
    );
    expect(reopened.content.html).toContain(
      `${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="false"`,
    );
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      `${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="true"`,
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      `${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="false"`,
    );
    const reopenedDocument = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    expect(
      reopenedDocument.querySelector(
        `aside[data-note-kind="footnote"] span[${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="true"]`,
      ),
    ).not.toBeNull();
    expect(
      reopenedDocument.querySelector(
        `aside[data-note-kind="endnote"] span[${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="false"]`,
      ),
    ).not.toBeNull();
  });

  test('imports, exports, accepts, and rejects native hidden-text formatting revisions', async () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p><s:r>
        <s:rPr>
          <s:vanish/>
          <s:rPrChange s:id="9" s:author="Ada Reviewer" s:date="2026-08-24T03:00:00.000Z">
            <s:rPr><s:vanish s:val="0"/></s:rPr>
          </s:rPrChange>
        </s:rPr>
        <s:t>strict revision</s:t>
      </s:r></s:p></s:body></s:document>
    `);
    const importedMarker = markDocxRunFormatting(strict).runs[0];
    expect(importedMarker?.formatting.hiddenText).toBe(true);
    expect(importedMarker?.change?.before).toContain('"hiddenText":false');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = '[{"type":"textStyle","attrs":{"hiddenText":false}}]';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-hidden" data-change-author="Ada Reviewer" data-change-date="2026-08-24T03:15:00.000Z">${hiddenSpan(
      true,
      'Changed hidden text',
    )}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed hidden text',
    );
    if (!run) throw new Error('Expected changed hidden-text run.');
    const properties = directChild(run, 'rPr');
    const current = directChild(properties ?? run, 'vanish');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    const previous = directChild(prior ?? run, 'vanish');
    if (!current || !change || !prior || !previous) {
      throw new Error('Expected current and prior hidden-text values.');
    }
    expect(attribute(current, 'val')).toBeNull();
    expect(attribute(previous, 'val')).toBe('0');

    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'formatting-hidden',
          author: 'Reviewer',
          date: '2026-08-24T03:30:00.000Z',
        }),
      }),
      content: `<section data-document-section="true"><p>${hiddenSpan(
        false,
        'Tracked hidden text',
      )}</p></section>`,
    });
    editor.commands.setTextSelection(textRange(editor, 'Tracked hidden text'));
    expect(editor.commands.setDocumentHiddenText(true)).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'formatting-hidden',
        kind: 'formatting',
        text: 'Tracked hidden text',
      }),
    ]);
    expect(editor.getHTML()).toContain('&quot;hiddenText&quot;:false');
    expect(editor.commands.rejectDocumentChange('formatting-hidden')).toBe(
      true,
    );
    expect(editor.getAttributes('textStyle').hiddenText).toBe(false);

    editor.commands.setTextSelection(textRange(editor, 'Tracked hidden text'));
    expect(editor.commands.setDocumentHiddenText(true)).toBe(true);
    expect(editor.commands.acceptDocumentChange('formatting-hidden')).toBe(
      true,
    );
    expect(editor.getAttributes('textStyle').hiddenText).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function hiddenSpan(hiddenText: boolean, content: string): string {
  return `<span ${DOCUMENT_HIDDEN_TEXT_ATTRIBUTE}="${String(
    hiddenText,
  )}">${content}</span>`;
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
