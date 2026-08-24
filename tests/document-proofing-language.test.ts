import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import {
  DOCUMENT_NO_PROOF_ATTRIBUTE,
  DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE,
  documentProofingDomAttributes,
  normalizeDocumentLanguageTag,
  normalizeDocumentProofingLanguages,
  parseDocumentProofingLanguages,
} from '../src/internal/features/work/work-document-proofing';
import { DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE } from '../src/internal/features/work/work-document-script-fonts';
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

describe('document proofing language', () => {
  test('normalizes bounded native language slots and projects semantic DOM attributes', () => {
    expect(normalizeDocumentLanguageTag('en-US')).toBe('en-US');
    expect(normalizeDocumentLanguageTag('zh-Hans-CN')).toBe('zh-Hans-CN');
    expect(normalizeDocumentLanguageTag('x-none')).toBe('x-none');
    for (const invalid of [
      '',
      ' en-US',
      'en_US',
      'en--US',
      'toolonglanguage-US',
      'en-US<script>',
      1,
      null,
      undefined,
    ]) {
      expect(normalizeDocumentLanguageTag(invalid)).toBeNull();
    }

    const languages = {
      latin: 'en-US',
      eastAsia: 'zh-CN',
      bidi: 'ar-SA',
    } as const;
    expect(normalizeDocumentProofingLanguages(languages)).toEqual(languages);
    expect(parseDocumentProofingLanguages(JSON.stringify(languages))).toEqual(
      languages,
    );
    expect(documentProofingDomAttributes(languages, true, 'eastAsia')).toEqual({
      [DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE]: JSON.stringify(languages),
      [DOCUMENT_NO_PROOF_ATTRIBUTE]: 'true',
      [DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE]: 'eastAsia',
      lang: 'zh-CN',
      spellcheck: 'false',
    });
    expect(
      documentProofingDomAttributes(languages, false, 'complexScript'),
    ).toEqual({
      [DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE]: JSON.stringify(languages),
      [DOCUMENT_NO_PROOF_ATTRIBUTE]: 'false',
      [DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE]: 'complexScript',
      lang: 'ar-SA',
    });
  });

  test('authors languages and explicit proofing state with one-step Undo', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Proofing language</p>',
    });
    editor.commands.selectAll();

    expect(
      editor.commands.setDocumentProofingLanguages({
        latin: 'en-US',
        eastAsia: 'zh-CN',
        bidi: 'ar-SA',
      }),
    ).toBe(true);
    expect(editor.commands.setDocumentNoProof(true)).toBe(true);
    expect(editor.getAttributes('textStyle')).toMatchObject({
      noProof: true,
      proofingLanguages: JSON.stringify({
        latin: 'en-US',
        eastAsia: 'zh-CN',
        bidi: 'ar-SA',
      }),
    });
    expect(editor.getHTML()).toContain(
      `${DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE}=`,
    );
    expect(editor.getHTML()).toContain(`${DOCUMENT_NO_PROOF_ATTRIBUTE}="true"`);

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getAttributes('textStyle').noProof).toBeNull();
    expect(editor.getAttributes('textStyle').proofingLanguages).toContain(
      'en-US',
    );
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).not.toContain(
      DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE,
    );
  });

  test('resolves defaults, styles, and independent direct language slots', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="ProofingStyle"/></w:pPr>
        <w:r><w:t>styled</w:t></w:r>
        <w:r><w:rPr><w:lang w:eastAsia="ja-JP"/><w:noProof w:val="0"/></w:rPr><w:t>direct</w:t></w:r>
      </w:p>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:noProof/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:styleId="ProofingStyle">
          <w:rPr><w:lang w:val="fr-FR"/><w:noProof w:val="false"/></w:rPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxRunFormatting(document, styles);
    expect(
      markers.runs.map(({ formatting }) => formatting.proofingLanguages),
    ).toEqual([
      { latin: 'fr-FR', eastAsia: 'zh-CN', bidi: 'ar-SA' },
      { latin: 'fr-FR', eastAsia: 'ja-JP', bidi: 'ar-SA' },
    ]);
    expect(markers.runs.map(({ formatting }) => formatting.noProof)).toEqual([
      false,
      false,
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
          `span[${DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE}]`,
        ),
      ).map((span) =>
        parseDocumentProofingLanguages(
          span.getAttribute(DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE),
        ),
      ),
    ).toEqual([
      { latin: 'fr-FR', eastAsia: 'zh-CN', bidi: 'ar-SA' },
      { latin: 'fr-FR', eastAsia: 'ja-JP', bidi: 'ar-SA' },
    ]);
  });

  test('accepts strict values and rejects malformed, duplicate, or namespace-spoofed properties', () => {
    const strict = parseXml(`
      <s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body><s:p>
        <s:r><s:rPr><s:lang s:val="en-GB" s:eastAsia="ja-JP" s:bidi="he-IL"/><s:noProof/></s:rPr><s:t>strict enabled</s:t></s:r>
        <s:r><s:rPr><s:lang s:val="fr-FR"/><s:noProof s:val="off"/></s:rPr><s:t>strict disabled</s:t></s:r>
      </s:p></s:body></s:document>
    `);
    expect(
      markDocxRunFormatting(strict).runs.map(({ formatting }) => ({
        languages: formatting.proofingLanguages,
        noProof: formatting.noProof,
      })),
    ).toEqual([
      {
        languages: { latin: 'en-GB', eastAsia: 'ja-JP', bidi: 'he-IL' },
        noProof: true,
      },
      { languages: { latin: 'fr-FR' }, noProof: false },
    ]);

    const rejected = wordXml(`
      <w:p xmlns:evil="https://example.test/evil">
        <w:r><w:rPr><evil:lang evil:val="en-US"/></w:rPr><w:t>spoofed element</w:t></w:r>
        <w:r><w:rPr><w:lang evil:val="en-US"/></w:rPr><w:t>spoofed attribute</w:t></w:r>
        <w:r><w:rPr><w:lang w:val="en_US"/></w:rPr><w:t>invalid tag</w:t></w:r>
        <w:r><w:rPr><w:lang w:val="en-US"/><w:lang w:val="fr-FR"/></w:rPr><w:t>duplicate language</w:t></w:r>
        <w:r><w:rPr><w:lang w:val="en-US"><w:b/></w:lang></w:rPr><w:t>nested language</w:t></w:r>
        <w:r><w:rPr><w:noProof w:val="yes"/></w:rPr><w:t>invalid proofing</w:t></w:r>
        <w:r><w:rPr><w:noProof/><w:noProof w:val="0"/></w:rPr><w:t>duplicate proofing</w:t></w:r>
      </w:p>
    `);
    expect(
      markDocxRunFormatting(rejected).runs.map(({ formatting }) => ({
        languages: formatting.proofingLanguages,
        noProof: formatting.noProof,
      })),
    ).toEqual(
      Array.from({ length: 7 }, () => ({
        languages: undefined,
        noProof: undefined,
      })),
    );
  });

  test('exports and reopens language and proofing state across editable Word stories', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const languages = {
      latin: 'en-US',
      eastAsia: 'zh-CN',
      bidi: 'ar-SA',
    } as const;
    artifact.content.html = [
      `<p>${proofingSpan(languages, true, 'body excluded')} ${proofingSpan(
        { latin: 'fr-FR' },
        false,
        'body checked',
      )}`,
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="proofing-footnote">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="proofing-endnote">1</sup></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="proofing-footnote">',
      `<p>${proofingSpan({ latin: 'de-DE' }, true, 'footnote excluded')}</p></aside>`,
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="proofing-endnote">',
      `<p>${proofingSpan({ eastAsia: 'ja-JP' }, false, 'endnote checked')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${proofingSpan(
          { latin: 'en-GB' },
          true,
          'header excluded',
        )}</p>`,
        footerHtml: `<p>${proofingSpan(
          { bidi: 'he-IL' },
          false,
          'footer checked',
        )}</p>`,
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
    expect(storyXml).toMatch(
      /<w:lang\b[^>]*w:val="en-US"[^>]*w:eastAsia="zh-CN"[^>]*w:bidi="ar-SA"/,
    );
    expect(storyXml).toMatch(/<w:noProof\/?\s*>/);
    expect(storyXml).toMatch(/<w:noProof\b[^>]*w:val="false"/);

    const reopened = await importOfficeFile(
      new File([blob], 'proofing-language.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(reopened.content.html).toContain(
      `${DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE}=`,
    );
    expect(reopened.content.html).toContain(
      `${DOCUMENT_NO_PROOF_ATTRIBUTE}="true"`,
    );
    expect(reopened.content.html).toContain(
      `${DOCUMENT_NO_PROOF_ATTRIBUTE}="false"`,
    );
    expect(reopened.content.pageChrome?.default.headerHtml).toContain('en-GB');
    expect(reopened.content.pageChrome?.default.footerHtml).toContain('he-IL');
    const reopenedDocument = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    expect(
      reopenedDocument.querySelector(
        `aside[data-note-kind="footnote"] span[${DOCUMENT_NO_PROOF_ATTRIBUTE}="true"]`,
      ),
    ).not.toBeNull();
    expect(
      reopenedDocument.querySelector(
        `aside[data-note-kind="endnote"] span[${DOCUMENT_NO_PROOF_ATTRIBUTE}="false"]`,
      ),
    ).not.toBeNull();
  });

  test('imports and exports native proofing formatting revisions', async () => {
    const source = wordXml(`
      <w:p><w:r><w:rPr>
        <w:noProof/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>
        <w:rPrChange w:id="31" w:author="Ada Reviewer" w:date="2026-08-24T05:00:00.000Z">
          <w:rPr><w:noProof w:val="0"/><w:lang w:val="fr-FR" w:eastAsia="ja-JP"/></w:rPr>
        </w:rPrChange>
      </w:rPr><w:t>proofing revision</w:t></w:r></w:p>
    `);
    const marker = markDocxRunFormatting(source).runs[0];
    expect(marker?.formatting).toMatchObject({
      noProof: true,
      proofingLanguages: { latin: 'en-US', eastAsia: 'zh-CN' },
    });
    expect(marker?.change?.before ?? '').toContain('fr-FR');
    expect(marker?.change?.before ?? '').toContain('noProof');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = JSON.stringify([
      {
        type: 'textStyle',
        attrs: {
          noProof: false,
          proofingLanguages: JSON.stringify({
            latin: 'fr-FR',
            eastAsia: 'ja-JP',
          }),
        },
      },
    ]);
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="proofing-change" data-change-author="Ada Reviewer" data-change-date="2026-08-24T05:30:00.000Z">${proofingSpan(
      { latin: 'en-US', eastAsia: 'zh-CN' },
      true,
      'changed proofing',
    )}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(document, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'changed proofing',
    );
    if (!run) throw new Error('Expected changed proofing run.');
    const properties = directChild(run, 'rPr');
    const change = directChild(properties ?? run, 'rPrChange');
    const prior = directChild(change ?? run, 'rPr');
    expect(
      attribute(directChild(properties ?? run, 'lang') ?? run, 'val'),
    ).toBe('en-US');
    expect(attribute(directChild(prior ?? run, 'lang') ?? run, 'val')).toBe(
      'fr-FR',
    );
    expect(attribute(directChild(prior ?? run, 'noProof') ?? run, 'val')).toBe(
      '0',
    );
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function proofingSpan(
  languages: Record<string, string>,
  noProof: boolean,
  content: string,
): string {
  const attributes = documentProofingDomAttributes(languages, noProof);
  return `<span ${Object.entries(attributes)
    .map(([name, value]) => `${name}='${value}'`)
    .join(' ')}>${content}</span>`;
}
