import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import {
  parseDocumentParagraphFormatting,
  serializeDocumentParagraphFormatting,
} from '../src/internal/features/work/work-document-paragraph-format-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  applyImportedDocxParagraphFormattingChangeMarkers,
  isSupportedDocxParagraphFormattingChange,
  markDocxParagraphFormattingChanges,
} from '../src/internal/features/work/work-docx-paragraph-format-change-import';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

describe('DOCX paragraph-formatting revisions', () => {
  test('imports a bounded native paragraph-property revision as node review metadata', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:jc w:val="center"/>
          <w:spacing w:before="360" w:after="180" w:line="360" w:lineRule="auto"/>
          <w:pPrChange w:id="9" w:author="Ada Reviewer" w:date="2026-08-18T09:30:00Z">
            <w:pPr>
              <w:keepNext w:val="0"/>
              <w:keepLines/>
              <w:bidi/>
              <w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/>
              <w:ind w:left="720" w:right="180" w:hanging="180"/>
              <w:jc w:val="right"/>
              <w:outlineLvl w:val="2"/>
            </w:pPr>
          </w:pPrChange>
        </w:pPr>
        <w:r><w:t>Changed paragraph</w:t></w:r>
      </w:p>
    `);

    const markers = markDocxParagraphFormattingChanges(document);
    expect(markers.paragraphs).toHaveLength(1);
    const change = markers.paragraphs[0];
    if (!change) throw new Error('Expected a paragraph change marker.');
    expect(change).toMatchObject({
      id: 'docx-paragraph-format-change-9',
      author: 'Ada Reviewer',
      date: '2026-08-18T09:30:00.000Z',
    });
    expect(parseDocumentParagraphFormatting(change.before)).toMatchObject({
      textAlign: 'right',
      paragraphDirection: 'rtl',
      indentLevel: 2,
      rightIndent: 12,
      firstLineIndent: -12,
      spaceBefore: 6,
      spaceAfter: 3,
      lineHeight: '1',
      lineRule: 'auto',
      keepLines: true,
      keepWithNext: false,
      outlineLevel: 2,
    });

    const html = new DOMParser().parseFromString(
      `<p style="text-align: center">${change.marker}Changed paragraph</p>`,
      'text/html',
    );
    applyImportedDocxParagraphFormattingChangeMarkers(html, markers);
    const paragraph = html.querySelector('p');
    expect(paragraph?.dataset.changeKind).toBe('paragraph-formatting');
    expect(paragraph?.dataset.changeId).toBe('docx-paragraph-format-change-9');
    expect(html.body.textContent).not.toContain('__A3S_');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html.body.innerHTML,
    });
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        kind: 'paragraph-formatting',
        text: 'Changed paragraph',
      }),
    ]);
    editor.destroy();
  });

  test('requires one strict, supported, uniquely keyed previous property set', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:s="${STRICT_WORD_NAMESPACE}" xmlns:evil="https://example.test/evil">
        <w:body><w:p><w:pPr>
          <w:pPrChange w:id="1" w:author="Valid"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>
          <w:pPrChange w:id="2" w:author="Duplicate"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>
        </w:pPr></w:p>
        <w:p><w:pPr><w:pPrChange w:id="3" w:author="Unsupported"><w:pPr><w:numPr/></w:pPr></w:pPrChange></w:pPr></w:p>
        <w:p><w:pPr><w:pPrChange w:id="4" w:author="Spoofed"><evil:pPr><evil:jc/></evil:pPr></w:pPrChange></w:pPr></w:p>
        <w:p><evil:pPr><w:pPrChange w:id="5" w:author="Spoofed parent"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></evil:pPr></w:p>
        <s:p><s:pPr><s:pPrChange s:id="6" s:author="Strict"><s:pPr><s:jc s:val="center"/></s:pPr></s:pPrChange></s:pPr></s:p>
        </w:body>
      </w:document>
    `);
    const changes = descendants(document, 'pPrChange');

    expect(changes.map(isSupportedDocxParagraphFormattingChange)).toEqual([
      false,
      false,
      false,
      false,
      false,
      true,
    ]);
    expect(markDocxParagraphFormattingChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-format-change-6',
        author: 'Strict',
      }),
    ]);
  });

  test('exports and reopens native paragraph-property revisions without marker leakage', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = serializeDocumentParagraphFormatting({
      textAlign: 'right',
      paragraphDirection: 'rtl',
      indentLevel: 2,
      rightIndent: 12,
      firstLineIndent: -12,
      spaceBefore: 6,
      spaceAfter: 3,
      lineHeight: '1.25',
      lineRule: 'auto',
      keepLines: true,
      keepWithNext: false,
      pageBreakBefore: false,
      widowControl: true,
      contextualSpacing: true,
      outlineLevel: 2,
      tabStops: [{ position: 96, alignment: 'right', leader: 'dot' }],
      paragraphBorders: {
        bottom: {
          style: 'single',
          color: { value: '#336699' },
          size: 8,
          space: 1,
        },
      },
      paragraphShading: {
        pattern: 'clear',
        fill: { value: '#ddeeff' },
      },
      defaultCollapsed: false,
    });
    artifact.content.html = `<section data-document-section="true"><p data-document-change="true" data-change-kind="paragraph-formatting" data-change-before='${before}' data-change-id="paragraph-9" data-change-author="Ada Reviewer" data-change-date="2026-08-18T09:30:00.000Z" style="text-align: center" data-office-space-before="18">Changed paragraph</p></section>`;
    artifact.content.trackChanges = true;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).not.toContain('__A3S_WORK_PARAGRAPH_FORMAT_CHANGE_');
    expect(xml).toMatch(
      /<w:pPrChange\b[^>]*w:id="1"[^>]*w:author="Ada Reviewer"[^>]*w:date="2026-08-18T09:30:00.000Z"/,
    );
    expect(xml).toMatch(
      /<w:pPrChange\b[^>]*>[\s\S]*?<w:pPr>[\s\S]*?<w:keepLines\b[^>]*>[\s\S]*?<w:pBdr>[\s\S]*?<w:bottom\b[^>]*w:color="336699"[\s\S]*?<w:shd\b[^>]*w:fill="DDEEFF"[\s\S]*?<w:tabs>[\s\S]*?<w:tab\b[^>]*w:val="right"[^>]*w:pos="1440"[\s\S]*?<w:jc\b[^>]*w:val="right"/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'paragraph-formatting-revision.docx', {
        type: blob.type,
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-change-kind="paragraph-formatting"',
    );
    expect(reopened.content.html).toContain(
      'data-change-author="Ada Reviewer"',
    );
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc)[0];
    expect(change).toMatchObject({ kind: 'paragraph-formatting' });
    expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
    expect(editor.getHTML()).toContain('text-align: right');
    expect(editor.getText()).toContain('Changed paragraph');
    editor.destroy();
  });

  test('separates supported paragraph revisions from structural warnings', async () => {
    const supported = await revisionCompatibility(`
      <w:p><w:pPr><w:pPrChange w:id="7" w:author="Ada Reviewer">
        <w:pPr><w:jc w:val="right"/><w:spacing w:before="120"/></w:pPr>
      </w:pPrChange></w:pPr><w:r><w:t>Supported</w:t></w:r></w:p>
    `);
    expect(supported.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.paragraph-formatting',
        severity: 'info',
      }),
    );
    expect(
      supported.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);

    const malformed = await revisionCompatibility(`
      <w:p><w:pPr><w:pPrChange w:id="7" w:author="Ada Reviewer">
        <w:pPr><w:numPr/></w:pPr>
      </w:pPrChange></w:pPr><w:r><w:t>Unsupported</w:t></w:r></w:p>
    `);
    expect(malformed.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.structural',
        severity: 'warning',
      }),
    );
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

async function revisionCompatibility(
  body: string,
): ReturnType<typeof analyzeDocxCompatibility> {
  const archive = new JSZip();
  archive.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
  const bytes = await archive.generateAsync({ type: 'arraybuffer' });
  return analyzeDocxCompatibility(
    new File([bytes], 'paragraph-formatting-revision.docx'),
    [],
  );
}
