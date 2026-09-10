import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  applyImportedDocxParagraphBreakChangeMarkers,
  createDocxExternalHyperlinkTargets,
  createDocxImageEmbedTargets,
  EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  inspectDocxParagraphBreakMarkChanges,
  isIsolatedDocxParagraphBreakMarkChange,
  markDocxParagraphBreakChanges,
} from '../src/internal/features/work/work-docx-paragraph-mark-change-import';
import {
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORDPROCESSING_DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DRAWINGML_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PICTURE_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/picture';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IMAGE_RELATIONSHIP_TYPE = `${RELATIONSHIP_NAMESPACE}/image`;

const INLINE_PICTURE_DRAWING = [
  `<w:drawing xmlns:wp="${WORDPROCESSING_DRAWING_NAMESPACE}"`,
  ` xmlns:a="${DRAWINGML_NAMESPACE}"`,
  ` xmlns:pic="${PICTURE_NAMESPACE}"`,
  ` xmlns:r="${RELATIONSHIP_NAMESPACE}">`,
  '<wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">',
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>',
  '</a:graphicData></a:graphic></wp:inline>',
  '</w:drawing>',
].join('');

describe('DOCX paragraph-break merge/split revisions', () => {
  test('imports an eligible mark-only merge as a reviewable paragraph-break change', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p><w:pPr><w:rPr><w:del w:id="6" w:author="Ada" w:date="2026-09-05T01:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(inspectDocxParagraphBreakMarkChanges(document)).toEqual([
      expect.objectContaining({ kind: 'merge', author: 'Ada' }),
    ]);

    const markers = markDocxParagraphBreakChanges(document);
    expect(markers.paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
    expect(descendants(document, 'del')).toHaveLength(0);

    const html = new DOMParser().parseFromString(
      `<section><p>${markers.paragraphs[0]?.marker}Alpha</p><p>Bravo</p></section>`,
      'text/html',
    );
    applyImportedDocxParagraphBreakChangeMarkers(html, markers);
    const paragraph = html.querySelector('p');
    expect(paragraph?.dataset.paragraphBreakChange).toBe('true');
    expect(paragraph?.dataset.paragraphBreakKind).toBe('merge');
    expect(paragraph?.dataset.paragraphBreakAuthor).toBe('Ada');
    expect(html.body.textContent).not.toContain('__A3S_');
  });

  test('accepts a merge by joining the next text-only paragraph and exports mark-free body', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true">',
      '<p data-paragraph-break-change="true" data-paragraph-break-kind="merge" data-paragraph-break-id="break-merge-1" data-paragraph-break-author="Ada Reviewer" data-paragraph-break-date="2026-09-05T01:00:00.000Z">Alpha</p>',
      '<p>Bravo</p>',
      '</section>',
    ].join('');
    artifact.content.trackChanges = true;

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: artifact.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc).find(
      (candidate) => candidate.kind === 'paragraph-break',
    );
    expect(change).toEqual(
      expect.objectContaining({
        kind: 'paragraph-break',
        author: 'Ada Reviewer',
        text: 'Alpha',
      }),
    );
    expect(editor.commands.acceptDocumentChange(change?.id ?? '')).toBe(true);
    expect(editor.getText()).toContain('AlphaBravo');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);

    artifact.content.html = editor.getHTML();
    editor.destroy();

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('Alpha');
    expect(xml).toContain('Bravo');
    expect(xml).not.toMatch(/<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:del\b/);
  });

  test('rejects a split by joining into the previous paragraph and round-trips mark-only ins while pending', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true">',
      '<p>Alpha</p>',
      '<p data-paragraph-break-change="true" data-paragraph-break-kind="split" data-paragraph-break-id="break-split-1" data-paragraph-break-author="Ada Reviewer" data-paragraph-break-date="2026-09-05T01:00:00.000Z">Bravo</p>',
      '</section>',
    ].join('');
    artifact.content.trackChanges = true;

    const pendingBlob = await createArtifactBlob(artifact);
    const pendingArchive = await JSZip.loadAsync(
      await pendingBlob.arrayBuffer(),
    );
    const pendingXml =
      (await pendingArchive.file('word/document.xml')?.async('text')) ?? '';
    expect(pendingXml).toMatch(
      /<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:ins\b[^>]*w:author="Ada Reviewer"/,
    );
    expect(pendingXml).not.toMatch(
      /<w:ins\b[^>]*>[\s\S]*?<w:r>[\s\S]*?<w:t>Bravo/,
    );

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: artifact.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc).find(
      (candidate) => candidate.kind === 'paragraph-break',
    );
    expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
    expect(editor.getText().replace(/\n/g, '')).toContain('AlphaBravo');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    editor.destroy();
  });

  test('admits relationship-free bookmarks and internal hyperlinks in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="16" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:bookmarkStart w:id="1" w:name="AlphaMark"/>
            <w:r><w:t>Alpha</w:t><w:br/><w:t>line</w:t></w:r>
            <w:bookmarkEnd w:id="1"/>
          </w:p>
          <w:p>
            <w:hyperlink w:anchor="AlphaMark" w:tooltip="Jump">
              <w:r><w:t>Bravo</w:t></w:r>
            </w:hyperlink>
            <w:r><w:rPr><w:i/></w:rPr></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(inspectDocxParagraphBreakMarkChanges(document)).toEqual([
      expect.objectContaining({ kind: 'merge', author: 'Ada' }),
    ]);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('admits tabs and hyphen glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="26" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r>
              <w:t>Alpha</w:t><w:tab/><w:noBreakHyphen/><w:softHyphen/><w:t>mark</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r><w:t>Bravo</w:t><w:tab/><w:t>next</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('rejects attributed tab glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="27" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha</w:t><w:tab w:val="left"/><w:t>mark</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('admits empty carriage-return glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="28" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r>
              <w:t>Alpha</w:t><w:cr/><w:t>mark</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r><w:t>Bravo</w:t><w:cr/><w:t>next</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('rejects attributed carriage-return glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="29" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha</w:t><w:cr w:val="1"/><w:t>mark</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('admits empty lastRenderedPageBreak glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="30" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r>
              <w:t>Alpha</w:t><w:lastRenderedPageBreak/><w:t>mark</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r><w:t>Bravo</w:t><w:lastRenderedPageBreak/><w:t>next</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('rejects attributed lastRenderedPageBreak glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="31" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha</w:t><w:lastRenderedPageBreak w:val="1"/><w:t>mark</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('admits empty page-number and date-field glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="32" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r>
              <w:t>Alpha</w:t><w:pgNum/><w:dayLong/><w:monthLong/><w:yearLong/><w:t>mark</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r><w:t>Bravo</w:t><w:pgNum/><w:t>next</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('rejects attributed page-number glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="33" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha</w:t><w:pgNum w:val="1"/><w:t>mark</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('admits empty short date-field glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="34" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r>
              <w:t>Alpha</w:t><w:dayShort/><w:monthShort/><w:yearShort/><w:t>mark</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:r><w:t>Bravo</w:t><w:dayShort/><w:t>next</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('rejects attributed short date-field glyphs in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="35" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha</w:t><w:dayShort w:val="1"/><w:t>mark</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('admits relationship-bound external hyperlinks in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="17" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:hyperlink r:id="rId9">
              <w:r><w:t>Alpha</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const externalHyperlinks = createDocxExternalHyperlinkTargets([
      {
        id: 'rId9',
        target: 'mailto:review@a3s.dev',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        targetMode: 'External',
      },
    ]);
    const mark = descendants(document, 'del')[0];
    expect(
      mark && isIsolatedDocxParagraphBreakMarkChange(mark, externalHyperlinks),
    ).toBe(true);
    expect(
      markDocxParagraphBreakChanges(document, externalHyperlinks).paragraphs,
    ).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('admits inline DrawingML pictures in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="21" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha </w:t></w:r>
            <w:r>${INLINE_PICTURE_DRAWING}</w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const imageEmbeds = createDocxImageEmbedTargets([
      { id: 'rId1', type: IMAGE_RELATIONSHIP_TYPE },
    ]);
    const mark = descendants(document, 'del')[0];
    expect(
      mark &&
        isIsolatedDocxParagraphBreakMarkChange(
          mark,
          EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
          imageEmbeds,
        ),
    ).toBe(true);
    expect(
      markDocxParagraphBreakChanges(
        document,
        EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
        imageEmbeds,
      ).paragraphs,
    ).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('admits picture-only paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="31" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r>${INLINE_PICTURE_DRAWING}</w:r>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const imageEmbeds = createDocxImageEmbedTargets([
      { id: 'rId1', type: IMAGE_RELATIONSHIP_TYPE },
    ]);
    const mark = descendants(document, 'del')[0];
    expect(
      mark &&
        isIsolatedDocxParagraphBreakMarkChange(
          mark,
          EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
          imageEmbeds,
        ),
    ).toBe(true);
    expect(
      markDocxParagraphBreakChanges(
        document,
        EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
        imageEmbeds,
      ).paragraphs,
    ).toEqual([
      expect.objectContaining({
        kind: 'merge',
        author: 'Ada',
      }),
    ]);
  });

  test('rejects unresolved relationship-bound hyperlinks in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="17" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:hyperlink r:id="rId9">
              <w:r><w:t>Alpha</w:t></w:r>
            </w:hyperlink>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(inspectDocxParagraphBreakMarkChanges(document)).toEqual([]);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('keeps ineligible neighbors as diagnostics instead of reviewable breaks', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="19" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:r><w:t>Alpha</w:t></w:r>
          </w:p>
          <w:p>
            <w:hyperlink r:id="rId9">
              <w:r><w:t>Bravo</w:t></w:r>
            </w:hyperlink>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(true);
    expect(inspectDocxParagraphBreakMarkChanges(document)).toEqual([
      expect.objectContaining({ kind: 'merge', author: 'Ada' }),
    ]);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('rejects relationship-spoofed bookmarks in paragraph-break bodies', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:del w:id="18" w:author="Ada" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:bookmarkStart w:id="2" w:name="Spoof" r:id="rId1"/>
            <w:r><w:t>Alpha</w:t></w:r>
            <w:bookmarkEnd w:id="2"/>
          </w:p>
          <w:p><w:r><w:t>Bravo</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'del')[0];
    expect(mark && isIsolatedDocxParagraphBreakMarkChange(mark)).toBe(false);
    expect(markDocxParagraphBreakChanges(document).paragraphs).toEqual([]);
  });

  test('imports native mark-only merge through the DOCX pipeline as a reviewable change', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:rPr><w:del w:id="6" w:author="Ada" w:date="2026-09-05T01:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Bravo</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
    );
    archive.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    );
    archive.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
    );
    const bytes = await archive.generateAsync({ type: 'uint8array' });
    const reopened = await importOfficeFile(
      new File([bytes], 'paragraph-break-merge.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    expect(
      collectDocumentChanges(editor.state.doc).some(
        (change) => change.kind === 'paragraph-break',
      ),
    ).toBe(true);
    editor.destroy();
  });
});
