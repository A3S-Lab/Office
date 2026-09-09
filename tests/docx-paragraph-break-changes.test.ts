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
