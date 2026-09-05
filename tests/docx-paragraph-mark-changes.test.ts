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
  applyImportedDocxParagraphMarkChangeMarkers,
  isSupportedDocxParagraphMarkChange,
  markDocxParagraphMarkChanges,
} from '../src/internal/features/work/work-docx-paragraph-mark-change-import';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

describe('DOCX paragraph-mark revisions', () => {
  test.each([
    ['ins', 'insertion'],
    ['del', 'deletion'],
  ] as const)('imports a native paragraph-mark %s revision', (tag, kind) => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:${tag} w:id="17" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
          <w:rFonts w:hint="eastAsia"/>
        </w:rPr></w:pPr>
        <w:${tag} w:id="18" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"><w:r><w:${tag === 'del' ? 'delText' : 't'}>${kind === 'deletion' ? 'Removed' : 'Added'} paragraph</w:${tag === 'del' ? 'delText' : 't'}></w:r></w:${tag}>
      </w:p>
    `);
    const markers = markDocxParagraphMarkChanges(document);
    expect(markers.paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-mark-change-17',
        kind,
        author: 'Ada Reviewer',
        date: '2026-09-05T01:00:00.000Z',
      }),
    ]);
    const marker = markers.paragraphs[0];
    if (!marker) throw new Error('Expected a paragraph-mark marker.');

    const html = new DOMParser().parseFromString(
      `<p>${marker.marker}${kind === 'deletion' ? '<del data-document-change="true" data-change-kind="deletion">Removed paragraph</del>' : '<ins data-document-change="true" data-change-kind="insertion">Added paragraph</ins>'}</p>`,
      'text/html',
    );
    applyImportedDocxParagraphMarkChangeMarkers(html, markers);
    const paragraph = html.querySelector('p');
    expect(paragraph?.dataset.documentBlockChange).toBe('true');
    expect(paragraph?.dataset.blockChangeKind).toBe(kind);
    expect(paragraph?.dataset.blockChangeId).toBe(
      'docx-paragraph-mark-change-17',
    );
    expect(html.body.textContent).not.toContain('__A3S_');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: `<section data-document-section="true">${html.body.innerHTML}</section>`,
    });
    const changes = collectDocumentChanges(editor.state.doc);
    expect(changes).toEqual([
      expect.objectContaining({
        kind,
        text: kind === 'deletion' ? 'Removed paragraph' : 'Added paragraph',
      }),
    ]);
    const change = changes[0];
    if (!change) throw new Error('Expected one block revision.');
    expect(
      kind === 'insertion'
        ? editor.commands.acceptDocumentChange(change.id)
        : editor.commands.rejectDocumentChange(change.id),
    ).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.getText()).toContain(
      kind === 'deletion' ? 'Removed paragraph' : 'Added paragraph',
    );
    editor.destroy();
  });

  test('accepts strict namespace and rejects ambiguous or spoofed paragraph marks', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:s="${STRICT_WORD_NAMESPACE}" xmlns:evil="https://example.test/evil">
        <w:body>
          <s:p><s:pPr><s:rPr><s:ins s:id="1" s:author="Strict"/></s:rPr></s:pPr><s:ins s:id="11" s:author="Strict"><s:r><s:t>Strict</s:t></s:r></s:ins></s:p>
          <w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="One"/><w:del w:id="3" w:author="Two"/></w:rPr></w:pPr></w:p>
          <w:p><w:pPr><w:rPr><evil:ins evil:id="4" evil:author="Spoofed"/></w:rPr></w:pPr></w:p>
          <w:p><w:pPr><w:rPr><w:ins w:id="bad" w:author="Malformed"/></w:rPr></w:pPr></w:p>
          <w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Boundary"/></w:rPr></w:pPr><w:r><w:t>Only the paragraph mark changed</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const changes = descendants(document, 'ins').filter(
      (change) => change.parentElement?.localName === 'rPr',
    );
    expect(changes.map(isSupportedDocxParagraphMarkChange)).toEqual([
      true,
      false,
      false,
      false,
      false,
    ]);
    expect(markDocxParagraphMarkChanges(document).paragraphs).toEqual([
      expect.objectContaining({ id: 'docx-paragraph-mark-change-1' }),
    ]);
  });

  test.each([
    ['insertion', 'ins', 'Inserted paragraph', 'reject'],
    ['deletion', 'del', 'Deleted paragraph', 'accept'],
  ] as const)('exports and reopens paragraph-mark %s revisions with atomic block decisions', async (kind, tag, changedText, decision) => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true">',
      `<p data-document-block-change="true" data-block-change-kind="${kind}" data-block-change-id="block-${kind}" data-block-change-author="Ada Reviewer" data-block-change-date="2026-09-05T01:00:00.000Z">`,
      `<${tag} data-document-change="true" data-change-kind="${kind}" data-change-id="text-${kind}" data-change-author="Ada Reviewer" data-change-date="2026-09-05T01:00:00.000Z">${changedText}</${tag}>`,
      '</p><p>Stable paragraph</p></section>',
    ].join('');
    artifact.content.trackChanges = true;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).not.toContain('__A3S_WORK_PARAGRAPH_MARK_CHANGE_EXPORT_');
    expect(xml).toMatch(
      new RegExp(
        `<w:pPr>[\\s\\S]*?<w:rPr>[\\s\\S]*?<w:${tag}\\b[^>]*w:id="1"[^>]*w:author="Ada Reviewer"[^>]*w:date="2026-09-05T01:00:00.000Z"`,
      ),
    );

    const reopened = await importOfficeFile(
      new File([blob], 'paragraph-mark-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc).find(
      (candidate) => candidate.kind === kind,
    );
    expect(change).toEqual(
      expect.objectContaining({
        text: changedText,
        author: 'Ada Reviewer',
      }),
    );
    expect(
      decision === 'accept'
        ? editor.commands.acceptDocumentChange(change?.id ?? '')
        : editor.commands.rejectDocumentChange(change?.id ?? ''),
    ).toBe(true);
    expect(editor.getText()).toContain('Stable paragraph');
    expect(editor.getText()).not.toContain(changedText);
    editor.destroy();
  });

  test('refuses to export an isolated paragraph-mark boundary as a whole-paragraph revision', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p data-document-block-change="true" data-block-change-kind="deletion" data-block-change-id="boundary-only" data-block-change-author="Ada Reviewer" data-block-change-date="2026-09-05T01:00:00.000Z">Untracked paragraph body</p></section>';
    artifact.content.trackChanges = true;

    await expect(createArtifactBlob(artifact)).rejects.toThrow(
      'invalid paragraph-mark revision',
    );
  });

  test('reports supported paragraph-mark revisions without a structural warning', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:rPr><w:ins w:id="2" w:author="Ada"/></w:rPr></w:pPr><w:ins w:id="3" w:author="Ada"><w:r><w:t>Added</w:t></w:r></w:ins></w:p></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'paragraph-mark-revision.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.paragraph-mark',
        severity: 'info',
      }),
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}
