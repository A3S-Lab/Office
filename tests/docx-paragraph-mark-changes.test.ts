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
  inspectDocxParagraphBreakMarkChanges,
  isIsolatedDocxParagraphBreakMarkChange,
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

  test('reports isolated paragraph-break mark revisions with a dedicated diagnostic', async () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Boundary" w:date="2026-09-05T01:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Only the paragraph mark changed</w:t></w:r></w:p>
          <w:p><w:pPr><w:rPr><w:del w:id="6" w:author="Boundary" w:date="2026-09-05T01:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Merge candidate</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const markRevisions = [
      ...descendants(document, 'ins'),
      ...descendants(document, 'del'),
    ].filter((revision) => revision.parentElement?.localName === 'rPr');
    expect(markRevisions.map(isSupportedDocxParagraphMarkChange)).toEqual([
      false,
      false,
    ]);
    expect(markRevisions.map(isIsolatedDocxParagraphBreakMarkChange)).toEqual([
      true,
      true,
    ]);
    expect(inspectDocxParagraphBreakMarkChanges(document)).toEqual([
      expect.objectContaining({ kind: 'split', author: 'Boundary' }),
      expect.objectContaining({ kind: 'merge', author: 'Boundary' }),
    ]);

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:pPr><w:rPr><w:ins w:id="5" w:author="Boundary" w:date="2026-09-05T01:00:00Z"/></w:rPr></w:pPr><w:r><w:t>Only the paragraph mark changed</w:t></w:r></w:p></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'paragraph-break-revision.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.paragraph-break',
      }),
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.paragraph-mark'),
    ).toBe(false);
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);
  });

  test('imports mixed multi-wrapper paragraph-mark revisions as one atomic block', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="21" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:ins w:id="22" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold </w:t></w:r></w:ins>
        <w:ins w:id="23" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"><w:r><w:t>plain</w:t></w:r></w:ins>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(true);
    const markers = markDocxParagraphMarkChanges(document);
    expect(markers.paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-mark-change-21',
        kind: 'insertion',
        author: 'Ada Reviewer',
      }),
    ]);

    const html = new DOMParser().parseFromString(
      `<p>${markers.paragraphs[0]?.marker}<ins data-document-change="true" data-change-kind="insertion" data-change-author="Ada Reviewer" data-change-date="2026-09-05T01:00:00.000Z"><strong>Bold </strong></ins><ins data-document-change="true" data-change-kind="insertion" data-change-author="Ada Reviewer" data-change-date="2026-09-05T01:00:00.000Z">plain</ins></p>`,
      'text/html',
    );
    applyImportedDocxParagraphMarkChangeMarkers(html, markers);
    expect(html.querySelector('p')?.dataset.documentBlockChange).toBe('true');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: `<section data-document-section="true">${html.body.innerHTML}</section>`,
    });
    const changes = collectDocumentChanges(editor.state.doc);
    expect(changes).toEqual([
      expect.objectContaining({
        kind: 'insertion',
        text: 'Bold plain',
        author: 'Ada Reviewer',
      }),
    ]);
    expect(editor.commands.rejectDocumentChange(changes[0]?.id ?? '')).toBe(
      true,
    );
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.getText()).not.toContain('Bold');
    editor.destroy();
  });

  test('rejects mixed untracked body siblings as whole-paragraph mark revisions', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="31" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:r><w:t>Keep </w:t></w:r>
        <w:ins w:id="32" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"><w:r><w:t>Added</w:t></w:r></w:ins>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(false);
    expect(markDocxParagraphMarkChanges(document).paragraphs).toEqual([]);
  });

  test('admits empty and rPr-only untracked runs beside whole-paragraph mark revisions', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="81" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:r><w:rPr><w:b/></w:rPr></w:r>
        <w:ins w:id="82" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
          <w:r><w:t>Marked</w:t></w:r>
        </w:ins>
        <w:r/>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(true);
    expect(markDocxParagraphMarkChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-mark-change-81',
        kind: 'insertion',
      }),
    ]);
  });

  test('admits soft text-wrapping breaks inside whole-paragraph mark revisions', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="41" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:ins w:id="42" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
          <w:r><w:t>Line one</w:t><w:br/><w:t>Line two</w:t></w:r>
        </w:ins>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(true);
    expect(markDocxParagraphMarkChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-mark-change-41',
        kind: 'insertion',
      }),
    ]);
  });

  test('admits relationship-free internal hyperlinks inside whole-paragraph mark revisions', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="51" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:ins w:id="52" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
          <w:hyperlink w:anchor="TargetBookmark" w:tooltip="Jump">
            <w:r><w:t>Linked paragraph</w:t></w:r>
          </w:hyperlink>
        </w:ins>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(true);
    expect(markDocxParagraphMarkChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-mark-change-51',
        kind: 'insertion',
      }),
    ]);
  });

  test('admits relationship-free bookmarks inside and beside whole-paragraph mark revisions', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="61" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:bookmarkStart w:id="7" w:name="MarkedParagraph"/>
        <w:ins w:id="62" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
          <w:bookmarkStart w:id="8" w:name="InnerMark"/>
          <w:r><w:t>Bookmarked paragraph</w:t></w:r>
          <w:bookmarkEnd w:id="8"/>
        </w:ins>
        <w:bookmarkEnd w:id="7"/>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(true);
    expect(markDocxParagraphMarkChanges(document).paragraphs).toEqual([
      expect.objectContaining({
        id: 'docx-paragraph-mark-change-61',
        kind: 'insertion',
      }),
    ]);
  });

  test('rejects relationship-bound hyperlinks inside whole-paragraph mark revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:ins w:id="53" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:ins w:id="54" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
              <w:hyperlink r:id="rId5">
                <w:r><w:t>External</w:t></w:r>
              </w:hyperlink>
            </w:ins>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(false);
  });

  test('rejects relationship-spoofed bookmarks beside whole-paragraph mark revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:pPr><w:rPr>
              <w:ins w:id="71" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
            </w:rPr></w:pPr>
            <w:bookmarkStart w:id="9" w:name="Spoof" r:id="rId1"/>
            <w:ins w:id="72" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
              <w:r><w:t>Spoofed</w:t></w:r>
            </w:ins>
            <w:bookmarkEnd w:id="9"/>
          </w:p>
        </w:body>
      </w:document>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(false);
  });

  test('rejects page breaks inside whole-paragraph mark revisions', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:rPr>
          <w:ins w:id="43" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z"/>
        </w:rPr></w:pPr>
        <w:ins w:id="44" w:author="Ada Reviewer" w:date="2026-09-05T01:00:00Z">
          <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
        </w:ins>
      </w:p>
    `);
    const mark = descendants(document, 'ins').find(
      (revision) => revision.parentElement?.localName === 'rPr',
    );
    expect(mark && isSupportedDocxParagraphMarkChange(mark)).toBe(false);
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

  test('exports and reopens multi-wrapper mixed paragraph-mark insertions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true">',
      '<p data-document-block-change="true" data-block-change-kind="insertion" data-block-change-id="block-mixed" data-block-change-author="Ada Reviewer" data-block-change-date="2026-09-05T01:00:00.000Z">',
      '<ins data-document-change="true" data-change-kind="insertion" data-change-id="text-a" data-change-author="Ada Reviewer" data-change-date="2026-09-05T01:00:00.000Z">Bold </ins>',
      '<ins data-document-change="true" data-change-kind="insertion" data-change-id="text-b" data-change-author="Ada Reviewer" data-change-date="2026-09-05T01:00:00.000Z">plain</ins>',
      '</p><p>Stable paragraph</p></section>',
    ].join('');
    artifact.content.trackChanges = true;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(
      /<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:ins\b[^>]*w:author="Ada Reviewer"/,
    );
    expect(xml.match(/<w:ins\b/g)?.length).toBeGreaterThanOrEqual(2);

    const reopened = await importOfficeFile(
      new File([blob], 'paragraph-mark-mixed.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc).find(
      (candidate) => candidate.kind === 'insertion',
    );
    expect(change).toEqual(
      expect.objectContaining({
        text: 'Bold plain',
        author: 'Ada Reviewer',
      }),
    );
    expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
    expect(editor.getText()).not.toContain('Bold');
    expect(editor.getText()).toContain('Stable paragraph');
    editor.destroy();
  });

  test('exports paragraph-mark revisions that include soft line breaks', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true">',
      '<p data-document-block-change="true" data-block-change-kind="insertion" data-block-change-id="block-br" data-block-change-author="Ada Reviewer" data-block-change-date="2026-09-05T01:00:00.000Z">',
      '<ins data-document-change="true" data-change-kind="insertion" data-change-id="text-br" data-change-author="Ada Reviewer" data-change-date="2026-09-05T01:00:00.000Z">Line one<br>Line two</ins>',
      '</p></section>',
    ].join('');
    artifact.content.trackChanges = true;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(
      /<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:ins\b[^>]*w:author="Ada Reviewer"/,
    );
    expect(xml).toMatch(/<w:br\b/);
    expect(xml).toContain('Line one');
    expect(xml).toContain('Line two');
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
