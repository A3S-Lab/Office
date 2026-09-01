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
  parseDocumentNumberingChange,
  serializeDocumentNumberingChange,
} from '../src/internal/features/work/work-document-numbering-changes';
import {
  applyImportedDocxNumberingChangeMarkers,
  isSupportedDocxNumberingChange,
  markDocxNumberingChanges,
} from '../src/internal/features/work/work-docx-numbering-change-import';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

describe('DOCX ordered-list numbering revisions', () => {
  test('imports one contiguous native numbering revision as an atomic review card', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:numPr>
        <w:ilvl w:val="0"/><w:numId w:val="42"/>
        <w:numberingChange w:id="17" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:3:1:."/>
      </w:numPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
      <w:p><w:pPr><w:numPr>
        <w:ilvl w:val="0"/><w:numId w:val="42"/>
        <w:numberingChange w:id="18" w:author="Ada Reviewer" w:date="2026-09-01T09:30:00Z" w:original="%1:4:1:."/>
      </w:numPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
    `);

    const markers = markDocxNumberingChanges(document);
    expect(markers.groups).toHaveLength(1);
    const group = markers.groups[0];
    if (!group) throw new Error('Expected a numbering-change group.');
    expect(group).toMatchObject({
      id: 'docx-numbering-change-17',
      author: 'Ada Reviewer',
      date: '2026-09-01T09:30:00.000Z',
      start: 3,
      level: 0,
      format: 1,
      suffix: '.',
    });
    expect(group.markers).toHaveLength(2);

    const html = new DOMParser().parseFromString(
      [
        '<ol start="3" type="a" data-office-numbering-id="42" data-office-numbering-level="0">',
        `<li><p>${group.markers[0]}First</p></li>`,
        `<li><p>${group.markers[1]}Second</p></li>`,
        '</ol>',
      ].join(''),
      'text/html',
    );
    applyImportedDocxNumberingChangeMarkers(html, markers);

    const list = html.querySelector('ol');
    expect(list?.dataset.changeKind).toBe('numbering');
    expect(list?.dataset.changeId).toBe('docx-numbering-change-17');
    expect(list?.dataset.changeAuthor).toBe('Ada Reviewer');
    expect(html.body.textContent).not.toContain('__A3S_');
    expect(parseDocumentNumberingChange(list?.dataset.changeBefore)).toEqual(
      expect.objectContaining({
        start: 3,
        type: 'I',
        level: 0,
        originalFormat: 1,
        originalSuffix: '.',
      }),
    );

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html.body.innerHTML,
    });
    const changes = collectDocumentChanges(editor.state.doc);
    expect(changes).toEqual([
      expect.objectContaining({
        id: 'docx-numbering-change-17',
        kind: 'numbering',
        text: 'FirstSecond',
      }),
    ]);
    expect(editor.commands.rejectDocumentChange(changes[0]?.id ?? '')).toBe(
      true,
    );
    expect(editor.getHTML()).toContain('start="3"');
    expect(editor.getHTML()).toContain('type="I"');
    expect(editor.getHTML()).not.toContain('data-change-kind="numbering"');
    expect(editor.commands.undo()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);
    editor.destroy();
  });

  test('tracks an ordered-list style change and keeps accept or reject atomic', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind) => ({
          id: `${kind}-1`,
          author: 'Local Reviewer',
          date: '2026-09-01T10:00:00.000Z',
        }),
      }),
      content:
        '<ol start="4" type="I"><li><p>First</p></li><li><p>Second</p></li></ol>',
    });
    editor.commands.setTextSelection(textPosition(editor, 'First'));

    expect(editor.commands.applyDocumentOrderedList('lower-alpha')).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'numbering-1',
        kind: 'numbering',
        author: 'Local Reviewer',
      }),
    ]);
    expect(editor.getHTML()).toContain('data-change-kind="numbering"');
    expect(editor.getHTML()).toContain('start="4"');
    expect(editor.getHTML()).toContain('type="a"');

    expect(editor.commands.rejectDocumentChange('numbering-1')).toBe(true);
    expect(editor.getHTML()).toContain('start="4"');
    expect(editor.getHTML()).toContain('type="I"');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.commands.undo()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);

    expect(editor.commands.acceptDocumentChange('numbering-1')).toBe(true);
    expect(editor.getHTML()).toContain('start="4"');
    expect(editor.getHTML()).toContain('type="a"');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    editor.destroy();
  });

  test('tracks and restores an ordered-list start value', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind) => ({
          id: `${kind}-start`,
          author: 'Local Reviewer',
          date: '2026-09-01T10:15:00.000Z',
        }),
      }),
      content: '<ol start="3"><li><p>First</p></li></ol>',
    });
    editor.commands.setTextSelection(textPosition(editor, 'First'));

    expect(editor.commands.setDocumentNumberingStart(9)).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'numbering-start',
        kind: 'numbering',
      }),
    ]);
    expect(editor.getHTML()).toContain('start="9"');

    expect(editor.commands.rejectDocumentChange('numbering-start')).toBe(true);
    expect(editor.getHTML()).toContain('start="3"');
    expect(editor.getHTML()).not.toContain('data-change-kind="numbering"');
    editor.destroy();
  });

  test('exports and reopens native numbering revisions without marker leakage', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = serializeDocumentNumberingChange({
      start: 4,
      type: 'I',
      level: 0,
      originalFormat: 1,
      originalSuffix: '.',
    });
    artifact.content.html = [
      '<section data-document-section="true">',
      `<ol start="4" type="a" data-document-change="true" data-change-kind="numbering" data-change-before='${before}' data-change-id="numbering-9" data-change-author="Ada Reviewer" data-change-date="2026-09-01T10:30:00.000Z">`,
      '<li><p>First</p></li><li><p>Second</p></li>',
      '</ol></section>',
    ].join('');
    artifact.content.trackChanges = false;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    const settings =
      (await archive.file('word/settings.xml')?.async('text')) ?? '';
    expect(settings).toMatch(/<w:trackRevisions\b/);
    expect(xml).not.toContain('__A3S_WORK_NUMBERING_CHANGE_');
    expect(xml).toMatch(
      /<w:numberingChange\b[^>]*w:id="1"[^>]*w:author="Ada Reviewer"[^>]*w:original="%1:4:1:\."/,
    );
    expect(xml).toMatch(
      /<w:numberingChange\b[^>]*w:id="2"[^>]*w:author="Ada Reviewer"[^>]*w:original="%1:5:1:\."/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'numbering-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="numbering"');
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc)[0];
    expect(change).toMatchObject({ kind: 'numbering' });
    expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
    expect(editor.getHTML()).toContain('start="4"');
    expect(editor.getHTML()).toContain('type="I"');
    editor.destroy();

    artifact.content.html = artifact.content.html.replace(
      'Ada Reviewer',
      'Ada&#10;Reviewer',
    );
    await expect(createArtifactBlob(artifact)).rejects.toThrow(
      'invalid numbering revision',
    );
  });

  test('fails closed for malformed, duplicated, spoofed, or unsupported native forms', async () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:s="${STRICT_WORD_NAMESPACE}" xmlns:evil="https://example.test/evil">
        <w:body>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>
            <w:numberingChange w:id="1" w:author="Valid" w:original="%1:1:0:."/>
          </w:numPr></w:pPr></w:p>
          <s:p><s:pPr><s:numPr><s:ilvl s:val="0"/><s:numId s:val="2"/>
            <s:numberingChange s:id="2" s:author="Strict" s:original="%1:1:2:."/>
          </s:numPr></s:pPr></s:p>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/>
            <w:numberingChange w:id="3" w:author="First" w:original="%1:1:0:."/>
            <w:numberingChange w:id="4" w:author="Duplicate" w:original="%1:1:0:."/>
          </w:numPr></w:pPr></w:p>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/>
            <w:numberingChange w:id="5" w:author="Malformed" w:original="not-a-definition"/>
          </w:numPr></w:pPr></w:p>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/>
            <w:numberingChange w:id="6" w:author="Unsupported" w:original="%1:1:23:•"/>
          </w:numPr></w:pPr></w:p>
          <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="6"/>
            <w:numberingChange w:id="7" w:author="Complex" w:original="%1:1:0:.%2:1:4:)"/>
          </w:numPr></w:pPr></w:p>
          <w:p><w:pPr><evil:numPr><evil:numberingChange evil:id="8" evil:author="Spoofed" evil:original="%1:1:0:."/></evil:numPr></w:pPr></w:p>
        </w:body>
      </w:document>
    `);
    const changes = descendants(document, 'numberingChange');
    expect(changes.map(isSupportedDocxNumberingChange)).toEqual([
      true,
      true,
      false,
      false,
      false,
      false,
      false,
      false,
    ]);
    expect(markDocxNumberingChanges(document).groups).toHaveLength(2);

    const conflictingHtml = new DOMParser().parseFromString(
      '<ol><li><p>__A3S_WORK_NUMBERING_CHANGE_1__First</p></li><li><p>__A3S_WORK_NUMBERING_CHANGE_2__Second</p></li></ol>',
      'text/html',
    );
    applyImportedDocxNumberingChangeMarkers(conflictingHtml, {
      groups: [
        {
          markers: ['__A3S_WORK_NUMBERING_CHANGE_1__'],
          id: 'first',
          author: 'Ada',
          date: '',
          start: 1,
          level: 0,
          format: 0,
          suffix: '.',
        },
        {
          markers: ['__A3S_WORK_NUMBERING_CHANGE_2__'],
          id: 'second',
          author: 'Grace',
          date: '',
          start: 2,
          level: 0,
          format: 0,
          suffix: '.',
        },
      ],
    });
    expect(conflictingHtml.querySelector('ol')?.dataset.changeKind).toBe(
      undefined,
    );
    expect(conflictingHtml.body.textContent).toBe('FirstSecond');

    const supported = await revisionCompatibility(`
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>
        <w:numberingChange w:id="1" w:author="Ada" w:original="%1:1:0:."/>
      </w:numPr></w:pPr><w:r><w:t>Supported</w:t></w:r></w:p>
    `);
    expect(supported.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.numbering',
        severity: 'info',
      }),
    );
    expect(
      supported.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);

    const unsupported = await revisionCompatibility(`
      <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>
        <w:numberingChange w:id="1" w:author="Ada" w:original="broken"/>
      </w:numPr></w:pPr><w:r><w:t>Unsupported</w:t></w:r></w:p>
    `);
    expect(unsupported.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.structural',
        severity: 'warning',
      }),
    );
  });
});

function textPosition(editor: Editor, text: string): number {
  let match = -1;
  editor.state.doc.descendants((node, position) => {
    if (match >= 0 || !node.isText || !node.text?.includes(text)) return;
    match = position + node.text.indexOf(text);
  });
  if (match < 0) throw new Error(`Text not found: ${text}`);
  return match;
}

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
    new File([bytes], 'numbering-revision.docx'),
    [],
  );
}
