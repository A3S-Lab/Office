import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import * as Y from 'yjs';
import {
  createArtifact,
  createArtifactBlob,
  createOfficeCollaborationSession,
  createOfficeDocumentCollaborationBinding,
  importOfficeFile,
  initializeOfficeDocumentCollaboration,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  isSupportedDocxMoveChange,
  markDocxTextChanges,
  supportedDocxMovePairCount,
} from '../src/internal/features/work/work-docx-change-import';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

describe('DOCX move revisions', () => {
  test('recognizes a bounded paired text move and rejects malformed shapes', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:s="${STRICT_WORD_NAMESPACE}" xmlns:evil="https://example.test/evil">
        <w:body>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
            <w:r><w:t>middle</w:t></w:r>
          </w:p>
          <s:p><s:moveTo s:id="7" s:author="Ada" s:date="2026-09-01T00:00:00Z"><s:r><s:t>old</s:t></s:r></s:moveTo></s:p>
          <w:p><w:moveFrom w:id="8" w:author="Ada"><w:r><w:delText>unpaired</w:delText></w:r></w:moveFrom></w:p>
          <w:p><w:moveTo w:id="9" w:author="Ada"><w:r><w:t><evil:bad/></w:t></w:r></w:moveTo></w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([
      true,
      true,
      true,
      false,
    ]);
    expect(supportedDocxMovePairCount(document)).toBe(1);

    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        id: 'docx-move-7',
        kind: 'move',
        moveRole: 'from',
        author: 'Ada',
      }),
      expect.objectContaining({
        id: 'docx-move-7',
        kind: 'move',
        moveRole: 'to',
      }),
    ]);
    expect(document.documentElement.textContent).toContain(
      '__A3S_WORK_CHANGE_START_',
    );
    expect(descendants(document, 'moveFrom')).toHaveLength(1);
    expect(descendants(document, 'moveTo')).toHaveLength(1);
  });

  test('reviews both sides atomically in the editor', () => {
    const html = [
      '<p>',
      '<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-id="move-1" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</del>',
      ' middle ',
      '<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-id="move-1" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</ins>',
      '</p>',
    ].join('');
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html,
    });
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'move-1',
        kind: 'move',
        text: 'old',
        author: 'Ada',
      }),
    ]);

    expect(editor.commands.rejectDocumentChange('move-1')).toBe(true);
    expect(editor.getText()).toBe('old middle ');
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.commands.acceptDocumentChange('move-1')).toBe(true);
    expect(editor.getText()).toContain('middle old');
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
    editor.destroy();
  });

  test('resolves a move when the destination appears before the source', () => {
    const html = [
      '<p>目标：<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-id="move-reversed" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</ins></p>',
      '<p>源：<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-id="move-reversed" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</del></p>',
    ].join('');
    const accepted = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html,
    });
    expect(accepted.commands.acceptDocumentChange('move-reversed')).toBe(true);
    expect(accepted.getText()).toBe('目标：old\n\n源：');
    accepted.destroy();

    const rejected = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html,
    });
    expect(rejected.commands.rejectDocumentChange('move-reversed')).toBe(true);
    expect(rejected.getText()).toBe('目标：\n\n源：old');
    rejected.destroy();
  });

  test('round-trips native moveFrom and moveTo wrappers', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-id="move-7" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</del>',
      ' middle ',
      '<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-id="move-7" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</ins>',
      '</p></section>',
    ].join('');
    artifact.content.trackChanges = true;
    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(/<w:moveFrom\b[^>]*w:id="1"[^>]*w:author="Ada"/);
    expect(xml).toMatch(/<w:moveTo\b[^>]*w:id="1"[^>]*w:author="Ada"/);
    expect(xml).not.toContain('w:id="-1"');
    expect(xml).not.toContain('data-change-move-role');

    const reopened = await importOfficeFile(
      new File([blob], 'move-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="move"');
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    const change = collectDocumentChanges(editor.state.doc).find(
      (candidate) => candidate.kind === 'move',
    );
    expect(change).toEqual(expect.objectContaining({ author: 'Ada' }));
    expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
    expect(editor.getText()).toContain('old middle');
    editor.destroy();
  });

  test('assigns one stable native date when a browser move omits its date', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-id="move-no-date" data-change-author="Ada">old</del>',
      ' middle ',
      '<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-id="move-no-date" data-change-author="Ada">old</ins>',
      '</p></section>',
    ].join('');
    artifact.content.trackChanges = true;
    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    const dates = Array.from(xml.matchAll(/w:date="([^"]*)"/g)).map(
      (match) => match[1],
    );
    expect(dates.filter(Boolean)).toHaveLength(2);
    expect(new Set(dates.filter(Boolean)).size).toBe(1);
  });

  test('records one immutable collaboration decision for both move sides', () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const document = new Y.Doc();
    const session = createOfficeCollaborationSession({
      actor: { id: 'grace', name: 'Grace Editor' },
      artifactId: 'document-move-decision',
      document,
      kind: 'document',
      mode: 'edit',
    });
    initializeOfficeDocumentCollaboration(session, {
      ...artifact.content,
      html: [
        '<section data-document-section="true"><p>',
        '<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-id="move-collaboration-1" data-change-actor-id="ada" data-change-author="Ada Reviewer" data-change-date="2026-09-01T00:00:00.000Z">old</del>',
        ' middle ',
        '<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-id="move-collaboration-1" data-change-actor-id="ada" data-change-author="Ada Reviewer" data-change-date="2026-09-01T00:00:00.000Z">old</ins>',
        '</p></section>',
      ].join(''),
      model: undefined,
      trackChanges: true,
    });
    const binding = createOfficeDocumentCollaborationBinding(session);
    const editor = new Editor({ extensions: binding.extensions });

    expect(
      binding.decideChanges(editor, ['move-collaboration-1'], 'accept', {
        decidedAt: '2026-09-01T00:01:00.000Z',
      }),
    ).toBe(true);
    expect(editor.getText()).toContain('middle old');
    expect(binding.content().changeDecisions).toEqual([
      {
        id: 'move:move-collaboration-1',
        changeId: 'move-collaboration-1',
        changeKind: 'move',
        suggestedByActorId: 'ada',
        suggestedBy: 'Ada Reviewer',
        suggestedAt: '2026-09-01T00:00:00.000Z',
        text: 'old',
        decision: 'accept',
        decidedByActorId: 'grace',
        decidedBy: 'Grace Editor',
        decidedAt: '2026-09-01T00:01:00.000Z',
      },
    ]);

    editor.destroy();
    binding.destroy();
    session.destroy();
    document.destroy();
  });

  test('reports supported moves without a structural warning', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:moveFrom w:id="2" w:author="Ada"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveTo w:id="2" w:author="Ada"><w:r><w:t>old</w:t></w:r></w:moveTo></w:p></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'move-revision.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.move',
        severity: 'info',
      }),
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);
  });

  test('reports unpaired move sides as structural compatibility issues', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:moveFrom w:id="3" w:author="Ada"><w:r><w:delText>old</w:delText></w:r></w:moveFrom></w:p></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'unpaired-move-revision.docx'),
      [],
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(true);
  });

  test('admits soft breaks and relationship-free hyperlinks inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:br/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:hyperlink w:anchor="Target" w:tooltip="Jump">
                <w:r><w:t>old</w:t><w:br/><w:t>line</w:t></w:r>
              </w:hyperlink>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
    expect(markDocxTextChanges(document).changes).toEqual([
      expect.objectContaining({ kind: 'move', moveRole: 'from' }),
      expect.objectContaining({ kind: 'move', moveRole: 'to' }),
    ]);
  });

  test('admits tabs and hyphen glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="17" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:tab/>
                <w:noBreakHyphen/><w:softHyphen/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="17" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:tab/>
                <w:noBreakHyphen/><w:softHyphen/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects attributed tab glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="18" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:tab w:val="left"/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="18" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:t>old</w:t><w:tab w:val="left"/><w:t>line</w:t></w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([false, false]);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('admits empty carriage-return glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="19" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:cr/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="19" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:cr/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects attributed carriage-return glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="20" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:cr w:val="1"/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="20" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:t>old</w:t><w:cr w:val="1"/><w:t>line</w:t></w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([false, false]);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('admits empty lastRenderedPageBreak glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="21" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:lastRenderedPageBreak/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="21" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:lastRenderedPageBreak/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects attributed lastRenderedPageBreak glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="22" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:lastRenderedPageBreak w:val="1"/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="22" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:t>old</w:t><w:lastRenderedPageBreak w:val="1"/><w:t>line</w:t></w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([false, false]);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('admits empty page-number and date-field glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="23" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:pgNum/><w:dayLong/><w:monthLong/><w:yearLong/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="23" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:pgNum/><w:dayLong/><w:monthLong/><w:yearLong/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects attributed page-number glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="24" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:pgNum w:val="1"/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="24" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:t>old</w:t><w:pgNum w:val="1"/><w:t>line</w:t></w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([false, false]);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('admits empty short date-field glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="25" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:dayShort/><w:monthShort/><w:yearShort/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="25" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:dayShort/><w:monthShort/><w:yearShort/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects attributed short date-field glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="26" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:yearShort w:val="1"/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="26" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:t>old</w:t><w:yearShort w:val="1"/><w:t>line</w:t></w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([false, false]);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('admits empty footnoteRef glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="27" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:footnoteRef/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="27" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:footnoteRef/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('admits empty endnoteRef glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="29" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:endnoteRef/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="29" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:endnoteRef/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('admits empty annotationRef glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="30" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:annotationRef/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="30" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:annotationRef/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('admits empty separator glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="31" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:separator/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="31" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:separator/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('admits empty continuationSeparator glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="32" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:delText>old</w:delText><w:continuationSeparator/><w:delText>line</w:delText>
              </w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="32" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r>
                <w:t>old</w:t><w:continuationSeparator/><w:t>line</w:t>
              </w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects attributed footnoteRef glyphs inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="28" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText><w:footnoteRef w:val="1"/><w:delText>line</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="28" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:t>old</w:t><w:footnoteRef w:val="1"/><w:t>line</w:t></w:r>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([false, false]);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('admits relationship-free bookmarks inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFrom w:id="8" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:bookmarkStart w:id="3" w:name="Moved"/>
              <w:r><w:delText>old</w:delText></w:r>
              <w:bookmarkEnd w:id="3"/>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="8" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:bookmarkStart w:id="4" w:name="Moved"/>
              <w:r><w:t>old</w:t></w:r>
              <w:bookmarkEnd w:id="4"/>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    const moves = [
      ...descendants(document, 'moveFrom'),
      ...descendants(document, 'moveTo'),
    ];
    expect(moves.map(isSupportedDocxMoveChange)).toEqual([true, true]);
    expect(supportedDocxMovePairCount(document)).toBe(1);
  });

  test('rejects relationship-spoofed bookmarks inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:moveFrom w:id="8" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="8" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:bookmarkStart w:id="4" w:name="Spoof" r:id="rId1"/>
              <w:r><w:t>old</w:t></w:r>
              <w:bookmarkEnd w:id="4"/>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('rejects relationship-bound hyperlinks inside move revisions', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:r><w:delText>old</w:delText></w:r>
            </w:moveFrom>
          </w:p>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z">
              <w:hyperlink r:id="rId5"><w:r><w:t>old</w:t></w:r></w:hyperlink>
            </w:moveTo>
          </w:p>
        </w:body>
      </w:document>
    `);
    expect(supportedDocxMovePairCount(document)).toBe(0);
  });

  test('exports move revisions that include soft line breaks', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-id="move-br" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old<br>line</del>',
      ' middle ',
      '<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-id="move-br" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old<br>line</ins>',
      '</p></section>',
    ].join('');
    artifact.content.trackChanges = true;
    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(/<w:moveFrom\b/);
    expect(xml).toMatch(/<w:moveTo\b/);
    expect(xml).toMatch(/<w:br\b/);
  });

  test('imports companion move-range bookmarks without a structural warning', async () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:p>
            <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
            <w:moveFromRangeEnd w:id="0"/>
          </w:p>
          <w:p>
            <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
            <w:moveToRangeEnd w:id="0"/>
          </w:p>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'moveToRangeEnd')).toHaveLength(0);

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="0"/></w:p><w:p><w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="0"/></w:p></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'move-range-companion.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.move',
        severity: 'info',
      }),
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.move-range'),
    ).toBe(false);
  });

  test('round-trips companion move-range bookmarks with native move wrappers', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<del data-document-change="true" data-change-kind="move" data-change-move-role="from" data-change-move-range-id="0" data-change-move-range-name="move0" data-change-id="move-7" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</del>',
      ' middle ',
      '<ins data-document-change="true" data-change-kind="move" data-change-move-role="to" data-change-move-range-id="0" data-change-move-range-name="move0" data-change-id="move-7" data-change-author="Ada" data-change-date="2026-09-01T00:00:00.000Z">old</ins>',
      '</p></section>',
    ].join('');
    artifact.content.trackChanges = true;
    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(
      /<w:moveFromRangeStart\b[^>]*w:id="0"[^>]*w:name="move0"/,
    );
    expect(xml).toMatch(/<w:moveFrom\b[^>]*w:author="Ada"/);
    expect(xml).toMatch(/<w:moveFromRangeEnd\b[^>]*w:id="0"/);
    expect(xml).toMatch(
      /<w:moveToRangeStart\b[^>]*w:id="0"[^>]*w:name="move0"/,
    );
    expect(xml).toMatch(/<w:moveToRangeEnd\b[^>]*w:id="0"/);

    const reopened = await importOfficeFile(
      new File([blob], 'move-range-roundtrip.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-move-range-id="0"');
    expect(reopened.content.html).toContain(
      'data-change-move-range-name="move0"',
    );
  });

  test('imports same-section cross-paragraph companion move-range bookmarks', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
          </w:p>
          <w:moveFromRangeEnd w:id="0"/>
          <w:p><w:r><w:t>stable</w:t></w:r></w:p>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(document.documentElement.textContent).toContain('stable');
  });

  test('rejects move-range bookmarks that sandwich extra sibling content', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
          </w:p>
          <w:p><w:r><w:t>extra</w:t></w:r></w:p>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('imports companion move-range bookmarks that cross a section boundary', async () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
          </w:p>
          <w:moveFromRangeEnd w:id="0"/>
          <w:p><w:pPr><w:sectPr/></w:pPr></w:p>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'sectPr')).toHaveLength(1);

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:p><w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom></w:p><w:moveFromRangeEnd w:id="0"/><w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:p><w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo></w:p><w:moveToRangeEnd w:id="0"/></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'cross-section-move-range.docx'),
      [],
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.move-range'),
    ).toBe(false);
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.move',
      }),
    );
  });

  test('rejects companion move-range bookmarks that sandwich a section break with the move', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
          </w:p>
          <w:p><w:pPr><w:sectPr/></w:pPr></w:p>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('imports companion move-range bookmarks around a single-cell table move', async () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:trPr/>
              <w:tc>
                <w:tcPr/>
                <w:p>
                  <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'tbl')).toHaveLength(1);

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:tbl><w:tblPr/><w:tblGrid><w:tblGridCol/></w:tblGrid><w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom></w:p></w:tc></w:tr></w:tbl><w:moveFromRangeEnd w:id="0"/><w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:p><w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo></w:p><w:moveToRangeEnd w:id="0"/></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'single-cell-table-move-range.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.move',
        severity: 'info',
      }),
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.move-range'),
    ).toBe(false);
  });

  test('imports companion move-range bookmarks when both sides use single-cell tables', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:p>
                  <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:p>
                  <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'tbl')).toHaveLength(2);
  });

  test('imports companion move-range bookmarks around a multi-cell table with sibling cell text', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:p>
                  <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                </w:p>
              </w:tc>
              <w:tc>
                <w:tcPr/>
                <w:p><w:r><w:t>extra</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'tbl')).toHaveLength(1);
  });

  test('rejects companion move-range bookmarks when a sibling cell carries a tracked revision', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:p>
                  <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                </w:p>
              </w:tc>
              <w:tc>
                <w:tcPr/>
                <w:p>
                  <w:ins w:id="3" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>extra</w:t></w:r></w:ins>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('imports companion move-range bookmarks around a one-level nested table move', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:tbl>
                  <w:tblPr/>
                  <w:tblGrid><w:tblGridCol/><w:tblGridCol/></w:tblGrid>
                  <w:tr>
                    <w:tc>
                      <w:tcPr/>
                      <w:p>
                        <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                      </w:p>
                    </w:tc>
                    <w:tc>
                      <w:tcPr/>
                      <w:p><w:r><w:t>extra</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'tbl')).toHaveLength(2);
  });

  test('imports companion move-range bookmarks when a nested table sits beside the move', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:tbl>
                  <w:tblPr/>
                  <w:tblGrid><w:tblGridCol/></w:tblGrid>
                  <w:tr>
                    <w:tc>
                      <w:tcPr/>
                      <w:p><w:r><w:t>nested</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
                <w:p>
                  <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'tbl')).toHaveLength(2);
  });

  test('rejects companion move-range bookmarks when a beside nested table carries a tracked revision', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:tbl>
                  <w:tblPr/>
                  <w:tblGrid><w:tblGridCol/></w:tblGrid>
                  <w:tr>
                    <w:tc>
                      <w:tcPr/>
                      <w:p>
                        <w:ins w:id="9" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>nested</w:t></w:r></w:ins>
                      </w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
                <w:p>
                  <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                </w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('rejects companion move-range bookmarks that sandwich deeper nested tables', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:tblGridCol/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:tbl>
                  <w:tblPr/>
                  <w:tblGrid><w:tblGridCol/></w:tblGrid>
                  <w:tr>
                    <w:tc>
                      <w:tcPr/>
                      <w:tbl>
                        <w:tblPr/>
                        <w:tblGrid><w:tblGridCol/></w:tblGrid>
                        <w:tr>
                          <w:tc>
                            <w:tcPr/>
                            <w:p>
                              <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                            </w:p>
                          </w:tc>
                        </w:tr>
                      </w:tbl>
                    </w:tc>
                  </w:tr>
                </w:tbl>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('rejects companion move-range bookmarks that sandwich an SDT with a nested table', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:sdt>
            <w:sdtContent>
              <w:tbl>
                <w:tblPr/>
                <w:tblGrid><w:tblGridCol/></w:tblGrid>
                <w:tr>
                  <w:tc>
                    <w:tcPr/>
                    <w:tbl>
                      <w:tblPr/>
                      <w:tblGrid><w:tblGridCol/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr/>
                          <w:p>
                            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                          </w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                  </w:tc>
                </w:tr>
              </w:tbl>
            </w:sdtContent>
          </w:sdt>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('imports companion move-range bookmarks around a simple SDT paragraph move', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:sdt>
            <w:sdtPr><w:alias w:val="moved"/></w:sdtPr>
            <w:sdtContent>
              <w:p>
                <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
              </w:p>
            </w:sdtContent>
          </w:sdt>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'sdt')).toHaveLength(1);
  });

  test('imports companion move-range bookmarks around a simple SDT table move', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:sdt>
            <w:sdtContent>
              <w:tbl>
                <w:tblPr/>
                <w:tblGrid><w:tblGridCol/></w:tblGrid>
                <w:tr>
                  <w:tc>
                    <w:tcPr/>
                    <w:p>
                      <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                    </w:p>
                  </w:tc>
                </w:tr>
              </w:tbl>
            </w:sdtContent>
          </w:sdt>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes).toEqual([
      expect.objectContaining({
        kind: 'move',
        moveRole: 'from',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
      expect.objectContaining({
        kind: 'move',
        moveRole: 'to',
        moveRangeId: '0',
        moveRangeName: 'move0',
      }),
    ]);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(0);
    expect(descendants(document, 'sdt')).toHaveLength(1);
    expect(descendants(document, 'tbl')).toHaveLength(1);
  });

  test('rejects companion move-range bookmarks that sandwich nested SDTs', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:sdt>
            <w:sdtContent>
              <w:sdt>
                <w:sdtContent>
                  <w:p>
                    <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
                  </w:p>
                </w:sdtContent>
              </w:sdt>
            </w:sdtContent>
          </w:sdt>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('rejects companion move-range bookmarks when an SDT sits beside the move', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:sdt>
            <w:sdtContent>
              <w:p><w:r><w:t>extra</w:t></w:r></w:p>
            </w:sdtContent>
          </w:sdt>
          <w:p>
            <w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom>
          </w:p>
          <w:moveFromRangeEnd w:id="0"/>
          <w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/>
          <w:p>
            <w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo>
          </w:p>
          <w:moveToRangeEnd w:id="0"/>
        </w:body>
      </w:document>
    `);
    const markers = markDocxTextChanges(document);
    expect(markers.changes.every((change) => !change.moveRangeId)).toBe(true);
    expect(descendants(document, 'moveFromRangeStart')).toHaveLength(1);
  });

  test('reports unpaired move-range markers when companions cross a section only as diagnostics for leftover shapes', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:moveFromRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:p><w:moveFrom w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>old</w:delText></w:r></w:moveFrom></w:p><w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:moveFromRangeEnd w:id="0"/><w:moveToRangeStart w:id="0" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="move0"/><w:p><w:moveTo w:id="7" w:author="Ada" w:date="2026-09-01T00:00:00Z"><w:r><w:t>old</w:t></w:r></w:moveTo></w:p><w:moveToRangeEnd w:id="0"/></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'section-sandwich-move-range.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.move-range',
      }),
    );
  });

  test('reports unpaired move-range markers with a dedicated diagnostic', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:moveFromRangeStart w:id="9" w:author="Ada" w:date="2026-09-01T00:00:00Z" w:name="orphan"/><w:r><w:t>loose</w:t></w:r><w:moveFromRangeEnd w:id="9"/></w:p></w:body></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'unpaired-move-range.docx'),
      [],
    );
    expect(report.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.move-range',
      }),
    );
    expect(
      report.issues.some(({ code }) => code === 'docx.revisions.structural'),
    ).toBe(false);
  });
});
