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
});
