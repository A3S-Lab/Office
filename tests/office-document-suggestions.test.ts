import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createArtifact,
  createOfficeCollaborationSession,
  createOfficeDocumentCollaborationBinding,
  type DocumentContent,
  initializeOfficeDocumentCollaboration,
  readOfficeDocumentCollaboration,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import type { WorkDocumentChangeDecision } from '../src/internal/features/work/work-types';

test('creates attributed deletion and replacement suggestions while preserving canonical text', () => {
  const document = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-suggestion-replacement',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    trackChanges: false,
  });
  const suggester = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada Reviewer' },
    artifactId: 'document-suggestion-replacement',
    document,
    kind: 'document',
    mode: 'suggest',
  });
  let sequence = 0;
  const binding = createOfficeDocumentCollaborationBinding(suggester, {
    workExtensions: {
      createChange: (kind) => ({
        actorId: 'ada',
        author: 'Ada Reviewer',
        date: '2026-08-17T09:00:00.000Z',
        id: `suggestion-${kind}-${++sequence}`,
      }),
      isTracking: () => true,
    },
  });
  const editor = new Editor({ extensions: binding.extensions });

  const replacement = documentTextRange(editor, 'Shared');
  expect(
    editor.commands.replaceDocumentTextWithTrackedChange(
      replacement.from,
      replacement.to,
      'Reviewed',
    ),
  ).toBe(true);
  const deletion = documentTextRange(editor, 'document');
  expect(
    editor.commands.replaceDocumentTextWithTrackedChange(
      deletion.from,
      deletion.to,
      '',
    ),
  ).toBe(true);

  expect(collectDocumentChanges(editor.state.doc)).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        actorId: 'ada',
        author: 'Ada Reviewer',
        kind: 'deletion',
        text: 'Shared',
      }),
      expect.objectContaining({
        actorId: 'ada',
        author: 'Ada Reviewer',
        kind: 'insertion',
        text: 'Reviewed',
      }),
      expect.objectContaining({
        actorId: 'ada',
        author: 'Ada Reviewer',
        kind: 'deletion',
        text: 'document',
      }),
    ]),
  );
  expect(editor.getHTML()).toContain('>Shared</del>');
  expect(editor.getHTML()).toContain('>Reviewed</ins>');
  expect(editor.getHTML()).toContain('>document</del>');

  editor.destroy();
  binding.destroy();
});

test('prevents a suggester from withdrawing or rewriting another actor suggestion', () => {
  const document = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-external-suggestion',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    html: [
      '<p>',
      '<ins data-document-change="true" data-change-kind="insertion" data-change-id="grace-insertion" data-change-actor-id="grace" data-change-author="Grace" data-change-date="2026-08-17T09:10:00.000Z">proposal</ins>',
      ' ',
      '<del data-document-change="true" data-change-kind="deletion" data-change-id="grace-deletion" data-change-actor-id="grace" data-change-author="Grace" data-change-date="2026-08-17T09:11:00.000Z">canonical</del>',
      ' ',
      '<ins data-document-change="true" data-change-kind="insertion" data-change-id="ada-insertion" data-change-actor-id="ada" data-change-author="Ada" data-change-date="2026-08-17T09:12:00.000Z">draft</ins>',
      '</p>',
    ].join(''),
    model: undefined,
    trackChanges: false,
  });
  const suggester = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-external-suggestion',
    document,
    kind: 'document',
    mode: 'suggest',
  });
  let sequence = 0;
  const binding = createOfficeDocumentCollaborationBinding(suggester, {
    workExtensions: {
      createChange: (kind) => ({
        actorId: 'ada',
        author: 'Ada',
        date: '2026-08-17T09:13:00.000Z',
        id: `ada-rewrite-${kind}-${++sequence}`,
      }),
      isTracking: () => true,
    },
  });
  const editor = new Editor({ extensions: binding.extensions });
  const original = editor.getHTML();

  expect(editor.commands.rejectDocumentChange('grace-insertion')).toBe(true);
  expect(editor.getHTML()).toBe(original);
  expect(editor.commands.rejectDocumentChange('grace-deletion')).toBe(true);
  expect(editor.getHTML()).toBe(original);

  const external = documentTextRange(editor, 'proposal');
  expect(
    editor.commands.replaceDocumentTextWithTrackedChange(
      external.from,
      external.to,
      'rewritten',
    ),
  ).toBe(true);
  expect(editor.getHTML()).toBe(original);

  expect(editor.commands.rejectDocumentChange('ada-insertion')).toBe(true);
  expect(editor.getHTML()).toContain('data-change-id="grace-insertion"');
  expect(editor.getHTML()).toContain('data-change-id="grace-deletion"');
  expect(editor.getHTML()).not.toContain('data-change-id="ada-insertion"');
  expect(editor.getText()).not.toContain('draft');

  editor.destroy();
  binding.destroy();
});

test('keeps existing formatting revisions immutable while admitting text suggestions', () => {
  const document = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-formatting-suggestion-boundary',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    html: [
      '<p>',
      '<span data-document-change="true" data-change-kind="formatting" data-change-before="[]" data-change-id="editor-formatting" data-change-actor-id="grace" data-change-author="Grace" data-change-date="2026-08-17T09:15:00.000Z"><strong>Shared</strong></span>',
      ' document',
      '</p>',
    ].join(''),
    model: undefined,
  });
  const suggester = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-formatting-suggestion-boundary',
    document,
    kind: 'document',
    mode: 'suggest',
  });
  const binding = createOfficeDocumentCollaborationBinding(suggester, {
    workExtensions: {
      createChange: (kind) => ({
        actorId: 'ada',
        author: 'Ada',
        date: '2026-08-17T09:16:00.000Z',
        id: `ada-${kind}`,
      }),
      isTracking: () => true,
    },
  });
  const editor = new Editor({ extensions: binding.extensions });
  const originalFormatting = editor.getHTML();
  const shared = documentTextRange(editor, 'Shared');

  expect(editor.chain().setTextSelection(shared).toggleBold().run()).toBe(true);
  expect(editor.getHTML()).toBe(originalFormatting);

  const target = documentTextRange(editor, 'document');
  expect(
    editor.commands.replaceDocumentTextWithTrackedChange(
      target.from,
      target.to,
      'proposal',
    ),
  ).toBe(true);
  expect(collectDocumentChanges(editor.state.doc)).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        id: 'editor-formatting',
        kind: 'formatting',
      }),
      expect.objectContaining({ kind: 'deletion', text: 'document' }),
      expect.objectContaining({ kind: 'insertion', text: 'proposal' }),
    ]),
  );

  editor.destroy();
  binding.destroy();
});

test('persists a formatting decision without deleting canonical text', () => {
  const document = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-formatting-decision',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    html: '<p><span data-document-change="true" data-change-kind="formatting" data-change-before="[]" data-change-id="formatting-bold" data-change-actor-id="ada" data-change-author="Ada" data-change-date="2026-08-17T09:17:00.000Z"><strong>Shared</strong></span> document</p>',
    model: undefined,
  });
  const editorSession = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-formatting-decision',
    document,
    kind: 'document',
  });
  const binding = createOfficeDocumentCollaborationBinding(editorSession);
  const editor = new Editor({ extensions: binding.extensions });

  expect(binding.decideChanges(editor, ['formatting-bold'], 'reject')).toBe(
    true,
  );
  expect(editor.getText()).toContain('Shared document');
  expect(editor.getHTML()).not.toContain('<strong>');
  expect(binding.content().changeDecisions).toEqual([
    expect.objectContaining({
      changeId: 'formatting-bold',
      changeKind: 'formatting',
      decision: 'reject',
      text: 'Shared',
    }),
  ]);

  editor.destroy();
  binding.destroy();
});

test('deduplicates an identical final-decision retry from disconnected editors', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-decision-retry',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-decision-retry',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();
  const decision = decisionFixture();

  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    changeDecisions: [decision],
  });
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    changeDecisions: [decision],
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.content().changeDecisions).toEqual([decision]);
  expect(secondBinding.content()).toEqual(firstBinding.content());

  firstBinding.destroy();
  secondBinding.destroy();
});

test('rejects a stale conflicting final decision after synchronization', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-decision-conflict',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'linus', name: 'Linus Editor' },
    artifactId: 'document-decision-conflict',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const staleSecondBefore = secondBinding.content();
  const accepted = decisionFixture();

  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    changeDecisions: [accepted],
  });
  applyMissingUpdate(firstDocument, secondDocument);

  const rejected: WorkDocumentChangeDecision = {
    ...accepted,
    decision: 'reject',
    decidedByActorId: 'linus',
    decidedBy: 'Linus Editor',
    decidedAt: '2026-08-17T09:21:00.000Z',
  };
  expect(() =>
    secondBinding.updateSidecars(staleSecondBefore, {
      ...staleSecondBefore,
      changeDecisions: [rejected],
    }),
  ).toThrow(
    /concurrently assigned to different records|different final decision/i,
  );
  expect(secondBinding.content().changeDecisions).toEqual([accepted]);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('reads a legacy Document session without decision roots', () => {
  const legacy = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-legacy-decisions',
    document: legacy,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, documentFixture());
  legacy.share.delete(bootstrap.rootName('document.change-decisions'));
  legacy.share.delete(bootstrap.rootName('document.change-decision-order'));

  const restored = new Y.Doc();
  Y.applyUpdate(restored, Y.encodeStateAsUpdate(legacy));
  const session = createOfficeCollaborationSession({
    artifactId: 'document-legacy-decisions',
    document: restored,
    kind: 'document',
  });

  expect(readOfficeDocumentCollaboration(session)).toMatchObject({
    html: expect.stringContaining('Shared document'),
    type: 'document',
  });
  expect(readOfficeDocumentCollaboration(session).changeDecisions).toBe(
    undefined,
  );
  expect(
    restored.getMap(session.rootName('document.change-decisions')).size,
  ).toBe(0);
  expect(
    restored.getArray(session.rootName('document.change-decision-order'))
      .length,
  ).toBe(0);
});

function documentFixture(): DocumentContent {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a Document fixture.');
  }
  return {
    ...artifact.content,
    html: '<p>Shared document</p>',
    model: undefined,
  };
}

function decisionFixture(): WorkDocumentChangeDecision {
  return {
    id: 'insertion:suggestion-stable',
    changeId: 'suggestion-stable',
    changeKind: 'insertion',
    suggestedByActorId: 'ada',
    suggestedBy: 'Ada Reviewer',
    suggestedAt: '2026-08-17T09:19:00.000Z',
    text: 'proposed wording',
    decision: 'accept',
    decidedByActorId: 'grace',
    decidedBy: 'Grace Editor',
    decidedAt: '2026-08-17T09:20:00.000Z',
  };
}

function documentTextRange(
  editor: Editor,
  text: string,
): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  editor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Unable to find "${text}" in the document.`);
  return range;
}

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}

function applyMissingUpdate(source: Y.Doc, target: Y.Doc): void {
  Y.applyUpdate(
    target,
    Y.encodeStateAsUpdate(source, Y.encodeStateVector(target)),
    'test-network',
  );
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}
