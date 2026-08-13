import { expect, test } from '@rstest/core';
import { Editor, Extension, Node } from '@tiptap/core';
import * as Y from 'yjs';
import {
  createArtifact,
  createOfficeCollaborationSession,
  createOfficeDocumentCollaborationBinding,
  type OfficeCollaborationOrigin,
  type DocumentContent,
  initializeOfficeDocumentCollaboration,
  readOfficeCollaborationMetadata,
  readOfficeDocumentCollaboration,
} from '../src/core';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';

test('initializes a structured Document collaboration document', () => {
  const source = documentFixture();
  const session = createOfficeCollaborationSession({
    artifactId: 'document-1',
    kind: 'document',
  });

  const initialized = initializeOfficeDocumentCollaboration(session, source);

  expect(initialized.initialized).toBe(true);
  expect(initialized.content.model?.root.type).toBe('doc');
  expect(initialized.content.html).toContain('Shared document');
  expect(initialized.content.pageColor).toBe('#F8FAFC');
  expect(initialized.content.trackChanges).toBe(true);
  expect(initialized.content.comments).toEqual(source.comments);
  expect(initialized.content.bibliography).toEqual(source.bibliography);
  expect(readOfficeCollaborationMetadata(session)).toMatchObject({
    artifactId: 'document-1',
    initialized: true,
    kind: 'document',
  });

  expect(
    initializeOfficeDocumentCollaboration(session, {
      ...source,
      html: '<p>Must not overwrite</p>',
    }),
  ).toMatchObject({
    initialized: false,
    content: { pageColor: '#F8FAFC' },
  });
  expect(readOfficeDocumentCollaboration(session).html).toContain(
    'Shared document',
  );
});

test('rejects unattributed Document roots before writing metadata', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-unattributed',
    kind: 'document',
  });
  session.document
    .getMap(session.rootName('document.options'))
    .set('pageColor', '#FFFFFF');

  expect(() =>
    initializeOfficeDocumentCollaboration(session, documentFixture()),
  ).toThrow(/sidecars contain data without initialized metadata/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
});

test('rejects an unsupported structured model before bootstrap writes', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-invalid-model',
    kind: 'document',
  });
  const prepared = createWorkDocumentModelFromContent(documentFixture());
  if (!prepared.model) throw new Error('Expected a structured fixture.');

  expect(() =>
    initializeOfficeDocumentCollaboration(session, {
      ...prepared,
      model: {
        ...prepared.model,
        root: {
          type: 'doc',
          content: [{ type: 'unsupportedDocumentNode' }],
        },
      },
    }),
  ).toThrow(/Office document schema/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
  expect(
    session.document.getArray(session.rootName('bootstrap.initializers'))
      .length,
  ).toBe(0);
});

test('detects concurrent independent Document bootstrap', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'document-bootstrap-race',
    document: firstDocument,
    kind: 'document',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'document-bootstrap-race',
    document: secondDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  initializeOfficeDocumentCollaboration(second, {
    ...documentFixture(),
    html: '<p>Concurrent seed</p>',
    model: undefined,
  });

  exchangeUpdates(firstDocument, secondDocument);

  expect(() => readOfficeDocumentCollaboration(first)).toThrow(
    /Multiple clients initialized/,
  );
  expect(() => readOfficeDocumentCollaboration(second)).toThrow(
    /Multiple clients initialized/,
  );
});

test('deduplicates identical stable-ID retries from disconnected clients', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'document-comment-retry',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'document-comment-retry',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const retriedComment = {
    id: 'comment-retried',
    author: 'Ada',
    date: '2026-08-13T02:00:00.000Z',
    text: 'Exactly once in the canonical snapshot.',
    resolved: false,
  };
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    comments: [...(firstBefore.comments ?? []), retriedComment],
  });
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    comments: [...(secondBefore.comments ?? []), retriedComment],
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(
    converged.comments?.filter(({ id }) => id === retriedComment.id),
  ).toEqual([retriedComment]);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('rejects conflicting stable-ID assignments after offline synchronization', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'document-comment-id-collision',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'document-comment-id-collision',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();
  const collisionBase = {
    id: 'comment-collision',
    author: 'Ada',
    date: '2026-08-13T02:30:00.000Z',
    resolved: false,
  };

  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    comments: [
      ...(firstBefore.comments ?? []),
      { ...collisionBase, text: 'First assignment' },
    ],
  });
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    comments: [
      ...(secondBefore.comments ?? []),
      { ...collisionBase, text: 'Conflicting assignment' },
    ],
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(() => firstBinding.content()).toThrow(
    /concurrently assigned to different records/,
  );
  expect(() => secondBinding.content()).toThrow(
    /concurrently assigned to different records/,
  );

  firstBinding.destroy();
  secondBinding.destroy();
});

test('merges conflict-local Document sidecars from concurrent clients', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-sidecars',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-sidecars',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();
  if (!secondBefore.bibliography) {
    throw new Error('Expected the Document bibliography fixture.');
  }

  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    comments: [
      ...(firstBefore.comments ?? []),
      {
        id: 'comment-ada',
        author: 'Ada',
        date: '2026-08-13T01:00:00.000Z',
        text: 'Ada comment',
        resolved: false,
      },
    ],
    pageColor: '#FFF7ED',
  });
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    bibliography: {
      ...secondBefore.bibliography,
      sources: [
        ...secondBefore.bibliography.sources,
        {
          id: 'source-grace',
          tag: 'Lovelace1843',
          sourceType: 'bookSection',
          title: 'Notes on the Analytical Engine',
        },
      ],
    },
    trackChanges: false,
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.pageColor).toBe('#FFF7ED');
  expect(converged.trackChanges).toBe(false);
  expect(converged.comments?.map((comment) => comment.id)).toContain(
    'comment-ada',
  );
  expect(converged.bibliography?.sources.map((source) => source.id)).toContain(
    'source-grace',
  );

  firstBinding.destroy();
  secondBinding.destroy();
});

test('preserves a concurrent bibliography addition when another client removes its snapshot', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'document-bibliography-delete-add',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'document-bibliography-delete-add',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();
  if (!secondBefore.bibliography) {
    throw new Error('Expected the Document bibliography fixture.');
  }

  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    bibliography: undefined,
  });
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    bibliography: {
      ...secondBefore.bibliography,
      sources: [
        ...secondBefore.bibliography.sources,
        {
          id: 'source-concurrent',
          tag: 'Hopper1952',
          sourceType: 'report',
          title: 'The Education of a Computer',
        },
      ],
    },
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.bibliography?.sources).toEqual([
    expect.objectContaining({ id: 'source-concurrent' }),
  ]);

  firstBinding.updateSidecars(converged, {
    ...converged,
    bibliography: converged.bibliography
      ? { ...converged.bibliography, sources: [] }
      : undefined,
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  expect(firstBinding.content().bibliography?.sources).toEqual([]);
  expect(secondBinding.content()).toEqual(firstBinding.content());

  firstBinding.destroy();
  secondBinding.destroy();
});

test('does not resurrect a remotely removed comment from an unrelated stale snapshot', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'document-comment-delete-option',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'document-comment-delete-option',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.updateSidecars(firstBefore, { ...firstBefore, comments: [] });
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    pageColor: '#E0F2FE',
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.comments).toEqual([]);
  expect(converged.pageColor).toBe('#E0F2FE');

  firstBinding.destroy();
  secondBinding.destroy();
});

test('rejects an edit to a comment concurrently removed by another client', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'document-comment-delete-edit',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'document-comment-delete-edit',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.updateSidecars(firstBefore, { ...firstBefore, comments: [] });
  Y.applyUpdate(
    secondDocument,
    Y.encodeStateAsUpdate(firstDocument, Y.encodeStateVector(secondDocument)),
    'test-network',
  );
  expect(() =>
    secondBinding.updateSidecars(stale, {
      ...stale,
      comments: stale.comments?.map((comment) => ({
        ...comment,
        text: 'Stale edit',
      })),
    }),
  ).toThrow(/removed before this change/);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('keeps Document sidecar undo local to one client', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-local-undo',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-local-undo',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    pageColor: '#FEF3C7',
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.canUndo()).toBe(true);
  expect(secondBinding.canUndo()).toBe(false);
  const secondBefore = secondBinding.content();
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    trackChanges: false,
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  const converged = firstBinding.content();
  expect(converged.pageColor).toBe('#F8FAFC');
  expect(converged.trackChanges).toBe(false);
  expect(secondBinding.content()).toEqual(converged);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('synchronizes TipTap edits and only undoes the local client', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-editor-undo',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-editor-undo',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstChanges: Array<{ local: boolean; origin: unknown }> = [];
  const secondChanges: Array<{ local: boolean; origin: unknown }> = [];
  firstBinding.subscribe(({ local, origin }) =>
    firstChanges.push({ local, origin }),
  );
  secondBinding.subscribe(({ local, origin }) =>
    secondChanges.push({ local, origin }),
  );
  const firstEditor = new Editor({ extensions: firstBinding.extensions });
  const secondEditor = new Editor({ extensions: secondBinding.extensions });

  firstEditor
    .chain()
    .setTextSelection(firstEditor.state.doc.content.size - 2)
    .insertContent(' by Ada')
    .run();
  await flushMicrotasks();
  expect(firstBinding.canUndo()).toBe(true);
  expect(secondBinding.canUndo()).toBe(false);
  expect(firstChanges).toContainEqual({
    local: true,
    origin: firstBinding.origin,
  });

  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  expect(secondEditor.getText()).toContain('by Ada');
  expect(secondChanges.some((change) => change.local === false)).toBe(true);

  secondBinding.stopCapturing();
  secondEditor
    .chain()
    .setTextSelection(secondEditor.state.doc.content.size - 2)
    .insertContent(' and Grace')
    .run();
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstEditor.getText()).toContain('and Grace');
  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  expect(firstEditor.getText()).not.toContain('by Ada');
  expect(firstEditor.getText()).toContain('and Grace');
  expect(secondEditor.getText()).toBe(firstEditor.getText());

  firstEditor.destroy();
  secondEditor.destroy();
  firstBinding.destroy();
  secondBinding.destroy();
});

test('drops a queued Document snapshot once destruction is requested', async () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-destroy-queued-change',
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(session, documentFixture());
  const binding = createOfficeDocumentCollaborationBinding(session);
  const editor = new Editor({ extensions: binding.extensions });
  const changes: unknown[] = [];
  binding.subscribe((change) => changes.push(change));

  editor.commands.insertContent('Queued before teardown');
  binding.destroy();
  await flushMicrotasks();

  expect(changes).toEqual([]);
  expect(() => binding.content()).toThrow(/binding has been destroyed/);

  editor.destroy();
  await flushMicrotasks();
});

test('rejects a forged Document binding origin at runtime', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-invalid-origin',
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(session, documentFixture());

  expect(() =>
    createOfficeDocumentCollaborationBinding(session, {
      origin: {
        protocol: 'a3s.office.collaboration',
        kind: 'untrusted',
      } as unknown as OfficeCollaborationOrigin,
    }),
  ).toThrow(/origin kind 'untrusted' is invalid/);
});

test('rejects canonical Document mutation outside edit mode', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-read-only',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, documentFixture());
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'document-read-only',
    document,
    kind: 'document',
    mode: 'view',
  });
  const binding = createOfficeDocumentCollaborationBinding(readOnly);
  const before = binding.content();

  expect(() =>
    binding.updateSidecars(before, { ...before, pageColor: '#000000' }),
  ).toThrow(/cannot modify canonical content/);
  expect(() => binding.undo()).toThrow(/cannot modify canonical content/);
  expect(binding.content().pageColor).toBe('#F8FAFC');

  binding.destroy();
});

test('rejects custom shared nodes without an Office schema migration', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-custom-node',
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(session, documentFixture());
  const customNode = Node.create({
    name: 'hostSharedNode',
    group: 'block',
    parseHTML: () => [{ tag: 'host-shared-node' }],
    renderHTML: () => ['host-shared-node', 0],
  });

  expect(() =>
    createOfficeDocumentCollaborationBinding(session, {
      additionalExtensions: [customNode],
    }),
  ).toThrow(/versioned Office schema migration/);
});

test('rejects behavior extensions that add shared schema attributes', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'document-global-attribute',
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(session, documentFixture());
  const attributeExtension = Extension.create({
    name: 'hostGlobalAttribute',
    addGlobalAttributes: () => [
      {
        types: ['paragraph'],
        attributes: { hostValue: { default: null } },
      },
    ],
  });

  expect(() =>
    createOfficeDocumentCollaborationBinding(session, {
      additionalExtensions: [attributeExtension],
    }),
  ).toThrow(/schema-changing extension/);
});

function documentFixture(): DocumentContent {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a Document fixture.');
  }
  return {
    ...artifact.content,
    html: '<p>Shared document</p>',
    pageColor: '#F8FAFC',
    trackChanges: true,
    comments: [
      {
        id: 'comment-1',
        author: 'Ada',
        date: '2026-08-13T00:00:00.000Z',
        text: 'Keep this boundary.',
        resolved: false,
        replies: [
          {
            id: 'reply-1',
            author: 'Grace',
            date: '2026-08-13T00:01:00.000Z',
            text: 'Agreed.',
          },
        ],
      },
    ],
    bibliography: {
      style: 'apa',
      styleName: 'APA 7th',
      selectedStyle: 'apa',
      sources: [
        {
          id: 'source-1',
          tag: 'Turing1936',
          sourceType: 'journalArticle',
          title: 'On Computable Numbers',
          year: '1936',
          additionalFields: { doi: '10.1112/plms/s2-42.1.230' },
        },
      ],
    },
  };
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

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}
