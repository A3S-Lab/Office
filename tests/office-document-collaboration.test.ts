import { expect, test } from '@rstest/core';
import { Editor, Extension, Node } from '@tiptap/core';
import * as Y from 'yjs';
import {
  createArtifact,
  createOfficeCollaborationSession,
  createOfficeDocumentCollaborationBinding,
  type DocumentContent,
  initializeOfficeDocumentCollaboration,
  type OfficeCollaborationOrigin,
  readOfficeCollaborationMetadata,
  readOfficeDocumentCollaboration,
} from '../src/core';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';
import { serializeDocumentNumberingChange } from '../src/internal/features/work/work-document-numbering-changes';
import { serializeDocumentParagraphFormatting } from '../src/internal/features/work/work-document-paragraph-format-changes';

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

test('allows actor-attributed comment-mode review mutations only', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-comment-mode-sidecars',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, documentFixture());
  const reviewer = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-comment-mode-sidecars',
    document,
    kind: 'document',
    mode: 'comment',
  });
  const binding = createOfficeDocumentCollaborationBinding(reviewer);
  const before = binding.content();
  const comment = {
    id: 'comment-by-ada',
    actorId: 'ada',
    author: 'Ada',
    date: '2026-08-17T00:00:00.000Z',
    text: 'Please keep this wording.',
    resolved: false,
  };

  expect(
    binding.updateSidecars(before, {
      ...before,
      comments: [...(before.comments ?? []), comment],
    }),
  ).toBe(true);
  expect(binding.content().comments).toContainEqual(comment);

  const withComment = binding.content();
  expect(
    binding.updateSidecars(withComment, {
      ...withComment,
      comments: withComment.comments?.map((candidate) =>
        candidate.id === comment.id
          ? {
              ...candidate,
              resolved: true,
              replies: [
                {
                  id: 'reply-by-ada',
                  actorId: 'ada',
                  author: 'Ada',
                  date: '2026-08-17T00:01:00.000Z',
                  text: 'Resolved after review.',
                },
              ],
            }
          : candidate,
      ),
    }),
  ).toBe(true);
  expect(
    binding
      .content()
      .comments?.find((candidate) => candidate.id === comment.id),
  ).toMatchObject({
    actorId: 'ada',
    resolved: true,
    replies: [expect.objectContaining({ actorId: 'ada' })],
  });

  const current = binding.content();
  expect(() =>
    binding.updateSidecars(current, {
      ...current,
      pageColor: '#000000',
    }),
  ).toThrow(/comment.*mode|review records/i);
  expect(() =>
    binding.updateSidecars(current, {
      ...current,
      comments: current.comments?.map((candidate) =>
        candidate.id === comment.id
          ? { ...candidate, text: 'Rewritten after creation.' }
          : candidate,
      ),
    }),
  ).toThrow(/comment.*mode|review records/i);
  expect(() =>
    binding.updateSidecars(current, {
      ...current,
      comments: [
        ...(current.comments ?? []),
        {
          ...comment,
          id: 'comment-forged-actor',
          actorId: 'grace',
        },
      ],
    }),
  ).toThrow(/actor/i);
  expect(() =>
    binding.updateSidecars(current, {
      ...current,
      comments: current.comments?.filter(
        (candidate) => candidate.id !== 'comment-1',
      ),
    }),
  ).toThrow(/own comment|actor/i);

  expect(
    binding.updateSidecars(current, {
      ...current,
      comments: current.comments?.filter(
        (candidate) => candidate.id !== comment.id,
      ),
    }),
  ).toBe(true);
  expect(
    binding
      .content()
      .comments?.some((candidate) => candidate.id === comment.id),
  ).toBe(false);

  binding.destroy();
});

test('allows comment anchors but rejects text edits in comment mode', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-comment-mode-anchor',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, documentFixture());
  const reviewer = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-comment-mode-anchor',
    document,
    kind: 'document',
    mode: 'comment',
  });
  const binding = createOfficeDocumentCollaborationBinding(reviewer);
  const before = binding.content();
  binding.updateSidecars(before, {
    ...before,
    comments: [
      ...(before.comments ?? []),
      {
        id: 'comment-anchor-by-ada',
        actorId: 'ada',
        author: 'Ada',
        date: '2026-08-17T00:00:00.000Z',
        text: 'Anchor this selection.',
        resolved: false,
      },
    ],
  });
  const editor = new Editor({ extensions: binding.extensions });

  expect(
    editor.commands.insertDocumentComment({
      id: 'comment-anchor-by-ada',
      range: { from: 1, to: 7 },
    }),
  ).toBe(true);
  expect(editor.getHTML()).toContain('data-comment-id="comment-anchor-by-ada"');

  const textBefore = editor.getText();
  editor.chain().setTextSelection(2).insertContent('forbidden').run();
  expect(editor.getText()).toBe(textBefore);
  editor.chain().setTextSelection({ from: 1, to: 7 }).toggleBold().run();
  expect(editor.getHTML()).not.toContain('<strong>');

  expect(editor.commands.removeDocumentComment('comment-anchor-by-ada')).toBe(
    true,
  );
  expect(editor.getHTML()).not.toContain(
    'data-comment-id="comment-anchor-by-ada"',
  );

  editor.destroy();
  binding.destroy();
});

test('converges comment-mode review across disconnected Document peers', async () => {
  const authorDocument = new Y.Doc();
  const author = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-comment-peer-convergence',
    document: authorDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(author, documentFixture());
  const reviewerDocument = cloneDocument(authorDocument);
  const reviewer = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-comment-peer-convergence',
    document: reviewerDocument,
    kind: 'document',
    mode: 'comment',
  });
  const authorBinding = createOfficeDocumentCollaborationBinding(author);
  const reviewerBinding = createOfficeDocumentCollaborationBinding(reviewer);
  const authorEditor = new Editor({ extensions: authorBinding.extensions });
  const reviewerEditor = new Editor({ extensions: reviewerBinding.extensions });
  reviewerEditor.commands.setTextSelection(
    reviewerEditor.state.doc.content.size - 1,
  );
  const reviewerSelection = {
    from: reviewerEditor.state.selection.from,
    to: reviewerEditor.state.selection.to,
  };
  const comment = {
    id: 'comment-live-ada',
    actorId: 'ada',
    author: 'Ada',
    date: '2026-08-17T01:00:00.000Z',
    text: 'Review this exact range.',
    resolved: false,
  };
  const authorBefore = authorBinding.content();

  authorBinding.updateSidecars(authorBefore, {
    ...authorBefore,
    comments: [...(authorBefore.comments ?? []), comment],
  });
  expect(
    authorEditor.commands.insertDocumentComment({
      id: comment.id,
      range: { from: 1, to: 7 },
    }),
  ).toBe(true);
  exchangeUpdates(authorDocument, reviewerDocument);
  await flushMicrotasks();

  expect(reviewerBinding.content().comments).toContainEqual(comment);
  expect(reviewerEditor.getHTML()).toContain(`data-comment-id="${comment.id}"`);
  expect(reviewerEditor.state.selection).toMatchObject({
    from: reviewerSelection.from,
    to: reviewerSelection.to,
  });

  const reviewerBefore = reviewerBinding.content();
  reviewerBinding.updateSidecars(reviewerBefore, {
    ...reviewerBefore,
    comments: reviewerBefore.comments?.map((candidate) =>
      candidate.id === comment.id
        ? {
            ...candidate,
            resolved: true,
            replies: [
              {
                id: 'reply-live-grace',
                actorId: 'grace',
                author: 'Grace',
                date: '2026-08-17T01:01:00.000Z',
                text: 'Reviewed and resolved.',
              },
            ],
          }
        : candidate,
    ),
  });
  const reviewerText = reviewerEditor.getText();
  reviewerEditor.chain().setTextSelection(2).insertContent('forbidden').run();
  expect(reviewerEditor.getText()).toBe(reviewerText);
  exchangeUpdates(authorDocument, reviewerDocument);
  await flushMicrotasks();

  expect(authorBinding.content().comments?.[1]).toMatchObject({
    id: comment.id,
    resolved: true,
    replies: [expect.objectContaining({ actorId: 'grace' })],
  });

  const viewDocument = cloneDocument(authorDocument);
  const viewer = createOfficeCollaborationSession({
    actor: { id: 'lin', name: 'Lin' },
    artifactId: 'document-comment-peer-convergence',
    document: viewDocument,
    kind: 'document',
    mode: 'view',
  });
  const viewBinding = createOfficeDocumentCollaborationBinding(viewer);
  const viewEditor = new Editor({ extensions: viewBinding.extensions });
  const viewBefore = viewBinding.content();
  expect(() =>
    viewBinding.updateSidecars(viewBefore, {
      ...viewBefore,
      comments: viewBefore.comments?.map((candidate) => ({
        ...candidate,
        resolved: false,
      })),
    }),
  ).toThrow(/cannot modify canonical content|review records/);
  const viewText = viewEditor.getText();
  viewEditor.chain().setTextSelection(2).insertContent('forbidden').run();
  expect(viewEditor.getText()).toBe(viewText);

  viewEditor.destroy();
  viewBinding.destroy();
  authorEditor.destroy();
  reviewerEditor.destroy();
  authorBinding.destroy();
  reviewerBinding.destroy();
});

test('undoes a local comment anchor without deleting a remote reply', async () => {
  const firstDocument = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-comment-undo',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, documentFixture());
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-comment-undo',
    document: firstDocument,
    kind: 'document',
    mode: 'comment',
  });
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-comment-undo',
    document: secondDocument,
    kind: 'document',
    mode: 'comment',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstEditor = new Editor({ extensions: firstBinding.extensions });
  const secondEditor = new Editor({ extensions: secondBinding.extensions });
  const comment = {
    id: 'comment-undo-ada',
    actorId: 'ada',
    author: 'Ada',
    date: '2026-08-17T02:00:00.000Z',
    text: 'Local comment with a future remote reply.',
    resolved: false,
  };
  const firstBefore = firstBinding.content();
  firstBinding.updateSidecars(firstBefore, {
    ...firstBefore,
    comments: [...(firstBefore.comments ?? []), comment],
  });
  firstEditor.commands.insertDocumentComment({
    id: comment.id,
    range: { from: 1, to: 7 },
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const secondBefore = secondBinding.content();
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    comments: secondBefore.comments?.map((candidate) =>
      candidate.id === comment.id
        ? {
            ...candidate,
            replies: [
              ...(candidate.replies ?? []),
              {
                id: 'reply-undo-grace',
                actorId: 'grace',
                author: 'Grace',
                date: '2026-08-17T02:01:00.000Z',
                text: 'This remote reply must survive.',
              },
            ],
          }
        : candidate,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const retained = firstBinding
    .content()
    .comments?.find(({ id }) => id === comment.id);
  expect(retained?.replies).toContainEqual(
    expect.objectContaining({
      id: 'reply-undo-grace',
      actorId: 'grace',
    }),
  );
  expect(firstEditor.getHTML()).not.toContain(
    `data-comment-id="${comment.id}"`,
  );
  expect(secondBinding.content()).toEqual(firstBinding.content());

  firstEditor.destroy();
  secondEditor.destroy();
  firstBinding.destroy();
  secondBinding.destroy();
});

test('undoes a local text edit without removing a remote comment', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'document-edit-undo-remote-comment',
    document: firstDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(first, documentFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'document-edit-undo-remote-comment',
    document: secondDocument,
    kind: 'document',
    mode: 'comment',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const firstEditor = new Editor({ extensions: firstBinding.extensions });
  const secondEditor = new Editor({ extensions: secondBinding.extensions });

  firstEditor
    .chain()
    .setTextSelection(firstEditor.state.doc.content.size - 1)
    .insertContent(' local edit')
    .run();
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  const comment = {
    id: 'comment-remote-grace',
    actorId: 'grace',
    author: 'Grace',
    date: '2026-08-17T03:00:00.000Z',
    text: 'Remote review must survive edit undo.',
    resolved: false,
  };
  const secondBefore = secondBinding.content();
  secondBinding.updateSidecars(secondBefore, {
    ...secondBefore,
    comments: [...(secondBefore.comments ?? []), comment],
  });
  secondEditor.commands.insertDocumentComment({
    id: comment.id,
    range: { from: 1, to: 7 },
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstEditor.getText()).not.toContain('local edit');
  expect(firstBinding.content().comments).toContainEqual(comment);
  expect(firstEditor.getHTML()).toContain(`data-comment-id="${comment.id}"`);
  expect(secondBinding.content()).toEqual(firstBinding.content());

  firstEditor.destroy();
  secondEditor.destroy();
  firstBinding.destroy();
  secondBinding.destroy();
});

test('allows actor-attributed suggestion-mode text revisions only', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'document-suggestion-permissions',
    document,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(writable, {
    ...documentFixture(),
    trackChanges: false,
  });
  const suggester = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada Reviewer' },
    artifactId: 'document-suggestion-permissions',
    document,
    kind: 'document',
    mode: 'suggest',
  });
  const binding = createOfficeDocumentCollaborationBinding(suggester, {
    workExtensions: {
      createChange: () => ({
        actorId: 'ada',
        author: 'Ada Reviewer',
        date: '2026-08-17T08:00:00.000Z',
        id: 'suggestion-ada-1',
      }),
      isTracking: () => true,
    },
  });
  const editor = new Editor({ extensions: binding.extensions });
  const insertionPoint = lastDocumentTextPosition(editor);

  expect(
    editor.commands.replaceDocumentTextWithTrackedChange(
      insertionPoint,
      insertionPoint,
      ' proposed',
    ),
  ).toBe(true);
  expect(editor.getHTML()).toContain('data-change-actor-id="ada"');
  expect(editor.getHTML()).toContain('data-change-id="suggestion-ada-1"');
  expect(editor.getHTML()).toContain('> proposed</ins>');

  const beforeForgedEdit = editor.getHTML();
  editor
    .chain()
    .setTextSelection(2)
    .insertContent('FORGED CANONICAL EDIT')
    .run();
  expect(editor.getHTML()).toBe(beforeForgedEdit);
  expect(editor.commands.acceptDocumentChange('suggestion-ada-1')).toBe(true);
  expect(editor.getHTML()).toBe(beforeForgedEdit);

  const before = binding.content();
  expect(() =>
    binding.updateSidecars(before, { ...before, trackChanges: true }),
  ).toThrow(/suggest.*mode|canonical content/i);
  expect(() =>
    binding.decideChanges(editor, ['suggestion-ada-1'], 'accept'),
  ).toThrow(/suggest.*mode|decision/i);

  editor.destroy();
  binding.destroy();
});

test('converges final edit-mode suggestion decisions and clears stale suggester undo', async () => {
  const suggestionDocument = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-suggestion-decisions',
    document: suggestionDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    trackChanges: false,
  });
  const suggester = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada Reviewer' },
    artifactId: 'document-suggestion-decisions',
    document: suggestionDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const suggestionBinding = createOfficeDocumentCollaborationBinding(
    suggester,
    {
      workExtensions: {
        createChange: () => ({
          actorId: 'ada',
          author: 'Ada Reviewer',
          date: '2026-08-17T08:10:00.000Z',
          id: 'suggestion-final-1',
        }),
        isTracking: () => true,
      },
    },
  );
  const suggestionEditor = new Editor({
    extensions: suggestionBinding.extensions,
  });
  const insertionPoint = lastDocumentTextPosition(suggestionEditor);
  suggestionEditor.commands.replaceDocumentTextWithTrackedChange(
    insertionPoint,
    insertionPoint,
    ' accepted proposal',
  );
  expect(suggestionBinding.canUndo()).toBe(true);

  const editorDocument = cloneDocument(suggestionDocument);
  const editorSession = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-suggestion-decisions',
    document: editorDocument,
    kind: 'document',
    mode: 'edit',
  });
  const editorBinding = createOfficeDocumentCollaborationBinding(editorSession);
  const editor = new Editor({ extensions: editorBinding.extensions });

  expect(
    editorBinding.decideChanges(editor, ['suggestion-final-1'], 'accept', {
      decidedAt: '2026-08-17T08:11:00.000Z',
    }),
  ).toBe(true);
  expect(editor.getText()).toContain('accepted proposal');
  expect(editor.getHTML()).not.toContain('suggestion-final-1');
  expect(editorBinding.content().changeDecisions).toEqual([
    {
      id: 'insertion:suggestion-final-1',
      changeId: 'suggestion-final-1',
      changeKind: 'insertion',
      suggestedByActorId: 'ada',
      suggestedBy: 'Ada Reviewer',
      suggestedAt: '2026-08-17T08:10:00.000Z',
      text: ' accepted proposal',
      decision: 'accept',
      decidedByActorId: 'grace',
      decidedBy: 'Grace Editor',
      decidedAt: '2026-08-17T08:11:00.000Z',
    },
  ]);
  expect(editorBinding.canUndo()).toBe(false);

  exchangeUpdates(suggestionDocument, editorDocument);
  await flushMicrotasks();
  expect(suggestionEditor.getText()).toContain('accepted proposal');
  expect(suggestionEditor.getHTML()).not.toContain('suggestion-final-1');
  expect(suggestionBinding.content().changeDecisions).toEqual(
    editorBinding.content().changeDecisions,
  );
  expect(suggestionBinding.canUndo()).toBe(false);
  expect(suggestionBinding.undo()).toBe(false);

  const decision = editorBinding.content();
  expect(() =>
    editorBinding.updateSidecars(decision, {
      ...decision,
      changeDecisions: decision.changeDecisions?.map((record) => ({
        ...record,
        decision: 'reject',
      })),
    }),
  ).toThrow(/decision.*immutable|cannot rewrite/i);

  suggestionEditor.destroy();
  editor.destroy();
  suggestionBinding.destroy();
  editorBinding.destroy();
});

test('preserves paragraph-formatting revisions in suggest mode and converges immutable edit decisions', async () => {
  const before = serializeDocumentParagraphFormatting({
    textAlign: 'right',
    paragraphDirection: null,
    indentLevel: 0,
    rightIndent: 0,
    firstLineIndent: 0,
    spaceBefore: 6,
    spaceAfter: null,
    lineHeight: null,
    lineRule: null,
    autoLineHeight: null,
    keepLines: null,
    keepWithNext: null,
    pageBreakBefore: null,
    widowControl: null,
    contextualSpacing: null,
    outlineLevel: null,
    tabStops: null,
    paragraphBorders: null,
    paragraphShading: null,
    defaultCollapsed: null,
  });
  const sharedDocument = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-paragraph-formatting-decision',
    document: sharedDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    html: `<section data-document-section="true"><p data-document-change="true" data-change-kind="paragraph-formatting" data-change-id="paragraph-review-1" data-change-actor-id="ada" data-change-author="Ada Reviewer" data-change-date="2026-08-18T10:00:00.000Z" data-change-before='${before}' style="text-align: center">Shared paragraph</p></section>`,
  });

  const suggestSession = createOfficeCollaborationSession({
    actor: { id: 'lin', name: 'Lin Suggester' },
    artifactId: 'document-paragraph-formatting-decision',
    document: sharedDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const suggestBinding =
    createOfficeDocumentCollaborationBinding(suggestSession);
  const suggestEditor = new Editor({ extensions: suggestBinding.extensions });
  const beforeForbiddenFormatting = suggestEditor.getHTML();
  const paragraphText = lastDocumentTextPosition(suggestEditor);
  suggestEditor
    .chain()
    .setTextSelection(paragraphText)
    .setTextAlign('justify')
    .run();
  expect(suggestEditor.getHTML()).toBe(beforeForbiddenFormatting);

  const editorDocument = cloneDocument(sharedDocument);
  const editSession = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-paragraph-formatting-decision',
    document: editorDocument,
    kind: 'document',
    mode: 'edit',
  });
  const editBinding = createOfficeDocumentCollaborationBinding(editSession);
  const editEditor = new Editor({ extensions: editBinding.extensions });
  expect(
    editBinding.decideChanges(editEditor, ['paragraph-review-1'], 'reject', {
      decidedAt: '2026-08-18T10:01:00.000Z',
    }),
  ).toBe(true);
  expect(editEditor.getHTML()).toContain('text-align: right');
  expect(editEditor.getHTML()).not.toContain('paragraph-review-1');
  expect(editBinding.content().changeDecisions).toEqual([
    {
      id: 'paragraph-formatting:paragraph-review-1',
      changeId: 'paragraph-review-1',
      changeKind: 'paragraph-formatting',
      suggestedByActorId: 'ada',
      suggestedBy: 'Ada Reviewer',
      suggestedAt: '2026-08-18T10:00:00.000Z',
      text: 'Shared paragraph',
      decision: 'reject',
      decidedByActorId: 'grace',
      decidedBy: 'Grace Editor',
      decidedAt: '2026-08-18T10:01:00.000Z',
    },
  ]);

  exchangeUpdates(sharedDocument, editorDocument);
  await flushMicrotasks();
  expect(suggestEditor.getHTML()).toContain('text-align: right');
  expect(suggestEditor.getHTML()).not.toContain('paragraph-review-1');
  expect(suggestBinding.content().changeDecisions).toEqual(
    editBinding.content().changeDecisions,
  );

  suggestEditor.destroy();
  editEditor.destroy();
  suggestBinding.destroy();
  editBinding.destroy();
});

test('converges an immutable ordered-list numbering decision', async () => {
  const before = serializeDocumentNumberingChange({
    start: 4,
    type: 'I',
    level: 0,
    originalFormat: 1,
    originalSuffix: '.',
  });
  const sharedDocument = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    artifactId: 'document-numbering-decision',
    document: sharedDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, {
    ...documentFixture(),
    html: `<section data-document-section="true"><ol start="4" type="a" data-document-change="true" data-change-kind="numbering" data-change-id="numbering-review-1" data-change-actor-id="ada" data-change-author="Ada Reviewer" data-change-date="2026-09-01T10:00:00.000Z" data-change-before='${before}'><li><p>Shared list</p></li></ol></section>`,
    model: undefined,
  });

  const observerDocument = cloneDocument(sharedDocument);
  const observerSession = createOfficeCollaborationSession({
    actor: { id: 'lin', name: 'Lin Observer' },
    artifactId: 'document-numbering-decision',
    document: observerDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const observerBinding =
    createOfficeDocumentCollaborationBinding(observerSession);
  const observerEditor = new Editor({ extensions: observerBinding.extensions });

  const editorDocument = cloneDocument(sharedDocument);
  const editSession = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace Editor' },
    artifactId: 'document-numbering-decision',
    document: editorDocument,
    kind: 'document',
    mode: 'edit',
  });
  const editBinding = createOfficeDocumentCollaborationBinding(editSession);
  const editEditor = new Editor({ extensions: editBinding.extensions });
  expect(
    editBinding.decideChanges(editEditor, ['numbering-review-1'], 'reject', {
      decidedAt: '2026-09-01T10:01:00.000Z',
    }),
  ).toBe(true);
  expect(editEditor.getHTML()).toContain('type="I"');
  expect(editEditor.getHTML()).not.toContain('numbering-review-1');
  expect(editBinding.content().changeDecisions).toEqual([
    {
      id: 'numbering:numbering-review-1',
      changeId: 'numbering-review-1',
      changeKind: 'numbering',
      suggestedByActorId: 'ada',
      suggestedBy: 'Ada Reviewer',
      suggestedAt: '2026-09-01T10:00:00.000Z',
      text: 'Shared list',
      decision: 'reject',
      decidedByActorId: 'grace',
      decidedBy: 'Grace Editor',
      decidedAt: '2026-09-01T10:01:00.000Z',
    },
  ]);

  exchangeUpdates(observerDocument, editorDocument);
  await flushMicrotasks();
  expect(observerEditor.getHTML()).toContain('type="I"');
  expect(observerEditor.getHTML()).not.toContain('numbering-review-1');
  expect(observerBinding.content().changeDecisions).toEqual(
    editBinding.content().changeDecisions,
  );

  observerEditor.destroy();
  editEditor.destroy();
  observerBinding.destroy();
  editBinding.destroy();
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

function lastDocumentTextPosition(editor: Editor): number {
  let position = 1;
  editor.state.doc.descendants((node, offset) => {
    if (node.isText) position = offset + node.nodeSize;
  });
  return position;
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
