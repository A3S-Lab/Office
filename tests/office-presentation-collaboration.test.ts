import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficePresentationCollaborationBinding,
  initializeOfficePresentationCollaboration,
  readOfficePresentationCollaboration,
} from '../src/core';
import { presentationCollaborationFixture as presentationFixture } from './fixtures/presentation-collaboration';

test('initializes typed Presentation roots and never stores one deck blob', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'presentation-1',
    kind: 'presentation',
  });

  expect(
    initializeOfficePresentationCollaboration(session, presentationFixture()),
  ).toMatchObject({ initialized: true });
  expect(readOfficePresentationCollaboration(session)).toEqual(
    presentationFixture(),
  );
  expect(
    session.document.getArray(session.rootName('presentation.slide-order'))
      .length,
  ).toBe(2);
  expect(
    session.document.getMap(session.rootName('presentation.slides')).size,
  ).toBe(2);
  expect(
    session.document.share.has(session.rootName('presentation.content')),
  ).toBe(false);
});

test('keeps empty-comment metadata internal to the shared model', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'presentation-empty-comments',
    kind: 'presentation',
  });
  const fixture = presentationFixture();
  const content = {
    ...fixture,
    slides: fixture.slides.map((slide, index) =>
      index === 0 ? { ...slide, comments: [] } : slide,
    ),
  };

  initializeOfficePresentationCollaboration(session, content);

  expect(readOfficePresentationCollaboration(session)).toEqual(content);
  expect(
    readOfficePresentationCollaboration(session).slides[0],
  ).not.toHaveProperty('commentsPresent');
});

test('rejects malformed nested roots without mutating them during read', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'presentation-pure-read',
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(session, presentationFixture());
  const slides = session.document.getMap(
    session.rootName('presentation.slides'),
  );
  const slide = slides.get('slide-1');
  expect(slide).toBeInstanceOf(Y.Map);
  (slide as Y.Map<unknown>).delete('elements');
  let transactions = 0;
  const countTransaction = () => {
    transactions += 1;
  };
  session.document.on('afterTransaction', countTransaction);

  expect(() => readOfficePresentationCollaboration(session)).toThrow(
    /shared Presentation collaboration slide elements is invalid/,
  );
  expect(transactions).toBe(0);
  session.document.off('afterTransaction', countTransaction);
});

test('rejects duplicate identities before bootstrap writes metadata', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'presentation-duplicate-id',
    kind: 'presentation',
  });
  const fixture = presentationFixture();

  expect(() =>
    initializeOfficePresentationCollaboration(session, {
      ...fixture,
      slides: [fixture.slides[0], { ...fixture.slides[1], id: 'slide-1' }],
    }),
  ).toThrow(/unique slide ID/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
});

test('rejects duplicate comment identities before bootstrap writes metadata', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'presentation-duplicate-comment-id',
    kind: 'presentation',
  });
  const fixture = presentationFixture();
  const comment = {
    id: 'comment-1',
    author: 'Ada',
    date: '2026-08-13T04:00:00.000Z',
    text: 'Review',
    x: 20,
    y: 20,
  };

  expect(() =>
    initializeOfficePresentationCollaboration(session, {
      ...fixture,
      slides: fixture.slides.map((slide, index) =>
        index === 0
          ? { ...slide, comments: [comment, { ...comment, text: 'Again' }] }
          : slide,
      ),
    }),
  ).toThrow(/unique comment in slide 'slide-1' ID/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
});

test('detects concurrent independent Presentation bootstrap', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-bootstrap-race',
    document: firstDocument,
    kind: 'presentation',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-bootstrap-race',
    document: secondDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  initializeOfficePresentationCollaboration(second, {
    ...presentationFixture(),
    width: 16,
  });

  exchangeUpdates(firstDocument, secondDocument);

  expect(() => readOfficePresentationCollaboration(first)).toThrow(
    /Multiple clients initialized/,
  );
  expect(() => readOfficePresentationCollaboration(second)).toThrow(
    /Multiple clients initialized/,
  );
});

test('merges edits to separate slides and separate scene objects', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'presentation-convergence',
    document: firstDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'presentation-convergence',
    document: secondDocument,
    kind: 'presentation',
  });
  const firstBinding = createOfficePresentationCollaborationBinding(first);
  const secondBinding = createOfficePresentationCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(firstBefore, {
    ...firstBefore,
    slides: firstBefore.slides.map((slide) =>
      slide.id === 'slide-1'
        ? {
            ...slide,
            elements: slide.elements.map((element) =>
              element.id === 'element-title'
                ? { ...element, x: 18, y: 22 }
                : element,
            ),
          }
        : slide,
    ),
  });
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    slides: secondBefore.slides.map((slide) =>
      slide.id === 'slide-2'
        ? { ...slide, background: '#EEF2FF', notes: 'Remote speaker notes' }
        : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.slides[0].elements[0]).toMatchObject({ x: 18, y: 22 });
  expect(converged.slides[1]).toMatchObject({
    background: '#EEF2FF',
    notes: 'Remote speaker notes',
  });

  firstBinding.destroy();
  secondBinding.destroy();
});

test('does not delete remote records when applying an unrelated stale snapshot', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-stale-snapshot',
    document: firstDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-stale-snapshot',
    document: secondDocument,
    kind: 'presentation',
  });
  const firstBinding = createOfficePresentationCollaborationBinding(first);
  const secondBinding = createOfficePresentationCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();
  const remoteSlide = {
    ...firstBefore.slides[1],
    id: 'slide-remote',
    name: 'Remote slide',
    elements: firstBefore.slides[1].elements.map((element) => ({
      ...element,
      id: 'element-remote',
    })),
  };

  firstBinding.replace(firstBefore, {
    ...firstBefore,
    slides: [...firstBefore.slides, remoteSlide],
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  secondBinding.replace(stale, {
    ...stale,
    slides: stale.slides.map((slide) =>
      slide.id === 'slide-1' ? { ...slide, background: '#FEF3C7' } : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.slides.map(({ id }) => id)).toContain('slide-remote');
  expect(converged.slides[0].background).toBe('#FEF3C7');

  firstBinding.destroy();
  secondBinding.destroy();
});

test('merges concurrent comments on one slide by stable ID', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-comments',
    document: firstDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-comments',
    document: secondDocument,
    kind: 'presentation',
  });
  const firstBinding = createOfficePresentationCollaborationBinding(first);
  const secondBinding = createOfficePresentationCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();
  const adaComment = {
    id: 'comment-ada',
    author: 'Ada',
    date: '2026-08-13T04:00:00.000Z',
    text: 'Ada review',
    x: 20,
    y: 20,
  };
  const graceComment = {
    id: 'comment-grace',
    author: 'Grace',
    date: '2026-08-13T04:01:00.000Z',
    text: 'Grace review',
    x: 30,
    y: 30,
  };

  firstBinding.replace(firstBefore, {
    ...firstBefore,
    slides: firstBefore.slides.map((slide) =>
      slide.id === 'slide-1' ? { ...slide, comments: [adaComment] } : slide,
    ),
  });
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    slides: secondBefore.slides.map((slide) =>
      slide.id === 'slide-1' ? { ...slide, comments: [graceComment] } : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.slides[0].comments).toEqual(
    expect.arrayContaining([adaComment, graceComment]),
  );

  firstBinding.destroy();
  secondBinding.destroy();
});

test('rejects an edit to a scene object concurrently removed by another client', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-delete-edit',
    document: firstDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-delete-edit',
    document: secondDocument,
    kind: 'presentation',
  });
  const firstBinding = createOfficePresentationCollaborationBinding(first);
  const secondBinding = createOfficePresentationCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(firstBefore, {
    ...firstBefore,
    slides: firstBefore.slides.map((slide) =>
      slide.id === 'slide-1' ? { ...slide, elements: [] } : slide,
    ),
  });
  Y.applyUpdate(
    secondDocument,
    Y.encodeStateAsUpdate(firstDocument, Y.encodeStateVector(secondDocument)),
    'test-network',
  );

  expect(() =>
    secondBinding.replace(stale, {
      ...stale,
      slides: stale.slides.map((slide) =>
        slide.id === 'slide-1'
          ? {
              ...slide,
              elements: slide.elements.map((element) => ({
                ...element,
                text: 'Stale edit',
              })),
            }
          : slide,
      ),
    }),
  ).toThrow(/removed before this change/);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('keeps Presentation undo local to one client', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-local-undo',
    document: firstDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-local-undo',
    document: secondDocument,
    kind: 'presentation',
  });
  const firstBinding = createOfficePresentationCollaborationBinding(first);
  const secondBinding = createOfficePresentationCollaborationBinding(second);

  const firstBefore = firstBinding.content();
  firstBinding.replace(firstBefore, {
    ...firstBefore,
    slides: firstBefore.slides.map((slide) =>
      slide.id === 'slide-1' ? { ...slide, name: 'Ada title' } : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  const secondBefore = secondBinding.content();
  secondBinding.stopCapturing();
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    slides: secondBefore.slides.map((slide) =>
      slide.id === 'slide-2' ? { ...slide, name: 'Grace title' } : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  const converged = firstBinding.content();
  expect(converged.slides[0].name).toBe('Cover');
  expect(converged.slides[1].name).toBe('Grace title');
  expect(secondBinding.content()).toEqual(converged);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('undoes a local Presentation comment without removing remote records', async () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'presentation-comment-undo',
    document: firstDocument,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(first, presentationFixture());
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'presentation-comment-undo',
    document: secondDocument,
    kind: 'presentation',
  });
  const firstBinding = createOfficePresentationCollaborationBinding(first);
  const secondBinding = createOfficePresentationCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  firstBinding.replace(firstBefore, {
    ...firstBefore,
    slides: firstBefore.slides.map((slide) =>
      slide.id === 'slide-1'
        ? {
            ...slide,
            comments: [
              {
                id: 'comment-local',
                author: 'Ada',
                date: '2026-08-13T05:00:00.000Z',
                text: 'Local review',
                x: 20,
                y: 20,
              },
            ],
          }
        : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  const secondBefore = secondBinding.content();
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    slides: secondBefore.slides.map((slide) =>
      slide.id === 'slide-2' ? { ...slide, notes: 'Remote note' } : slide,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  await flushMicrotasks();
  const converged = firstBinding.content();
  expect(converged.slides[0].comments).toBeUndefined();
  expect(converged.slides[1].notes).toBe('Remote note');
  expect(secondBinding.content()).toEqual(converged);

  firstBinding.destroy();
  secondBinding.destroy();
});

test('rejects Presentation mutation outside edit mode', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    artifactId: 'presentation-view',
    document,
    kind: 'presentation',
  });
  initializeOfficePresentationCollaboration(writable, presentationFixture());
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'presentation-view',
    document,
    kind: 'presentation',
    mode: 'view',
  });
  const binding = createOfficePresentationCollaborationBinding(readOnly);
  const before = binding.content();

  expect(() => binding.replace(before, { ...before, width: 14 })).toThrow(
    /cannot modify canonical content/,
  );
  expect(() => binding.undo()).toThrow(/cannot modify canonical content/);

  binding.destroy();
});

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
