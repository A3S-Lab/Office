import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeMarkdownCollaborationBinding,
  initializeOfficeMarkdownCollaboration,
  OfficeCollaborationError,
  readOfficeCollaborationMetadata,
  readOfficeMarkdownCollaboration,
} from '../src/core';

test('initializes one versioned Markdown collaboration document', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-1',
    document,
    kind: 'markdown',
  });

  expect(
    initializeOfficeMarkdownCollaboration(session, {
      type: 'markdown',
      markdown: '# Shared notes',
    }),
  ).toEqual({
    initialized: true,
    content: { type: 'markdown', markdown: '# Shared notes' },
  });
  expect(readOfficeCollaborationMetadata(session)).toEqual({
    protocol: 'a3s.office.collaboration',
    version: 1,
    artifactId: 'notes-1',
    kind: 'markdown',
    initialized: true,
  });

  expect(
    initializeOfficeMarkdownCollaboration(session, {
      type: 'markdown',
      markdown: '# Must not overwrite',
    }),
  ).toEqual({
    initialized: false,
    content: { type: 'markdown', markdown: '# Shared notes' },
  });
});

test('requires an explicit synchronized bootstrap before binding', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-uninitialized',
    kind: 'markdown',
  });

  expect(() => createOfficeMarkdownCollaborationBinding(session)).toThrow(
    /has not been initialized/,
  );
  expect(readOfficeCollaborationMetadata(session)).toBeNull();
});

test('leaves no partial metadata when bootstrap validation fails', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-invalid-bootstrap',
    kind: 'markdown',
  });
  session.document
    .getText(session.rootName('markdown.source'))
    .insert(0, 'Unattributed source');

  expect(() =>
    initializeOfficeMarkdownCollaboration(session, {
      type: 'markdown',
      markdown: 'Seed',
    }),
  ).toThrow(/contains data without initialized metadata/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
  expect(
    session.document.getArray(session.rootName('bootstrap.initializers'))
      .length,
  ).toBe(0);
});

test('detects concurrent bootstrap instead of choosing one initial value', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'notes-concurrent-bootstrap',
    document: firstDocument,
    kind: 'markdown',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'notes-concurrent-bootstrap',
    document: secondDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'First seed',
  });
  initializeOfficeMarkdownCollaboration(second, {
    type: 'markdown',
    markdown: 'Second seed',
  });

  exchangeUpdates(firstDocument, secondDocument);

  expect(() => readOfficeCollaborationMetadata(first)).toThrow(
    /Multiple clients initialized/,
  );
  expect(() => readOfficeMarkdownCollaboration(second)).toThrow(
    /Multiple clients initialized/,
  );
});

test('rejects canonical changes outside edit mode and marks agent origins', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    actor: { id: 'agent-1', kind: 'agent', name: 'Coding agent' },
    artifactId: 'notes-permissions',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(writable, {
    type: 'markdown',
    markdown: 'Authorized content',
  });
  expect(writable.localOrigin).toMatchObject({
    actorId: 'agent-1',
    kind: 'agent',
  });

  const readOnly = createOfficeCollaborationSession({
    artifactId: 'notes-permissions',
    document,
    kind: 'markdown',
    mode: 'view',
  });
  const binding = createOfficeMarkdownCollaborationBinding(readOnly);
  expect(() => binding.replace('Unauthorized content')).toThrow(
    /cannot modify canonical content/,
  );
  expect(() => readOnly.transact(() => undefined)).toThrow(
    /cannot modify canonical content/,
  );
  expect(readOfficeMarkdownCollaboration(readOnly).markdown).toBe(
    'Authorized content',
  );
});

test('converges concurrent Markdown edits without replacing the shared document', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'notes-2',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Base',
  });

  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'notes-2',
    document: secondDocument,
    kind: 'markdown',
  });
  const firstBinding = createOfficeMarkdownCollaborationBinding(first);
  const secondBinding = createOfficeMarkdownCollaborationBinding(second);

  firstBinding.replace('Alpha Base');
  secondBinding.replace('Base Omega');
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.content()).toEqual(secondBinding.content());
  expect(firstBinding.content().markdown).toContain('Alpha');
  expect(firstBinding.content().markdown).toContain('Omega');
  expect(readOfficeCollaborationMetadata(first)?.artifactId).toBe('notes-2');
});

test('tracks only local Markdown operations in each undo manager', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'notes-3',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Shared',
  });
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'notes-3',
    document: secondDocument,
    kind: 'markdown',
  });
  const firstBinding = createOfficeMarkdownCollaborationBinding(first);
  const secondBinding = createOfficeMarkdownCollaborationBinding(second);

  firstBinding.replace('Shared by Ada');
  exchangeUpdates(firstDocument, secondDocument);
  expect(firstBinding.canUndo()).toBe(true);
  expect(secondBinding.canUndo()).toBe(false);

  secondBinding.replace('Shared by Ada and Grace');
  exchangeUpdates(firstDocument, secondDocument);
  expect(firstBinding.canUndo()).toBe(true);
  expect(secondBinding.canUndo()).toBe(true);

  expect(secondBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  expect(readOfficeMarkdownCollaboration(first).markdown).toBe('Shared by Ada');
  expect(firstBinding.canUndo()).toBe(true);
});

test('reports locality relative to one binding origin', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-locality',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Shared',
  });
  const first = createOfficeMarkdownCollaborationBinding(session);
  const second = createOfficeMarkdownCollaborationBinding(session);
  const firstLocality: boolean[] = [];
  const secondLocality: boolean[] = [];
  first.subscribe((change) => firstLocality.push(change.local));
  second.subscribe((change) => secondLocality.push(change.local));

  first.replace('Shared by first');

  expect(firstLocality).toEqual([true]);
  expect(secondLocality).toEqual([false]);
  expect(first.canUndo()).toBe(true);
  expect(second.canUndo()).toBe(false);
});

test('does not destroy a host-owned Y.Doc with the session', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-4',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Host owned',
  });
  session.destroy();

  expect(document.getText('a3s.office.markdown.source').toString()).toBe(
    'Host owned',
  );
  document.getMap('still-alive').set('value', true);
  expect(document.getMap('still-alive').get('value')).toBe(true);
});

test('rejects a shared document attached with another identity or kind', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-5',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Identity bound',
  });

  expect(() =>
    createOfficeCollaborationSession({
      artifactId: 'notes-6',
      document,
      kind: 'markdown',
    }),
  ).toThrowError(OfficeCollaborationError);
  expect(() =>
    createOfficeCollaborationSession({
      artifactId: 'notes-5',
      document,
      kind: 'document',
    }),
  ).toThrow(/contains 'markdown' content/);
});

test('preserves UTF-16 pairs when applying a bounded Markdown replacement', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-emoji',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'A😀B',
  });
  const binding = createOfficeMarkdownCollaborationBinding(session);

  binding.replace('A😃B');

  expect(binding.content().markdown).toBe('A😃B');
});

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
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
