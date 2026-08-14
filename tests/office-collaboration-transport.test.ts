import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeCollaborationTransportBinding,
  createOfficeMarkdownCollaborationBinding,
  initializeOfficeMarkdownCollaboration,
  type OfficeCollaborationTransport,
  type OfficeCollaborationTransportMessage,
  readOfficeMarkdownCollaboration,
} from '../src/core';

test('synchronizes an initialized Yjs artifact over a host-owned transport', () => {
  const transport = new MemoryTransport();
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'transport-notes',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: '# Shared over a host channel',
  });
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'transport-notes',
    document: new Y.Doc(),
    kind: 'markdown',
  });
  const firstTransport = createOfficeCollaborationTransportBinding(
    first,
    transport,
    { autoSynchronize: false },
  );
  const secondTransport = createOfficeCollaborationTransportBinding(
    second,
    transport,
    { autoSynchronize: false },
  );

  secondTransport.synchronize();

  expect(readOfficeMarkdownCollaboration(second)).toEqual(
    readOfficeMarkdownCollaboration(first),
  );
  expect(transport.messages.map((message) => message.type)).toEqual([
    'sync-step-1',
    'sync-step-2',
  ]);

  firstTransport.destroy();
  secondTransport.destroy();
});

test('preserves typed actor and operation origins on incremental updates', () => {
  const transport = new MemoryTransport();
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'agent-7', kind: 'agent', name: 'Coding agent' },
    artifactId: 'transport-origin',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Shared',
  });
  const second = createOfficeCollaborationSession({
    actor: { id: 'reviewer', name: 'Reviewer' },
    artifactId: 'transport-origin',
    document: cloneDocument(firstDocument),
    kind: 'markdown',
  });
  const firstTransport = createOfficeCollaborationTransportBinding(
    first,
    transport,
    { autoSynchronize: false },
  );
  const secondTransport = createOfficeCollaborationTransportBinding(
    second,
    transport,
    { autoSynchronize: false },
  );
  const remoteOrigins: unknown[] = [];
  second.document.on('update', (_update, origin) => remoteOrigins.push(origin));

  const origin = first.createOrigin('agent', 'operation-42');
  first.transact(() => {
    first.document
      .getText(first.rootName('markdown.source'))
      .insert('Shared'.length, ' by agent');
  }, origin);

  expect(readOfficeMarkdownCollaboration(second).markdown).toBe(
    'Shared by agent',
  );
  expect(remoteOrigins).toEqual([
    {
      protocol: 'a3s.office.collaboration',
      kind: 'agent',
      actorId: 'agent-7',
      operationId: 'operation-42',
    },
  ]);
  expect(transport.messages).toHaveLength(1);
  expect(transport.messages[0]?.origin).toEqual(origin);

  firstTransport.destroy();
  secondTransport.destroy();
});

test('reconverges offline peers after both sides request a fresh state-vector sync', () => {
  const transport = new MemoryTransport();
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'transport-reconnect',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Base',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'transport-reconnect',
    document: cloneDocument(firstDocument),
    kind: 'markdown',
  });
  const firstTransport = createOfficeCollaborationTransportBinding(
    first,
    transport,
    { autoSynchronize: false },
  );
  const secondTransport = createOfficeCollaborationTransportBinding(
    second,
    transport,
    { autoSynchronize: false },
  );
  const firstContent = createOfficeMarkdownCollaborationBinding(first);
  const secondContent = createOfficeMarkdownCollaborationBinding(second);

  transport.deliver = false;
  firstContent.replace('Alpha Base');
  secondContent.replace('Base Omega');
  transport.deliver = true;
  const beforeReconnect = transport.messages.length;

  firstTransport.synchronize();
  secondTransport.synchronize();

  expect(firstContent.content()).toEqual(secondContent.content());
  expect(firstContent.content().markdown).toContain('Alpha');
  expect(firstContent.content().markdown).toContain('Omega');
  expect(
    transport.messages.slice(beforeReconnect).map((message) => message.type),
  ).toEqual(['sync-step-1', 'sync-step-2', 'sync-step-1', 'sync-step-2']);

  firstContent.destroy();
  secondContent.destroy();
  firstTransport.destroy();
  secondTransport.destroy();
});

test('does not echo a remotely applied update back into the room', () => {
  const transport = new MemoryTransport();
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'transport-no-loop',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Before',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'transport-no-loop',
    document: cloneDocument(firstDocument),
    kind: 'markdown',
  });
  const firstTransport = createOfficeCollaborationTransportBinding(
    first,
    transport,
    { autoSynchronize: false },
  );
  const secondTransport = createOfficeCollaborationTransportBinding(
    second,
    transport,
    { autoSynchronize: false },
  );
  const firstContent = createOfficeMarkdownCollaborationBinding(first);

  firstContent.replace('After');

  expect(transport.messages).toHaveLength(1);
  expect(transport.messages[0]?.type).toBe('update');
  expect(readOfficeMarkdownCollaboration(second).markdown).toBe('After');

  firstContent.destroy();
  firstTransport.destroy();
  secondTransport.destroy();
});

test('rejects cross-artifact, oversized, and malformed transport messages', () => {
  const transport = new MemoryTransport();
  const session = createOfficeCollaborationSession({
    artifactId: 'transport-bounds',
    kind: 'markdown',
  });
  const binding = createOfficeCollaborationTransportBinding(
    session,
    transport,
    { autoSynchronize: false, maxPayloadBytes: 8 },
  );
  const base = {
    protocol: 'a3s.office.collaboration',
    version: 1,
    artifactId: session.artifactId,
    artifactKind: session.kind,
    namespace: session.namespace,
    senderClientId: session.document.clientID + 1,
    type: 'update',
  } as const;

  expect(() =>
    transport.publish({
      ...base,
      artifactId: 'another-artifact',
      payload: new Uint8Array([0]),
    }),
  ).toThrow(/belongs to another artifact/);
  expect(() =>
    transport.publish({ ...base, payload: new Uint8Array(9) }),
  ).toThrow(/configured limit is 8 bytes/);
  expect(() =>
    transport.publish({ ...base, payload: new Uint8Array([255]) }),
  ).toThrow(/not a valid Yjs v1 message/);
  expect(session.document.getMap('untouched').size).toBe(0);

  binding.destroy();
  expect(() => binding.synchronize()).toThrow(/has been destroyed/);
  expect(transport.listenerCount).toBe(0);
  session.destroy();
  expect(() =>
    createOfficeCollaborationTransportBinding(session, transport),
  ).toThrow(/session has been destroyed/);
});

class MemoryTransport implements OfficeCollaborationTransport {
  readonly messages: OfficeCollaborationTransportMessage[] = [];
  readonly #listeners = new Set<(message: unknown) => void>();
  deliver = true;

  get listenerCount(): number {
    return this.#listeners.size;
  }

  publish(message: OfficeCollaborationTransportMessage): void {
    this.messages.push(message);
    if (!this.deliver) return;
    for (const listener of [...this.#listeners]) listener(message);
  }

  subscribe(listener: (message: unknown) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }
}

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}
