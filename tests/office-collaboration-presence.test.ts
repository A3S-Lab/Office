import { expect, test } from '@rstest/core';
import {
  applyAwarenessUpdate,
  Awareness,
  encodeAwarenessUpdate,
} from 'y-protocols/awareness';
import * as Y from 'yjs';
import {
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  OFFICE_COLLABORATION_PRESENCE_FIELD,
  type OfficeArtifactKind,
  type OfficeCollaborationPresenceLocation,
} from '../src/core';

test('publishes typed local presence and observes remote Awareness updates', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const firstAwareness = new Awareness(firstDocument);
  const secondAwareness = new Awareness(secondDocument);
  const firstSession = createOfficeCollaborationSession({
    actor: {
      id: 'ada',
      name: 'Ada',
      color: '  #7c3aed  ',
      avatarUrl: '  https://example.test/ada.png  ',
    },
    artifactId: 'presence-markdown',
    awareness: firstAwareness,
    document: firstDocument,
    kind: 'markdown',
  });
  const secondSession = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'presence-markdown',
    awareness: secondAwareness,
    document: secondDocument,
    kind: 'markdown',
    mode: 'view',
  });
  const first = createOfficeCollaborationPresence(firstSession, {
    location: { kind: 'markdown', anchor: 2, head: 8 },
  });
  const second = createOfficeCollaborationPresence(secondSession);
  const observed: number[] = [];
  second.subscribe((snapshot) => observed.push(snapshot.participants.length));

  relayAwareness(firstAwareness, secondAwareness);

  const participants = second.snapshot().participants;
  expect(participants).toHaveLength(2);
  expect(participants.find((participant) => participant.local)).toMatchObject({
    actor: { id: 'grace', name: 'Grace' },
    activity: 'active',
    clientId: secondDocument.clientID,
    mode: 'view',
  });
  expect(participants.find((participant) => !participant.local)).toMatchObject({
    actor: {
      id: 'ada',
      name: 'Ada',
      kind: 'human',
      color: '#7c3aed',
      avatarUrl: 'https://example.test/ada.png',
    },
    activity: 'active',
    clientId: firstDocument.clientID,
    location: { kind: 'markdown', anchor: 2, head: 8 },
    mode: 'edit',
  });
  expect(observed.at(-1)).toBe(2);

  first.update({ activity: 'idle', location: null });
  relayAwareness(firstAwareness, secondAwareness);
  const remote = second
    .snapshot()
    .participants.find((participant) => !participant.local);
  expect(remote).toMatchObject({ activity: 'idle' });
  expect(remote?.location).toBeUndefined();

  destroyPresenceFixture(
    first,
    second,
    firstAwareness,
    secondAwareness,
    firstDocument,
    secondDocument,
  );
});

test('validates locations for every Office artifact kind', () => {
  const fixtures: Array<{
    kind: OfficeArtifactKind;
    location: OfficeCollaborationPresenceLocation;
  }> = [
    {
      kind: 'document',
      location: { kind: 'document', anchor: 4, head: 12 },
    },
    {
      kind: 'markdown',
      location: { kind: 'markdown', anchor: 1, head: 3 },
    },
    {
      kind: 'spreadsheet',
      location: {
        kind: 'spreadsheet',
        sheetId: 'sheet-1',
        ranges: [
          {
            startRow: 1,
            startColumn: 2,
            endRow: 3,
            endColumn: 4,
          },
        ],
        activeCell: { row: 1, column: 2 },
      },
    },
    {
      kind: 'presentation',
      location: {
        kind: 'presentation',
        slideId: 'slide-1',
        elementIds: ['shape-1', 'shape-2'],
      },
    },
    {
      kind: 'pdf',
      location: { kind: 'pdf', pageIndex: 2, annotationId: 'annotation-1' },
    },
  ];

  for (const fixture of fixtures) {
    const document = new Y.Doc();
    const awareness = new Awareness(document);
    const session = createOfficeCollaborationSession({
      actor: { id: `actor-${fixture.kind}`, name: 'Editor' },
      artifactId: `presence-${fixture.kind}`,
      awareness,
      document,
      kind: fixture.kind,
    });
    const presence = createOfficeCollaborationPresence(session);

    expect(presence.update({ location: fixture.location }).location).toEqual(
      fixture.location,
    );
    expect(Object.isFrozen(presence.local().location)).toBe(true);

    presence.destroy();
    awareness.destroy();
    document.destroy();
  }
});

test('rejects mismatched, malformed, and unbounded local presence', () => {
  const document = new Y.Doc();
  const awareness = new Awareness(document);
  const session = createOfficeCollaborationSession({
    actor: { id: 'editor', name: 'Editor' },
    artifactId: 'presence-bounds',
    awareness,
    document,
    kind: 'spreadsheet',
  });
  const presence = createOfficeCollaborationPresence(session);

  expect(() =>
    presence.update({
      activity: 'away',
      location: {
        kind: 'pdf',
        pageIndex: 0,
      } as OfficeCollaborationPresenceLocation,
    }),
  ).toThrow(/invalid for 'spreadsheet'/);
  expect(presence.local()).toMatchObject({
    activity: 'active',
    location: undefined,
  });
  expect(() =>
    presence.update({
      location: {
        kind: 'spreadsheet',
        sheetId: 'sheet-1',
        ranges: Array.from({ length: 65 }, () => ({
          startRow: 0,
          startColumn: 0,
          endRow: 0,
          endColumn: 0,
        })),
      },
    }),
  ).toThrow(/invalid for 'spreadsheet'/);
  expect(() => presence.update({ activity: 'busy' as 'active' })).toThrow(
    /activity 'busy' is invalid/,
  );

  presence.destroy();
  awareness.destroy();
  document.destroy();
});

test('ignores untrusted remote presence without poisoning valid participants', () => {
  const localDocument = new Y.Doc();
  const remoteDocument = new Y.Doc();
  const localAwareness = new Awareness(localDocument);
  const remoteAwareness = new Awareness(remoteDocument);
  const session = createOfficeCollaborationSession({
    actor: { id: 'local', name: 'Local editor' },
    artifactId: 'presence-untrusted',
    awareness: localAwareness,
    document: localDocument,
    kind: 'document',
  });
  const presence = createOfficeCollaborationPresence(session);
  remoteAwareness.setLocalStateField(OFFICE_COLLABORATION_PRESENCE_FIELD, {
    protocol: 'a3s.office.collaboration',
    version: 1,
    artifactId: 'another-artifact',
    artifactKind: 'document',
    namespace: 'a3s.office',
    presenceId: 'remote:1',
    actor: { id: 'remote', name: 'Remote', kind: 'human' },
    mode: 'edit',
    activity: 'active',
  });

  relayAwareness(remoteAwareness, localAwareness);

  expect(presence.snapshot().participants).toHaveLength(1);
  expect(presence.snapshot().participants[0]).toMatchObject({
    actor: { id: 'local' },
    local: true,
  });

  presence.destroy();
  localAwareness.destroy();
  remoteAwareness.destroy();
  localDocument.destroy();
  remoteDocument.destroy();
});

test('enforces one presence owner and clears only its own Awareness field', () => {
  const document = new Y.Doc();
  const awareness = new Awareness(document);
  const session = createOfficeCollaborationSession({
    actor: { id: 'owner', name: 'Owner' },
    artifactId: 'presence-owner',
    awareness,
    document,
    kind: 'pdf',
  });
  const presence = createOfficeCollaborationPresence(session);

  expect(() => createOfficeCollaborationPresence(session)).toThrow(
    /already has an A3S Office presence owner/,
  );
  awareness.setLocalStateField(OFFICE_COLLABORATION_PRESENCE_FIELD, {
    presenceId: 'replacement',
  });
  expect(() => presence.local()).toThrow(/replaced by another owner/);
  presence.destroy();
  expect(
    awareness.getLocalState()?.[OFFICE_COLLABORATION_PRESENCE_FIELD],
  ).toEqual({ presenceId: 'replacement' });

  awareness.destroy();
  document.destroy();
});

test('requires actor-bound Awareness from the same Y.Doc', () => {
  const document = new Y.Doc();
  const otherDocument = new Y.Doc();
  const otherAwareness = new Awareness(otherDocument);
  const mismatched = createOfficeCollaborationSession({
    actor: { id: 'actor', name: 'Actor' },
    artifactId: 'presence-mismatch',
    awareness: otherAwareness,
    document,
    kind: 'markdown',
  });
  expect(() => createOfficeCollaborationPresence(mismatched)).toThrow(
    /must belong to the session Y.Doc/,
  );

  const awareness = new Awareness(document);
  const anonymous = createOfficeCollaborationSession({
    artifactId: 'presence-anonymous',
    awareness,
    document,
    kind: 'markdown',
  });
  expect(() => createOfficeCollaborationPresence(anonymous)).toThrow(
    /requires a typed local actor/,
  );

  awareness.destroy();
  otherAwareness.destroy();
  document.destroy();
  otherDocument.destroy();
});

function relayAwareness(source: Awareness, target: Awareness): void {
  applyAwarenessUpdate(
    target,
    encodeAwarenessUpdate(source, [source.clientID]),
    'test-transport',
  );
}

function destroyPresenceFixture(
  first: { destroy(): void },
  second: { destroy(): void },
  firstAwareness: Awareness,
  secondAwareness: Awareness,
  firstDocument: Y.Doc,
  secondDocument: Y.Doc,
): void {
  first.destroy();
  second.destroy();
  firstAwareness.destroy();
  secondAwareness.destroy();
  firstDocument.destroy();
  secondDocument.destroy();
}
