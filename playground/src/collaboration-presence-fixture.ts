import {
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  type DocumentContent,
  initializeOfficeDocumentCollaboration,
  type OfficeCollaborationPresence,
  type OfficeCollaborationSession,
} from '@a3s-lab/office/core';
import { useEffect, useRef, useState } from 'react';
import {
  Awareness,
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
} from 'y-protocols/awareness';
import * as Y from 'yjs';

export interface PlaygroundCollaborationPresenceFixture {
  readonly collaboration: OfficeCollaborationSession;
  readonly presence: OfficeCollaborationPresence;
}

export function usePlaygroundCollaborationPresenceFixture(
  options:
    | {
        artifactId: string;
        content: DocumentContent;
      }
    | undefined,
): PlaygroundCollaborationPresenceFixture | undefined {
  const initialOptions = useRef(options);
  const [fixture, setFixture] =
    useState<OwnedPlaygroundCollaborationPresenceFixture>();

  useEffect(() => {
    const fixtureOptions = initialOptions.current;
    if (!fixtureOptions) return;
    const nextFixture = createCollaborationPresenceFixture(fixtureOptions);
    setFixture(nextFixture);
    return () => nextFixture.destroy();
  }, []);

  return fixture;
}

interface OwnedPlaygroundCollaborationPresenceFixture
  extends PlaygroundCollaborationPresenceFixture {
  destroy(): void;
}

function createCollaborationPresenceFixture({
  artifactId,
  content,
}: {
  artifactId: string;
  content: DocumentContent;
}): OwnedPlaygroundCollaborationPresenceFixture {
  const localDocument = new Y.Doc();
  const remoteDocument = new Y.Doc();
  const localAwareness = new Awareness(localDocument);
  const remoteAwareness = new Awareness(remoteDocument);
  const collaboration = createOfficeCollaborationSession({
    actor: { id: 'playground-user', name: '林澄', color: '#047857' },
    artifactId,
    awareness: localAwareness,
    document: localDocument,
    kind: 'document',
  });
  const remoteSession = createOfficeCollaborationSession({
    actor: {
      id: 'playground-agent',
      name: 'A3S Agent',
      color: '#6d28d9',
      kind: 'agent',
    },
    artifactId,
    awareness: remoteAwareness,
    document: remoteDocument,
    kind: 'document',
    mode: 'suggest',
  });

  initializeOfficeDocumentCollaboration(collaboration, content);
  const presence = createOfficeCollaborationPresence(collaboration);
  const remotePresence = createOfficeCollaborationPresence(remoteSession, {
    location: { kind: 'document', anchor: 3, head: 9 },
  });
  applyAwarenessUpdate(
    localAwareness,
    encodeAwarenessUpdate(remoteAwareness, [remoteAwareness.clientID]),
    'playground-e2e',
  );

  return {
    collaboration,
    presence,
    destroy() {
      presence.destroy();
      remotePresence.destroy();
      collaboration.destroy();
      remoteSession.destroy();
      localAwareness.destroy();
      remoteAwareness.destroy();
      localDocument.destroy();
      remoteDocument.destroy();
    },
  };
}
