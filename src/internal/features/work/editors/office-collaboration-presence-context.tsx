import { createContext, type ReactNode, useContext } from 'react';
import {
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationSession,
} from '../../../collaboration/office-collaboration';
import type { WorkOfficeCollaborationPresence } from '../../../collaboration/office-collaboration-presence';

const OfficeCollaborationPresenceContext = createContext<
  WorkOfficeCollaborationPresence | undefined
>(undefined);

export function OfficeCollaborationPresenceProvider({
  children,
  presence,
}: {
  children: ReactNode;
  presence?: WorkOfficeCollaborationPresence;
}) {
  return (
    <OfficeCollaborationPresenceContext.Provider value={presence}>
      {children}
    </OfficeCollaborationPresenceContext.Provider>
  );
}

export function useOfficeCollaborationPresence():
  | WorkOfficeCollaborationPresence
  | undefined {
  return useContext(OfficeCollaborationPresenceContext);
}

export function assertOfficeCollaborationPresencePairing({
  expectedKind,
  presence,
  session,
}: {
  expectedKind: WorkOfficeCollaborationSession['kind'];
  presence?: WorkOfficeCollaborationPresence;
  session?: WorkOfficeCollaborationSession;
}): void {
  if (!presence) return;
  if (!session) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.presence_invalid',
      'Office editor presence must be paired with its collaboration session.',
    );
  }

  const local = presence.local();
  if (session.kind !== expectedKind || local.artifactKind !== expectedKind) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.kind_mismatch',
      `Office editor presence and collaboration must both target '${expectedKind}'.`,
    );
  }
  if (local.artifactId !== session.artifactId) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.artifact_mismatch',
      'Office editor presence belongs to another collaboration artifact.',
    );
  }

  const snapshot = presence.snapshot();
  const actor = session.actor;
  if (
    local.namespace !== session.namespace ||
    snapshot.localClientId !== session.document.clientID ||
    !actor ||
    local.actor.id !== actor.id ||
    local.actor.kind !== (actor.kind ?? 'human') ||
    local.mode !== session.mode
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.presence_invalid',
      'Office editor presence does not belong to the supplied collaboration session.',
    );
  }
}
