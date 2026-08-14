import {
  createContext,
  type ReactNode,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
} from 'react';
import {
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationSession,
} from '../../../collaboration/office-collaboration';
import type {
  WorkOfficeCollaborationParticipant,
  WorkOfficeCollaborationPresence,
} from '../../../collaboration/office-collaboration-presence';

type OfficeCollaborationLocationNavigator = (
  participant: WorkOfficeCollaborationParticipant,
) => boolean;

interface OfficeCollaborationPresenceContextValue {
  presence?: WorkOfficeCollaborationPresence;
  navigateToParticipant: OfficeCollaborationLocationNavigator;
  registerLocationNavigator: (
    navigator: OfficeCollaborationLocationNavigator,
  ) => () => void;
}

const OfficeCollaborationPresenceContext =
  createContext<OfficeCollaborationPresenceContextValue | null>(null);

export function OfficeCollaborationPresenceProvider({
  children,
  presence,
}: {
  children: ReactNode;
  presence?: WorkOfficeCollaborationPresence;
}) {
  const navigatorRef = useRef<OfficeCollaborationLocationNavigator | null>(
    null,
  );
  const registerLocationNavigator = useCallback(
    (navigator: OfficeCollaborationLocationNavigator) => {
      navigatorRef.current = navigator;
      return () => {
        if (navigatorRef.current === navigator) navigatorRef.current = null;
      };
    },
    [],
  );
  const navigateToParticipant = useCallback(
    (participant: WorkOfficeCollaborationParticipant) =>
      navigatorRef.current?.(participant) ?? false,
    [],
  );
  const value = useMemo<OfficeCollaborationPresenceContextValue>(
    () => ({
      navigateToParticipant,
      presence,
      registerLocationNavigator,
    }),
    [navigateToParticipant, presence, registerLocationNavigator],
  );
  return (
    <OfficeCollaborationPresenceContext.Provider value={value}>
      {children}
    </OfficeCollaborationPresenceContext.Provider>
  );
}

export function useOfficeCollaborationPresence():
  | WorkOfficeCollaborationPresence
  | undefined {
  return useContext(OfficeCollaborationPresenceContext)?.presence;
}

export function useOfficeCollaborationLocationNavigator(
  navigator: OfficeCollaborationLocationNavigator,
): void {
  const context = useContext(OfficeCollaborationPresenceContext);
  useEffect(() => {
    if (!context?.presence) return;
    return context.registerLocationNavigator(navigator);
  }, [context, navigator]);
}

export function useOfficeCollaborationParticipantNavigation():
  | OfficeCollaborationLocationNavigator
  | undefined {
  const context = useContext(OfficeCollaborationPresenceContext);
  return context?.presence ? context.navigateToParticipant : undefined;
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
