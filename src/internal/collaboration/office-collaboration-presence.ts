import {
  WORK_OFFICE_COLLABORATION_PROTOCOL,
  WORK_OFFICE_COLLABORATION_VERSION,
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationActor,
  type WorkOfficeCollaborationActorKind,
  type WorkOfficeCollaborationMode,
  type WorkOfficeCollaborationSession,
} from './office-collaboration';

export const WORK_OFFICE_COLLABORATION_PRESENCE_FIELD = 'a3sOffice' as const;

const MAX_PRESENCE_RANGES = 64;
const MAX_PRESENCE_ELEMENT_IDS = 128;

export type WorkOfficeCollaborationPresenceActivity =
  | 'active'
  | 'idle'
  | 'away';

export interface WorkOfficeTextPresenceLocation {
  readonly kind: 'document' | 'markdown';
  /** Zero-based model position; Markdown positions are UTF-16 offsets. */
  readonly anchor: number;
  readonly head: number;
}

export interface WorkOfficeSpreadsheetPresenceCell {
  readonly row: number;
  readonly column: number;
}

export interface WorkOfficeSpreadsheetPresenceRange {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

export interface WorkOfficeSpreadsheetPresenceLocation {
  readonly kind: 'spreadsheet';
  readonly sheetId: string;
  readonly ranges: readonly WorkOfficeSpreadsheetPresenceRange[];
  readonly activeCell?: WorkOfficeSpreadsheetPresenceCell;
}

export interface WorkOfficePresentationPresenceLocation {
  readonly kind: 'presentation';
  readonly slideId: string;
  readonly elementIds: readonly string[];
}

export interface WorkOfficePdfPresenceLocation {
  readonly kind: 'pdf';
  /** Zero-based page index. */
  readonly pageIndex: number;
  readonly annotationId?: string;
}

export type WorkOfficeCollaborationPresenceLocation =
  | WorkOfficeTextPresenceLocation
  | WorkOfficeSpreadsheetPresenceLocation
  | WorkOfficePresentationPresenceLocation
  | WorkOfficePdfPresenceLocation;

export interface WorkOfficeCollaborationPresenceActor {
  readonly id: string;
  readonly name: string;
  readonly color?: string;
  readonly avatarUrl?: string;
  readonly kind: WorkOfficeCollaborationActorKind;
}

export interface WorkOfficeCollaborationPresenceState {
  readonly protocol: typeof WORK_OFFICE_COLLABORATION_PROTOCOL;
  readonly version: typeof WORK_OFFICE_COLLABORATION_VERSION;
  readonly artifactId: string;
  readonly artifactKind: WorkOfficeCollaborationSession['kind'];
  readonly namespace: string;
  readonly presenceId: string;
  readonly actor: WorkOfficeCollaborationPresenceActor;
  readonly mode: WorkOfficeCollaborationMode;
  readonly activity: WorkOfficeCollaborationPresenceActivity;
  readonly location?: WorkOfficeCollaborationPresenceLocation;
}

export interface WorkOfficeCollaborationParticipant
  extends WorkOfficeCollaborationPresenceState {
  readonly clientId: number;
  readonly local: boolean;
}

export interface WorkOfficeCollaborationPresenceSnapshot {
  readonly localClientId: number;
  readonly participants: readonly WorkOfficeCollaborationParticipant[];
}

export interface WorkOfficeCollaborationPresenceUpdate {
  readonly activity?: WorkOfficeCollaborationPresenceActivity;
  /** `null` clears the local location; omission keeps the current location. */
  readonly location?: WorkOfficeCollaborationPresenceLocation | null;
}

export interface WorkOfficeCollaborationPresence {
  local(): WorkOfficeCollaborationPresenceState;
  snapshot(): WorkOfficeCollaborationPresenceSnapshot;
  update(
    update: WorkOfficeCollaborationPresenceUpdate,
  ): WorkOfficeCollaborationPresenceState;
  subscribe(
    listener: (snapshot: WorkOfficeCollaborationPresenceSnapshot) => void,
  ): () => void;
  destroy(): void;
}

export function createWorkOfficeCollaborationPresence(
  session: WorkOfficeCollaborationSession,
  initial: WorkOfficeCollaborationPresenceUpdate = {},
): WorkOfficeCollaborationPresence {
  return new WorkOfficeCollaborationPresenceImpl(session, initial);
}

let nextPresenceSequence = 0;

class WorkOfficeCollaborationPresenceImpl
  implements WorkOfficeCollaborationPresence
{
  readonly #session: WorkOfficeCollaborationSession;
  readonly #awareness: NonNullable<WorkOfficeCollaborationSession['awareness']>;
  readonly #actor: WorkOfficeCollaborationPresenceActor;
  readonly #presenceId: string;
  readonly #listeners = new Set<
    (snapshot: WorkOfficeCollaborationPresenceSnapshot) => void
  >();
  #activity: WorkOfficeCollaborationPresenceActivity;
  #location: WorkOfficeCollaborationPresenceLocation | undefined;
  #destroyed = false;

  constructor(
    session: WorkOfficeCollaborationSession,
    initial: WorkOfficeCollaborationPresenceUpdate,
  ) {
    session.metadata();
    if (!session.awareness) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.presence_unavailable',
        'Office collaboration presence requires host-provided Yjs Awareness.',
      );
    }
    if (!session.actor) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.presence_unavailable',
        'Office collaboration presence requires a typed local actor.',
      );
    }
    if (
      !isClientId(session.awareness.clientID) ||
      session.awareness.clientID !== session.document.clientID
    ) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.presence_unavailable',
        'Office collaboration Awareness must belong to the session Y.Doc.',
      );
    }
    const existing =
      session.awareness.getLocalState()?.[
        WORK_OFFICE_COLLABORATION_PRESENCE_FIELD
      ];
    if (existing !== undefined && existing !== null) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.presence_conflict',
        'The local Awareness state already has an A3S Office presence owner.',
      );
    }

    this.#session = session;
    this.#awareness = session.awareness;
    this.#actor = copiedActor(session.actor);
    this.#presenceId = createPresenceId(session.document.clientID);
    this.#activity = normalizedActivity(initial.activity ?? 'active');
    this.#location =
      initial.location === undefined || initial.location === null
        ? undefined
        : normalizedLocalLocation(initial.location, session.kind);

    this.#awareness.on('change', this.#onAwarenessChange);
    try {
      this.publish();
    } catch (error) {
      this.#awareness.off('change', this.#onAwarenessChange);
      throw error;
    }
  }

  local(): WorkOfficeCollaborationPresenceState {
    this.ensureActive();
    this.ensureOwnership();
    return this.state();
  }

  snapshot(): WorkOfficeCollaborationPresenceSnapshot {
    this.ensureActive();
    const participants: WorkOfficeCollaborationParticipant[] = [];
    for (const [clientId, awarenessState] of this.#awareness.getStates()) {
      if (!isClientId(clientId) || !isRecord(awarenessState)) continue;
      const state = validatedPresenceState(
        awarenessState[WORK_OFFICE_COLLABORATION_PRESENCE_FIELD],
        this.#session,
      );
      if (!state) continue;
      participants.push(
        Object.freeze({
          ...state,
          clientId,
          local: clientId === this.#awareness.clientID,
        }),
      );
    }
    participants.sort((left, right) => left.clientId - right.clientId);
    return Object.freeze({
      localClientId: this.#awareness.clientID,
      participants: Object.freeze(participants),
    });
  }

  update(
    update: WorkOfficeCollaborationPresenceUpdate,
  ): WorkOfficeCollaborationPresenceState {
    this.ensureActive();
    this.ensureOwnership();
    if (!isRecord(update)) {
      throw invalidPresence('Presence updates must be plain objects.');
    }
    const nextActivity =
      update.activity === undefined
        ? this.#activity
        : normalizedActivity(update.activity);
    const nextLocation = Object.hasOwn(update, 'location')
      ? update.location === null
        ? undefined
        : normalizedLocalLocation(update.location, this.#session.kind)
      : this.#location;
    this.#activity = nextActivity;
    this.#location = nextLocation;
    this.publish();
    return this.state();
  }

  subscribe(
    listener: (snapshot: WorkOfficeCollaborationPresenceSnapshot) => void,
  ): () => void {
    this.ensureActive();
    if (typeof listener !== 'function') {
      throw invalidPresence('Presence listeners must be functions.');
    }
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#awareness.off('change', this.#onAwarenessChange);
    this.#listeners.clear();
    const current =
      this.#awareness.getLocalState()?.[
        WORK_OFFICE_COLLABORATION_PRESENCE_FIELD
      ];
    if (isRecord(current) && current.presenceId === this.#presenceId) {
      this.#awareness.setLocalStateField(
        WORK_OFFICE_COLLABORATION_PRESENCE_FIELD,
        null,
      );
    }
  }

  readonly #onAwarenessChange = (): void => {
    if (this.#destroyed) return;
    const snapshot = this.snapshot();
    for (const listener of this.#listeners) listener(snapshot);
  };

  private publish(): void {
    this.#awareness.setLocalStateField(
      WORK_OFFICE_COLLABORATION_PRESENCE_FIELD,
      this.state(),
    );
  }

  private state(): WorkOfficeCollaborationPresenceState {
    return Object.freeze({
      protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
      version: WORK_OFFICE_COLLABORATION_VERSION,
      artifactId: this.#session.artifactId,
      artifactKind: this.#session.kind,
      namespace: this.#session.namespace,
      presenceId: this.#presenceId,
      actor: this.#actor,
      mode: this.#session.mode,
      activity: this.#activity,
      location: this.#location,
    });
  }

  private ensureOwnership(): void {
    const current =
      this.#awareness.getLocalState()?.[
        WORK_OFFICE_COLLABORATION_PRESENCE_FIELD
      ];
    if (!isRecord(current) || current.presenceId !== this.#presenceId) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.presence_conflict',
        'The local A3S Office presence field was replaced by another owner.',
      );
    }
  }

  private ensureActive(): void {
    if (!this.#destroyed) return;
    throw new WorkOfficeCollaborationError(
      'office.collaboration.binding_destroyed',
      'The Office collaboration presence controller has been destroyed.',
    );
  }
}

function validatedPresenceState(
  value: unknown,
  session: WorkOfficeCollaborationSession,
): WorkOfficeCollaborationPresenceState | null {
  if (
    !isRecord(value) ||
    value.protocol !== WORK_OFFICE_COLLABORATION_PROTOCOL ||
    value.version !== WORK_OFFICE_COLLABORATION_VERSION ||
    value.artifactId !== session.artifactId ||
    value.artifactKind !== session.kind ||
    value.namespace !== session.namespace ||
    !isIdentifier(value.presenceId) ||
    !isMode(value.mode) ||
    !isActivity(value.activity)
  ) {
    return null;
  }
  const actor = validatedActor(value.actor);
  if (!actor) return null;
  const location =
    value.location === undefined
      ? undefined
      : validatedLocation(value.location, session.kind);
  if (value.location !== undefined && location === null) return null;
  return Object.freeze({
    protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
    version: WORK_OFFICE_COLLABORATION_VERSION,
    artifactId: session.artifactId,
    artifactKind: session.kind,
    namespace: session.namespace,
    presenceId: value.presenceId,
    actor,
    mode: value.mode,
    activity: value.activity,
    location: location ?? undefined,
  });
}

function normalizedLocalLocation(
  value: unknown,
  kind: WorkOfficeCollaborationSession['kind'],
): WorkOfficeCollaborationPresenceLocation {
  const location = validatedLocation(value, kind);
  if (!location) {
    throw invalidPresence(
      `The Office collaboration presence location is invalid for '${kind}'.`,
    );
  }
  return location;
}

function validatedLocation(
  value: unknown,
  kind: WorkOfficeCollaborationSession['kind'],
): WorkOfficeCollaborationPresenceLocation | null {
  if (!isRecord(value) || value.kind !== kind) return null;
  if (kind === 'document' || kind === 'markdown') {
    if (!isPosition(value.anchor) || !isPosition(value.head)) return null;
    return Object.freeze({ kind, anchor: value.anchor, head: value.head });
  }
  if (kind === 'spreadsheet') {
    if (
      !isIdentifier(value.sheetId) ||
      !Array.isArray(value.ranges) ||
      value.ranges.length < 1 ||
      value.ranges.length > MAX_PRESENCE_RANGES
    ) {
      return null;
    }
    const ranges: WorkOfficeSpreadsheetPresenceRange[] = [];
    for (const candidate of value.ranges) {
      const range = validatedSpreadsheetRange(candidate);
      if (!range) return null;
      ranges.push(range);
    }
    const activeCell =
      value.activeCell === undefined
        ? undefined
        : validatedSpreadsheetCell(value.activeCell);
    if (value.activeCell !== undefined && !activeCell) return null;
    return Object.freeze({
      kind,
      sheetId: value.sheetId,
      ranges: Object.freeze(ranges),
      activeCell: activeCell ?? undefined,
    });
  }
  if (kind === 'presentation') {
    if (
      !isIdentifier(value.slideId) ||
      !Array.isArray(value.elementIds) ||
      value.elementIds.length > MAX_PRESENCE_ELEMENT_IDS
    ) {
      return null;
    }
    const elementIds: string[] = [];
    const seen = new Set<string>();
    for (const candidate of value.elementIds) {
      if (!isIdentifier(candidate) || seen.has(candidate)) return null;
      seen.add(candidate);
      elementIds.push(candidate);
    }
    return Object.freeze({
      kind,
      slideId: value.slideId,
      elementIds: Object.freeze(elementIds),
    });
  }
  if (!isPosition(value.pageIndex)) return null;
  if (value.annotationId !== undefined && !isIdentifier(value.annotationId)) {
    return null;
  }
  return Object.freeze({
    kind: 'pdf',
    pageIndex: value.pageIndex,
    annotationId: value.annotationId as string | undefined,
  });
}

function validatedSpreadsheetRange(
  value: unknown,
): WorkOfficeSpreadsheetPresenceRange | null {
  if (
    !isRecord(value) ||
    !isPosition(value.startRow) ||
    !isPosition(value.startColumn) ||
    !isPosition(value.endRow) ||
    !isPosition(value.endColumn) ||
    value.startRow > value.endRow ||
    value.startColumn > value.endColumn
  ) {
    return null;
  }
  return Object.freeze({
    startRow: value.startRow,
    startColumn: value.startColumn,
    endRow: value.endRow,
    endColumn: value.endColumn,
  });
}

function validatedSpreadsheetCell(
  value: unknown,
): WorkOfficeSpreadsheetPresenceCell | null {
  if (!isRecord(value) || !isPosition(value.row) || !isPosition(value.column)) {
    return null;
  }
  return Object.freeze({ row: value.row, column: value.column });
}

function validatedActor(
  value: unknown,
): WorkOfficeCollaborationPresenceActor | null {
  if (
    !isRecord(value) ||
    !isIdentifier(value.id) ||
    !isIdentifier(value.name) ||
    !isActorKind(value.kind) ||
    !isOptionalBoundedString(value.color, 64) ||
    !isOptionalBoundedString(value.avatarUrl, 2_048)
  ) {
    return null;
  }
  return Object.freeze({
    id: value.id,
    name: value.name,
    kind: value.kind,
    color: value.color as string | undefined,
    avatarUrl: value.avatarUrl as string | undefined,
  });
}

function copiedActor(
  actor: WorkOfficeCollaborationActor,
): WorkOfficeCollaborationPresenceActor {
  return Object.freeze({
    id: actor.id,
    name: actor.name,
    kind: actor.kind ?? 'human',
    color: actor.color,
    avatarUrl: actor.avatarUrl,
  });
}

function normalizedActivity(
  value: unknown,
): WorkOfficeCollaborationPresenceActivity {
  if (isActivity(value)) return value;
  throw invalidPresence(
    `The Office collaboration presence activity '${String(value)}' is invalid.`,
  );
}

function createPresenceId(clientId: number): string {
  nextPresenceSequence =
    nextPresenceSequence >= Number.MAX_SAFE_INTEGER
      ? 1
      : nextPresenceSequence + 1;
  return `${clientId}:${nextPresenceSequence.toString(36)}`;
}

function invalidPresence(message: string): WorkOfficeCollaborationError {
  return new WorkOfficeCollaborationError(
    'office.collaboration.presence_invalid',
    message,
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isIdentifier(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    value.length >= 1 &&
    value.length <= 256 &&
    value.trim() === value
  );
}

function isOptionalBoundedString(
  value: unknown,
  maximumLength: number,
): boolean {
  return (
    value === undefined ||
    (typeof value === 'string' &&
      value.length >= 1 &&
      value.length <= maximumLength &&
      value.trim() === value)
  );
}

function isPosition(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0;
}

function isClientId(value: unknown): value is number {
  return isPosition(value);
}

function isActivity(
  value: unknown,
): value is WorkOfficeCollaborationPresenceActivity {
  return value === 'active' || value === 'idle' || value === 'away';
}

function isMode(value: unknown): value is WorkOfficeCollaborationMode {
  return (
    value === 'view' ||
    value === 'comment' ||
    value === 'suggest' ||
    value === 'edit'
  );
}

function isActorKind(
  value: unknown,
): value is NonNullable<WorkOfficeCollaborationActor['kind']> {
  return value === 'human' || value === 'agent' || value === 'system';
}
