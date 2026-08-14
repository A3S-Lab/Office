import * as Y from 'yjs';
import type { WorkArtifactKind } from '../features/work/work-types';

export const WORK_OFFICE_COLLABORATION_PROTOCOL =
  'a3s.office.collaboration' as const;
export const WORK_OFFICE_COLLABORATION_VERSION = 1 as const;
export const WORK_OFFICE_COLLABORATION_NAMESPACE = 'a3s.office' as const;

export type WorkOfficeCollaborationMode =
  | 'view'
  | 'comment'
  | 'suggest'
  | 'edit';

export type WorkOfficeCollaborationActorKind = 'human' | 'agent' | 'system';

export interface WorkOfficeCollaborationActor {
  id: string;
  name: string;
  color?: string;
  avatarUrl?: string;
  kind?: WorkOfficeCollaborationActorKind;
}

export type WorkOfficeCollaborationOriginKind =
  | 'bootstrap'
  | 'editor'
  | 'agent'
  | 'import'
  | 'system';

export interface WorkOfficeCollaborationOrigin {
  readonly protocol: typeof WORK_OFFICE_COLLABORATION_PROTOCOL;
  readonly kind: WorkOfficeCollaborationOriginKind;
  readonly actorId?: string;
  readonly operationId?: string;
}

/**
 * Structural subset of y-protocols Awareness used by Office surfaces.
 * Providers keep ownership of the Awareness instance and its transport.
 */
export interface WorkOfficeCollaborationAwareness {
  readonly clientID: number;
  getLocalState(): Record<string, unknown> | null;
  getStates(): Map<number, Record<string, unknown>>;
  setLocalStateField(field: string, value: unknown): void;
  on(
    event: 'change' | 'update',
    listener: (changes: {
      added: number[];
      updated: number[];
      removed: number[];
    }) => void,
  ): void;
  off(
    event: 'change' | 'update',
    listener: (changes: {
      added: number[];
      updated: number[];
      removed: number[];
    }) => void,
  ): void;
}

export interface WorkOfficeCollaborationMetadata {
  protocol: typeof WORK_OFFICE_COLLABORATION_PROTOCOL;
  version: typeof WORK_OFFICE_COLLABORATION_VERSION;
  artifactId: string;
  kind: WorkArtifactKind;
  initialized: boolean;
}

export interface WorkOfficeCollaborationSessionOptions {
  artifactId: string;
  kind: WorkArtifactKind;
  document?: Y.Doc;
  awareness?: WorkOfficeCollaborationAwareness;
  actor?: WorkOfficeCollaborationActor;
  mode?: WorkOfficeCollaborationMode;
  namespace?: string;
}

export interface WorkOfficeCollaborationSession {
  readonly artifactId: string;
  readonly kind: WorkArtifactKind;
  readonly document: Y.Doc;
  readonly awareness?: WorkOfficeCollaborationAwareness;
  readonly actor?: WorkOfficeCollaborationActor;
  readonly mode: WorkOfficeCollaborationMode;
  readonly namespace: string;
  readonly ownsDocument: boolean;
  readonly localOrigin: WorkOfficeCollaborationOrigin;
  metadata(): WorkOfficeCollaborationMetadata | null;
  rootName(suffix: string): string;
  createOrigin(
    kind: WorkOfficeCollaborationOriginKind,
    operationId?: string,
  ): WorkOfficeCollaborationOrigin;
  transact<T>(
    operation: (transaction: Y.Transaction) => T,
    origin?: WorkOfficeCollaborationOrigin,
  ): T;
  destroy(): void;
}

export type WorkOfficeCollaborationErrorCode =
  | 'office.collaboration.actor_invalid'
  | 'office.collaboration.artifact_mismatch'
  | 'office.collaboration.binding_destroyed'
  | 'office.collaboration.bootstrap_ambiguous'
  | 'office.collaboration.bootstrap_invalid'
  | 'office.collaboration.content_invalid'
  | 'office.collaboration.identifier_invalid'
  | 'office.collaboration.kind_invalid'
  | 'office.collaboration.kind_mismatch'
  | 'office.collaboration.metadata_invalid'
  | 'office.collaboration.metadata_missing'
  | 'office.collaboration.mode_invalid'
  | 'office.collaboration.namespace_invalid'
  | 'office.collaboration.not_initialized'
  | 'office.collaboration.origin_invalid'
  | 'office.collaboration.permission_denied'
  | 'office.collaboration.presence_conflict'
  | 'office.collaboration.presence_invalid'
  | 'office.collaboration.presence_unavailable'
  | 'office.collaboration.root_invalid'
  | 'office.collaboration.session_destroyed'
  | 'office.collaboration.transport_identity_mismatch'
  | 'office.collaboration.transport_invalid'
  | 'office.collaboration.transport_message_invalid'
  | 'office.collaboration.transport_message_too_large';

export class WorkOfficeCollaborationError extends Error {
  readonly code: WorkOfficeCollaborationErrorCode;

  constructor(code: WorkOfficeCollaborationErrorCode, message: string) {
    super(message);
    this.name = 'WorkOfficeCollaborationError';
    this.code = code;
  }
}

class WorkOfficeCollaborationSessionImpl
  implements WorkOfficeCollaborationSession
{
  readonly artifactId: string;
  readonly kind: WorkArtifactKind;
  readonly document: Y.Doc;
  readonly awareness?: WorkOfficeCollaborationAwareness;
  readonly actor?: WorkOfficeCollaborationActor;
  readonly mode: WorkOfficeCollaborationMode;
  readonly namespace: string;
  readonly ownsDocument: boolean;
  readonly localOrigin: WorkOfficeCollaborationOrigin;
  #destroyed = false;

  constructor(options: WorkOfficeCollaborationSessionOptions) {
    this.artifactId = normalizedIdentifier(options.artifactId, 'artifact ID');
    this.kind = normalizedArtifactKind(options.kind);
    this.document = options.document ?? new Y.Doc();
    this.awareness = options.awareness;
    this.actor = options.actor ? normalizedActor(options.actor) : undefined;
    this.mode = normalizedMode(options.mode ?? 'edit');
    this.namespace = normalizedNamespace(
      options.namespace ?? WORK_OFFICE_COLLABORATION_NAMESPACE,
    );
    this.ownsDocument = options.document === undefined;
    this.localOrigin = this.createOrigin(
      this.actor?.kind === 'agent'
        ? 'agent'
        : this.actor?.kind === 'system'
          ? 'system'
          : 'editor',
    );
    assertSessionMetadata(this);
  }

  metadata(): WorkOfficeCollaborationMetadata | null {
    this.ensureActive();
    return readWorkOfficeCollaborationMetadata(this);
  }

  rootName(suffix: string): string {
    this.ensureActive();
    const normalized = normalizedRootSuffix(suffix);
    return `${this.namespace}.${normalized}`;
  }

  createOrigin(
    kind: WorkOfficeCollaborationOriginKind,
    operationId?: string,
  ): WorkOfficeCollaborationOrigin {
    this.ensureActive();
    const normalizedKind = normalizedOriginKind(kind);
    return Object.freeze({
      protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
      kind: normalizedKind,
      actorId: this.actor?.id,
      operationId:
        operationId === undefined
          ? undefined
          : normalizedIdentifier(operationId, 'operation ID'),
    });
  }

  transact<T>(
    operation: (transaction: Y.Transaction) => T,
    origin: WorkOfficeCollaborationOrigin = this.localOrigin,
  ): T {
    this.ensureActive();
    assertWorkOfficeCollaborationEditable(this);
    assertWorkOfficeCollaborationOrigin(origin);
    let result: T | undefined;
    this.document.transact((transaction) => {
      result = operation(transaction);
    }, origin);
    return result as T;
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    if (this.ownsDocument) this.document.destroy();
  }

  private ensureActive(): void {
    if (!this.#destroyed) return;
    throw new WorkOfficeCollaborationError(
      'office.collaboration.session_destroyed',
      'The Office collaboration session has been destroyed.',
    );
  }
}

export function createWorkOfficeCollaborationSession(
  options: WorkOfficeCollaborationSessionOptions,
): WorkOfficeCollaborationSession {
  return new WorkOfficeCollaborationSessionImpl(options);
}

export function readWorkOfficeCollaborationMetadata(
  session: WorkOfficeCollaborationSession,
): WorkOfficeCollaborationMetadata | null {
  const metadata = session.document.getMap<unknown>(
    session.rootName('metadata'),
  );
  if (metadata.size === 0) {
    if (collaborationInitializers(session).length > 0) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.metadata_missing',
        'The shared document contains bootstrap state without Office collaboration metadata.',
      );
    }
    return null;
  }
  const protocol = metadata.get('protocol');
  const version = metadata.get('version');
  const artifactId = metadata.get('artifactId');
  const kind = metadata.get('kind');
  const initialized = metadata.get('initialized');
  if (
    protocol !== WORK_OFFICE_COLLABORATION_PROTOCOL ||
    version !== WORK_OFFICE_COLLABORATION_VERSION ||
    typeof artifactId !== 'string' ||
    !isWorkArtifactKind(kind) ||
    typeof initialized !== 'boolean'
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.metadata_invalid',
      'The shared document contains invalid or unsupported Office collaboration metadata.',
    );
  }
  const result = { protocol, version, artifactId, kind, initialized };
  assertMetadataMatches(session, result);
  assertBootstrapState(session, result);
  return result;
}

export function initializeWorkOfficeCollaborationMetadata(
  session: WorkOfficeCollaborationSession,
): WorkOfficeCollaborationMetadata {
  const existing = readWorkOfficeCollaborationMetadata(session);
  if (existing) {
    assertMetadataMatches(session, existing);
    return existing;
  }
  const metadata = session.document.getMap<unknown>(
    session.rootName('metadata'),
  );
  metadata.set('protocol', WORK_OFFICE_COLLABORATION_PROTOCOL);
  metadata.set('version', WORK_OFFICE_COLLABORATION_VERSION);
  metadata.set('artifactId', session.artifactId);
  metadata.set('kind', session.kind);
  metadata.set('initialized', false);
  return {
    protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
    version: WORK_OFFICE_COLLABORATION_VERSION,
    artifactId: session.artifactId,
    kind: session.kind,
    initialized: false,
  };
}

export function markWorkOfficeCollaborationInitialized(
  session: WorkOfficeCollaborationSession,
): void {
  const metadata = readWorkOfficeCollaborationMetadata(session);
  if (!metadata) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.metadata_missing',
      'Office collaboration metadata must be initialized before its content.',
    );
  }
  assertMetadataMatches(session, metadata);
  session.document
    .getMap<unknown>(session.rootName('metadata'))
    .set('initialized', true);
}

export function registerWorkOfficeCollaborationInitializer(
  session: WorkOfficeCollaborationSession,
): void {
  const initializers = collaborationInitializers(session);
  if (initializers.length > 0) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The shared document already contains an Office collaboration initializer.',
    );
  }
  initializers.push([
    `${session.document.clientID}:${session.actor?.id ?? 'anonymous'}`,
  ]);
}

export function assertWorkOfficeCollaborationEditable(
  session: WorkOfficeCollaborationSession,
): void {
  if (session.mode === 'edit') return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `The '${session.mode}' collaboration mode cannot modify canonical content.`,
  );
}

function assertSessionMetadata(session: WorkOfficeCollaborationSession): void {
  readWorkOfficeCollaborationMetadata(session);
}

function assertBootstrapState(
  session: WorkOfficeCollaborationSession,
  metadata: WorkOfficeCollaborationMetadata,
): void {
  const count = collaborationInitializers(session).length;
  if (count > 1) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'Multiple clients initialized the shared document before synchronization completed.',
    );
  }
  if (metadata.initialized && count !== 1) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_invalid',
      'Initialized Office collaboration metadata must have exactly one initializer.',
    );
  }
}

function collaborationInitializers(
  session: WorkOfficeCollaborationSession,
): Y.Array<string> {
  return session.document.getArray<string>(
    session.rootName('bootstrap.initializers'),
  );
}

function assertMetadataMatches(
  session: WorkOfficeCollaborationSession,
  metadata: WorkOfficeCollaborationMetadata,
): void {
  if (metadata.artifactId !== session.artifactId) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.artifact_mismatch',
      `The shared document belongs to artifact '${metadata.artifactId}', not '${session.artifactId}'.`,
    );
  }
  if (metadata.kind !== session.kind) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.kind_mismatch',
      `The shared document contains '${metadata.kind}' content, not '${session.kind}'.`,
    );
  }
}

function normalizedIdentifier(value: string, label: string): string {
  if (typeof value !== 'string') {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.identifier_invalid',
      `The ${label} must be a string containing between 1 and 256 characters.`,
    );
  }
  const normalized = value.trim();
  if (!normalized || normalized.length > 256) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.identifier_invalid',
      `The ${label} must contain between 1 and 256 characters.`,
    );
  }
  return normalized;
}

export function assertWorkOfficeCollaborationOrigin(
  origin: WorkOfficeCollaborationOrigin,
): void {
  if (!origin || origin.protocol !== WORK_OFFICE_COLLABORATION_PROTOCOL) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.origin_invalid',
      'Office collaboration transactions require a versioned collaboration origin.',
    );
  }
  normalizedOriginKind(origin.kind);
  if (origin.actorId !== undefined) {
    normalizedIdentifier(origin.actorId, 'origin actor ID');
  }
  if (origin.operationId !== undefined) {
    normalizedIdentifier(origin.operationId, 'operation ID');
  }
}

function normalizedActor(
  actor: WorkOfficeCollaborationActor,
): WorkOfficeCollaborationActor {
  const kind = actor.kind ?? 'human';
  if (kind !== 'human' && kind !== 'agent' && kind !== 'system') {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.actor_invalid',
      `The collaboration actor kind '${String(kind)}' is invalid.`,
    );
  }
  return Object.freeze({
    id: normalizedIdentifier(actor.id, 'actor ID'),
    name: normalizedIdentifier(actor.name, 'actor name'),
    color: normalizedOptionalActorField(actor.color, 'color', 64),
    avatarUrl: normalizedOptionalActorField(
      actor.avatarUrl,
      'avatar URL',
      2_048,
    ),
    kind,
  });
}

function normalizedOptionalActorField(
  value: string | undefined,
  label: string,
  maximumLength: number,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'string') {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.actor_invalid',
      `The collaboration actor ${label} must be a string when provided.`,
    );
  }
  const normalized = value.trim();
  if (!normalized || normalized.length > maximumLength) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.actor_invalid',
      `The collaboration actor ${label} must contain between 1 and ${maximumLength} characters.`,
    );
  }
  return normalized;
}

function normalizedArtifactKind(value: WorkArtifactKind): WorkArtifactKind {
  if (isWorkArtifactKind(value)) return value;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.kind_invalid',
    `The collaboration artifact kind '${String(value)}' is invalid.`,
  );
}

function normalizedMode(
  value: WorkOfficeCollaborationMode,
): WorkOfficeCollaborationMode {
  if (
    value === 'view' ||
    value === 'comment' ||
    value === 'suggest' ||
    value === 'edit'
  ) {
    return value;
  }
  throw new WorkOfficeCollaborationError(
    'office.collaboration.mode_invalid',
    `The collaboration mode '${String(value)}' is invalid.`,
  );
}

function normalizedOriginKind(
  value: WorkOfficeCollaborationOriginKind,
): WorkOfficeCollaborationOriginKind {
  if (
    value === 'bootstrap' ||
    value === 'editor' ||
    value === 'agent' ||
    value === 'import' ||
    value === 'system'
  ) {
    return value;
  }
  throw new WorkOfficeCollaborationError(
    'office.collaboration.origin_invalid',
    `The collaboration origin kind '${String(value)}' is invalid.`,
  );
}

function normalizedNamespace(value: string): string {
  const normalized = normalizedIdentifier(value, 'namespace');
  if (!/^[a-zA-Z0-9](?:[a-zA-Z0-9._-]*[a-zA-Z0-9])?$/.test(normalized)) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.namespace_invalid',
      'The collaboration namespace may contain only letters, digits, dots, underscores, and hyphens.',
    );
  }
  return normalized;
}

function normalizedRootSuffix(value: string): string {
  if (typeof value !== 'string') {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.root_invalid',
      'The collaboration root suffix must be a string.',
    );
  }
  const normalized = value.trim();
  if (!/^[a-z][a-z0-9.-]*$/.test(normalized)) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.root_invalid',
      `The collaboration root suffix '${value}' is invalid.`,
    );
  }
  return normalized;
}

function isWorkArtifactKind(value: unknown): value is WorkArtifactKind {
  return (
    value === 'document' ||
    value === 'markdown' ||
    value === 'spreadsheet' ||
    value === 'presentation' ||
    value === 'pdf'
  );
}
