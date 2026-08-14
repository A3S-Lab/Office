import * as Y from 'yjs';
import {
  assertWorkOfficeCollaborationOrigin,
  WORK_OFFICE_COLLABORATION_PROTOCOL,
  WORK_OFFICE_COLLABORATION_VERSION,
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationOrigin,
  type WorkOfficeCollaborationSession,
} from './office-collaboration';

export const WORK_OFFICE_COLLABORATION_DEFAULT_MAX_TRANSPORT_PAYLOAD_BYTES =
  64 * 1_024 * 1_024;
export const WORK_OFFICE_COLLABORATION_MAX_TRANSPORT_PAYLOAD_BYTES =
  256 * 1_024 * 1_024;

export type WorkOfficeCollaborationTransportMessageType =
  | 'sync-step-1'
  | 'sync-step-2'
  | 'update';

/**
 * A structured-clone-safe envelope around standard Yjs v1 state vectors and
 * updates. The host may encode this envelope for its own wire protocol.
 */
export interface WorkOfficeCollaborationTransportMessage {
  readonly protocol: typeof WORK_OFFICE_COLLABORATION_PROTOCOL;
  readonly version: typeof WORK_OFFICE_COLLABORATION_VERSION;
  readonly artifactId: string;
  readonly artifactKind: WorkOfficeCollaborationSession['kind'];
  readonly namespace: string;
  readonly senderClientId: number;
  readonly type: WorkOfficeCollaborationTransportMessageType;
  readonly payload: Uint8Array;
  /** Present only for incremental updates whose Yjs transaction had a typed origin. */
  readonly origin?: WorkOfficeCollaborationOrigin;
}

/**
 * Host-owned room channel. It owns connectivity, buffering, authentication,
 * authorization, persistence, and delivery ordering.
 */
export interface WorkOfficeCollaborationTransport {
  publish(message: WorkOfficeCollaborationTransportMessage): void;
  subscribe(listener: (message: unknown) => void): () => void;
}

export interface WorkOfficeCollaborationTransportBindingOptions {
  /**
   * Sends one SyncStep1 immediately after subscribing. Disable when the host
   * must wait for a room connection, then call `synchronize()` on connect.
   */
  autoSynchronize?: boolean;
  maxPayloadBytes?: number;
}

export interface WorkOfficeCollaborationTransportBinding {
  /** Send a fresh state vector. Call this after every transport reconnect. */
  synchronize(): void;
  destroy(): void;
}

export function createWorkOfficeCollaborationTransportBinding(
  session: WorkOfficeCollaborationSession,
  transport: WorkOfficeCollaborationTransport,
  options: WorkOfficeCollaborationTransportBindingOptions = {},
): WorkOfficeCollaborationTransportBinding {
  return new WorkOfficeCollaborationTransportBindingImpl(
    session,
    transport,
    options,
  );
}

class WorkOfficeCollaborationTransportBindingImpl
  implements WorkOfficeCollaborationTransportBinding
{
  readonly #session: WorkOfficeCollaborationSession;
  readonly #transport: WorkOfficeCollaborationTransport;
  readonly #maxPayloadBytes: number;
  readonly #remoteOrigins = new WeakSet<object>();
  #unsubscribe: (() => void) | undefined;
  #destroyed = false;

  constructor(
    session: WorkOfficeCollaborationSession,
    transport: WorkOfficeCollaborationTransport,
    options: WorkOfficeCollaborationTransportBindingOptions,
  ) {
    session.metadata();
    assertTransport(transport);
    this.#session = session;
    this.#transport = transport;
    this.#maxPayloadBytes = normalizedMaximumPayloadBytes(
      options.maxPayloadBytes ??
        WORK_OFFICE_COLLABORATION_DEFAULT_MAX_TRANSPORT_PAYLOAD_BYTES,
    );

    session.document.on('update', this.#onDocumentUpdate);
    try {
      const unsubscribe = transport.subscribe(this.#onTransportMessage);
      if (typeof unsubscribe !== 'function') {
        throw new WorkOfficeCollaborationError(
          'office.collaboration.transport_invalid',
          'The Office collaboration transport subscribe function must return an unsubscribe function.',
        );
      }
      this.#unsubscribe = unsubscribe;
      if (options.autoSynchronize !== false) this.synchronize();
    } catch (error) {
      session.document.off('update', this.#onDocumentUpdate);
      this.#unsubscribe?.();
      this.#unsubscribe = undefined;
      this.#destroyed = true;
      throw error;
    }
  }

  synchronize(): void {
    this.ensureActive();
    this.publish('sync-step-1', Y.encodeStateVector(this.#session.document));
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#session.document.off('update', this.#onDocumentUpdate);
    const unsubscribe = this.#unsubscribe;
    this.#unsubscribe = undefined;
    unsubscribe?.();
  }

  readonly #onDocumentUpdate = (update: Uint8Array, origin: unknown): void => {
    if (this.#destroyed || isRemoteOrigin(this.#remoteOrigins, origin)) return;
    this.publish('update', update, copiedLocalOrigin(origin));
  };

  readonly #onTransportMessage = (value: unknown): void => {
    if (this.#destroyed) return;
    const message = validatedTransportMessage(
      value,
      this.#session,
      this.#maxPayloadBytes,
    );
    if (message.senderClientId === this.#session.document.clientID) return;

    try {
      if (message.type === 'sync-step-1') {
        this.publish(
          'sync-step-2',
          Y.encodeStateAsUpdate(this.#session.document, message.payload),
        );
        return;
      }
      const remoteOrigin = message.origin ?? remoteSystemOrigin();
      this.#remoteOrigins.add(remoteOrigin);
      Y.applyUpdate(this.#session.document, message.payload, remoteOrigin);
      this.#session.metadata();
    } catch (error) {
      if (error instanceof WorkOfficeCollaborationError) throw error;
      throw invalidTransportMessage(
        `The '${message.type}' payload is not a valid Yjs v1 message: ${errorMessage(error)}`,
      );
    }
  };

  private publish(
    type: WorkOfficeCollaborationTransportMessageType,
    payload: Uint8Array,
    origin?: WorkOfficeCollaborationOrigin,
  ): void {
    assertPayloadSize(payload, this.#maxPayloadBytes);
    const message = Object.freeze({
      protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
      version: WORK_OFFICE_COLLABORATION_VERSION,
      artifactId: this.#session.artifactId,
      artifactKind: this.#session.kind,
      namespace: this.#session.namespace,
      senderClientId: this.#session.document.clientID,
      type,
      payload: payload.slice(),
      origin,
    }) satisfies WorkOfficeCollaborationTransportMessage;
    this.#transport.publish(message);
  }

  private ensureActive(): void {
    if (!this.#destroyed) return;
    throw new WorkOfficeCollaborationError(
      'office.collaboration.binding_destroyed',
      'The Office collaboration transport binding has been destroyed.',
    );
  }
}

function validatedTransportMessage(
  value: unknown,
  session: WorkOfficeCollaborationSession,
  maximumPayloadBytes: number,
): WorkOfficeCollaborationTransportMessage {
  if (!isRecord(value)) {
    throw invalidTransportMessage(
      'Office collaboration transport messages must be plain objects.',
    );
  }
  if (
    value.protocol !== WORK_OFFICE_COLLABORATION_PROTOCOL ||
    value.version !== WORK_OFFICE_COLLABORATION_VERSION
  ) {
    throw invalidTransportMessage(
      'The Office collaboration transport protocol or version is unsupported.',
    );
  }
  if (
    value.artifactId !== session.artifactId ||
    value.artifactKind !== session.kind ||
    value.namespace !== session.namespace
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.transport_identity_mismatch',
      'The Office collaboration transport message belongs to another artifact, kind, or namespace.',
    );
  }
  if (!isClientId(value.senderClientId)) {
    throw invalidTransportMessage(
      'The Office collaboration transport sender client ID is invalid.',
    );
  }
  if (!isTransportMessageType(value.type)) {
    throw invalidTransportMessage(
      `The Office collaboration transport message type '${String(value.type)}' is invalid.`,
    );
  }
  if (!(value.payload instanceof Uint8Array)) {
    throw invalidTransportMessage(
      'The Office collaboration transport payload must be a Uint8Array.',
    );
  }
  assertPayloadSize(value.payload, maximumPayloadBytes);
  const origin = copiedTransportOrigin(value.origin);
  if (value.type !== 'update' && origin !== undefined) {
    throw invalidTransportMessage(
      'Only incremental Office collaboration update messages may carry a transaction origin.',
    );
  }
  return Object.freeze({
    protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
    version: WORK_OFFICE_COLLABORATION_VERSION,
    artifactId: session.artifactId,
    artifactKind: session.kind,
    namespace: session.namespace,
    senderClientId: value.senderClientId,
    type: value.type,
    payload: value.payload.slice(),
    origin,
  });
}

function copiedLocalOrigin(
  value: unknown,
): WorkOfficeCollaborationOrigin | undefined {
  try {
    return copiedTransportOrigin(value);
  } catch {
    return undefined;
  }
}

function copiedTransportOrigin(
  value: unknown,
): WorkOfficeCollaborationOrigin | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) {
    throw invalidTransportMessage(
      'The Office collaboration transport origin must be a typed origin object.',
    );
  }
  try {
    assertWorkOfficeCollaborationOrigin(
      value as unknown as WorkOfficeCollaborationOrigin,
    );
  } catch {
    throw invalidTransportMessage(
      'The Office collaboration transport origin is invalid.',
    );
  }
  return Object.freeze({
    protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
    kind: value.kind as WorkOfficeCollaborationOrigin['kind'],
    actorId: value.actorId as string | undefined,
    operationId: value.operationId as string | undefined,
  });
}

function remoteSystemOrigin(): WorkOfficeCollaborationOrigin {
  return Object.freeze({
    protocol: WORK_OFFICE_COLLABORATION_PROTOCOL,
    kind: 'system',
  });
}

function isRemoteOrigin(origins: WeakSet<object>, value: unknown): boolean {
  return typeof value === 'object' && value !== null && origins.has(value);
}

function assertTransport(value: WorkOfficeCollaborationTransport): void {
  if (
    !value ||
    typeof value.publish !== 'function' ||
    typeof value.subscribe !== 'function'
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.transport_invalid',
      'An Office collaboration transport must provide publish and subscribe functions.',
    );
  }
}

function normalizedMaximumPayloadBytes(value: number): number {
  if (
    !Number.isSafeInteger(value) ||
    value < 1 ||
    value > WORK_OFFICE_COLLABORATION_MAX_TRANSPORT_PAYLOAD_BYTES
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.transport_invalid',
      `The Office collaboration transport payload limit must be an integer between 1 and ${WORK_OFFICE_COLLABORATION_MAX_TRANSPORT_PAYLOAD_BYTES} bytes.`,
    );
  }
  return value;
}

function assertPayloadSize(
  payload: Uint8Array,
  maximumPayloadBytes: number,
): void {
  if (payload.byteLength <= maximumPayloadBytes) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.transport_message_too_large',
    `The Office collaboration transport payload is ${payload.byteLength} bytes; the configured limit is ${maximumPayloadBytes} bytes.`,
  );
}

function invalidTransportMessage(
  message: string,
): WorkOfficeCollaborationError {
  return new WorkOfficeCollaborationError(
    'office.collaboration.transport_message_invalid',
    message,
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isClientId(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) >= 0;
}

function isTransportMessageType(
  value: unknown,
): value is WorkOfficeCollaborationTransportMessageType {
  return (
    value === 'sync-step-1' || value === 'sync-step-2' || value === 'update'
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
