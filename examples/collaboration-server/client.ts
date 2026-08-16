import {
  createOfficeCollaborationTransportBinding,
  type OfficeCollaborationSession,
  type OfficeCollaborationTransport,
  type OfficeCollaborationTransportBinding,
  type OfficeCollaborationTransportMessage,
} from '@a3s-lab/office/core';
import {
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
  type Awareness,
  removeAwarenessStates,
} from 'y-protocols/awareness';

const PROTOCOL = 'a3s.office.collaboration' as const;
const VERSION = 1 as const;
const REMOTE_AWARENESS_ORIGIN = Object.freeze({
  protocol: PROTOCOL,
  kind: 'remote-awareness',
});

export type CollaborationConnectionStatus =
  | 'connecting'
  | 'connected'
  | 'reconnecting'
  | 'disconnected'
  | 'error';

export interface A3sBootCollaborationRoomOptions {
  /** The short-lived URL returned by POST /api/collaboration/tickets. */
  webSocketUrl: string;
  session: OfficeCollaborationSession;
  awareness?: Awareness;
  reconnect?: boolean;
  onStatus?: (status: CollaborationConnectionStatus, detail?: string) => void;
}

export interface A3sBootCollaborationRoom {
  readonly transport: OfficeCollaborationTransport;
  destroy(): void;
}

/**
 * Connect one A3S Office Y.Doc and optional Awareness instance to the runnable
 * A3S Boot backend example. Reconnects use a fresh two-way state-vector sync,
 * so offline document changes do not need an additional client-side queue.
 */
export function connectA3sBootCollaborationRoom(
  options: A3sBootCollaborationRoomOptions,
): A3sBootCollaborationRoom {
  return new A3sBootCollaborationRoomImpl(options);
}

class A3sBootCollaborationRoomImpl implements A3sBootCollaborationRoom {
  readonly transport: OfficeCollaborationTransport;
  readonly #options: A3sBootCollaborationRoomOptions;
  readonly #session: OfficeCollaborationSession;
  readonly #awareness: Awareness | undefined;
  readonly #listeners = new Set<(message: unknown) => void>();
  readonly #remoteClientIds = new Set<number>();
  readonly #binding: OfficeCollaborationTransportBinding;
  #socket: WebSocket | undefined;
  #reconnectTimer: ReturnType<typeof setTimeout> | undefined;
  #reconnectAttempt = 0;
  #ready = false;
  #destroyed = false;

  constructor(options: A3sBootCollaborationRoomOptions) {
    if (
      !options.webSocketUrl.startsWith('ws://') &&
      !options.webSocketUrl.startsWith('wss://')
    ) {
      throw new TypeError('webSocketUrl must use ws:// or wss://.');
    }
    if (
      options.awareness &&
      options.awareness.clientID !== options.session.document.clientID
    ) {
      throw new TypeError(
        'Awareness must belong to the collaboration session Y.Doc.',
      );
    }
    this.#options = options;
    this.#session = options.session;
    this.#awareness = options.awareness;
    this.transport = Object.freeze({
      publish: (message: OfficeCollaborationTransportMessage) =>
        this.publish(message),
      subscribe: (listener: (message: unknown) => void) =>
        this.subscribe(listener),
    });
    this.#binding = createOfficeCollaborationTransportBinding(
      this.#session,
      this.transport,
      { autoSynchronize: false },
    );
    this.#awareness?.on('update', this.#onAwarenessUpdate);
    this.connect('connecting');
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#ready = false;
    if (this.#reconnectTimer !== undefined) {
      clearTimeout(this.#reconnectTimer);
      this.#reconnectTimer = undefined;
    }
    this.#awareness?.off('update', this.#onAwarenessUpdate);
    this.clearRemoteAwareness();
    this.#binding.destroy();
    this.#listeners.clear();
    this.#socket?.close(1000, 'client destroyed');
    this.#socket = undefined;
    this.#options.onStatus?.('disconnected');
  }

  private connect(status: 'connecting' | 'reconnecting'): void {
    if (this.#destroyed) return;
    this.#options.onStatus?.(status);
    const socket = new WebSocket(this.#options.webSocketUrl);
    this.#socket = socket;
    socket.addEventListener('open', () => {
      if (this.#destroyed || socket !== this.#socket) return;
      this.send('collaboration.hello', {
        protocol: PROTOCOL,
        version: VERSION,
        senderClientId: this.#session.document.clientID,
      });
    });
    socket.addEventListener('message', (event) => {
      if (this.#destroyed || socket !== this.#socket) return;
      try {
        this.receive(event.data);
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error);
        this.#options.onStatus?.('error', detail);
        socket.close(1002, 'invalid collaboration message');
      }
    });
    socket.addEventListener('close', () => {
      if (socket !== this.#socket) return;
      this.#socket = undefined;
      this.#ready = false;
      this.clearRemoteAwareness();
      if (this.#destroyed || this.#options.reconnect === false) {
        this.#options.onStatus?.('disconnected');
        return;
      }
      const delay = Math.min(10_000, 250 * 2 ** this.#reconnectAttempt);
      this.#reconnectAttempt += 1;
      this.#reconnectTimer = setTimeout(() => {
        this.#reconnectTimer = undefined;
        this.connect('reconnecting');
      }, delay);
    });
    socket.addEventListener('error', () => {
      if (socket === this.#socket) {
        this.#options.onStatus?.('error', 'WebSocket transport error');
      }
    });
  }

  private publish(message: OfficeCollaborationTransportMessage): void {
    if (!this.#ready) return;
    this.send('collaboration.document', {
      protocol: message.protocol,
      version: message.version,
      artifactId: message.artifactId,
      artifactKind: message.artifactKind,
      namespace: message.namespace,
      senderClientId: message.senderClientId,
      type: message.type,
      payloadBase64: bytesToBase64(message.payload),
      ...(message.origin ? { origin: message.origin } : {}),
    });
  }

  private subscribe(listener: (message: unknown) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  private receive(value: unknown): void {
    if (typeof value !== 'string') {
      throw new TypeError('The collaboration server must send JSON text.');
    }
    const message: unknown = JSON.parse(value);
    if (!isRecord(message) || typeof message.event !== 'string') {
      throw new TypeError('Invalid A3S Boot WebSocket envelope.');
    }
    const data = message.data;
    switch (message.event) {
      case 'collaboration.ready':
        this.receiveReady(data);
        return;
      case 'collaboration.document':
        this.receiveDocument(data);
        return;
      case 'collaboration.awareness':
        this.receiveAwareness(data);
        return;
      case 'collaboration.peer-left':
        this.receivePeerLeft(data);
        return;
      case 'collaboration.error': {
        const detail =
          isRecord(data) && typeof data.message === 'string'
            ? data.message
            : 'Collaboration server rejected a message.';
        this.#options.onStatus?.('error', detail);
        return;
      }
      default:
        return;
    }
  }

  private receiveReady(value: unknown): void {
    const actor = this.#session.actor;
    if (
      !isRecord(value) ||
      value.protocol !== PROTOCOL ||
      value.version !== VERSION ||
      value.artifactId !== this.#session.artifactId ||
      value.artifactKind !== this.#session.kind ||
      value.namespace !== this.#session.namespace ||
      !actor ||
      value.actorId !== actor.id ||
      value.actorName !== actor.name ||
      value.actorKind !== actor.kind ||
      value.mode !== this.#session.mode ||
      typeof value.senderClientId !== 'number'
    ) {
      throw new TypeError(
        'The collaboration.ready identity does not match the local Office session.',
      );
    }
    this.#ready = true;
    this.#reconnectAttempt = 0;
    this.#options.onStatus?.('connected');
    // The server also sends its SyncStep1 to edit/comment clients. Together
    // these handshakes repair both download and upload gaps after an offline
    // period.
    this.#binding.synchronize();
    this.sendLocalAwareness();
  }

  private receiveDocument(value: unknown): void {
    if (
      !isRecord(value) ||
      value.protocol !== PROTOCOL ||
      value.version !== VERSION ||
      typeof value.artifactId !== 'string' ||
      typeof value.artifactKind !== 'string' ||
      typeof value.namespace !== 'string' ||
      typeof value.senderClientId !== 'number' ||
      typeof value.type !== 'string' ||
      typeof value.payloadBase64 !== 'string'
    ) {
      throw new TypeError('Invalid collaboration.document message.');
    }
    const decoded = {
      protocol: PROTOCOL,
      version: VERSION,
      artifactId: value.artifactId,
      artifactKind: value.artifactKind,
      namespace: value.namespace,
      senderClientId: value.senderClientId,
      type: value.type,
      payload: base64ToBytes(value.payloadBase64),
      ...(isRecord(value.origin) ? { origin: value.origin } : {}),
    };
    for (const listener of this.#listeners) listener(decoded);
  }

  private receiveAwareness(value: unknown): void {
    if (!this.#awareness) return;
    if (
      !isRecord(value) ||
      value.protocol !== PROTOCOL ||
      value.version !== VERSION ||
      typeof value.senderClientId !== 'number' ||
      typeof value.payloadBase64 !== 'string'
    ) {
      throw new TypeError('Invalid collaboration.awareness message.');
    }
    this.#remoteClientIds.add(value.senderClientId);
    applyAwarenessUpdate(
      this.#awareness,
      base64ToBytes(value.payloadBase64),
      REMOTE_AWARENESS_ORIGIN,
    );
  }

  private receivePeerLeft(value: unknown): void {
    if (
      !this.#awareness ||
      !isRecord(value) ||
      typeof value.senderClientId !== 'number'
    )
      return;
    this.#remoteClientIds.delete(value.senderClientId);
    removeAwarenessStates(
      this.#awareness,
      [value.senderClientId],
      REMOTE_AWARENESS_ORIGIN,
    );
  }

  readonly #onAwarenessUpdate = (
    changes: { added: number[]; updated: number[]; removed: number[] },
    origin: unknown,
  ): void => {
    if (!this.#awareness || origin === REMOTE_AWARENESS_ORIGIN) return;
    const clientIds = [
      ...changes.added,
      ...changes.updated,
      ...changes.removed,
    ];
    if (!clientIds.includes(this.#awareness.clientID)) return;
    this.sendAwarenessUpdate([this.#awareness.clientID]);
  };

  private sendLocalAwareness(): void {
    if (!this.#awareness?.getLocalState()) return;
    this.sendAwarenessUpdate([this.#awareness.clientID]);
  }

  private sendAwarenessUpdate(clientIds: number[]): void {
    if (!this.#awareness || !this.#ready) return;
    this.#session.metadata();
    this.send('collaboration.awareness', {
      protocol: PROTOCOL,
      version: VERSION,
      artifactId: this.#session.artifactId,
      artifactKind: this.#session.kind,
      namespace: this.#session.namespace,
      senderClientId: this.#awareness.clientID,
      payloadBase64: bytesToBase64(
        encodeAwarenessUpdate(this.#awareness, clientIds),
      ),
    });
  }

  private clearRemoteAwareness(): void {
    if (!this.#awareness || this.#remoteClientIds.size === 0) return;
    removeAwarenessStates(
      this.#awareness,
      [...this.#remoteClientIds],
      REMOTE_AWARENESS_ORIGIN,
    );
    this.#remoteClientIds.clear();
  }

  private send(event: string, data: unknown): void {
    if (this.#socket?.readyState !== WebSocket.OPEN) return;
    this.#socket.send(JSON.stringify({ event, data }));
  }
}

function bytesToBase64(value: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < value.length; offset += 0x8000) {
    binary += String.fromCharCode(...value.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
