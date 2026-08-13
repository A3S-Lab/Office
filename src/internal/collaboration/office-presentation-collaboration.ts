import * as Y from 'yjs';
import type { WorkPresentationContent } from '../features/work/work-types';
import {
  assertWorkOfficeCollaborationEditable,
  assertWorkOfficeCollaborationOrigin,
  initializeWorkOfficeCollaborationMetadata,
  markWorkOfficeCollaborationInitialized,
  readWorkOfficeCollaborationMetadata,
  registerWorkOfficeCollaborationInitializer,
  type WorkOfficeCollaborationOrigin,
  type WorkOfficeCollaborationSession,
  WorkOfficeCollaborationError,
} from './office-collaboration';
import { workOfficeCollaborationJsonEqual as jsonEqual } from './office-collaboration-json';
import {
  assertWorkOfficePresentationRootsEmpty,
  initializeWorkOfficePresentationRoots,
  patchWorkOfficePresentationRoots,
  readWorkOfficePresentationRoots,
  validateWorkOfficePresentationContent,
  workOfficePresentationRoots,
  workOfficePresentationUndoScope,
} from './office-presentation-collaboration-model';

export interface WorkOfficePresentationCollaborationChange {
  content: WorkPresentationContent;
  local: boolean;
  origin: unknown;
}

export interface WorkOfficePresentationCollaborationBindingOptions {
  captureTimeoutMs?: number;
  origin?: WorkOfficeCollaborationOrigin;
}

export interface WorkOfficePresentationCollaborationBinding {
  readonly origin: WorkOfficeCollaborationOrigin;
  content(): WorkPresentationContent;
  replace(
    previous: WorkPresentationContent,
    next: WorkPresentationContent,
  ): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): boolean;
  redo(): boolean;
  stopCapturing(): void;
  subscribe(
    listener: (change: WorkOfficePresentationCollaborationChange) => void,
  ): () => void;
  subscribeError(listener: (error: unknown) => void): () => void;
  subscribeHistory(listener: () => void): () => void;
  destroy(): void;
}

export function initializeWorkOfficePresentationCollaboration(
  session: WorkOfficeCollaborationSession,
  content: WorkPresentationContent,
): { initialized: boolean; content: WorkPresentationContent } {
  assertPresentationSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const prepared = validateWorkOfficePresentationContent(content);
  const existing = readWorkOfficeCollaborationMetadata(session);
  if (existing?.initialized) {
    return {
      initialized: false,
      content: readWorkOfficePresentationCollaboration(session),
    };
  }
  const roots = presentationRoots(session);
  assertWorkOfficePresentationRootsEmpty(roots);
  const origin = session.createOrigin('bootstrap');
  session.transact(() => {
    const metadata = initializeWorkOfficeCollaborationMetadata(session);
    if (metadata.initialized) return;
    registerWorkOfficeCollaborationInitializer(session);
    initializeWorkOfficePresentationRoots(roots, prepared);
    markWorkOfficeCollaborationInitialized(session);
  }, origin);
  return {
    initialized: true,
    content: readWorkOfficePresentationCollaboration(session),
  };
}

export function readWorkOfficePresentationCollaboration(
  session: WorkOfficeCollaborationSession,
): WorkPresentationContent {
  assertInitializedPresentationSession(session);
  return readWorkOfficePresentationRoots(presentationRoots(session));
}

export function replaceWorkOfficePresentationCollaboration(
  session: WorkOfficeCollaborationSession,
  previous: WorkPresentationContent,
  next: WorkPresentationContent,
  origin: WorkOfficeCollaborationOrigin = session.localOrigin,
): boolean {
  assertInitializedPresentationSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const before = validateWorkOfficePresentationContent(previous);
  const after = validateWorkOfficePresentationContent(next);
  if (jsonEqual(before, after)) return false;
  session.transact(() => {
    patchWorkOfficePresentationRoots(presentationRoots(session), before, after);
  }, origin);
  return true;
}

export function createWorkOfficePresentationCollaborationBinding(
  session: WorkOfficeCollaborationSession,
  options: WorkOfficePresentationCollaborationBindingOptions = {},
): WorkOfficePresentationCollaborationBinding {
  assertInitializedPresentationSession(session);
  return new WorkOfficePresentationCollaborationBindingImpl(session, options);
}

class WorkOfficePresentationCollaborationBindingImpl
  implements WorkOfficePresentationCollaborationBinding
{
  readonly origin: WorkOfficeCollaborationOrigin;
  readonly #session: WorkOfficeCollaborationSession;
  readonly #undoManager: Y.UndoManager;
  readonly #roots: ReturnType<typeof presentationRoots>;
  readonly #undoScope: ReturnType<typeof workOfficePresentationUndoScope>;
  readonly #listeners = new Set<
    (change: WorkOfficePresentationCollaborationChange) => void
  >();
  readonly #errorListeners = new Set<(error: unknown) => void>();
  readonly #historyListeners = new Set<() => void>();
  #destroyed = false;
  #pendingChange = false;
  #pendingLocal = true;
  #pendingOrigin: unknown;

  constructor(
    session: WorkOfficeCollaborationSession,
    options: WorkOfficePresentationCollaborationBindingOptions,
  ) {
    this.#session = session;
    this.#roots = presentationRoots(session);
    this.origin =
      options.origin ?? session.createOrigin(session.localOrigin.kind);
    assertWorkOfficeCollaborationOrigin(this.origin);
    this.#undoScope = workOfficePresentationUndoScope(this.#roots);
    this.#undoManager = new Y.UndoManager(this.#undoScope, {
      captureTimeout: options.captureTimeoutMs ?? 500,
      trackedOrigins: new Set([this.origin]),
    });
    session.document.on('afterTransaction', this.#onTransaction);
    this.#undoManager.on('stack-item-added', this.#onHistoryChange);
    this.#undoManager.on('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.on('stack-item-updated', this.#onHistoryChange);
  }

  content(): WorkPresentationContent {
    this.ensureActive();
    return readWorkOfficePresentationRoots(this.#roots);
  }

  replace(
    previous: WorkPresentationContent,
    next: WorkPresentationContent,
  ): boolean {
    this.ensureActive();
    return replaceWorkOfficePresentationCollaboration(
      this.#session,
      previous,
      next,
      this.origin,
    );
  }

  canUndo(): boolean {
    this.ensureActive();
    return this.#undoManager.canUndo();
  }

  canRedo(): boolean {
    this.ensureActive();
    return this.#undoManager.canRedo();
  }

  undo(): boolean {
    this.ensureActive();
    assertWorkOfficeCollaborationEditable(this.#session);
    return this.#undoManager.undo() !== null;
  }

  redo(): boolean {
    this.ensureActive();
    assertWorkOfficeCollaborationEditable(this.#session);
    return this.#undoManager.redo() !== null;
  }

  stopCapturing(): void {
    this.ensureActive();
    this.#undoManager.stopCapturing();
  }

  subscribe(
    listener: (change: WorkOfficePresentationCollaborationChange) => void,
  ): () => void {
    this.ensureActive();
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  subscribeError(listener: (error: unknown) => void): () => void {
    this.ensureActive();
    this.#errorListeners.add(listener);
    return () => this.#errorListeners.delete(listener);
  }

  subscribeHistory(listener: () => void): () => void {
    this.ensureActive();
    this.#historyListeners.add(listener);
    return () => this.#historyListeners.delete(listener);
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#pendingChange = false;
    this.#session.document.off('afterTransaction', this.#onTransaction);
    this.#undoManager.off('stack-item-added', this.#onHistoryChange);
    this.#undoManager.off('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.off('stack-item-updated', this.#onHistoryChange);
    this.#undoManager.destroy();
    this.#listeners.clear();
    this.#errorListeners.clear();
    this.#historyListeners.clear();
  }

  readonly #onTransaction = (transaction: Y.Transaction): void => {
    const changedParents = new Set<unknown>(
      transaction.changedParentTypes.keys(),
    );
    if (
      this.#destroyed ||
      !this.#undoScope.some((root) => changedParents.has(root))
    ) {
      return;
    }
    const local =
      transaction.origin === this.origin ||
      transaction.origin === this.#undoManager;
    this.#pendingLocal = this.#pendingChange
      ? this.#pendingLocal && local
      : local;
    this.#pendingOrigin = transaction.origin;
    if (this.#pendingChange) return;
    this.#pendingChange = true;
    queueMicrotask(this.#flushChange);
  };

  readonly #flushChange = (): void => {
    if (this.#destroyed || !this.#pendingChange) return;
    const local = this.#pendingLocal;
    const origin = this.#pendingOrigin;
    this.#pendingChange = false;
    this.#pendingLocal = true;
    this.#pendingOrigin = undefined;
    try {
      const change = {
        content: this.content(),
        local,
        origin,
      } satisfies WorkOfficePresentationCollaborationChange;
      for (const listener of this.#listeners) listener(change);
    } catch (error) {
      for (const listener of this.#errorListeners) listener(error);
    }
  };

  readonly #onHistoryChange = (): void => {
    for (const listener of this.#historyListeners) listener();
  };

  private ensureActive(): void {
    if (!this.#destroyed) return;
    throw new WorkOfficeCollaborationError(
      'office.collaboration.binding_destroyed',
      'The Presentation collaboration binding has been destroyed.',
    );
  }
}

function presentationRoots(session: WorkOfficeCollaborationSession) {
  return workOfficePresentationRoots(
    session.document,
    session.rootName.bind(session),
  );
}

function assertPresentationSession(
  session: WorkOfficeCollaborationSession,
): void {
  if (session.kind === 'presentation') return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.kind_mismatch',
    `A Presentation collaboration binding cannot use a '${session.kind}' session.`,
  );
}

function assertInitializedPresentationSession(
  session: WorkOfficeCollaborationSession,
): void {
  assertPresentationSession(session);
  const metadata = readWorkOfficeCollaborationMetadata(session);
  if (metadata?.initialized) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.not_initialized',
    'The Presentation collaboration session has not been initialized.',
  );
}
