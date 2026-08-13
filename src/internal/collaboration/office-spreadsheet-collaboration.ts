import * as Y from 'yjs';
import type { WorkSpreadsheetContent } from '../features/work/work-types';
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
  assertWorkOfficeSpreadsheetRootsEmpty,
  initializeWorkOfficeSpreadsheetRoots,
  patchWorkOfficeSpreadsheetRoots,
  readWorkOfficeSpreadsheetRoots,
  validateWorkOfficeSpreadsheetContent,
  workOfficeSpreadsheetRoots,
  workOfficeSpreadsheetUndoScope,
} from './office-spreadsheet-collaboration-model';

export interface WorkOfficeSpreadsheetCollaborationChange {
  content: WorkSpreadsheetContent;
  local: boolean;
  origin: unknown;
}

export interface WorkOfficeSpreadsheetCollaborationBindingOptions {
  captureTimeoutMs?: number;
  origin?: WorkOfficeCollaborationOrigin;
}

export interface WorkOfficeSpreadsheetCollaborationBinding {
  readonly origin: WorkOfficeCollaborationOrigin;
  content(): WorkSpreadsheetContent;
  replace(
    previous: WorkSpreadsheetContent,
    next: WorkSpreadsheetContent,
  ): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): boolean;
  redo(): boolean;
  stopCapturing(): void;
  subscribe(
    listener: (change: WorkOfficeSpreadsheetCollaborationChange) => void,
  ): () => void;
  subscribeError(listener: (error: unknown) => void): () => void;
  subscribeHistory(listener: () => void): () => void;
  destroy(): void;
}

export function initializeWorkOfficeSpreadsheetCollaboration(
  session: WorkOfficeCollaborationSession,
  content: WorkSpreadsheetContent,
): { initialized: boolean; content: WorkSpreadsheetContent } {
  assertSpreadsheetSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const prepared = validateWorkOfficeSpreadsheetContent(content);
  const existing = readWorkOfficeCollaborationMetadata(session);
  if (existing?.initialized) {
    return {
      initialized: false,
      content: readWorkOfficeSpreadsheetCollaboration(session),
    };
  }
  const roots = spreadsheetRoots(session);
  assertWorkOfficeSpreadsheetRootsEmpty(roots);
  const origin = session.createOrigin('bootstrap');
  session.transact(() => {
    const metadata = initializeWorkOfficeCollaborationMetadata(session);
    if (metadata.initialized) return;
    registerWorkOfficeCollaborationInitializer(session);
    initializeWorkOfficeSpreadsheetRoots(roots, prepared);
    markWorkOfficeCollaborationInitialized(session);
  }, origin);
  return {
    initialized: true,
    content: readWorkOfficeSpreadsheetCollaboration(session),
  };
}

export function readWorkOfficeSpreadsheetCollaboration(
  session: WorkOfficeCollaborationSession,
): WorkSpreadsheetContent {
  assertInitializedSpreadsheetSession(session);
  return readWorkOfficeSpreadsheetRoots(spreadsheetRoots(session));
}

export function replaceWorkOfficeSpreadsheetCollaboration(
  session: WorkOfficeCollaborationSession,
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
  origin: WorkOfficeCollaborationOrigin = session.localOrigin,
): boolean {
  assertInitializedSpreadsheetSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const before = validateWorkOfficeSpreadsheetContent(previous);
  const after = validateWorkOfficeSpreadsheetContent(next);
  if (jsonEqual(before, after)) return false;
  session.transact(() => {
    patchWorkOfficeSpreadsheetRoots(spreadsheetRoots(session), before, after);
  }, origin);
  return true;
}

export function createWorkOfficeSpreadsheetCollaborationBinding(
  session: WorkOfficeCollaborationSession,
  options: WorkOfficeSpreadsheetCollaborationBindingOptions = {},
): WorkOfficeSpreadsheetCollaborationBinding {
  assertInitializedSpreadsheetSession(session);
  return new WorkOfficeSpreadsheetCollaborationBindingImpl(session, options);
}

class WorkOfficeSpreadsheetCollaborationBindingImpl
  implements WorkOfficeSpreadsheetCollaborationBinding
{
  readonly origin: WorkOfficeCollaborationOrigin;
  readonly #session: WorkOfficeCollaborationSession;
  readonly #undoManager: Y.UndoManager;
  readonly #roots: ReturnType<typeof spreadsheetRoots>;
  readonly #undoScope: ReturnType<typeof workOfficeSpreadsheetUndoScope>;
  readonly #listeners = new Set<
    (change: WorkOfficeSpreadsheetCollaborationChange) => void
  >();
  readonly #errorListeners = new Set<(error: unknown) => void>();
  readonly #historyListeners = new Set<() => void>();
  #destroyed = false;
  #pendingChange = false;
  #pendingLocal = true;
  #pendingOrigin: unknown;

  constructor(
    session: WorkOfficeCollaborationSession,
    options: WorkOfficeSpreadsheetCollaborationBindingOptions,
  ) {
    this.#session = session;
    this.#roots = spreadsheetRoots(session);
    this.origin =
      options.origin ?? session.createOrigin(session.localOrigin.kind);
    assertWorkOfficeCollaborationOrigin(this.origin);
    this.#undoScope = workOfficeSpreadsheetUndoScope(this.#roots);
    this.#undoManager = new Y.UndoManager(this.#undoScope, {
      captureTimeout: options.captureTimeoutMs ?? 500,
      trackedOrigins: new Set([this.origin]),
    });
    session.document.on('afterTransaction', this.#onTransaction);
    this.#undoManager.on('stack-item-added', this.#onHistoryChange);
    this.#undoManager.on('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.on('stack-item-updated', this.#onHistoryChange);
  }

  content(): WorkSpreadsheetContent {
    this.ensureActive();
    return readWorkOfficeSpreadsheetRoots(this.#roots);
  }

  replace(
    previous: WorkSpreadsheetContent,
    next: WorkSpreadsheetContent,
  ): boolean {
    this.ensureActive();
    return replaceWorkOfficeSpreadsheetCollaboration(
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
    listener: (change: WorkOfficeSpreadsheetCollaborationChange) => void,
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
      } satisfies WorkOfficeSpreadsheetCollaborationChange;
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
      'The Spreadsheet collaboration binding has been destroyed.',
    );
  }
}

function spreadsheetRoots(session: WorkOfficeCollaborationSession) {
  return workOfficeSpreadsheetRoots(
    session.document,
    session.rootName.bind(session),
  );
}

function assertSpreadsheetSession(
  session: WorkOfficeCollaborationSession,
): void {
  if (session.kind === 'spreadsheet') return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.kind_mismatch',
    `A Spreadsheet collaboration binding cannot use a '${session.kind}' session.`,
  );
}

function assertInitializedSpreadsheetSession(
  session: WorkOfficeCollaborationSession,
): void {
  assertSpreadsheetSession(session);
  const metadata = readWorkOfficeCollaborationMetadata(session);
  if (metadata?.initialized) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.not_initialized',
    'The Spreadsheet collaboration session has not been initialized.',
  );
}
