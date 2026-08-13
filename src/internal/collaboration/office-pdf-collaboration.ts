import * as Y from 'yjs';
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
  assertWorkOfficePdfRootsEmpty,
  initializeWorkOfficePdfRoots,
  patchWorkOfficePdfRoots,
  readWorkOfficePdfRoots,
  validateWorkOfficePdfCollaborationContent,
  workOfficePdfIrreversibleScope,
  workOfficePdfRoots,
  workOfficePdfUndoScope,
} from './office-pdf-collaboration-model';
import type {
  WorkPdfCollaborationContent,
  WorkPdfCollaborationSource,
} from './office-pdf-collaboration-types';

export type * from './office-pdf-collaboration-types';

export interface WorkOfficePdfCollaborationChange {
  content: WorkPdfCollaborationContent;
  local: boolean;
  origin: unknown;
}

export interface WorkOfficePdfCollaborationBindingOptions {
  captureTimeoutMs?: number;
  origin?: WorkOfficeCollaborationOrigin;
}

export interface WorkOfficePdfCollaborationBinding {
  readonly origin: WorkOfficeCollaborationOrigin;
  content(): WorkPdfCollaborationContent;
  replace(
    previous: WorkPdfCollaborationContent,
    next: WorkPdfCollaborationContent,
  ): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): boolean;
  redo(): boolean;
  stopCapturing(): void;
  subscribe(
    listener: (change: WorkOfficePdfCollaborationChange) => void,
  ): () => void;
  subscribeError(listener: (error: unknown) => void): () => void;
  subscribeHistory(listener: () => void): () => void;
  destroy(): void;
}

export function initializeWorkOfficePdfCollaboration(
  session: WorkOfficeCollaborationSession,
  content: WorkPdfCollaborationContent,
): { initialized: boolean; content: WorkPdfCollaborationContent } {
  assertPdfSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const prepared = validateWorkOfficePdfCollaborationContent(content);
  const existing = readWorkOfficeCollaborationMetadata(session);
  if (existing?.initialized) {
    const current = readWorkOfficePdfCollaboration(session);
    if (!jsonEqual(current.source, prepared.source)) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.artifact_mismatch',
        `The PDF source does not match collaboration artifact '${session.artifactId}'.`,
      );
    }
    return { initialized: false, content: current };
  }
  const roots = pdfRoots(session);
  assertWorkOfficePdfRootsEmpty(roots);
  const origin = session.createOrigin('bootstrap');
  session.transact(() => {
    const metadata = initializeWorkOfficeCollaborationMetadata(session);
    if (metadata.initialized) return;
    registerWorkOfficeCollaborationInitializer(session);
    initializeWorkOfficePdfRoots(roots, prepared);
    markWorkOfficeCollaborationInitialized(session);
  }, origin);
  return {
    initialized: true,
    content: readWorkOfficePdfCollaboration(session),
  };
}

export function readWorkOfficePdfCollaboration(
  session: WorkOfficeCollaborationSession,
): WorkPdfCollaborationContent {
  assertInitializedPdfSession(session);
  return readWorkOfficePdfRoots(pdfRoots(session));
}

export function assertWorkOfficePdfCollaborationSource(
  session: WorkOfficeCollaborationSession,
  source: WorkPdfCollaborationSource,
): void {
  const expected = readWorkOfficePdfCollaboration(session).source;
  if (jsonEqual(expected, source)) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.artifact_mismatch',
    `The PDF source does not match collaboration artifact '${session.artifactId}'.`,
  );
}

export function replaceWorkOfficePdfCollaboration(
  session: WorkOfficeCollaborationSession,
  previous: WorkPdfCollaborationContent,
  next: WorkPdfCollaborationContent,
  origin: WorkOfficeCollaborationOrigin = session.localOrigin,
): boolean {
  assertInitializedPdfSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const before = validateWorkOfficePdfCollaborationContent(previous);
  const after = validateWorkOfficePdfCollaborationContent(next);
  if (jsonEqual(before, after)) return false;
  session.transact(() => {
    patchWorkOfficePdfRoots(pdfRoots(session), before, after);
  }, origin);
  return true;
}

export function createWorkOfficePdfCollaborationBinding(
  session: WorkOfficeCollaborationSession,
  options: WorkOfficePdfCollaborationBindingOptions = {},
): WorkOfficePdfCollaborationBinding {
  assertInitializedPdfSession(session);
  return new WorkOfficePdfCollaborationBindingImpl(session, options);
}

class WorkOfficePdfCollaborationBindingImpl
  implements WorkOfficePdfCollaborationBinding
{
  readonly origin: WorkOfficeCollaborationOrigin;
  readonly #session: WorkOfficeCollaborationSession;
  readonly #undoManager: Y.UndoManager;
  readonly #roots: ReturnType<typeof pdfRoots>;
  readonly #undoScope: ReturnType<typeof workOfficePdfUndoScope>;
  readonly #irreversibleScope: ReturnType<
    typeof workOfficePdfIrreversibleScope
  >;
  readonly #changeScope: Array<Y.Map<unknown> | Y.Array<string>>;
  readonly #listeners = new Set<
    (change: WorkOfficePdfCollaborationChange) => void
  >();
  readonly #errorListeners = new Set<(error: unknown) => void>();
  readonly #historyListeners = new Set<() => void>();
  #destroyed = false;
  #pendingChange = false;
  #pendingLocal = true;
  #pendingOrigin: unknown;

  constructor(
    session: WorkOfficeCollaborationSession,
    options: WorkOfficePdfCollaborationBindingOptions,
  ) {
    this.#session = session;
    this.#roots = pdfRoots(session);
    this.origin =
      options.origin ?? session.createOrigin(session.localOrigin.kind);
    assertWorkOfficeCollaborationOrigin(this.origin);
    this.#undoScope = workOfficePdfUndoScope(this.#roots);
    this.#irreversibleScope = workOfficePdfIrreversibleScope(this.#roots);
    this.#changeScope = [...this.#undoScope, ...this.#irreversibleScope];
    this.#undoManager = new Y.UndoManager(this.#undoScope, {
      captureTimeout: options.captureTimeoutMs ?? 500,
      trackedOrigins: new Set([this.origin]),
    });
    session.document.on('afterTransaction', this.#onTransaction);
    this.#undoManager.on('stack-item-added', this.#onHistoryChange);
    this.#undoManager.on('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.on('stack-item-updated', this.#onHistoryChange);
  }

  content(): WorkPdfCollaborationContent {
    this.ensureActive();
    return readWorkOfficePdfRoots(this.#roots);
  }

  replace(
    previous: WorkPdfCollaborationContent,
    next: WorkPdfCollaborationContent,
  ): boolean {
    this.ensureActive();
    return replaceWorkOfficePdfCollaboration(
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
    listener: (change: WorkOfficePdfCollaborationChange) => void,
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
    if (this.#irreversibleScope.some((root) => changedParents.has(root))) {
      this.#undoManager.clear();
    }
    if (
      this.#destroyed ||
      !this.#changeScope.some((root) => changedParents.has(root))
    )
      return;
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
      } satisfies WorkOfficePdfCollaborationChange;
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
      'The PDF collaboration binding has been destroyed.',
    );
  }
}

function pdfRoots(session: WorkOfficeCollaborationSession) {
  return workOfficePdfRoots(session.document, session.rootName.bind(session));
}

function assertPdfSession(session: WorkOfficeCollaborationSession): void {
  if (session.kind === 'pdf') return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.kind_mismatch',
    `A PDF collaboration binding cannot use a '${session.kind}' session.`,
  );
}

function assertInitializedPdfSession(
  session: WorkOfficeCollaborationSession,
): void {
  assertPdfSession(session);
  const metadata = readWorkOfficeCollaborationMetadata(session);
  if (metadata?.initialized) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.not_initialized',
    'The PDF collaboration session has not been initialized.',
  );
}
