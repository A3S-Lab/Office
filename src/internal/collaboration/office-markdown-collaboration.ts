import * as Y from 'yjs';
import type { WorkMarkdownContent } from '../features/work/work-types';
import {
  assertWorkOfficeCollaborationEditable,
  initializeWorkOfficeCollaborationMetadata,
  markWorkOfficeCollaborationInitialized,
  readWorkOfficeCollaborationMetadata,
  registerWorkOfficeCollaborationInitializer,
  type WorkOfficeCollaborationOrigin,
  type WorkOfficeCollaborationSession,
  WorkOfficeCollaborationError,
} from './office-collaboration';

const MARKDOWN_SOURCE_ROOT = 'markdown.source';

export interface WorkOfficeMarkdownCollaborationChange {
  content: WorkMarkdownContent;
  local: boolean;
  origin: unknown;
}

export interface WorkOfficeMarkdownCollaborationBindingOptions {
  origin?: WorkOfficeCollaborationOrigin;
  captureTimeoutMs?: number;
}

export interface WorkOfficeMarkdownCollaborationBinding {
  readonly origin: WorkOfficeCollaborationOrigin;
  content(): WorkMarkdownContent;
  replace(markdown: string): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): boolean;
  redo(): boolean;
  stopCapturing(): void;
  subscribe(
    listener: (change: WorkOfficeMarkdownCollaborationChange) => void,
  ): () => void;
  subscribeHistory(listener: () => void): () => void;
  destroy(): void;
}

export function initializeWorkOfficeMarkdownCollaboration(
  session: WorkOfficeCollaborationSession,
  content: WorkMarkdownContent,
): { initialized: boolean; content: WorkMarkdownContent } {
  assertMarkdownSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const initialMarkdown = validatedMarkdown(content);
  const existing = readWorkOfficeCollaborationMetadata(session);
  if (existing?.initialized) {
    return {
      initialized: false,
      content: readWorkOfficeMarkdownCollaboration(session),
    };
  }
  const source = markdownSource(session);
  if (source.length > 0) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The Markdown collaboration source contains data without initialized metadata.',
    );
  }
  const origin = session.createOrigin('bootstrap');
  session.transact(() => {
    const metadata = initializeWorkOfficeCollaborationMetadata(session);
    if (metadata.initialized) return;
    registerWorkOfficeCollaborationInitializer(session);
    if (initialMarkdown) source.insert(0, initialMarkdown);
    markWorkOfficeCollaborationInitialized(session);
  }, origin);
  return {
    initialized: true,
    content: readWorkOfficeMarkdownCollaboration(session),
  };
}

export function readWorkOfficeMarkdownCollaboration(
  session: WorkOfficeCollaborationSession,
): WorkMarkdownContent {
  assertInitializedMarkdownSession(session);
  return { type: 'markdown', markdown: markdownSource(session).toString() };
}

export function replaceWorkOfficeMarkdownCollaboration(
  session: WorkOfficeCollaborationSession,
  markdown: string,
  origin: WorkOfficeCollaborationOrigin = session.localOrigin,
): boolean {
  assertInitializedMarkdownSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const source = markdownSource(session);
  const current = source.toString();
  if (current === markdown) return false;
  const replacement = markdownReplacement(current, markdown);
  session.transact(() => {
    if (replacement.deleteLength > 0) {
      source.delete(replacement.index, replacement.deleteLength);
    }
    if (replacement.insert)
      source.insert(replacement.index, replacement.insert);
  }, origin);
  return true;
}

export function createWorkOfficeMarkdownCollaborationBinding(
  session: WorkOfficeCollaborationSession,
  options: WorkOfficeMarkdownCollaborationBindingOptions = {},
): WorkOfficeMarkdownCollaborationBinding {
  assertInitializedMarkdownSession(session);
  return new WorkOfficeMarkdownCollaborationBindingImpl(session, options);
}

class WorkOfficeMarkdownCollaborationBindingImpl
  implements WorkOfficeMarkdownCollaborationBinding
{
  readonly origin: WorkOfficeCollaborationOrigin;
  readonly #session: WorkOfficeCollaborationSession;
  readonly #source: Y.Text;
  readonly #undoManager: Y.UndoManager;
  readonly #listeners = new Set<
    (change: WorkOfficeMarkdownCollaborationChange) => void
  >();
  readonly #historyListeners = new Set<() => void>();
  #destroyed = false;

  constructor(
    session: WorkOfficeCollaborationSession,
    options: WorkOfficeMarkdownCollaborationBindingOptions,
  ) {
    this.#session = session;
    this.#source = markdownSource(session);
    this.origin =
      options.origin ?? session.createOrigin(session.localOrigin.kind);
    this.#undoManager = new Y.UndoManager(this.#source, {
      captureTimeout: options.captureTimeoutMs ?? 1_000,
      trackedOrigins: new Set([this.origin]),
    });
    this.#source.observe(this.#onSourceChange);
    this.#undoManager.on('stack-item-added', this.#onHistoryChange);
    this.#undoManager.on('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.on('stack-item-updated', this.#onHistoryChange);
  }

  content(): WorkMarkdownContent {
    this.ensureActive();
    return { type: 'markdown', markdown: this.#source.toString() };
  }

  replace(markdown: string): boolean {
    this.ensureActive();
    return replaceWorkOfficeMarkdownCollaboration(
      this.#session,
      markdown,
      this.origin,
    );
  }

  canUndo(): boolean {
    this.ensureActive();
    return this.#undoManager.undoStack.length > 0;
  }

  canRedo(): boolean {
    this.ensureActive();
    return this.#undoManager.redoStack.length > 0;
  }

  undo(): boolean {
    this.ensureActive();
    assertWorkOfficeCollaborationEditable(this.#session);
    if (!this.canUndo()) return false;
    this.#undoManager.undo();
    return true;
  }

  redo(): boolean {
    this.ensureActive();
    assertWorkOfficeCollaborationEditable(this.#session);
    if (!this.canRedo()) return false;
    this.#undoManager.redo();
    return true;
  }

  stopCapturing(): void {
    this.ensureActive();
    this.#undoManager.stopCapturing();
  }

  subscribe(
    listener: (change: WorkOfficeMarkdownCollaborationChange) => void,
  ): () => void {
    this.ensureActive();
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  subscribeHistory(listener: () => void): () => void {
    this.ensureActive();
    this.#historyListeners.add(listener);
    return () => this.#historyListeners.delete(listener);
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#source.unobserve(this.#onSourceChange);
    this.#undoManager.off('stack-item-added', this.#onHistoryChange);
    this.#undoManager.off('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.off('stack-item-updated', this.#onHistoryChange);
    this.#undoManager.destroy();
    this.#listeners.clear();
    this.#historyListeners.clear();
  }

  readonly #onSourceChange = (event: Y.YTextEvent): void => {
    const change = {
      content: { type: 'markdown', markdown: this.#source.toString() },
      local: event.transaction.origin === this.origin,
      origin: event.transaction.origin,
    } satisfies WorkOfficeMarkdownCollaborationChange;
    for (const listener of this.#listeners) listener(change);
  };

  readonly #onHistoryChange = (): void => {
    for (const listener of this.#historyListeners) listener();
  };

  private ensureActive(): void {
    if (!this.#destroyed) return;
    throw new WorkOfficeCollaborationError(
      'office.collaboration.binding_destroyed',
      'The Markdown collaboration binding has been destroyed.',
    );
  }
}

function markdownSource(session: WorkOfficeCollaborationSession): Y.Text {
  return session.document.getText(session.rootName(MARKDOWN_SOURCE_ROOT));
}

function assertMarkdownSession(session: WorkOfficeCollaborationSession): void {
  if (session.kind === 'markdown') return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.kind_mismatch',
    `A Markdown collaboration binding cannot use a '${session.kind}' session.`,
  );
}

function validatedMarkdown(content: WorkMarkdownContent): string {
  if (content?.type === 'markdown' && typeof content.markdown === 'string') {
    return content.markdown;
  }
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    'Markdown collaboration initialization requires a valid Markdown content value.',
  );
}

function assertInitializedMarkdownSession(
  session: WorkOfficeCollaborationSession,
): void {
  assertMarkdownSession(session);
  const metadata = readWorkOfficeCollaborationMetadata(session);
  if (metadata?.initialized) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.not_initialized',
    'The Markdown collaboration session has not been initialized.',
  );
}

function markdownReplacement(
  current: string,
  next: string,
): { index: number; deleteLength: number; insert: string } {
  let prefix = 0;
  const maximumPrefix = Math.min(current.length, next.length);
  while (prefix < maximumPrefix && current[prefix] === next[prefix])
    prefix += 1;
  prefix = safeUtf16Boundary(current, prefix);
  prefix = safeUtf16Boundary(next, prefix);

  let suffix = 0;
  const maximumSuffix = Math.min(current.length - prefix, next.length - prefix);
  while (
    suffix < maximumSuffix &&
    current[current.length - suffix - 1] === next[next.length - suffix - 1]
  ) {
    suffix += 1;
  }
  const currentEnd = safeUtf16Boundary(current, current.length - suffix);
  const nextEnd = safeUtf16Boundary(next, next.length - suffix);
  return {
    index: prefix,
    deleteLength: Math.max(0, currentEnd - prefix),
    insert: next.slice(prefix, nextEnd),
  };
}

function safeUtf16Boundary(value: string, index: number): number {
  if (
    index > 0 &&
    index < value.length &&
    isHighSurrogate(value.charCodeAt(index - 1)) &&
    isLowSurrogate(value.charCodeAt(index))
  ) {
    return index - 1;
  }
  return index;
}

function isHighSurrogate(value: number): boolean {
  return value >= 0xd800 && value <= 0xdbff;
}

function isLowSurrogate(value: number): boolean {
  return value >= 0xdc00 && value <= 0xdfff;
}
