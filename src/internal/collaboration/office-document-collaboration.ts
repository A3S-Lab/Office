import {
  type Editor,
  Extension,
  type Extensions,
  getSchema,
} from '@tiptap/core';
import { Collaboration, isChangeOrigin } from '@tiptap/extension-collaboration';
import { Plugin } from '@tiptap/pm/state';
import { AddMarkStep, RemoveMarkStep } from '@tiptap/pm/transform';
import {
  defaultDeleteFilter,
  defaultProtectedNodes,
  prosemirrorJSONToYXmlFragment,
  ySyncPluginKey,
  yXmlFragmentToProsemirrorJSON,
} from '@tiptap/y-tiptap';
import * as Y from 'yjs';
import { createWorkDocumentExtensions } from '../features/work/work-document-extensions';
import { collectDocumentChanges } from '../features/work/work-document-changes';
import {
  createWorkDocumentModel,
  documentModelForContent,
} from '../features/work/work-document-model';
import {
  createWorkDocumentModelFromContent,
  serializeWorkDocumentNode,
  workDocumentSchema,
} from '../features/work/work-document-model-codec';
import { syncDocumentContentFromHtml } from '../features/work/work-document-section';
import type {
  WorkDocumentComment,
  WorkDocumentChangeDecisionAction,
  WorkDocumentContent,
  WorkDocumentNode,
} from '../features/work/work-types';
import {
  assertWorkOfficeCollaborationEditable,
  assertWorkOfficeCollaborationOrigin,
  assertWorkOfficeCollaborationWritable,
  initializeWorkOfficeCollaborationMetadata,
  markWorkOfficeCollaborationInitialized,
  readWorkOfficeCollaborationMetadata,
  registerWorkOfficeCollaborationInitializer,
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationOrigin,
  type WorkOfficeCollaborationSession,
} from './office-collaboration';
import {
  assertWorkOfficeDocumentSidecarsEmpty,
  initializeWorkOfficeDocumentSidecars,
  readWorkOfficeDocumentSidecars,
  updateWorkOfficeDocumentSidecars,
  validatedWorkOfficeDocumentSidecars,
  type WorkOfficeDocumentSidecars,
  workOfficeDocumentSidecarsChanged,
  workOfficeDocumentSidecarUndoScope,
  workOfficeDocumentDecisionRootsChanged,
} from './office-document-collaboration-sidecars';
import {
  createWorkOfficeDocumentChangeDecisions,
  workOfficeDocumentSuggestionTransactionAllowed,
} from './office-document-collaboration-suggestions';
import { jsonEqual } from './office-document-collaboration-sidecar-utils';

const DOCUMENT_CONTENT_ROOT = 'document.content';
const MAX_DOCUMENT_COMMENT_HISTORY = 100;

interface WorkOfficeDocumentCommentHistoryEntry {
  before: WorkDocumentComment[] | undefined;
  after: WorkDocumentComment[] | undefined;
  commentMembershipChanged: boolean;
}

export interface WorkOfficeDocumentCollaborationBindingOptions {
  /**
   * Behavior-only extensions are supported. Node and mark extensions must be
   * part of the versioned Office schema before they can enter shared content.
   */
  additionalExtensions?: Extensions;
  captureTimeoutMs?: number;
  onFirstRender?: () => void;
  origin?: WorkOfficeCollaborationOrigin;
  workExtensions?: Omit<
    Parameters<typeof createWorkDocumentExtensions>[0],
    'collaborative'
  >;
}

export interface WorkOfficeDocumentCollaborationChange {
  content: WorkDocumentContent;
  local: boolean;
  origin: unknown;
}

export interface WorkOfficeDocumentChangeDecisionOptions {
  decidedAt?: string;
}

export interface WorkOfficeDocumentCollaborationBinding {
  readonly extensions: Extensions;
  readonly fragment: Y.XmlFragment;
  readonly origin: WorkOfficeCollaborationOrigin;
  content(): WorkDocumentContent;
  updateSidecars(
    previous: WorkDocumentContent,
    next: WorkDocumentContent,
  ): boolean;
  decideChanges(
    editor: Editor,
    changeIds: readonly string[],
    decision: WorkDocumentChangeDecisionAction,
    options?: WorkOfficeDocumentChangeDecisionOptions,
  ): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): boolean;
  redo(): boolean;
  stopCapturing(): void;
  subscribe(
    listener: (change: WorkOfficeDocumentCollaborationChange) => void,
  ): () => void;
  subscribeError(listener: (error: unknown) => void): () => void;
  subscribeHistory(listener: () => void): () => void;
  destroy(): void;
}

const mountedDocumentBindings = new WeakMap<
  Y.Doc,
  WorkOfficeDocumentCollaborationBindingImpl
>();

type RestorableUndoManager = Y.UndoManager & {
  restore?: () => void;
};

export function initializeWorkOfficeDocumentCollaboration(
  session: WorkOfficeCollaborationSession,
  content: WorkDocumentContent,
): { initialized: boolean; content: WorkDocumentContent } {
  assertDocumentSession(session);
  assertWorkOfficeCollaborationEditable(session);
  const prepared = preparedDocumentContent(content);
  const initialRoot = preflightDocumentRoot(prepared);
  validatedWorkOfficeDocumentSidecars(prepared);
  const existing = readWorkOfficeCollaborationMetadata(session);
  if (existing?.initialized) {
    return {
      initialized: false,
      content: readWorkOfficeDocumentCollaboration(session),
    };
  }
  const fragment = workOfficeDocumentCollaborationFragment(session);
  if (fragment.length > 0) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The Document collaboration content contains data without initialized metadata.',
    );
  }
  assertWorkOfficeDocumentSidecarsEmpty(session);
  const origin = session.createOrigin('bootstrap');
  session.transact(() => {
    const metadata = initializeWorkOfficeCollaborationMetadata(session);
    if (metadata.initialized) return;
    registerWorkOfficeCollaborationInitializer(session);
    prosemirrorJSONToYXmlFragment(workDocumentSchema(), initialRoot, fragment);
    initializeWorkOfficeDocumentSidecars(session, prepared);
    markWorkOfficeCollaborationInitialized(session);
  }, origin);
  return {
    initialized: true,
    content: readWorkOfficeDocumentCollaboration(session),
  };
}

export function readWorkOfficeDocumentCollaboration(
  session: WorkOfficeCollaborationSession,
): WorkDocumentContent {
  assertInitializedDocumentSession(session);
  const fragment = workOfficeDocumentCollaborationFragment(session);
  if (fragment.length === 0) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      'The initialized Document collaboration content is empty.',
    );
  }
  let root: WorkDocumentNode;
  let html: string;
  try {
    root = workOfficeDocumentCollaborationNode(session);
    html = serializeWorkDocumentNode(root);
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `The shared Document collaboration content cannot be read by this Office schema: ${errorMessage(error)}`,
    );
  }
  const sidecars = readWorkOfficeDocumentSidecars(session);
  const base = documentContentFromSharedState(html, sidecars);
  const synchronized = syncDocumentContentFromHtml(base, html);
  return {
    ...synchronized,
    ...sidecars,
    model: createWorkDocumentModel(synchronized.html, root),
  };
}

export function createWorkOfficeDocumentCollaborationBinding(
  session: WorkOfficeCollaborationSession,
  options: WorkOfficeDocumentCollaborationBindingOptions = {},
): WorkOfficeDocumentCollaborationBinding {
  assertInitializedDocumentSession(session);
  return new WorkOfficeDocumentCollaborationBindingImpl(session, options);
}

class WorkOfficeDocumentCollaborationBindingImpl
  implements WorkOfficeDocumentCollaborationBinding
{
  readonly extensions: Extensions;
  readonly fragment: Y.XmlFragment;
  readonly origin: WorkOfficeCollaborationOrigin;
  readonly #session: WorkOfficeCollaborationSession;
  readonly #undoManager: Y.UndoManager;
  readonly #listeners = new Set<
    (change: WorkOfficeDocumentCollaborationChange) => void
  >();
  readonly #errorListeners = new Set<(error: unknown) => void>();
  readonly #historyListeners = new Set<() => void>();
  readonly #commentUndoStack: WorkOfficeDocumentCommentHistoryEntry[] = [];
  readonly #commentRedoStack: WorkOfficeDocumentCommentHistoryEntry[] = [];
  #applyingCommentHistory = false;
  #destroyed = false;
  #destroyRequested = false;
  #observersAttached = false;
  readonly #mountedEditors = new Set<Editor>();
  #pendingChange = false;
  #pendingLocal = true;
  #pendingOrigin: unknown;

  constructor(
    session: WorkOfficeCollaborationSession,
    options: WorkOfficeDocumentCollaborationBindingOptions,
  ) {
    this.#session = session;
    this.fragment = workOfficeDocumentCollaborationFragment(session);
    this.origin =
      options.origin ?? session.createOrigin(session.localOrigin.kind);
    assertWorkOfficeCollaborationOrigin(this.origin);
    const builtIns = createWorkDocumentExtensions({
      ...options.workExtensions,
      collaborative: true,
      rotateTrackedTextIdentities: () => session.mode !== 'suggest',
    });
    const additional = options.additionalExtensions ?? [];
    assertBehaviorOnlyExtensions(additional);
    assertUniqueExtensionNames([
      ...builtIns,
      ...additional,
      { name: 'collaboration' } as Extension,
      { name: 'officeDocumentCollaborationLifecycle' } as Extension,
      { name: 'officeDocumentCollaborationPermission' } as Extension,
    ]);
    validateSharedDocumentWithSchema(
      session,
      getSchema([...builtIns, ...additional]),
    );
    const existing = mountedDocumentBindings.get(session.document);
    if (existing) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        'Only one Document collaboration binding may use a local Y.Doc at a time. Give each editor client its own synchronized Y.Doc.',
      );
    }
    const undoScope =
      session.mode === 'comment' || session.mode === 'suggest'
        ? [this.fragment]
        : [this.fragment, ...workOfficeDocumentSidecarUndoScope(session)];
    this.#undoManager = new Y.UndoManager(undoScope, {
      captureTimeout: options.captureTimeoutMs ?? 500,
      captureTransaction: (transaction) =>
        transaction.meta.get('addToHistory') !== false,
      deleteFilter: (item) => defaultDeleteFilter(item, defaultProtectedNodes),
      trackedOrigins: new Set([this.origin]),
    });
    // TipTap's collaboration plugin owns this manager while mounted, so keep
    // the shared sidecar scopes and this binding's origin after a StrictMode
    // unmount/remount restores the manager.
    const restorable = this.#undoManager as RestorableUndoManager;
    const restore = restorable.restore?.bind(this.#undoManager);
    if (restore) {
      restorable.restore = () => {
        restore();
        this.#undoManager.addToScope(
          this.#session.mode === 'comment' || this.#session.mode === 'suggest'
            ? this.fragment
            : [
                this.fragment,
                ...workOfficeDocumentSidecarUndoScope(this.#session),
              ],
        );
        this.#undoManager.addTrackedOrigin(this.origin);
      };
    }
    this.extensions = [
      documentCollaborationLifecycleExtension(this),
      documentCollaborationPermissionExtension(session),
      ...builtIns,
      Collaboration.configure({
        fragment: this.fragment,
        onFirstRender: options.onFirstRender,
        yUndoOptions: { undoManager: this.#undoManager },
      }),
      ...additional,
    ];
    mountedDocumentBindings.set(session.document, this);
    session.document.on('beforeTransaction', this.#onBeforeTransaction);
    session.document.on('afterTransaction', this.#onTransaction);
    this.#undoManager.on('stack-item-added', this.#onHistoryChange);
    this.#undoManager.on('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.on('stack-item-updated', this.#onHistoryChange);
    this.#observersAttached = true;
  }

  content(): WorkDocumentContent {
    this.ensureActive();
    return readWorkOfficeDocumentCollaboration(this.#session);
  }

  updateSidecars(
    previous: WorkDocumentContent,
    next: WorkDocumentContent,
  ): boolean {
    this.ensureActive();
    const historyEntry =
      this.#session.mode === 'comment' && !this.#applyingCommentHistory
        ? createDocumentCommentHistoryEntry(previous, next)
        : null;
    if (historyEntry?.commentMembershipChanged) {
      this.#undoManager.stopCapturing();
    }
    const changed = updateWorkOfficeDocumentSidecars(
      this.#session,
      previous,
      next,
      this.origin,
    );
    if (changed && historyEntry) {
      appendBoundedCommentHistory(this.#commentUndoStack, historyEntry);
      this.#commentRedoStack.length = 0;
      this.#onHistoryChange();
    }
    return changed;
  }

  decideChanges(
    editor: Editor,
    changeIds: readonly string[],
    decision: WorkDocumentChangeDecisionAction,
    options: WorkOfficeDocumentChangeDecisionOptions = {},
  ): boolean {
    this.ensureActive();
    assertWorkOfficeCollaborationEditable(this.#session);
    if (editor.isDestroyed || changeIds.length === 0) return false;
    const requested = new Set(changeIds);
    const changes = collectDocumentChanges(editor.state.doc).filter((change) =>
      requested.has(change.id),
    );
    if (changes.length !== requested.size) return false;
    const before = this.content();
    const decisions = createWorkOfficeDocumentChangeDecisions(
      this.#session,
      changes,
      decision,
      options.decidedAt,
    );
    const existing = new Set(
      (before.changeDecisions ?? []).map(({ id }) => id),
    );
    if (decisions.some(({ id }) => existing.has(id))) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.permission_denied',
        'A tracked change with a final decision cannot be decided again.',
      );
    }
    const origin = this.#session.createOrigin(this.#session.localOrigin.kind);
    let handled = false;
    this.#session.transact(() => {
      handled =
        decision === 'accept'
          ? editor.commands.acceptDocumentChanges(changeIds)
          : editor.commands.rejectDocumentChanges(changeIds);
      if (!handled) return;
      updateWorkOfficeDocumentSidecars(
        this.#session,
        before,
        {
          ...before,
          changeDecisions: [...(before.changeDecisions ?? []), ...decisions],
        },
        origin,
      );
    }, origin);
    if (handled) {
      // A shared accept/reject decision is final. Clear older local history so
      // undo cannot resurrect a decided mark or remove accepted content while
      // leaving its append-only audit record behind.
      this.#undoManager.clear();
      this.#onHistoryChange();
    }
    return handled;
  }

  canUndo(): boolean {
    this.ensureActive();
    if (this.#session.mode === 'comment') {
      return this.#commentUndoStack.length > 0 || this.#undoManager.canUndo();
    }
    return this.#undoManager.canUndo();
  }

  canRedo(): boolean {
    this.ensureActive();
    if (this.#session.mode === 'comment') {
      return this.#commentRedoStack.length > 0 || this.#undoManager.canRedo();
    }
    return this.#undoManager.canRedo();
  }

  undo(): boolean {
    this.ensureActive();
    if (this.#session.mode === 'comment') {
      assertWorkOfficeCollaborationWritable(this.#session, 'document-comment');
      return this.applyCommentHistory('undo');
    }
    if (this.#session.mode === 'suggest') {
      assertWorkOfficeCollaborationWritable(
        this.#session,
        'document-suggestion',
      );
      return this.#undoManager.undo() !== null;
    }
    assertWorkOfficeCollaborationEditable(this.#session);
    return this.#undoManager.undo() !== null;
  }

  redo(): boolean {
    this.ensureActive();
    if (this.#session.mode === 'comment') {
      assertWorkOfficeCollaborationWritable(this.#session, 'document-comment');
      return this.applyCommentHistory('redo');
    }
    if (this.#session.mode === 'suggest') {
      assertWorkOfficeCollaborationWritable(
        this.#session,
        'document-suggestion',
      );
      return this.#undoManager.redo() !== null;
    }
    assertWorkOfficeCollaborationEditable(this.#session);
    return this.#undoManager.redo() !== null;
  }

  stopCapturing(): void {
    this.ensureActive();
    this.#undoManager.stopCapturing();
  }

  subscribe(
    listener: (change: WorkOfficeDocumentCollaborationChange) => void,
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
    if (this.#destroyed || this.#destroyRequested) return;
    this.#destroyRequested = true;
    this.detachObservers();
    if (this.#mountedEditors.size === 0) this.finishDestroy(true);
  }

  mount(editor: Editor): void {
    this.ensureActive();
    this.#mountedEditors.add(editor);
  }

  unmount(editor: Editor): void {
    this.#mountedEditors.delete(editor);
    if (this.#destroyRequested && this.#mountedEditors.size === 0) {
      queueMicrotask(() => {
        if (this.#destroyRequested && this.#mountedEditors.size === 0) {
          this.finishDestroy(true);
        }
      });
    }
  }

  readonly #onBeforeTransaction = (transaction: Y.Transaction): void => {
    if (transaction.origin === ySyncPluginKey) {
      transaction.origin = this.origin;
    }
  };

  readonly #onTransaction = (transaction: Y.Transaction): void => {
    if (
      this.#session.mode === 'suggest' &&
      workOfficeDocumentDecisionRootsChanged(this.#session, transaction)
    ) {
      this.#undoManager.clear();
      this.#onHistoryChange();
    }
    if (
      this.#destroyed ||
      this.#destroyRequested ||
      (!transactionTouchesRoot(transaction, this.fragment) &&
        !workOfficeDocumentSidecarsChanged(this.#session, transaction))
    ) {
      return;
    }
    const local =
      transaction.origin === ySyncPluginKey ||
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
    if (this.#destroyed || this.#destroyRequested || !this.#pendingChange) {
      this.#pendingChange = false;
      this.#pendingLocal = true;
      this.#pendingOrigin = undefined;
      return;
    }
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
      } satisfies WorkOfficeDocumentCollaborationChange;
      for (const listener of this.#listeners) listener(change);
    } catch (error) {
      for (const listener of this.#errorListeners) listener(error);
    }
  };

  readonly #onHistoryChange = (): void => {
    for (const listener of this.#historyListeners) listener();
  };

  private applyCommentHistory(direction: 'undo' | 'redo'): boolean {
    const sourceStack =
      direction === 'undo' ? this.#commentUndoStack : this.#commentRedoStack;
    const targetStack =
      direction === 'undo' ? this.#commentRedoStack : this.#commentUndoStack;
    const entry = sourceStack.pop();
    if (!entry) {
      return direction === 'undo'
        ? this.#undoManager.undo() !== null
        : this.#undoManager.redo() !== null;
    }
    const from = direction === 'undo' ? entry.after : entry.before;
    const to = direction === 'undo' ? entry.before : entry.after;
    const current = this.content();
    const comments = reconcileDocumentCommentHistory(
      current.comments,
      from,
      to,
    );
    let sidecarsChanged = false;
    this.#applyingCommentHistory = true;
    try {
      if (!jsonEqual(current.comments, comments)) {
        sidecarsChanged = this.updateSidecars(current, {
          ...current,
          comments,
        });
      }
    } finally {
      this.#applyingCommentHistory = false;
    }
    const anchorChanged = entry.commentMembershipChanged
      ? direction === 'undo'
        ? this.#undoManager.undo() !== null
        : this.#undoManager.redo() !== null
      : false;
    appendBoundedCommentHistory(targetStack, entry);
    this.#onHistoryChange();
    return sidecarsChanged || anchorChanged;
  }

  private finishDestroy(destroyUndoManager: boolean): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#mountedEditors.clear();
    if (mountedDocumentBindings.get(this.#session.document) === this) {
      mountedDocumentBindings.delete(this.#session.document);
    }
    this.detachObservers();
    if (destroyUndoManager) {
      (this.#undoManager as RestorableUndoManager).restore?.();
      this.#undoManager.destroy();
    }
    this.#listeners.clear();
    this.#errorListeners.clear();
    this.#historyListeners.clear();
    this.#commentUndoStack.length = 0;
    this.#commentRedoStack.length = 0;
  }

  private detachObservers(): void {
    if (!this.#observersAttached) return;
    this.#observersAttached = false;
    this.#session.document.off('beforeTransaction', this.#onBeforeTransaction);
    this.#session.document.off('afterTransaction', this.#onTransaction);
    this.#undoManager.off('stack-item-added', this.#onHistoryChange);
    this.#undoManager.off('stack-item-popped', this.#onHistoryChange);
    this.#undoManager.off('stack-item-updated', this.#onHistoryChange);
  }

  private ensureActive(): void {
    if (!this.#destroyed && !this.#destroyRequested) return;
    throw new WorkOfficeCollaborationError(
      'office.collaboration.binding_destroyed',
      'The Document collaboration binding has been destroyed.',
    );
  }
}

function createDocumentCommentHistoryEntry(
  previous: WorkDocumentContent,
  next: WorkDocumentContent,
): WorkOfficeDocumentCommentHistoryEntry | null {
  if (jsonEqual(previous.comments, next.comments)) return null;
  const before = cloneDocumentComments(previous.comments);
  const after = cloneDocumentComments(next.comments);
  return {
    before,
    after,
    commentMembershipChanged: !sameDocumentCommentMembership(before, after),
  };
}

function appendBoundedCommentHistory(
  stack: WorkOfficeDocumentCommentHistoryEntry[],
  entry: WorkOfficeDocumentCommentHistoryEntry,
): void {
  stack.push(entry);
  if (stack.length > MAX_DOCUMENT_COMMENT_HISTORY) stack.shift();
}

function sameDocumentCommentMembership(
  left: readonly WorkDocumentComment[] | undefined,
  right: readonly WorkDocumentComment[] | undefined,
): boolean {
  return jsonEqual(
    (left ?? []).map(({ id }) => id),
    (right ?? []).map(({ id }) => id),
  );
}

function reconcileDocumentCommentHistory(
  current: readonly WorkDocumentComment[] | undefined,
  from: readonly WorkDocumentComment[] | undefined,
  to: readonly WorkDocumentComment[] | undefined,
): WorkDocumentComment[] | undefined {
  const fromById = new Map(
    (from ?? []).map((comment) => [comment.id, comment]),
  );
  const toById = new Map((to ?? []).map((comment) => [comment.id, comment]));
  const result = cloneDocumentComments(current) ?? [];
  const resultById = () =>
    new Map(result.map((comment) => [comment.id, comment]));
  const orderedIds = [
    ...(from ?? []).map(({ id }) => id),
    ...(to ?? []).map(({ id }) => id),
  ].filter((id, index, ids) => ids.indexOf(id) === index);

  for (const id of orderedIds) {
    const source = fromById.get(id);
    const target = toById.get(id);
    const currentById = resultById();
    const existing = currentById.get(id);
    if (!source && target) {
      if (!existing) result.push(cloneDocumentComment(target));
      continue;
    }
    if (source && !target) {
      if (existing && jsonEqual(existing, source)) {
        result.splice(
          result.findIndex((comment) => comment.id === id),
          1,
        );
      }
      continue;
    }
    if (!source || !target || !existing) continue;
    const next = reconcileExistingDocumentComment(existing, source, target);
    const index = result.findIndex((comment) => comment.id === id);
    if (index >= 0) result[index] = next;
  }

  if (result.length > 0) return result;
  return to === undefined ? undefined : [];
}

function reconcileExistingDocumentComment(
  current: WorkDocumentComment,
  from: WorkDocumentComment,
  to: WorkDocumentComment,
): WorkDocumentComment {
  const next = cloneDocumentComment(current);
  if (from.resolved !== to.resolved && current.resolved === from.resolved) {
    next.resolved = to.resolved;
  }
  const fromReplies = new Map(
    (from.replies ?? []).map((reply) => [reply.id, reply]),
  );
  const toReplies = new Map(
    (to.replies ?? []).map((reply) => [reply.id, reply]),
  );
  const replies = (next.replies ?? []).map((reply) => ({ ...reply }));
  const replyIds = [
    ...(from.replies ?? []).map(({ id }) => id),
    ...(to.replies ?? []).map(({ id }) => id),
  ].filter((id, index, ids) => ids.indexOf(id) === index);
  for (const id of replyIds) {
    const source = fromReplies.get(id);
    const target = toReplies.get(id);
    const index = replies.findIndex((reply) => reply.id === id);
    if (!source && target && index < 0) {
      replies.push({ ...target });
    } else if (
      source &&
      !target &&
      index >= 0 &&
      jsonEqual(replies[index], source)
    ) {
      replies.splice(index, 1);
    }
  }
  if (replies.length > 0 || to.replies !== undefined) next.replies = replies;
  else delete next.replies;
  return next;
}

function cloneDocumentComments(
  comments: readonly WorkDocumentComment[] | undefined,
): WorkDocumentComment[] | undefined {
  return comments?.map(cloneDocumentComment);
}

function cloneDocumentComment(
  comment: WorkDocumentComment,
): WorkDocumentComment {
  return {
    ...comment,
    replies: comment.replies?.map((reply) => ({ ...reply })),
  };
}

export function workOfficeDocumentCollaborationFragment(
  session: WorkOfficeCollaborationSession,
): Y.XmlFragment {
  assertDocumentSession(session);
  return session.document.getXmlFragment(
    session.rootName(DOCUMENT_CONTENT_ROOT),
  );
}

export function workOfficeDocumentCollaborationNode(
  session: WorkOfficeCollaborationSession,
): WorkDocumentNode {
  assertInitializedDocumentSession(session);
  return yXmlFragmentToProsemirrorJSON(
    workOfficeDocumentCollaborationFragment(session),
  ) as unknown as WorkDocumentNode;
}

function preparedDocumentContent(
  content: WorkDocumentContent,
): WorkDocumentContent {
  if (
    !content ||
    content.type !== 'document' ||
    typeof content.html !== 'string'
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      'Document collaboration initialization requires a valid Document content value.',
    );
  }
  const prepared = documentModelForContent(content)
    ? content
    : createWorkDocumentModelFromContent(content);
  requiredDocumentModel(prepared);
  return prepared;
}

function requiredDocumentModel(content: WorkDocumentContent) {
  const model = documentModelForContent(content);
  if (model) return model;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    'Document collaboration initialization requires content supported by the Office document schema.',
  );
}

function preflightDocumentRoot(content: WorkDocumentContent): WorkDocumentNode {
  const root = requiredDocumentModel(content).root;
  const scratch = new Y.Doc();
  try {
    const schema = workDocumentSchema();
    schema.nodeFromJSON(root).check();
    prosemirrorJSONToYXmlFragment(
      schema,
      root,
      scratch.getXmlFragment('preflight'),
    );
    return root;
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `The initial structured content does not match the Office document schema: ${errorMessage(error)}`,
    );
  } finally {
    scratch.destroy();
  }
}

function documentContentFromSharedState(
  html: string,
  sidecars: WorkOfficeDocumentSidecars,
): WorkDocumentContent {
  return {
    type: 'document',
    html,
    pageSize: 'a4',
    ...sidecars,
  };
}

function validateSharedDocumentWithSchema(
  session: WorkOfficeCollaborationSession,
  schema: ReturnType<typeof getSchema>,
): void {
  try {
    schema.nodeFromJSON(workOfficeDocumentCollaborationNode(session)).check();
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `The shared Document collaboration content does not match the mounted editor schema: ${errorMessage(error)}`,
    );
  }
}

function assertUniqueExtensionNames(extensions: Extensions): void {
  const names = new Set<string>();
  for (const extension of extensions) {
    if (names.has(extension.name)) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The Document collaboration editor extension '${extension.name}' is registered more than once.`,
      );
    }
    names.add(extension.name);
  }
}

function assertBehaviorOnlyExtensions(extensions: Extensions): void {
  for (const extension of extensions) {
    if (
      extension.type === 'extension' &&
      !extension.config.addExtensions &&
      !extension.config.addGlobalAttributes &&
      !extension.config.extendNodeSchema &&
      !extension.config.extendMarkSchema
    ) {
      continue;
    }
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `Document collaboration cannot mount the schema-changing extension '${extension.name}'. Shared nodes, marks, and attributes must be added through a versioned Office schema migration.`,
    );
  }
}

function assertDocumentSession(session: WorkOfficeCollaborationSession): void {
  if (session.kind === 'document') return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.kind_mismatch',
    `A Document collaboration binding cannot use a '${session.kind}' session.`,
  );
}

function assertInitializedDocumentSession(
  session: WorkOfficeCollaborationSession,
): void {
  assertDocumentSession(session);
  const metadata = readWorkOfficeCollaborationMetadata(session);
  if (metadata?.initialized) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.not_initialized',
    'The Document collaboration session has not been initialized.',
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function documentCollaborationLifecycleExtension(
  binding: WorkOfficeDocumentCollaborationBindingImpl,
): Extension {
  return Extension.create({
    name: 'officeDocumentCollaborationLifecycle',
    onCreate() {
      binding.mount(this.editor);
    },
    onDestroy() {
      binding.unmount(this.editor);
    },
  });
}

function documentCollaborationPermissionExtension(
  session: WorkOfficeCollaborationSession,
): Extension {
  return Extension.create({
    name: 'officeDocumentCollaborationPermission',
    addProseMirrorPlugins() {
      return [
        new Plugin({
          filterTransaction(transaction) {
            return (
              session.mode === 'edit' ||
              !transaction.docChanged ||
              isChangeOrigin(transaction) ||
              (session.mode === 'comment' &&
                transaction.steps.length > 0 &&
                transaction.steps.every(
                  (step) =>
                    (step instanceof AddMarkStep ||
                      step instanceof RemoveMarkStep) &&
                    step.mark.type.name === 'documentComment',
                )) ||
              (session.mode === 'suggest' &&
                workOfficeDocumentSuggestionTransactionAllowed(
                  session,
                  transaction,
                ))
            );
          },
        }),
      ];
    },
  });
}

function transactionTouchesRoot(
  transaction: Y.Transaction,
  root: unknown,
): boolean {
  return new Set<unknown>(transaction.changedParentTypes.keys()).has(root);
}
