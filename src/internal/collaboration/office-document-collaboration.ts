import {
  type Editor,
  Extension,
  type Extensions,
  getSchema,
} from '@tiptap/core';
import { Collaboration, isChangeOrigin } from '@tiptap/extension-collaboration';
import { Plugin } from '@tiptap/pm/state';
import {
  defaultDeleteFilter,
  defaultProtectedNodes,
  prosemirrorJSONToYXmlFragment,
  ySyncPluginKey,
  yXmlFragmentToProsemirrorJSON,
} from '@tiptap/y-tiptap';
import * as Y from 'yjs';
import { createWorkDocumentExtensions } from '../features/work/work-document-extensions';
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
  WorkDocumentContent,
  WorkDocumentNode,
} from '../features/work/work-types';
import {
  assertWorkOfficeCollaborationEditable,
  assertWorkOfficeCollaborationOrigin,
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
} from './office-document-collaboration-sidecars';

const DOCUMENT_CONTENT_ROOT = 'document.content';

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

export interface WorkOfficeDocumentCollaborationBinding {
  readonly extensions: Extensions;
  readonly fragment: Y.XmlFragment;
  readonly origin: WorkOfficeCollaborationOrigin;
  content(): WorkDocumentContent;
  updateSidecars(
    previous: WorkDocumentContent,
    next: WorkDocumentContent,
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
    this.#undoManager = new Y.UndoManager(
      [this.fragment, ...workOfficeDocumentSidecarUndoScope(session)],
      {
        captureTimeout: options.captureTimeoutMs ?? 500,
        captureTransaction: (transaction) =>
          transaction.meta.get('addToHistory') !== false,
        deleteFilter: (item) =>
          defaultDeleteFilter(item, defaultProtectedNodes),
        trackedOrigins: new Set([this.origin]),
      },
    );
    // TipTap's collaboration plugin owns this manager while mounted, so keep
    // the shared sidecar scopes and this binding's origin after a StrictMode
    // unmount/remount restores the manager.
    const restorable = this.#undoManager as RestorableUndoManager;
    const restore = restorable.restore?.bind(this.#undoManager);
    if (restore) {
      restorable.restore = () => {
        restore();
        this.#undoManager.addToScope([
          this.fragment,
          ...workOfficeDocumentSidecarUndoScope(this.#session),
        ]);
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
    return updateWorkOfficeDocumentSidecars(
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
              isChangeOrigin(transaction)
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
