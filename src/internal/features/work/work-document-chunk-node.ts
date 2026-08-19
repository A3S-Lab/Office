import type { Editor } from '@tiptap/core';
import { mergeAttributes, Node } from '@tiptap/core';
import { Fragment, type Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  type EditorState,
  Plugin,
  PluginKey,
  Selection,
  TextSelection,
  type Transaction,
} from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import {
  documentChunkGeometry,
  documentChunkId,
  documentChunkPositions,
  documentChunkViewportRange,
  type DocumentChunkGeometry,
  type DocumentChunkPosition,
} from './work-document-chunk-geometry';
import { primeDocumentIntegrityFeatures } from './work-document-integrity-index';
import {
  DOCUMENT_LAZY_BLOCK_NODE,
  documentLazyChunkContent,
} from './work-document-lazy-model';
import {
  createDocumentLazyPreviewPool,
  type DocumentLazyPreviewLease,
  documentLazyPreviewPointerPosition,
  type DocumentLazyPreviewPool,
  mountDocumentLazyPreview,
  releaseDocumentLazyPreview,
  updateDocumentLazyPreviewSpacer,
} from './work-document-lazy-preview';
import { transferDocumentTextStatistics } from './work-document-text-statistics';
import type { WorkDocumentContent, WorkDocumentNode } from './work-types';

interface DocumentChunkOptions {
  getContent: () => WorkDocumentContent | null;
  trustInitialIntegrityFeatures: boolean;
}

interface DocumentChunkWindowState {
  chunkCount: number;
  decorations: DecorationSet;
  fallback: boolean;
  geometry: readonly DocumentChunkGeometry[];
  ids: readonly string[];
  mountedIds: ReadonlySet<string>;
  positions: ReadonlyMap<string, DocumentChunkPosition>;
}

interface DocumentChunkWindowMeta {
  visibleIds: readonly string[];
}

export interface DocumentChunkHydrationMeta {
  ids: readonly string[];
}

interface DocumentChunkPaginationMeta {
  breaks: readonly {
    position: number;
    spacerHeight: number;
  }[];
}

export interface DocumentChunkVisibleIdsMeta {
  visibleIds: readonly string[];
}

interface DocumentChunkControllerChange {
  id: string;
  next?: DocumentChunkNodeViewController;
  previous?: DocumentChunkNodeViewController;
}

interface DocumentChunkNodeViewController {
  element: HTMLElement;
  setPaginationExtraHeight: (height: number) => void;
}

interface DocumentChunkNodeViewRegistry {
  controllers: Map<string, DocumentChunkNodeViewController>;
  lazyPreviewPool: DocumentLazyPreviewPool;
  listeners: Set<(change: DocumentChunkControllerChange) => void>;
  paginationExtraHeights: Map<string, number>;
}

const INITIAL_DOCUMENT_CHUNK_WINDOW_SIZE = 2;
const DOCUMENT_CHUNK_WINDOW_ROOT_MARGIN = '0px';
export const DOCUMENT_CHUNK_PAGINATION_META = 'documentChunkPaginationGeometry';
export const DOCUMENT_CHUNK_VISIBLE_IDS_META = 'documentChunkVisibleIds';
export const DOCUMENT_CHUNK_HYDRATION_META = 'documentChunkLazyHydration';
export { documentChunkViewportRange } from './work-document-chunk-geometry';
const documentChunkWindowKey = new PluginKey<DocumentChunkWindowState>(
  'documentChunkWindow',
);
const documentChunkElementRegistries = new WeakMap<
  Editor,
  DocumentChunkNodeViewRegistry
>();
const documentChunkContentSources = new WeakMap<
  Editor,
  () => WorkDocumentContent | null
>();

export const DocumentChunk = Node.create<DocumentChunkOptions>({
  name: 'documentChunk',
  group: 'block',
  content: 'block+',
  defining: true,
  selectable: false,

  addOptions() {
    return {
      getContent: () => null,
      trustInitialIntegrityFeatures: false,
    };
  },

  addAttributes() {
    return {
      id: { default: '' },
      blockCount: { default: 0 },
      estimatedHeight: { default: 1 },
      integrityFeatures: { default: null, rendered: false },
      windowContainer: { default: false },
    };
  },

  parseHTML() {
    return [
      {
        tag: 'div[data-document-chunk]',
        getAttrs: (element) => {
          if (!(element instanceof HTMLElement)) return false;
          return {
            id: element.dataset.documentChunkId ?? '',
            blockCount: positiveInteger(element.dataset.documentChunkBlocks, 0),
            estimatedHeight: positiveNumber(
              element.dataset.documentChunkEstimatedHeight,
              1,
            ),
            windowContainer:
              element.dataset.documentChunkWindowContainer === 'true',
          };
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    return [
      'div',
      mergeAttributes(HTMLAttributes, documentChunkDomAttributes(node)),
      0,
    ];
  },

  addNodeView() {
    return ({ decorations, editor, getPos, node }) =>
      createDocumentChunkNodeView(editor, node, decorations, getPos);
  },

  addProseMirrorPlugins() {
    const editor = this.editor;
    const getContent = this.options.getContent;
    documentChunkContentSources.set(editor, getContent);
    const trustInitialIntegrityFeatures =
      this.options.trustInitialIntegrityFeatures;
    return [
      new Plugin<DocumentChunkWindowState>({
        key: documentChunkWindowKey,
        appendTransaction: (transactions, _previous, next) => {
          if (!transactions.some((transaction) => transaction.selectionSet)) {
            return null;
          }
          const selected = selectedDocumentChunkId(next);
          if (!selected) return null;
          const transaction = next.tr;
          return hydrateDocumentChunks(editor, transaction, next, [selected])
            .length
            ? transaction
            : null;
        },
        state: {
          init: (_configuration, state) => {
            if (trustInitialIntegrityFeatures) {
              primeDocumentIntegrityFeatures(state.doc);
            }
            return createDocumentChunkWindowState(state);
          },
          apply: (transaction, current, _previous, next) =>
            updateDocumentChunkWindowState(transaction, current, next),
        },
        props: {
          decorations: (state) =>
            documentChunkWindowKey.getState(state)?.decorations ?? null,
          handleDOMEvents: {
            pointerdown: (view, event) =>
              handleDocumentChunkLazyPreviewPointerDown(editor, view, event),
          },
          handleKeyDown: (view, event) =>
            handleDocumentChunkBoundaryKeyDown(editor, view, event),
        },
        view: (view) => createDocumentChunkIntersectionView(editor, view),
      }),
    ];
  },
});

function handleDocumentChunkBoundaryKeyDown(
  editor: Editor,
  view: EditorView,
  event: KeyboardEvent,
): boolean {
  const state = documentChunkWindowKey.getState(view.state);
  if (
    !state?.chunkCount ||
    event.altKey ||
    event.shiftKey ||
    (!event.ctrlKey && !event.metaKey)
  ) {
    return false;
  }
  const moveToStart =
    event.key === 'Home' || (event.metaKey && event.key === 'ArrowUp');
  const moveToEnd =
    event.key === 'End' || (event.metaKey && event.key === 'ArrowDown');
  if (!moveToStart && !moveToEnd) return false;
  const selection = moveToStart
    ? Selection.atStart(view.state.doc)
    : Selection.atEnd(view.state.doc);
  const transaction = view.state.tr.setSelection(selection);
  const selected = selectedDocumentChunkId({
    doc: view.state.doc,
    selection,
  });
  if (selected) {
    hydrateDocumentChunks(editor, transaction, view.state, [selected]);
  }
  const nextSelection = moveToStart
    ? Selection.atStart(transaction.doc)
    : Selection.atEnd(transaction.doc);
  view.dispatch(transaction.setSelection(nextSelection).scrollIntoView());
  return true;
}

function createDocumentChunkNodeView(
  editor: Editor,
  initialNode: ProseMirrorNode,
  initialDecorations: readonly Decoration[],
  getPos: () => number | undefined,
) {
  let node = initialNode;
  const id = documentChunkId(node);
  const mounted = documentChunkIsMounted(node, initialDecorations);
  let lazyPreview = mounted && documentChunkHasLazyContent(node);
  const registry = documentChunkElementRegistryFor(editor);
  let paginationExtraHeight = registry.paginationExtraHeights.get(id) ?? 0;
  const dom = document.createElement('div');
  let lazyPreviewLease: DocumentLazyPreviewLease | null = null;
  applyDocumentChunkNodeViewAttributes(
    dom,
    node,
    mounted,
    paginationExtraHeight,
  );
  if (lazyPreview) {
    lazyPreviewLease = mountDocumentLazyPreview(
      registry.lazyPreviewPool,
      dom,
      requiredDocumentLazyChunkContent(editor, id),
      (getPos() ?? 0) + 1,
      paginationExtraHeight,
    );
  }
  const controller: DocumentChunkNodeViewController = {
    element: dom,
    setPaginationExtraHeight(height) {
      if (height === paginationExtraHeight) return;
      paginationExtraHeight = height;
      applyDocumentChunkNodeViewAttributes(
        dom,
        node,
        mounted,
        paginationExtraHeight,
      );
      if (lazyPreview) {
        dom.setAttribute('contenteditable', 'false');
        updateDocumentLazyPreviewSpacer(lazyPreviewLease, height);
      }
    },
  };
  const unregister = registerDocumentChunkController(editor, id, controller);

  return {
    dom,
    ...(mounted && !lazyPreview ? { contentDOM: dom } : {}),
    destroy() {
      releaseDocumentLazyPreview(
        registry.lazyPreviewPool,
        dom,
        lazyPreviewLease,
      );
      lazyPreviewLease = null;
      unregister();
    },
    ignoreMutation() {
      return !mounted || lazyPreview;
    },
    update(
      nextNode: ProseMirrorNode,
      decorations: readonly Decoration[],
    ): boolean {
      if (nextNode.type !== node.type) return false;
      if (documentChunkId(nextNode) !== id) return false;
      if (mounted && documentChunkHasLazyContent(nextNode) !== lazyPreview) {
        return false;
      }
      if (documentChunkIsMounted(nextNode, decorations) !== mounted) {
        return false;
      }
      node = nextNode;
      lazyPreview = mounted && documentChunkHasLazyContent(node);
      applyDocumentChunkNodeViewAttributes(
        dom,
        node,
        mounted,
        paginationExtraHeight,
      );
      if (lazyPreview) dom.setAttribute('contenteditable', 'false');
      return true;
    },
  };
}

function documentChunkHasLazyContent(node: ProseMirrorNode): boolean {
  return (
    node.childCount === 1 &&
    node.firstChild?.type.name === DOCUMENT_LAZY_BLOCK_NODE
  );
}

function requiredDocumentLazyChunkContent(
  editor: Editor,
  chunkId: string,
): readonly WorkDocumentNode[] {
  const payload = documentLazyChunkContentForEditor(editor, chunkId);
  if (!payload?.length) {
    throw new Error(`The lazy document chunk payload "${chunkId}" is missing.`);
  }
  return payload;
}

function handleDocumentChunkLazyPreviewPointerDown(
  editor: Editor,
  view: EditorView,
  event: Event,
): boolean {
  const target = event.target;
  const chunk =
    target instanceof Element
      ? target.closest<HTMLElement>(
          '[data-document-chunk][data-document-lazy-preview="true"]',
        )
      : null;
  const id = chunk?.dataset.documentChunkId;
  if (!id) return false;
  const position = documentChunkWindowKey
    .getState(view.state)
    ?.positions.get(id);
  const logicalPosition = documentLazyPreviewPointerPosition(
    event,
    (position?.from ?? 0) + 2,
  );
  if (logicalPosition === null) return false;
  event.preventDefault();
  editor.commands.setTextSelection(logicalPosition);
  view.focus();
  return true;
}

function createDocumentChunkWindowState(
  state: EditorState,
): DocumentChunkWindowState {
  const geometry = documentChunkGeometry(state.doc, []);
  const ids = geometry.map(({ id }) => id);
  const positions = documentChunkPositions(state.doc);
  const fallback = typeof IntersectionObserver === 'undefined';
  const mountedIds = fallback
    ? new Set(ids)
    : new Set(ids.slice(0, INITIAL_DOCUMENT_CHUNK_WINDOW_SIZE));
  const selected = selectedDocumentChunkId(state);
  if (selected) mountedIds.add(selected);
  return {
    chunkCount: ids.length,
    decorations: documentChunkDecorations(state.doc, mountedIds, positions),
    fallback,
    geometry,
    ids,
    mountedIds,
    positions,
  };
}

function updateDocumentChunkWindowState(
  transaction: Transaction,
  current: DocumentChunkWindowState,
  next: EditorState,
): DocumentChunkWindowState {
  const meta = transaction.getMeta(documentChunkWindowKey) as
    | DocumentChunkWindowMeta
    | undefined;
  const paginationMeta = transaction.getMeta(DOCUMENT_CHUNK_PAGINATION_META) as
    | DocumentChunkPaginationMeta
    | undefined;
  const hydrationMeta = transaction.getMeta(DOCUMENT_CHUNK_HYDRATION_META) as
    | DocumentChunkHydrationMeta
    | undefined;
  const appendedTo = transaction.getMeta('appendedTransaction') as
    | Transaction
    | undefined;
  const sourcePaginationMeta = appendedTo?.getMeta(
    DOCUMENT_CHUNK_PAGINATION_META,
  ) as DocumentChunkPaginationMeta | undefined;
  const paginationBreaks =
    paginationMeta?.breaks ??
    sourcePaginationMeta?.breaks.map((pageBreak) => ({
      ...pageBreak,
      position: transaction.mapping.map(pageBreak.position, -1),
    }));
  if (
    !meta &&
    !paginationMeta &&
    !transaction.docChanged &&
    !transaction.selectionSet
  ) {
    return current;
  }
  const geometry = paginationBreaks
    ? documentChunkGeometry(next.doc, paginationBreaks)
    : transaction.docChanged && !hydrationMeta
      ? documentChunkGeometry(next.doc, [])
      : current.geometry;
  const ids =
    geometry === current.geometry ? current.ids : geometry.map(({ id }) => id);
  const positions =
    geometry === current.geometry
      ? current.positions
      : documentChunkPositions(next.doc);
  const known = geometry === current.geometry ? null : new Set(ids);
  const mountedIds = current.fallback
    ? new Set(ids)
    : new Set(
        Array.from(meta?.visibleIds ?? current.mountedIds).filter(
          (id) => !known || known.has(id),
        ),
      );
  const selected = selectedDocumentChunkId(next);
  if (selected) mountedIds.add(selected);
  if (
    !transaction.docChanged &&
    geometry === current.geometry &&
    setsEqual(mountedIds, current.mountedIds) &&
    ids.length === current.chunkCount
  ) {
    return current;
  }
  return {
    chunkCount: ids.length,
    decorations: documentChunkDecorations(next.doc, mountedIds, positions),
    fallback: current.fallback,
    geometry,
    ids,
    mountedIds,
    positions,
  };
}

function createDocumentChunkIntersectionView(
  editor: Editor,
  initialView: EditorView,
) {
  let view = initialView;
  let observer: IntersectionObserver | null = null;
  let scrollRoot: HTMLElement | null = null;
  let publishQueued = false;
  let mountFrame = 0;
  let synchronizedGeometry: readonly DocumentChunkGeometry[] | null = null;
  const observed = new Map<string, Element>();
  const intersecting = new Set<string>();
  const registry = documentChunkElementRegistryFor(editor);

  const dispatchVisibleWindow = (requestedIds: Iterable<string>) => {
    if (view.isDestroyed) return;
    const state = documentChunkWindowKey.getState(view.state);
    if (!state) return;
    const visibleIds = new Set(requestedIds);
    const selected = selectedDocumentChunkId(view.state);
    if (selected) visibleIds.add(selected);
    if (setsEqual(visibleIds, state.mountedIds)) return;
    const startedAt = globalThis.performance?.now?.() ?? Date.now();
    const transaction = view.state.tr;
    view.dispatch(
      transaction
        .setMeta(documentChunkWindowKey, {
          visibleIds: [...visibleIds],
        } satisfies DocumentChunkWindowMeta)
        .setMeta(DOCUMENT_CHUNK_VISIBLE_IDS_META, {
          visibleIds: [...visibleIds],
        } satisfies DocumentChunkVisibleIdsMeta)
        .setMeta('addToHistory', false)
        .setMeta('preventUpdate', true),
    );
    const duration =
      (globalThis.performance?.now?.() ?? Date.now()) - startedAt;
    const updates = Number(view.dom.dataset.documentWindowUpdates) || 0;
    const total = Number(view.dom.dataset.documentWindowTotalMs) || 0;
    const maximum = Number(view.dom.dataset.documentWindowMaxMs) || 0;
    view.dom.dataset.documentWindowUpdates = String(updates + 1);
    view.dom.dataset.documentWindowLastMs = String(roundDuration(duration));
    view.dom.dataset.documentWindowTotalMs = String(
      roundDuration(total + duration),
    );
    view.dom.dataset.documentWindowMaxMs = String(
      roundDuration(Math.max(maximum, duration)),
    );
  };
  const publishIntersectionWindow = () => {
    const state = documentChunkWindowKey.getState(view.state);
    if (!state) return;
    const visibleIds = new Set(
      Array.from(intersecting).filter((id) => registry.controllers.has(id)),
    );
    for (const id of [...intersecting]) {
      const index = state.ids.indexOf(id);
      if (index > 0) visibleIds.add(state.ids[index - 1] as string);
      if (index >= 0 && index + 1 < state.ids.length) {
        visibleIds.add(state.ids[index + 1] as string);
      }
    }
    dispatchVisibleWindow(visibleIds);
  };
  const publishGeometryWindow = () => {
    if (!scrollRoot || view.isDestroyed) return;
    const state = documentChunkWindowKey.getState(view.state);
    if (!state?.geometry.length) return;
    const rootRect = scrollRoot.getBoundingClientRect();
    const editorRect = view.dom.getBoundingClientRect();
    const scale =
      view.dom.offsetWidth > 0
        ? Math.max(0.01, editorRect.width / view.dom.offsetWidth)
        : 1;
    const viewportTop = Math.max(0, (rootRect.top - editorRect.top) / scale);
    const viewportBottom = Math.max(
      viewportTop,
      (rootRect.bottom - editorRect.top) / scale,
    );
    const { end, start } = documentChunkViewportRange(
      state.geometry,
      viewportTop,
      viewportBottom,
    );
    dispatchVisibleWindow(state.geometry.slice(start, end).map(({ id }) => id));
  };
  const queuePublication = (
    publish: typeof publishGeometryWindow | typeof publishIntersectionWindow,
  ) => {
    if (publishQueued) return;
    publishQueued = true;
    queueMicrotask(() => {
      publishQueued = false;
      publish();
    });
  };
  const handleIntersections: IntersectionObserverCallback = (entries) => {
    for (const entry of entries) {
      const element = entry.target as HTMLElement;
      const id = element.dataset.documentChunkId;
      if (!id) continue;
      if (entry.isIntersecting) intersecting.add(id);
      else intersecting.delete(id);
    }
    queuePublication(publishIntersectionWindow);
  };
  const handleScroll = () => queuePublication(publishGeometryWindow);
  const disconnectObserver = () => {
    observer?.disconnect();
    observer = null;
    observed.clear();
    intersecting.clear();
  };
  const ensureFallbackObserver = (): boolean => {
    const state = documentChunkWindowKey.getState(view.state);
    if (
      scrollRoot ||
      observer ||
      !state ||
      state.fallback ||
      state.chunkCount === 0
    ) {
      return false;
    }
    observer = new IntersectionObserver(handleIntersections, {
      root: null,
      rootMargin: DOCUMENT_CHUNK_WINDOW_ROOT_MARGIN,
      threshold: 0,
    });
    return true;
  };
  const synchronizeElements = () => {
    if (!observer) return;
    for (const [id, element] of observed) {
      if (registry.controllers.get(id)?.element === element) continue;
      observer.unobserve(element);
      observed.delete(id);
      intersecting.delete(id);
    }
    for (const [id, controller] of registry.controllers) {
      const element = controller.element;
      if (!id || observed.get(id) === element) continue;
      observed.set(id, element);
      observer.observe(element);
    }
  };
  const updateDatasets = () => {
    const state = documentChunkWindowKey.getState(view.state);
    view.dom.dataset.documentWindowed = String(Boolean(state?.chunkCount));
    view.dom.dataset.documentChunkCount = String(state?.chunkCount ?? 0);
    view.dom.dataset.documentMountedChunkCount = String(
      state?.mountedIds.size ?? 0,
    );
    if (state?.fallback && state.chunkCount) {
      view.dom.dataset.documentWindowingFallback = 'true';
    } else {
      delete view.dom.dataset.documentWindowingFallback;
    }
  };
  const ensureDriver = () => {
    const nextRoot = view.dom.closest<HTMLElement>('.work-document-scroll');
    if (nextRoot) {
      if (nextRoot !== scrollRoot) {
        scrollRoot?.removeEventListener('scroll', handleScroll);
        scrollRoot = nextRoot;
        scrollRoot.addEventListener('scroll', handleScroll, { passive: true });
      }
      disconnectObserver();
      queuePublication(publishGeometryWindow);
      return;
    }
    if (scrollRoot) {
      scrollRoot.removeEventListener('scroll', handleScroll);
      scrollRoot = null;
    }
    const created = ensureFallbackObserver();
    if (created) synchronizeElements();
  };
  const synchronize = () => {
    updateDatasets();
    const state = documentChunkWindowKey.getState(view.state);
    if (state && state.geometry !== synchronizedGeometry) {
      synchronizedGeometry = state.geometry;
      synchronizeDocumentChunkPaginationHeights(registry, state.geometry);
    }
    ensureDriver();
    if (observer) synchronizeElements();
  };
  const handleElementChange = ({
    id,
    next,
    previous,
  }: DocumentChunkControllerChange) => {
    if (previous && observed.get(id) === previous.element) {
      observer?.unobserve(previous.element);
      observed.delete(id);
      intersecting.delete(id);
    }
    if (next && observer) {
      observed.set(id, next.element);
      observer.observe(next.element);
    }
  };

  registry.listeners.add(handleElementChange);
  synchronize();
  queueMicrotask(synchronize);
  if (typeof requestAnimationFrame === 'function') {
    mountFrame = requestAnimationFrame(() => {
      mountFrame = 0;
      synchronize();
    });
  }
  return {
    destroy() {
      if (mountFrame) cancelAnimationFrame(mountFrame);
      registry.listeners.delete(handleElementChange);
      scrollRoot?.removeEventListener('scroll', handleScroll);
      scrollRoot = null;
      disconnectObserver();
    },
    update(nextView: EditorView) {
      view = nextView;
      synchronize();
    },
  };
}

export function documentLazyChunkContentForEditor(
  editor: Editor,
  chunkId: string,
) {
  const content = documentChunkContentSources.get(editor)?.() ?? null;
  return documentLazyChunkContent(content?.model, chunkId);
}

export function documentTransactionsOnlyHydrateChunks(
  transactions: readonly Transaction[],
): boolean {
  const changed = transactions.filter((transaction) => transaction.docChanged);
  return (
    changed.length > 0 &&
    changed.every((transaction) =>
      Boolean(transaction.getMeta(DOCUMENT_CHUNK_HYDRATION_META)),
    )
  );
}

function hydrateDocumentChunks(
  editor: Editor,
  transaction: Transaction,
  state: EditorState,
  requestedIds: Iterable<string>,
): string[] {
  const previousDocument = transaction.doc;
  const candidates = Array.from(new Set(requestedIds))
    .map((id) => ({
      id,
      position: documentChunkWindowKey.getState(state)?.positions.get(id),
    }))
    .filter(
      (
        candidate,
      ): candidate is {
        id: string;
        position: DocumentChunkPosition;
      } => Boolean(candidate.position),
    )
    .sort((left, right) => right.position.from - left.position.from);
  const hydratedIds: string[] = [];
  const originalAnchor = transaction.selection.anchor;
  const originalHead = transaction.selection.head;

  for (const { id, position } of candidates) {
    const chunk = transaction.doc.nodeAt(position.from);
    if (
      !chunk ||
      chunk.type.name !== 'documentChunk' ||
      chunk.childCount !== 1 ||
      chunk.firstChild?.type.name !== DOCUMENT_LAZY_BLOCK_NODE
    ) {
      continue;
    }
    const payload = documentLazyChunkContentForEditor(editor, id);
    if (!payload?.length) continue;
    let content: Fragment;
    try {
      content = Fragment.fromArray(
        payload.map((node) => transaction.doc.type.schema.nodeFromJSON(node)),
      );
    } catch (error) {
      throw new Error(`The lazy document chunk "${id}" cannot be hydrated.`, {
        cause: error,
      });
    }
    if (content.size !== chunk.content.size) {
      throw new Error(
        `The lazy document chunk "${id}" changed its logical size during hydration.`,
      );
    }
    transaction.replaceWith(position.from + 1, position.to - 1, content);
    hydratedIds.push(id);
  }

  if (!hydratedIds.length) return hydratedIds;
  transferDocumentTextStatistics(previousDocument, transaction.doc);
  transaction
    .setMeta(DOCUMENT_CHUNK_HYDRATION_META, {
      ids: hydratedIds,
    } satisfies DocumentChunkHydrationMeta)
    .setMeta('addToHistory', false)
    .setMeta('preventUpdate', true);
  const maximum = transaction.doc.content.size;
  if (originalAnchor <= maximum && originalHead <= maximum) {
    const anchor = transaction.doc.resolve(originalAnchor);
    const head = transaction.doc.resolve(originalHead);
    if (anchor.parent.inlineContent && head.parent.inlineContent) {
      transaction.setSelection(
        TextSelection.create(transaction.doc, originalAnchor, originalHead),
      );
    }
  }
  return hydratedIds;
}

function documentChunkDecorations(
  document: ProseMirrorNode,
  mountedIds: ReadonlySet<string>,
  positions: ReadonlyMap<string, DocumentChunkPosition>,
): DecorationSet {
  const decoratedIds = new Set<string>();
  for (const id of mountedIds) {
    const position = positions.get(id);
    if (!position) continue;
    decoratedIds.add(id);
    for (const ancestorId of position.ancestorIds) {
      decoratedIds.add(ancestorId);
    }
  }
  const decorations = Array.from(decoratedIds).flatMap((id) => {
    const position = positions.get(id);
    return position
      ? [
          Decoration.node(
            position.from,
            position.to,
            {},
            { documentChunkMounted: true },
          ),
        ]
      : [];
  });
  return DecorationSet.create(document, decorations);
}

export function selectedDocumentChunkId(
  state: Pick<EditorState, 'doc' | 'selection'>,
): string | null {
  const resolved = state.doc.resolve(state.selection.anchor);
  for (let depth = resolved.depth; depth >= 0; depth -= 1) {
    const node = resolved.node(depth);
    if (node.type.name === 'documentChunk') return documentChunkId(node);
  }
  return null;
}

export function documentChunkMountedIds(
  state: EditorState,
): readonly string[] | null {
  const chunkState = documentChunkWindowKey.getState(state);
  return chunkState?.chunkCount ? [...chunkState.mountedIds] : null;
}

function documentChunkIsMounted(
  _node: ProseMirrorNode,
  decorations: readonly Decoration[],
): boolean {
  return decorations.some(
    (decoration) => decoration.spec.documentChunkMounted === true,
  );
}

function documentChunkElementRegistryFor(
  editor: Editor,
): DocumentChunkNodeViewRegistry {
  let registry = documentChunkElementRegistries.get(editor);
  if (!registry) {
    registry = {
      controllers: new Map(),
      lazyPreviewPool: createDocumentLazyPreviewPool(),
      listeners: new Set(),
      paginationExtraHeights: new Map(),
    };
    documentChunkElementRegistries.set(editor, registry);
  }
  return registry;
}

function registerDocumentChunkController(
  editor: Editor,
  id: string,
  controller: DocumentChunkNodeViewController,
): () => void {
  const registry = documentChunkElementRegistryFor(editor);
  const previous = registry.controllers.get(id);
  registry.controllers.set(id, controller);
  const added = { id, next: controller, previous };
  for (const listener of registry.listeners) listener(added);
  return () => {
    if (registry.controllers.get(id) !== controller) return;
    registry.controllers.delete(id);
    const removed = { id, previous: controller };
    for (const listener of registry.listeners) listener(removed);
  };
}

function synchronizeDocumentChunkPaginationHeights(
  registry: DocumentChunkNodeViewRegistry,
  geometry: readonly DocumentChunkGeometry[],
): void {
  const heights = new Map<string, number>();
  for (const chunk of geometry) {
    heights.set(chunk.id, chunk.paginationExtraHeight);
    if (chunk.containerId) {
      heights.set(
        chunk.containerId,
        (heights.get(chunk.containerId) ?? 0) + chunk.paginationExtraHeight,
      );
    }
  }
  const nextIds = new Set<string>();
  for (const [id, height] of heights) {
    nextIds.add(id);
    registry.paginationExtraHeights.set(id, height);
    registry.controllers.get(id)?.setPaginationExtraHeight(height);
  }
  for (const id of [...registry.paginationExtraHeights.keys()]) {
    if (nextIds.has(id)) continue;
    registry.paginationExtraHeights.delete(id);
    registry.controllers.get(id)?.setPaginationExtraHeight(0);
  }
}

function documentChunkDomAttributes(
  node: ProseMirrorNode,
): Record<string, string> {
  return {
    'data-document-chunk': 'true',
    'data-document-chunk-id': documentChunkId(node),
    'data-document-chunk-blocks': String(
      positiveInteger(node.attrs.blockCount, node.childCount),
    ),
    'data-document-chunk-estimated-height': String(
      positiveNumber(node.attrs.estimatedHeight, 1),
    ),
    'data-document-chunk-window-container': String(
      node.attrs.windowContainer === true,
    ),
  };
}

function applyDocumentChunkNodeViewAttributes(
  dom: HTMLElement,
  node: ProseMirrorNode,
  mounted: boolean,
  paginationExtraHeight: number,
): void {
  for (const [name, value] of Object.entries(
    documentChunkDomAttributes(node),
  )) {
    dom.setAttribute(name, value);
  }
  dom.dataset.documentChunkMounted = String(mounted);
  dom.dataset.documentChunkPaginationExtraHeight = String(
    paginationExtraHeight,
  );
  dom.className = `work-document-chunk ${mounted ? 'mounted' : 'placeholder'}`;
  if (mounted) {
    dom.removeAttribute('aria-hidden');
    dom.removeAttribute('contenteditable');
    dom.style.removeProperty('height');
    dom.style.removeProperty('contain');
    return;
  }
  dom.setAttribute('aria-hidden', 'true');
  dom.setAttribute('contenteditable', 'false');
  dom.style.height = `${
    positiveNumber(node.attrs.estimatedHeight, 1) + paginationExtraHeight
  }px`;
  dom.style.contain = 'strict';
}

function setsEqual(
  left: ReadonlySet<string>,
  right: ReadonlySet<string>,
): boolean {
  if (left.size !== right.size) return false;
  for (const value of left) if (!right.has(value)) return false;
  return true;
}

function positiveInteger(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : fallback;
}

function positiveNumber(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

function roundDuration(value: number): number {
  return Math.round(value * 10) / 10;
}
