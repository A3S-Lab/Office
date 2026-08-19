import type { WorkDocumentModel, WorkDocumentNode } from './work-types';
import {
  createDocumentLazyHtmlProjection,
  type DocumentLazyHtmlProjection,
} from './work-document-lazy-html';

export const DOCUMENT_LAZY_BLOCK_NODE = 'documentLazyBlock';
export const DOCUMENT_LAZY_INITIAL_CHUNK_COUNT = 2;

export const DOCUMENT_LAZY_POSITION_BOUNDARY = '\ufffc';
const SIMPLE_LAZY_CONTAINER_TYPES = new Set([
  'paragraph',
  'table',
  'tableCell',
  'tableRow',
]);

export interface PreparedLazyDocumentEditorSource {
  lazyChunkCount: number;
  payloads: ReadonlyMap<string, readonly WorkDocumentNode[]>;
  root: WorkDocumentNode;
}

interface MutablePreparedLazyDocumentEditorSource {
  htmlProjection: DocumentLazyHtmlProjection | null;
  lazyChunkCount: number;
  payloads: Map<string, readonly WorkDocumentNode[]>;
  root: WorkDocumentNode | null;
}

const preparedModels = new WeakMap<
  WorkDocumentModel,
  MutablePreparedLazyDocumentEditorSource
>();

/**
 * Builds a compact editor-only projection for parser-authenticated plain DOCX
 * models. Every lazy leaf retains the exact node size and text offsets of its
 * complete subtree, while the canonical WorkDocumentModel remains untouched.
 */
export function prepareLazyDocumentEditorSource(
  model: WorkDocumentModel,
  allowCreate: boolean,
  html?: string,
): PreparedLazyDocumentEditorSource | null {
  const startedAt = lazyDocumentNow();
  const cacheHit = preparedModels.has(model);
  let prepared = preparedModels.get(model);
  if (!prepared) {
    if (!allowCreate) return null;
    const created = createPreparedLazyDocumentEditorSource(model.root);
    if (!created) return null;
    prepared = created;
    preparedModels.set(model, created);
  } else if (!prepared.root) {
    const htmlProjection = prepared.htmlProjection;
    const rebuilt = createPreparedLazyDocumentEditorSource(
      model.root,
      prepared.payloads,
    );
    if (!rebuilt) return null;
    rebuilt.htmlProjection = htmlProjection;
    prepared = rebuilt;
    preparedModels.set(model, rebuilt);
  }
  if (!prepared.htmlProjection && html) {
    prepared.htmlProjection = createDocumentLazyHtmlProjection(
      html,
      model.root,
    );
  }
  if (!prepared.root) return null;
  const result = {
    lazyChunkCount: prepared.lazyChunkCount,
    payloads: prepared.payloads,
    root: prepared.root,
  };
  recordLazyDocumentMeasure(
    'a3s-office.document.lazy-editor-source',
    startedAt,
    lazyDocumentNow(),
    { cacheHit },
  );
  return result;
}

export function documentLazyHtmlProjection(
  model: WorkDocumentModel | null | undefined,
): DocumentLazyHtmlProjection | null {
  return model ? (preparedModels.get(model)?.htmlProjection ?? null) : null;
}

export function invalidateDocumentLazyHtmlProjection(
  model: WorkDocumentModel | null | undefined,
): void {
  if (model) {
    const prepared = preparedModels.get(model);
    if (prepared) prepared.htmlProjection = null;
  }
}

export function documentLazyChunkContent(
  model: WorkDocumentModel | null | undefined,
  chunkId: string,
): readonly WorkDocumentNode[] | null {
  if (!model || !chunkId) return null;
  return preparedModels.get(model)?.payloads.get(chunkId) ?? null;
}

/**
 * Replaces editor-only lazy blocks with their canonical structured payloads.
 * Hydrated leaves update the payload registry so subsequent remounts retain
 * user edits without forcing every untouched subtree through ProseMirror.
 */
export function materializeLazyDocumentEditorRoot(
  root: WorkDocumentNode,
  model: WorkDocumentModel | null | undefined,
): WorkDocumentNode {
  const prepared = model ? preparedModels.get(model) : null;
  if (!prepared) return root;

  const visit = (node: WorkDocumentNode): WorkDocumentNode => {
    if (node.type === 'documentChunk' && node.attrs?.windowContainer !== true) {
      const id = documentChunkId(node);
      if (!id) return node;
      if (documentChunkIsLazy(node)) {
        const payload = prepared.payloads.get(id);
        if (!payload) {
          throw new Error(
            `The lazy document chunk payload "${id}" is missing.`,
          );
        }
        return { ...node, content: [...payload] };
      }
      if (node.content?.length) prepared.payloads.set(id, node.content);
      return node;
    }
    if (!node.content?.length) return node;
    let changed = false;
    const content = node.content.map((child) => {
      const next = visit(child);
      if (next !== child) changed = true;
      return next;
    });
    return changed ? { ...node, content } : node;
  };

  return visit(root);
}

/** Keeps the process-local lazy payload registry across controlled revisions. */
export function transferLazyDocumentModelState(
  previous: WorkDocumentModel | null | undefined,
  next: WorkDocumentModel,
): void {
  if (!previous) return;
  const prepared = preparedModels.get(previous);
  if (!prepared) return;
  preparedModels.set(next, {
    htmlProjection: prepared.htmlProjection,
    lazyChunkCount: prepared.lazyChunkCount,
    payloads: prepared.payloads,
    root: null,
  });
}

export function documentNodeUsesLazyPayload(node: WorkDocumentNode): boolean {
  return documentChunkIsLazy(node);
}

export function simpleDocumentNodeSize(node: WorkDocumentNode): number | null {
  if (node.type === 'text') {
    return typeof node.text === 'string' && node.text ? node.text.length : null;
  }
  if (!SIMPLE_LAZY_CONTAINER_TYPES.has(node.type)) return null;
  let size = 2;
  for (const child of node.content ?? []) {
    const childSize = simpleDocumentNodeSize(child);
    if (childSize === null) return null;
    size += childSize;
  }
  return size;
}

function createPreparedLazyDocumentEditorSource(
  root: WorkDocumentNode,
  previousPayloads?: ReadonlyMap<string, readonly WorkDocumentNode[]>,
): MutablePreparedLazyDocumentEditorSource | null {
  const payloads = new Map<string, readonly WorkDocumentNode[]>();
  let leafIndex = 0;
  let lazyChunkCount = 0;
  let unsupported = false;

  const visit = (node: WorkDocumentNode): WorkDocumentNode => {
    if (node.type === 'documentChunk' && node.attrs?.windowContainer !== true) {
      const id = documentChunkId(node);
      const payload = id
        ? (previousPayloads?.get(id) ?? node.content ?? [])
        : [];
      if (!id || !payload.length) {
        unsupported = true;
        return node;
      }
      payloads.set(id, payload);
      const currentIndex = leafIndex;
      leafIndex += 1;
      if (currentIndex < DOCUMENT_LAZY_INITIAL_CHUNK_COUNT) return node;
      const placeholder = lazyPlaceholderForContent(id, payload);
      if (!placeholder) {
        unsupported = true;
        return node;
      }
      lazyChunkCount += 1;
      return { ...node, content: [placeholder] };
    }
    if (!node.content?.length) return node;
    let changed = false;
    const content = node.content.map((child) => {
      const next = visit(child);
      if (next !== child) changed = true;
      return next;
    });
    return changed ? { ...node, content } : node;
  };

  const compactRoot = visit(root);
  if (
    unsupported ||
    leafIndex <= DOCUMENT_LAZY_INITIAL_CHUNK_COUNT ||
    lazyChunkCount === 0
  ) {
    return null;
  }
  return {
    htmlProjection: null,
    lazyChunkCount,
    payloads,
    root: compactRoot,
  };
}

function lazyPlaceholderForContent(
  chunkId: string,
  content: readonly WorkDocumentNode[],
): WorkDocumentNode | null {
  const tapeParts: string[] = [];
  const statistics = { contentSize: 0, paragraphCount: 0 };
  for (const node of content) {
    if (!appendDocumentNodePositionTape(node, tapeParts, statistics)) {
      return null;
    }
  }
  if (
    statistics.contentSize < 2 ||
    tapeParts[0] !== DOCUMENT_LAZY_POSITION_BOUNDARY ||
    tapeParts.at(-1) !== DOCUMENT_LAZY_POSITION_BOUNDARY
  ) {
    return null;
  }
  tapeParts[0] = '';
  tapeParts[tapeParts.length - 1] = '';
  const filler = tapeParts.join('');
  return {
    type: DOCUMENT_LAZY_BLOCK_NODE,
    attrs: {
      chunkId,
      contentSize: statistics.contentSize,
      paragraphCount: statistics.paragraphCount,
    },
    ...(filler
      ? { content: [{ type: 'text', text: filler }] }
      : { content: [] }),
  };
}

function appendDocumentNodePositionTape(
  node: WorkDocumentNode,
  parts: string[],
  statistics: { contentSize: number; paragraphCount: number },
): boolean {
  if (node.type === 'text') {
    if (typeof node.text !== 'string' || !node.text) return false;
    parts.push(node.text);
    statistics.contentSize += node.text.length;
    return true;
  }
  if (!SIMPLE_LAZY_CONTAINER_TYPES.has(node.type)) return false;
  if (node.type === 'paragraph') statistics.paragraphCount += 1;
  statistics.contentSize += 2;
  parts.push(DOCUMENT_LAZY_POSITION_BOUNDARY);
  for (const child of node.content ?? []) {
    if (!appendDocumentNodePositionTape(child, parts, statistics)) return false;
  }
  parts.push(DOCUMENT_LAZY_POSITION_BOUNDARY);
  return true;
}

function documentChunkIsLazy(node: WorkDocumentNode): boolean {
  return (
    node.type === 'documentChunk' &&
    node.attrs?.windowContainer !== true &&
    node.content?.length === 1 &&
    node.content[0]?.type === DOCUMENT_LAZY_BLOCK_NODE
  );
}

function documentChunkId(node: WorkDocumentNode): string | null {
  const id = node.attrs?.id;
  return typeof id === 'string' && id ? id : null;
}

function lazyDocumentNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function recordLazyDocumentMeasure(
  name: string,
  start: number,
  end: number,
  detail?: Record<string, unknown>,
): void {
  try {
    globalThis.performance?.measure(name, { detail, end, start });
  } catch {
    // User Timing diagnostics must never affect editor source preparation.
  }
}
