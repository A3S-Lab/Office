import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

export interface DocumentChunkGeometry {
  containerId?: string;
  end: number;
  from: number;
  id: string;
  paginationExtraHeight: number;
  start: number;
  to: number;
}

export interface DocumentChunkPosition {
  ancestorIds: readonly string[];
  from: number;
  to: number;
}

export function documentChunkViewportRange(
  geometry: readonly { end: number; start: number }[],
  viewportTop: number,
  viewportBottom: number,
): { end: number; start: number } {
  if (!geometry.length) return { end: 0, start: 0 };
  const first = firstDocumentChunkEndingAfter(geometry, viewportTop);
  const last = lastDocumentChunkStartingBefore(geometry, viewportBottom);
  return {
    end: Math.min(geometry.length, Math.max(first + 1, last + 1)),
    start: Math.max(0, first),
  };
}

export function documentChunkGeometry(
  document: ProseMirrorNode,
  breaks: readonly { position: number; spacerHeight: number }[],
): DocumentChunkGeometry[] {
  const geometry: DocumentChunkGeometry[] = [];
  const orderedBreaks = [...breaks].sort(
    (left, right) => left.position - right.position,
  );
  let breakIndex = 0;
  let top = 0;
  document.descendants((node, position, parent) => {
    if (node.type.name !== 'documentChunk') return true;
    if (node.attrs.windowContainer === true) return true;
    const nodeEnd = position + node.nodeSize;
    while (
      breakIndex < orderedBreaks.length &&
      (orderedBreaks[breakIndex]?.position ?? 0) <= position
    ) {
      breakIndex += 1;
    }
    let chunkBreakIndex = breakIndex;
    let paginationExtraHeight = 0;
    while (
      chunkBreakIndex < orderedBreaks.length &&
      (orderedBreaks[chunkBreakIndex]?.position ?? nodeEnd) < nodeEnd
    ) {
      paginationExtraHeight += Math.max(
        0,
        orderedBreaks[chunkBreakIndex]?.spacerHeight ?? 0,
      );
      chunkBreakIndex += 1;
    }
    breakIndex = chunkBreakIndex;
    const height =
      positiveNumber(node.attrs.estimatedHeight, 1) + paginationExtraHeight;
    geometry.push({
      ...(parent?.type.name === 'documentChunk' &&
      parent.attrs.windowContainer === true
        ? { containerId: documentChunkId(parent) }
        : {}),
      end: top + height,
      from: position,
      id: documentChunkId(node),
      paginationExtraHeight,
      start: top,
      to: nodeEnd,
    });
    top += height;
    return false;
  });
  return geometry;
}

export function documentChunkPositions(
  document: ProseMirrorNode,
): ReadonlyMap<string, DocumentChunkPosition> {
  const positions = new Map<string, DocumentChunkPosition>();
  const ancestorIdsByNode = new WeakMap<ProseMirrorNode, readonly string[]>();
  document.descendants((node, position, parent) => {
    const parentAncestorIds = parent
      ? (ancestorIdsByNode.get(parent) ?? [])
      : [];
    if (node.type.name !== 'documentChunk') {
      ancestorIdsByNode.set(node, parentAncestorIds);
      return true;
    }
    const ancestorIds =
      parent?.type.name === 'documentChunk'
        ? [...parentAncestorIds, documentChunkId(parent)]
        : parentAncestorIds;
    const id = documentChunkId(node);
    positions.set(id, {
      ancestorIds,
      from: position,
      to: position + node.nodeSize,
    });
    ancestorIdsByNode.set(node, ancestorIds);
    return node.attrs.windowContainer === true;
  });
  return positions;
}

export function documentChunkId(node: ProseMirrorNode): string {
  return typeof node.attrs.id === 'string' && node.attrs.id
    ? node.attrs.id
    : `document-chunk-${node.nodeSize}`;
}

function firstDocumentChunkEndingAfter(
  geometry: readonly { end: number }[],
  offset: number,
): number {
  let lower = 0;
  let upper = geometry.length - 1;
  let result = upper;
  while (lower <= upper) {
    const middle = Math.floor((lower + upper) / 2);
    const chunk = geometry[middle];
    if (chunk && chunk.end >= offset) {
      result = middle;
      upper = middle - 1;
    } else {
      lower = middle + 1;
    }
  }
  return result;
}

function lastDocumentChunkStartingBefore(
  geometry: readonly { start: number }[],
  offset: number,
): number {
  let lower = 0;
  let upper = geometry.length - 1;
  let result = 0;
  while (lower <= upper) {
    const middle = Math.floor((lower + upper) / 2);
    const chunk = geometry[middle];
    if (chunk && chunk.start <= offset) {
      result = middle;
      lower = middle + 1;
    } else {
      upper = middle - 1;
    }
  }
  return result;
}

function positiveNumber(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}
