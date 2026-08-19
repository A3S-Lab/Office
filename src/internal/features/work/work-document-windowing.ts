import type { WorkDocumentNode } from './work-types';

export const DOCUMENT_WINDOW_BLOCK_SIZE = 128;
export const DOCUMENT_WINDOW_BLOCK_THRESHOLD = 2_048;
export const DOCUMENT_TABLE_WINDOW_ROW_SIZE = 16;
export const DOCUMENT_TABLE_WINDOW_ROW_THRESHOLD = 512;
const DOCUMENT_TABLE_WINDOW_GROUP_SIZE = 32;

const DEFAULT_PARAGRAPH_LINE_HEIGHT = 21;
const DEFAULT_PARAGRAPH_AFTER = 8;
const DEFAULT_TABLE_ROW_HEIGHT = 22;
const DOCUMENT_TEXT_CHARACTERS_PER_LINE = 88;

export interface DocumentWindowingOptions {
  blockSize: number;
  blockThreshold: number;
  tableRowSize: number;
  tableRowThreshold: number;
  trustedIntegrityFeatures?: number;
}

const DEFAULT_DOCUMENT_WINDOWING_OPTIONS: DocumentWindowingOptions = {
  blockSize: DOCUMENT_WINDOW_BLOCK_SIZE,
  blockThreshold: DOCUMENT_WINDOW_BLOCK_THRESHOLD,
  tableRowSize: DOCUMENT_TABLE_WINDOW_ROW_SIZE,
  tableRowThreshold: DOCUMENT_TABLE_WINDOW_ROW_THRESHOLD,
};

/**
 * Adds stable semantic chunks to large section bodies. The complete
 * ProseMirror model remains in memory, while the chunk node view can omit
 * descendants outside the visible viewport without changing positions.
 */
export function windowDocumentModel(
  root: WorkDocumentNode,
  requested: Partial<DocumentWindowingOptions> = {},
): WorkDocumentNode {
  if (documentSectionModelUsesWindowing(root)) return root;
  const options = normalizeWindowingOptions(requested);
  let changed = false;
  const content = root.content?.map((node, sectionIndex) => {
    if (node.type !== 'documentSection' || !node.content?.length) return node;
    const needsBlockWindow = node.content.length >= options.blockThreshold;
    const needsTableWindow = node.content.some(
      (child) =>
        child.type === 'table' &&
        (child.content?.length ?? 0) >= options.tableRowThreshold,
    );
    if (!needsBlockWindow && !needsTableWindow) return node;
    changed = true;
    return {
      ...node,
      content: windowSectionContent(node.content, sectionIndex, options),
    };
  });
  return changed ? { ...root, content } : root;
}

/** Restores the canonical schema used by controlled HTML and DOCX export. */
export function materializeWindowedDocumentModel(
  root: WorkDocumentNode,
): WorkDocumentNode {
  if (!documentModelUsesWindowing(root)) return root;
  return {
    ...root,
    content: root.content?.map((node) =>
      node.type === 'documentSection'
        ? materializeWindowedSection(node)
        : materializeNestedWindowing(node),
    ),
  };
}

export function documentModelUsesWindowing(root: WorkDocumentNode): boolean {
  const pending = [root];
  while (pending.length) {
    const node = pending.pop();
    if (!node) continue;
    if (
      node.type === 'documentChunk' ||
      (node.type === 'table' && typeof node.attrs?.virtualTableId === 'string')
    ) {
      return true;
    }
    if (node.content) pending.push(...node.content);
  }
  return false;
}

/**
 * `windowDocumentModel` only inserts chunks directly beneath a section. Check
 * that invariant at the ownership boundary instead of traversing every text
 * node in a giant, known-unwindowed import before doing the actual work.
 */
function documentSectionModelUsesWindowing(root: WorkDocumentNode): boolean {
  for (const section of root.content ?? []) {
    if (
      section.type === 'documentSection' &&
      section.content?.some(
        (node) =>
          node.type === 'documentChunk' ||
          (node.type === 'table' &&
            typeof node.attrs?.virtualTableId === 'string'),
      )
    ) {
      return true;
    }
  }
  return false;
}

export function estimateDocumentNodeHeight(node: WorkDocumentNode): number {
  if (node.type === 'tableRow') return estimateDocumentTableRowHeight(node);
  if (node.type === 'table') {
    const rows = node.content ?? [];
    const rowsHeight = rows.reduce(
      (height, row) => height + estimateDocumentTableRowHeight(row),
      0,
    );
    return Math.max(
      DEFAULT_TABLE_ROW_HEIGHT,
      rowsHeight + (node.attrs?.officeImported ? 0 : 40),
    );
  }
  if (node.type === 'image') {
    const height = finitePositiveNumber(node.attrs?.height);
    return Math.max(DEFAULT_PARAGRAPH_LINE_HEIGHT, (height ?? 120) + 36);
  }
  if (node.type === 'horizontalRule' || node.type === 'pageBreak') return 38;
  if (node.type === 'heading') {
    const level = finitePositiveNumber(node.attrs?.level) ?? 1;
    const lineHeight = level <= 1 ? 32 : level === 2 ? 28 : 25;
    return estimatedDocumentNodeTextLines(node) * lineHeight + 18;
  }
  if (node.type === 'paragraph') {
    return (
      estimatedDocumentNodeTextLines(node) * DEFAULT_PARAGRAPH_LINE_HEIGHT +
      DEFAULT_PARAGRAPH_AFTER
    );
  }
  if (node.type === 'bulletList' || node.type === 'orderedList') {
    return Math.max(
      DEFAULT_PARAGRAPH_LINE_HEIGHT + DEFAULT_PARAGRAPH_AFTER,
      (node.content ?? []).reduce(
        (height, child) => height + estimateDocumentNodeHeight(child),
        0,
      ),
    );
  }
  if (node.type === 'listItem' || node.type === 'blockquote') {
    return Math.max(
      DEFAULT_PARAGRAPH_LINE_HEIGHT + DEFAULT_PARAGRAPH_AFTER,
      (node.content ?? []).reduce(
        (height, child) => height + estimateDocumentNodeHeight(child),
        0,
      ),
    );
  }
  if (node.type === 'documentChunk') {
    return (
      finitePositiveNumber(node.attrs?.estimatedHeight) ??
      (node.content ?? []).reduce(
        (height, child) => height + estimateDocumentNodeHeight(child),
        0,
      )
    );
  }
  const childrenHeight = (node.content ?? []).reduce(
    (height, child) => height + estimateDocumentNodeHeight(child),
    0,
  );
  return Math.max(
    DEFAULT_PARAGRAPH_LINE_HEIGHT + DEFAULT_PARAGRAPH_AFTER,
    childrenHeight,
  );
}

export function estimateDocumentTableRowHeight(row: WorkDocumentNode): number {
  const explicit = finitePositiveNumber(row.attrs?.rowHeight);
  let maximumLines = 1;
  for (const cell of row.content ?? []) {
    maximumLines = Math.max(maximumLines, estimatedDocumentNodeTextLines(cell));
  }
  const contentHeight = maximumLines * DEFAULT_PARAGRAPH_LINE_HEIGHT + 1;
  return Math.max(DEFAULT_TABLE_ROW_HEIGHT, explicit ?? 0, contentHeight);
}

function windowSectionContent(
  source: readonly WorkDocumentNode[],
  sectionIndex: number,
  options: DocumentWindowingOptions,
): WorkDocumentNode[] {
  const logicalBlocks = source.flatMap((node, blockIndex) =>
    splitLargeDocumentTable(node, sectionIndex, blockIndex, options),
  );
  const chunks: WorkDocumentNode[] = [];
  let pending: WorkDocumentNode[] = [];

  const flush = () => {
    if (!pending.length) return;
    const chunkIndex = chunks.length;
    chunks.push({
      type: 'documentChunk',
      attrs: {
        id: `document-chunk-${sectionIndex + 1}-${chunkIndex + 1}`,
        blockCount: pending.length,
        estimatedHeight: roundHeight(
          pending.reduce(
            (height, node) => height + estimateDocumentNodeHeight(node),
            0,
          ),
        ),
        ...(options.trustedIntegrityFeatures !== undefined
          ? { integrityFeatures: options.trustedIntegrityFeatures }
          : {}),
      },
      content: pending,
    });
    pending = [];
  };

  for (const node of logicalBlocks) {
    if (typeof node.attrs?.virtualTableId === 'string') {
      flush();
      pending = [node];
      flush();
      continue;
    }
    pending.push(node);
    if (pending.length >= options.blockSize) flush();
  }
  flush();
  return groupVirtualTableChunks(chunks, sectionIndex);
}

function groupVirtualTableChunks(
  chunks: WorkDocumentNode[],
  sectionIndex: number,
): WorkDocumentNode[] {
  const grouped: WorkDocumentNode[] = [];
  let index = 0;
  while (index < chunks.length) {
    const tableId = virtualTableChunkId(chunks[index]);
    if (!tableId) {
      grouped.push(chunks[index] as WorkDocumentNode);
      index += 1;
      continue;
    }
    let end = index + 1;
    while (
      end < chunks.length &&
      virtualTableChunkId(chunks[end]) === tableId
    ) {
      end += 1;
    }
    const run = chunks.slice(index, end);
    if (run.length <= DOCUMENT_TABLE_WINDOW_GROUP_SIZE) {
      grouped.push(...run);
      index = end;
      continue;
    }
    for (
      let offset = 0;
      offset < run.length;
      offset += DOCUMENT_TABLE_WINDOW_GROUP_SIZE
    ) {
      const leaves = run.slice(
        offset,
        offset + DOCUMENT_TABLE_WINDOW_GROUP_SIZE,
      );
      const groupIndex = Math.floor(offset / DOCUMENT_TABLE_WINDOW_GROUP_SIZE);
      const integrityFeatures = combinedDocumentChunkIntegrityFeatures(leaves);
      grouped.push({
        type: 'documentChunk',
        attrs: {
          id: `document-chunk-${sectionIndex + 1}-table-group-${grouped.length + 1}-${groupIndex + 1}`,
          blockCount: leaves.reduce(
            (count, leaf) => count + positiveInteger(leaf.attrs?.blockCount, 1),
            0,
          ),
          estimatedHeight: roundHeight(
            leaves.reduce(
              (height, leaf) =>
                height +
                (finitePositiveNumber(leaf.attrs?.estimatedHeight) ?? 1),
              0,
            ),
          ),
          windowContainer: true,
          ...(integrityFeatures !== null ? { integrityFeatures } : {}),
        },
        content: leaves,
      });
    }
    index = end;
  }
  return grouped;
}

function virtualTableChunkId(
  node: WorkDocumentNode | undefined,
): string | null {
  const table = node?.type === 'documentChunk' ? node.content?.[0] : undefined;
  return table?.type === 'table' &&
    typeof table.attrs?.virtualTableId === 'string'
    ? table.attrs.virtualTableId
    : null;
}

function splitLargeDocumentTable(
  node: WorkDocumentNode,
  sectionIndex: number,
  blockIndex: number,
  options: DocumentWindowingOptions,
): WorkDocumentNode[] {
  const rows = node.content ?? [];
  if (node.type !== 'table' || rows.length < options.tableRowThreshold) {
    return [node];
  }
  const sliceCount = Math.ceil(rows.length / options.tableRowSize);
  const tableId = `document-table-${sectionIndex + 1}-${blockIndex + 1}`;
  return Array.from({ length: sliceCount }, (_, index) => ({
    ...node,
    attrs: {
      ...(node.attrs ?? {}),
      virtualTableId: tableId,
      virtualTableIndex: index,
      virtualTableCount: sliceCount,
    },
    content: rows.slice(
      index * options.tableRowSize,
      (index + 1) * options.tableRowSize,
    ),
  }));
}

function materializeWindowedSection(
  section: WorkDocumentNode,
): WorkDocumentNode {
  const flattened = flattenDocumentChunks(section.content ?? []);
  const content: WorkDocumentNode[] = [];
  for (const node of flattened) {
    const tableId =
      node.type === 'table' && typeof node.attrs?.virtualTableId === 'string'
        ? node.attrs.virtualTableId
        : null;
    const previous = content.at(-1);
    if (
      tableId &&
      previous?.type === 'table' &&
      previous.attrs?.virtualTableId === tableId
    ) {
      previous.content = [...(previous.content ?? []), ...(node.content ?? [])];
      continue;
    }
    content.push(node);
  }
  return {
    ...section,
    content: content.map(stripVirtualTableAttributes),
  };
}

function flattenDocumentChunks(
  nodes: readonly WorkDocumentNode[],
): WorkDocumentNode[] {
  return nodes.flatMap((node) =>
    node.type === 'documentChunk'
      ? flattenDocumentChunks(node.content ?? [])
      : [materializeNestedWindowing(node)],
  );
}

function materializeNestedWindowing(node: WorkDocumentNode): WorkDocumentNode {
  if (!node.content?.length) return node;
  return {
    ...node,
    content: node.content.map(materializeNestedWindowing),
  };
}

function stripVirtualTableAttributes(node: WorkDocumentNode): WorkDocumentNode {
  if (node.type !== 'table' || !node.attrs?.virtualTableId) return node;
  const {
    virtualTableId: _virtualTableId,
    virtualTableIndex: _virtualTableIndex,
    virtualTableCount: _virtualTableCount,
    ...attrs
  } = node.attrs;
  return { ...node, attrs };
}

function normalizeWindowingOptions(
  requested: Partial<DocumentWindowingOptions>,
): DocumentWindowingOptions {
  const trustedIntegrityFeatures = nonNegativeInteger(
    requested.trustedIntegrityFeatures,
  );
  return {
    blockSize: positiveInteger(
      requested.blockSize,
      DEFAULT_DOCUMENT_WINDOWING_OPTIONS.blockSize,
    ),
    blockThreshold: positiveInteger(
      requested.blockThreshold,
      DEFAULT_DOCUMENT_WINDOWING_OPTIONS.blockThreshold,
    ),
    tableRowSize: positiveInteger(
      requested.tableRowSize,
      DEFAULT_DOCUMENT_WINDOWING_OPTIONS.tableRowSize,
    ),
    tableRowThreshold: positiveInteger(
      requested.tableRowThreshold,
      DEFAULT_DOCUMENT_WINDOWING_OPTIONS.tableRowThreshold,
    ),
    ...(trustedIntegrityFeatures !== null ? { trustedIntegrityFeatures } : {}),
  };
}

function combinedDocumentChunkIntegrityFeatures(
  chunks: readonly WorkDocumentNode[],
): number | null {
  let features = 0;
  for (const chunk of chunks) {
    const value = nonNegativeInteger(chunk.attrs?.integrityFeatures);
    if (value === null) return null;
    features |= value;
  }
  return features;
}

function positiveInteger(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0 ? number : fallback;
}

function nonNegativeInteger(value: unknown): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : null;
}

function finitePositiveNumber(value: unknown): number | null {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}

/**
 * Estimates wrapped lines without first concatenating every descendant text
 * node. Giant imported tables call this once per cell, so constructing a
 * temporary string and split array here would make windowing allocate in
 * proportion to the complete table even though only a line count is needed.
 */
function estimatedDocumentNodeTextLines(node: WorkDocumentNode): number {
  let completedLines = 0;
  let currentLineLength = 0;
  let hasText = false;

  const visit = (current: WorkDocumentNode): void => {
    if (typeof current.text === 'string') {
      const text = current.text;
      if (!text) return;
      hasText = true;
      let start = 0;
      let newline = text.indexOf('\n');
      while (newline >= 0) {
        currentLineLength += newline - start;
        completedLines += Math.max(
          1,
          Math.ceil(currentLineLength / DOCUMENT_TEXT_CHARACTERS_PER_LINE),
        );
        currentLineLength = 0;
        start = newline + 1;
        newline = text.indexOf('\n', start);
      }
      currentLineLength += text.length - start;
      return;
    }
    for (const child of current.content ?? []) visit(child);
  };

  visit(node);
  if (!hasText) return 1;
  return (
    completedLines +
    Math.max(
      1,
      Math.ceil(currentLineLength / DOCUMENT_TEXT_CHARACTERS_PER_LINE),
    )
  );
}

function roundHeight(value: number): number {
  return Math.max(1, Math.round(value * 100) / 100);
}
