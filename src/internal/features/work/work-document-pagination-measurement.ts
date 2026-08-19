import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type {
  OfficeKernelPageMetrics,
  OfficeKernelPageStyle,
  OfficeKernelTextLayoutParagraphResult,
} from '../../kernel/office-kernel-protocol';
import { millimetersToPixels } from './work-document-layout';
import { documentLazyChunkContentForEditor } from './work-document-chunk-node';
import { estimateLazyDocumentLayoutBlocks } from './work-document-lazy-pagination-measurement';
import { resolveDocumentPageMargins } from './work-document-page-margins';
import { resolveDocumentPageSize } from './work-document-page-size';
import { normalizeDocumentImageLayout } from './work-document-image-layout';
import { measureParagraphLineFragments } from './work-document-line-measurement';
import {
  documentBlockId,
  documentInlineOffsets,
  documentListChildId,
  documentNodeParagraphPagination,
  elementForNode,
  isDocumentListNode,
  measuredDocumentBlock,
  outerHeight,
  reusableDocumentChunkLayoutBlocks,
  reusableDocumentLayoutBlocks,
  shouldKeepDocumentBlockTogether,
  verticalBlockEnd,
  verticalBlockStart,
} from './work-document-pagination-dom';
import type {
  DocumentPaginationSection,
  DocumentPaginationSnapshot,
  MeasuredDocumentLayoutBlock,
} from './work-document-pagination-types';
import {
  documentSectionLayoutFromNodeAttributes,
  type DocumentSectionNodeAttributes,
} from './work-document-section';
import { measureDocumentTableRows } from './work-document-table-pagination';
import type { WorkDocumentSectionLayout } from './work-types';

export interface IncrementalDocumentLayoutMeasurementOptions {
  checkpoint?: (signal?: AbortSignal) => Promise<void>;
  signal?: AbortSignal;
}

const DOCUMENT_LAYOUT_MEASUREMENT_SLICE_MS = 32;

interface DocumentMeasurementCounts {
  measured: number;
  prefixOpen: boolean;
  reused: number;
  reusedPrefix: number;
}

export function measureDocumentLayoutBlocks(
  editor: Editor,
  previous: DocumentPaginationSnapshot | null = null,
  dirtyFrom = 0,
  textLayouts: ReadonlyMap<
    string,
    OfficeKernelTextLayoutParagraphResult
  > = new Map(),
  maximumFragmentedTableRowHeight = 1_000_000,
): DocumentPaginationSnapshot {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  const pageStyleByMetrics = new Map<string, OfficeKernelPageStyle>();
  const counts: DocumentMeasurementCounts = {
    measured: 0,
    prefixOpen: true,
    reused: 0,
    reusedPrefix: 0,
  };
  let unsupportedLayout = false;
  editor.state.doc.forEach((section, sectionPosition, sectionIndex) => {
    if (section.type.name !== 'documentSection') return;
    const layout = documentSectionLayoutFromNodeAttributes(
      section.attrs as Partial<DocumentSectionNodeAttributes>,
    );
    if (layout.columns.count > 1) unsupportedLayout = true;
    const page = documentPageMetrics(layout);
    const metricsKey = documentPageMetricsKey(page);
    let pageStyle = pageStyleByMetrics.get(metricsKey);
    if (!pageStyle) {
      pageStyle = {
        id: `document-page-style-${pageStyleByMetrics.size + 1}`,
        page,
      };
      pageStyleByMetrics.set(metricsKey, pageStyle);
    }
    measureSectionBlocks(
      editor,
      section,
      sectionPosition,
      sectionIndex,
      layout,
      pageStyle,
      previous?.blocks ?? [],
      dirtyFrom,
      textLayouts,
      Math.min(
        maximumFragmentedTableRowHeight,
        documentPageBodyHeight(pageStyle.page),
      ),
      counts,
      blocks,
    );
  });
  return {
    blocks,
    pageStyles: [...pageStyleByMetrics.values()],
    measuredBlockCount: counts.measured,
    reusedBlockCount: counts.reused,
    reusedPrefixBlockCount: counts.reusedPrefix,
    unsupportedLayout,
  };
}

/**
 * Measures the same canonical editor DOM as `measureDocumentLayoutBlocks`
 * while yielding between top-level blocks. The browser keeps input, media,
 * and animation tasks schedulable during large-document pagination without
 * weakening the layout result or replacing the controlled editor model.
 */
export async function measureDocumentLayoutBlocksIncrementally(
  editor: Editor,
  previous: DocumentPaginationSnapshot | null = null,
  dirtyFrom = 0,
  textLayouts: ReadonlyMap<
    string,
    OfficeKernelTextLayoutParagraphResult
  > = new Map(),
  maximumFragmentedTableRowHeight = 1_000_000,
  options: IncrementalDocumentLayoutMeasurementOptions = {},
): Promise<DocumentPaginationSnapshot> {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  const pageStyleByMetrics = new Map<string, OfficeKernelPageStyle>();
  const counts: DocumentMeasurementCounts = {
    measured: 0,
    prefixOpen: true,
    reused: 0,
    reusedPrefix: 0,
  };
  const checkpoint =
    options.checkpoint ?? createDocumentLayoutMeasurementCheckpoint();
  throwIfDocumentMeasurementAborted(options.signal);
  let unsupportedLayout = false;
  let sectionPosition = 0;
  for (
    let sectionIndex = 0;
    sectionIndex < editor.state.doc.childCount;
    sectionIndex += 1
  ) {
    const section = editor.state.doc.child(sectionIndex);
    if (section.type.name === 'documentSection') {
      const layout = documentSectionLayoutFromNodeAttributes(
        section.attrs as Partial<DocumentSectionNodeAttributes>,
      );
      if (layout.columns.count > 1) unsupportedLayout = true;
      const page = documentPageMetrics(layout);
      const metricsKey = documentPageMetricsKey(page);
      let pageStyle = pageStyleByMetrics.get(metricsKey);
      if (!pageStyle) {
        pageStyle = {
          id: `document-page-style-${pageStyleByMetrics.size + 1}`,
          page,
        };
        pageStyleByMetrics.set(metricsKey, pageStyle);
      }
      await measureSectionBlocksIncrementally(
        editor,
        section,
        sectionPosition,
        sectionIndex,
        layout,
        pageStyle,
        previous?.blocks ?? [],
        dirtyFrom,
        textLayouts,
        Math.min(
          maximumFragmentedTableRowHeight,
          documentPageBodyHeight(pageStyle.page),
        ),
        counts,
        blocks,
        checkpoint,
        options.signal,
      );
    }
    sectionPosition += section.nodeSize;
  }
  return {
    blocks,
    pageStyles: [...pageStyleByMetrics.values()],
    measuredBlockCount: counts.measured,
    reusedBlockCount: counts.reused,
    reusedPrefixBlockCount: counts.reusedPrefix,
    unsupportedLayout,
  };
}

export function documentPageMetrics(
  layout: WorkDocumentSectionLayout,
  physicalPage = 1,
): OfficeKernelPageMetrics {
  const pageSize = resolveDocumentPageSize(layout);
  const resolvedMargins = resolveDocumentPageMargins(layout, physicalPage);
  const margins = resolvedMargins.body;
  const marginTop = millimetersToPixels(margins.top);
  const marginBottom = millimetersToPixels(margins.bottom);
  return {
    width: pageSize.widthPixels,
    height: pageSize.heightPixels,
    marginTop,
    marginRight: millimetersToPixels(margins.right),
    marginBottom,
    marginLeft: millimetersToPixels(margins.left),
    headerHeight: Math.max(
      0,
      marginTop - millimetersToPixels(resolvedMargins.headerDistance),
    ),
    footerHeight: Math.max(
      0,
      marginBottom - millimetersToPixels(resolvedMargins.footerDistance),
    ),
    pageGap: 28,
  };
}

export function documentPageBodyHeight(page: OfficeKernelPageMetrics): number {
  return Math.max(1, page.height - page.marginTop - page.marginBottom);
}

export function documentPaginationSurfaceHeight(
  pageCount: number,
  page: OfficeKernelPageMetrics,
): number {
  const count = Math.max(1, Math.trunc(pageCount));
  return page.height * count + page.pageGap * (count - 1);
}

function measureSectionBlocks(
  editor: Editor,
  section: ProseMirrorNode,
  sectionPosition: number,
  sectionIndex: number,
  layout: WorkDocumentSectionLayout,
  pageStyle: OfficeKernelPageStyle,
  previous: readonly MeasuredDocumentLayoutBlock[],
  dirtyFrom: number,
  textLayouts: ReadonlyMap<string, OfficeKernelTextLayoutParagraphResult>,
  maximumFragmentedTableRowHeight: number,
  counts: DocumentMeasurementCounts,
  result: MeasuredDocumentLayoutBlock[],
): void {
  const sectionBlocks: MeasuredDocumentLayoutBlock[] = [];
  section.forEach((node, offset, index) => {
    measureSectionBlock(
      editor,
      node,
      offset,
      index,
      sectionPosition,
      previous,
      dirtyFrom,
      textLayouts,
      maximumFragmentedTableRowHeight,
      counts,
      sectionBlocks,
    );
  });

  finishMeasuredSection(
    section,
    sectionPosition,
    sectionIndex,
    layout,
    pageStyle,
    sectionBlocks,
    result,
  );
}

async function measureSectionBlocksIncrementally(
  editor: Editor,
  section: ProseMirrorNode,
  sectionPosition: number,
  sectionIndex: number,
  layout: WorkDocumentSectionLayout,
  pageStyle: OfficeKernelPageStyle,
  previous: readonly MeasuredDocumentLayoutBlock[],
  dirtyFrom: number,
  textLayouts: ReadonlyMap<string, OfficeKernelTextLayoutParagraphResult>,
  maximumFragmentedTableRowHeight: number,
  counts: DocumentMeasurementCounts,
  result: MeasuredDocumentLayoutBlock[],
  checkpoint: (signal?: AbortSignal) => Promise<void>,
  signal?: AbortSignal,
): Promise<void> {
  const sectionBlocks: MeasuredDocumentLayoutBlock[] = [];
  let offset = 0;
  for (let index = 0; index < section.childCount; index += 1) {
    await checkpoint(signal);
    throwIfDocumentMeasurementAborted(signal);
    const node = section.child(index);
    measureSectionBlock(
      editor,
      node,
      offset,
      index,
      sectionPosition,
      previous,
      dirtyFrom,
      textLayouts,
      maximumFragmentedTableRowHeight,
      counts,
      sectionBlocks,
    );
    offset += node.nodeSize;
  }

  finishMeasuredSection(
    section,
    sectionPosition,
    sectionIndex,
    layout,
    pageStyle,
    sectionBlocks,
    result,
  );
}

function measureSectionBlock(
  editor: Editor,
  node: ProseMirrorNode,
  offset: number,
  index: number,
  sectionPosition: number,
  previous: readonly MeasuredDocumentLayoutBlock[],
  dirtyFrom: number,
  textLayouts: ReadonlyMap<string, OfficeKernelTextLayoutParagraphResult>,
  maximumFragmentedTableRowHeight: number,
  counts: DocumentMeasurementCounts,
  sectionBlocks: MeasuredDocumentLayoutBlock[],
): void {
  const position = sectionPosition + offset + 1;
  const element = elementForNode(editor, position);
  if (!element) return;
  const id = documentBlockId(sectionPosition, index, position);
  if (node.type.name === 'documentChunk') {
    const reused = reusableDocumentChunkLayoutBlocks(
      previous,
      element,
      position,
      position + node.nodeSize,
      dirtyFrom,
    );
    if (reused.length) {
      recordReusedDocumentBlocks(counts, reused.length);
      sectionBlocks.push(...reused);
      return;
    }
    const estimated = estimatedDocumentChunkBlocks(
      editor,
      node,
      element,
      id,
      position,
    );
    recordMeasuredDocumentBlocks(counts, estimated.length);
    sectionBlocks.push(...estimated);
    return;
  }
  const paragraphPagination = documentNodeParagraphPagination(node);
  if (isDocumentListNode(node)) {
    const reused = reusableDocumentListLayoutBlocks(
      previous,
      id,
      position + node.nodeSize,
      dirtyFrom,
    );
    if (reused.length) {
      recordReusedDocumentBlocks(counts, reused.length);
      sectionBlocks.push(...reused);
      return;
    }
    const listBlocks = measureDocumentListBlocks(
      editor,
      node,
      element,
      id,
      position,
      textLayouts,
      maximumFragmentedTableRowHeight,
    );
    if (listBlocks.length) {
      recordMeasuredDocumentBlocks(counts, listBlocks.length);
      sectionBlocks.push(...listBlocks);
      return;
    }
  }
  const reused = reusableDocumentLayoutBlocks(
    previous,
    id,
    element,
    position + node.nodeSize,
    dirtyFrom,
  );
  if (reused.length) {
    recordReusedDocumentBlocks(counts, reused.length);
    sectionBlocks.push(...reused);
    return;
  }
  if (node.type.name === 'table') {
    const tableRows = measureDocumentTableRows(
      editor,
      node,
      element,
      id,
      position,
      maximumFragmentedTableRowHeight,
    );
    if (tableRows.length) {
      recordMeasuredDocumentBlocks(counts, tableRows.length);
      sectionBlocks.push(...tableRows);
      return;
    }
  }
  const lineFragments = measureParagraphLineFragments(
    editor,
    node,
    element,
    id,
    position,
    position + node.nodeSize,
    paragraphPagination,
    textLayouts.get(id),
  );
  if (lineFragments.length > 1) {
    recordMeasuredDocumentBlocks(counts, lineFragments.length);
    sectionBlocks.push(...lineFragments);
    return;
  }
  recordMeasuredDocumentBlocks(counts, 1);
  sectionBlocks.push(
    measuredDocumentBlock({
      block: {
        id,
        height: documentBlockFlowHeight(node, element),
        breakBefore: paragraphPagination.pageBreakBefore,
        breakAfter: node.type.name === 'pageBreak',
        keepTogether:
          paragraphPagination.keepLines ||
          shouldKeepDocumentBlockTogether(node),
        keepWithNext: paragraphPagination.keepWithNext,
      },
      element,
      from: position,
      to: position + node.nodeSize,
    }),
  );
}

function estimatedDocumentChunkBlocks(
  editor: Editor,
  chunk: ProseMirrorNode,
  element: HTMLElement,
  chunkId: string,
  chunkPosition: number,
): MeasuredDocumentLayoutBlock[] {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  chunk.forEach((node, offset, index) => {
    const position = chunkPosition + offset + 1;
    const id = `${chunkId}-child-${index}-${position}`;
    if (node.type.name === 'documentChunk') {
      blocks.push(
        ...estimatedDocumentChunkBlocks(editor, node, element, id, position),
      );
      return;
    }
    if (node.type.name === 'documentLazyBlock') {
      const sourceId = typeof chunk.attrs.id === 'string' ? chunk.attrs.id : '';
      const payload = documentLazyChunkContentForEditor(editor, sourceId);
      if (payload?.length) {
        blocks.push(
          ...estimateLazyDocumentLayoutBlocks(
            payload,
            element,
            chunkId,
            position,
          ),
        );
      }
      return;
    }
    if (node.type.name === 'table') {
      blocks.push(...estimatedDocumentTableRows(node, element, id, position));
      return;
    }
    const pagination = documentNodeParagraphPagination(node);
    const block = measuredDocumentBlock({
      block: {
        id,
        height: estimatedProseMirrorNodeHeight(node),
        breakBefore: pagination.pageBreakBefore,
        breakAfter: node.type.name === 'pageBreak',
        keepTogether:
          pagination.keepLines || shouldKeepDocumentBlockTogether(node),
        keepWithNext: pagination.keepWithNext,
      },
      element,
      from: position,
      to: position + node.nodeSize,
    });
    block.observeResize = false;
    blocks.push(block);
  });
  return blocks;
}

function estimatedDocumentTableRows(
  table: ProseMirrorNode,
  element: HTMLElement,
  blockId: string,
  tablePosition: number,
): MeasuredDocumentLayoutBlock[] {
  const rowCount = table.childCount;
  if (!rowCount) return [];
  const virtualTableId =
    typeof table.attrs.virtualTableId === 'string' && table.attrs.virtualTableId
      ? table.attrs.virtualTableId
      : blockId;
  const sliceIndex = Number.isSafeInteger(table.attrs.virtualTableIndex)
    ? Number(table.attrs.virtualTableIndex)
    : 0;
  const flowId = `${virtualTableId}-slice-${sliceIndex}`;
  const columnCount = estimatedDocumentTableColumnCount(table);
  const tableBreak = {
    tableId: flowId,
    columnCount,
    colgroupHtml: '',
    repeatedHeaderRowsHtml: [],
    repeatedHeaderOverlayHtml: '',
    repeatHeaderHeight: 0,
    tableWidth: 0,
    leadingCellOffsetLeft: 0,
  };
  const result: MeasuredDocumentLayoutBlock[] = [];
  table.forEach((row, offset, rowIndex) => {
    const position = tablePosition + offset + 1;
    result.push({
      block: {
        id: `${flowId}-row-${rowIndex}`,
        height: estimatedProseMirrorTableRowHeight(row),
        flowId,
        flowIndex: rowIndex,
        flowCount: rowCount,
        minimumFragmentsPerPage: 1,
      },
      element,
      from: position,
      to: position + row.nodeSize,
      inlineOffsetLeft: 0,
      inlineOffsetRight: 0,
      observeResize: false,
      selectionRanges: estimatedDocumentTableRowSelectionRanges(row, position),
      tableBreak,
    });
  });
  return result;
}

function estimatedDocumentTableRowSelectionRanges(
  row: ProseMirrorNode,
  rowPosition: number,
): Array<{ from: number; to: number }> {
  const ranges: Array<{ from: number; to: number }> = [];
  row.forEach((cell, offset) => {
    const cellPosition = rowPosition + offset + 1;
    ranges.push({
      from: cellPosition + 1,
      to: Math.max(cellPosition + 1, cellPosition + cell.nodeSize - 1),
    });
  });
  return ranges;
}

function estimatedDocumentTableColumnCount(table: ProseMirrorNode): number {
  let maximum = 1;
  table.forEach((row) => {
    let columns = 0;
    row.forEach((cell) => {
      const colspan = Number(cell.attrs.colspan);
      columns += Number.isSafeInteger(colspan) && colspan > 0 ? colspan : 1;
    });
    maximum = Math.max(maximum, columns);
  });
  return maximum;
}

function estimatedProseMirrorTableRowHeight(row: ProseMirrorNode): number {
  const explicit = Number(row.attrs.rowHeight);
  let maximumTextLength = 0;
  row.forEach((cell) => {
    maximumTextLength = Math.max(maximumTextLength, cell.textContent.length);
  });
  const lines = Math.max(1, Math.ceil(maximumTextLength / 88));
  return Math.max(
    22,
    Number.isFinite(explicit) && explicit > 0 ? explicit : 0,
    lines * 21 + 1,
  );
}

function estimatedProseMirrorNodeHeight(node: ProseMirrorNode): number {
  const lines = Math.max(1, Math.ceil(node.textContent.length / 88));
  if (node.type.name === 'heading') {
    const level = Number(node.attrs.level);
    const lineHeight = level <= 1 ? 32 : level === 2 ? 28 : 25;
    return lines * lineHeight + 18;
  }
  if (node.type.name === 'image') {
    const height = Number(node.attrs.height);
    return Math.max(
      21,
      (Number.isFinite(height) && height > 0 ? height : 120) + 36,
    );
  }
  if (node.type.name === 'horizontalRule' || node.type.name === 'pageBreak') {
    return 38;
  }
  if (node.type.name === 'bulletList' || node.type.name === 'orderedList') {
    return Math.max(29, node.childCount * 29);
  }
  return lines * 21 + 8;
}

function finishMeasuredSection(
  section: ProseMirrorNode,
  sectionPosition: number,
  sectionIndex: number,
  layout: WorkDocumentSectionLayout,
  pageStyle: OfficeKernelPageStyle,
  sectionBlocks: MeasuredDocumentLayoutBlock[],
  result: MeasuredDocumentLayoutBlock[],
): void {
  const last = sectionBlocks.at(-1);
  if (
    last &&
    layout.breakAfter !== 'continuous' &&
    layout.breakAfter !== 'nextColumn'
  ) {
    last.block.breakAfter = true;
  }
  const sectionMetadata: DocumentPaginationSection = {
    id:
      typeof section.attrs.id === 'string' && section.attrs.id
        ? section.attrs.id
        : `document-section-${sectionIndex + 1}`,
    index: sectionIndex,
    position: sectionPosition,
    layout,
    pageStyleId: pageStyle.id,
    page: pageStyle.page,
  };
  for (const block of sectionBlocks) {
    block.block.pageStyleId = pageStyle.id;
    block.section = sectionMetadata;
  }
  result.push(...sectionBlocks);
}

function recordReusedDocumentBlocks(
  counts: DocumentMeasurementCounts,
  count: number,
): void {
  if (count <= 0) return;
  counts.reused += count;
  if (counts.prefixOpen) counts.reusedPrefix += count;
}

function recordMeasuredDocumentBlocks(
  counts: DocumentMeasurementCounts,
  count: number,
): void {
  if (count <= 0) return;
  counts.measured += count;
  counts.prefixOpen = false;
}

function createDocumentLayoutMeasurementCheckpoint(): (
  signal?: AbortSignal,
) => Promise<void> {
  let deadline =
    documentMeasurementNow() + DOCUMENT_LAYOUT_MEASUREMENT_SLICE_MS;
  return async (signal) => {
    throwIfDocumentMeasurementAborted(signal);
    if (documentMeasurementNow() < deadline) return;
    await yieldDocumentMeasurement();
    throwIfDocumentMeasurementAborted(signal);
    deadline = documentMeasurementNow() + DOCUMENT_LAYOUT_MEASUREMENT_SLICE_MS;
  };
}

function documentMeasurementNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function yieldDocumentMeasurement(): Promise<void> {
  if (typeof MessageChannel !== 'undefined') {
    return new Promise((resolve) => {
      const channel = new MessageChannel();
      channel.port1.addEventListener(
        'message',
        () => {
          channel.port1.close();
          channel.port2.close();
          resolve();
        },
        { once: true },
      );
      channel.port1.start();
      channel.port2.postMessage(undefined);
    });
  }
  return new Promise((resolve) => globalThis.setTimeout(resolve, 0));
}

function throwIfDocumentMeasurementAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return;
  if (signal.reason instanceof Error) throw signal.reason;
  const error = new Error('Document layout measurement was aborted');
  error.name = 'AbortError';
  throw error;
}

function documentPageMetricsKey(page: OfficeKernelPageMetrics): string {
  return [
    page.width,
    page.height,
    page.marginTop,
    page.marginRight,
    page.marginBottom,
    page.marginLeft,
    page.headerHeight,
    page.footerHeight,
    page.pageGap,
  ].join(':');
}

function measureDocumentListBlocks(
  editor: Editor,
  list: ProseMirrorNode,
  listElement: HTMLElement,
  listId: string,
  listPosition: number,
  textLayouts: ReadonlyMap<string, OfficeKernelTextLayoutParagraphResult>,
  maximumFragmentedTableRowHeight: number,
): MeasuredDocumentLayoutBlock[] {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  let complete = true;
  list.forEach((item, itemOffset, itemIndex) => {
    if (item.type.name !== 'listItem') {
      complete = false;
      return;
    }
    const itemPosition = listPosition + itemOffset + 1;
    const itemElement = elementForNode(editor, itemPosition);
    if (!itemElement || itemElement.tagName.toLowerCase() !== 'li') {
      complete = false;
      return;
    }
    const itemBlocks: MeasuredDocumentLayoutBlock[] = [];
    item.forEach((node, offset, index) => {
      const position = itemPosition + offset + 1;
      const element = elementForNode(editor, position);
      if (!element) {
        complete = false;
        return;
      }
      const nestedList = isDocumentListNode(node);
      const id = documentListChildId(
        listId,
        itemIndex,
        node,
        index,
        nestedList,
      );
      if (nestedList) {
        const nestedBlocks = measureDocumentListBlocks(
          editor,
          node,
          element,
          id,
          position,
          textLayouts,
          maximumFragmentedTableRowHeight,
        );
        if (!nestedBlocks.length) {
          complete = false;
          return;
        }
        itemBlocks.push(...nestedBlocks);
        return;
      }
      itemBlocks.push(
        ...measureDocumentListItemBlock(
          editor,
          node,
          element,
          id,
          position,
          textLayouts.get(id),
          maximumFragmentedTableRowHeight,
        ),
      );
    });
    if (!itemBlocks.length) {
      complete = false;
      return;
    }
    fitDocumentBlocksToContainer(itemBlocks, itemElement);
    blocks.push(...itemBlocks);
  });
  if (!complete || !blocks.length) return [];
  fitDocumentBlocksToContainer(blocks, listElement);
  return blocks;
}

function measureDocumentListItemBlock(
  editor: Editor,
  node: ProseMirrorNode,
  element: HTMLElement,
  id: string,
  position: number,
  shapedLayout: OfficeKernelTextLayoutParagraphResult | undefined,
  maximumFragmentedTableRowHeight: number,
): MeasuredDocumentLayoutBlock[] {
  if (node.type.name === 'table') {
    const rows = measureDocumentTableRows(
      editor,
      node,
      element,
      id,
      position,
      maximumFragmentedTableRowHeight,
    );
    if (rows.length) return rows;
  }
  const pagination = documentNodeParagraphPagination(node);
  const lineFragments = measureParagraphLineFragments(
    editor,
    node,
    element,
    id,
    position,
    position + node.nodeSize,
    pagination,
    shapedLayout,
  );
  const { inlineOffsetLeft, inlineOffsetRight } = documentInlineOffsets(
    editor,
    element,
  );
  if (lineFragments.length > 1) {
    return lineFragments.map((fragment) => ({
      ...fragment,
      inlineOffsetLeft,
      inlineOffsetRight,
    }));
  }
  return [
    {
      ...measuredDocumentBlock({
        block: {
          id,
          height: documentBlockFlowHeight(node, element),
          breakBefore: pagination.pageBreakBefore,
          breakAfter: node.type.name === 'pageBreak',
          keepTogether:
            pagination.keepLines || shouldKeepDocumentBlockTogether(node),
          keepWithNext: pagination.keepWithNext,
        },
        element,
        from: position,
        to: position + node.nodeSize,
      }),
      inlineOffsetLeft,
      inlineOffsetRight,
    },
  ];
}

function fitDocumentBlocksToContainer(
  blocks: MeasuredDocumentLayoutBlock[],
  container: HTMLElement,
): void {
  const target = outerHeight(container);
  if (!(target > 0) || blocks.length === 0) return;
  const measured = blocks.reduce(
    (height, candidate) => height + candidate.block.height,
    0,
  );
  let difference = target - measured;
  if (Math.abs(difference) < 0.01) return;

  const first = blocks[0];
  const last = blocks.at(-1) as MeasuredDocumentLayoutBlock;
  if (difference > 0) {
    const style = getComputedStyle(container);
    const start = Math.min(difference, verticalBlockStart(style));
    first.block.height += start;
    difference -= start;
    const end = Math.min(difference, verticalBlockEnd(style));
    last.block.height += end;
    difference -= end;
    last.block.height += difference;
    return;
  }

  let remaining = -difference;
  const candidates = Array.from(
    new Set([last, first, ...blocks.slice(1, -1).reverse()]),
  );
  for (const candidate of candidates) {
    if (remaining <= 0) break;
    const removable = Math.max(0, candidate.block.height - 1);
    const amount = Math.min(removable, remaining);
    candidate.block.height -= amount;
    remaining -= amount;
  }
}

function documentBlockFlowHeight(
  node: ProseMirrorNode,
  element: HTMLElement,
): number {
  if (
    node.type.name === 'image' &&
    normalizeDocumentImageLayout(node.attrs.layout) === 'none'
  ) {
    return 1;
  }
  return Math.max(1, outerHeight(element));
}

function reusableDocumentListLayoutBlocks(
  previous: readonly MeasuredDocumentLayoutBlock[],
  listId: string,
  nodeTo: number,
  dirtyFrom: number,
): MeasuredDocumentLayoutBlock[] {
  if (nodeTo > dirtyFrom) return [];
  const prefix = `${listId}-item-`;
  return previous
    .filter(
      ({ block }) =>
        block.id.startsWith(prefix) || block.flowId?.startsWith(prefix),
    )
    .map((candidate) => ({
      ...candidate,
      block: { ...candidate.block },
    }));
}
