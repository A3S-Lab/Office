import { simpleDocumentNodeSize } from './work-document-lazy-model';
import { measuredDocumentBlock } from './work-document-pagination-dom';
import type { MeasuredDocumentLayoutBlock } from './work-document-pagination-types';
import {
  estimateDocumentNodeHeight,
  estimateDocumentTableRowHeight,
} from './work-document-windowing';
import type { WorkDocumentNode } from './work-types';

export function estimateLazyDocumentLayoutBlocks(
  nodes: readonly WorkDocumentNode[],
  element: HTMLElement,
  chunkId: string,
  firstPosition: number,
): MeasuredDocumentLayoutBlock[] {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  let position = firstPosition;
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (!node) continue;
    const nodeSize = simpleDocumentNodeSize(node);
    if (nodeSize === null) return [];
    const id = `${chunkId}-child-${index}-${position}`;
    if (node.type === 'table') {
      blocks.push(
        ...estimateLazyDocumentTableRows(node, element, id, position),
      );
    } else {
      const block = measuredDocumentBlock({
        block: {
          id,
          height: estimateDocumentNodeHeight(node),
          breakBefore: false,
          breakAfter: node.type === 'pageBreak',
          keepTogether: false,
          keepWithNext: false,
        },
        element,
        from: position,
        to: position + nodeSize,
      });
      block.observeResize = false;
      blocks.push(block);
    }
    position += nodeSize;
  }
  return blocks;
}

function estimateLazyDocumentTableRows(
  table: WorkDocumentNode,
  element: HTMLElement,
  blockId: string,
  tablePosition: number,
): MeasuredDocumentLayoutBlock[] {
  const rows = table.content ?? [];
  if (!rows.length) return [];
  const virtualTableId =
    typeof table.attrs?.virtualTableId === 'string' &&
    table.attrs.virtualTableId
      ? table.attrs.virtualTableId
      : blockId;
  const sliceIndex = Number.isSafeInteger(table.attrs?.virtualTableIndex)
    ? Number(table.attrs?.virtualTableIndex)
    : 0;
  const flowId = `${virtualTableId}-slice-${sliceIndex}`;
  const columnCount = estimateLazyDocumentTableColumnCount(table);
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
  let rowPosition = tablePosition + 1;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    if (!row) continue;
    const rowSize = simpleDocumentNodeSize(row);
    if (rowSize === null) return [];
    result.push({
      block: {
        id: `${flowId}-row-${rowIndex}`,
        height: estimateDocumentTableRowHeight(row),
        flowId,
        flowIndex: rowIndex,
        flowCount: rows.length,
        minimumFragmentsPerPage: 1,
      },
      element,
      from: rowPosition,
      to: rowPosition + rowSize,
      inlineOffsetLeft: 0,
      inlineOffsetRight: 0,
      observeResize: false,
      selectionRanges: estimateLazyDocumentTableRowSelectionRanges(
        row,
        rowPosition,
      ),
      tableBreak,
    });
    rowPosition += rowSize;
  }
  return result;
}

function estimateLazyDocumentTableRowSelectionRanges(
  row: WorkDocumentNode,
  rowPosition: number,
): Array<{ from: number; to: number }> {
  const ranges: Array<{ from: number; to: number }> = [];
  let cellPosition = rowPosition + 1;
  for (const cell of row.content ?? []) {
    const cellSize = simpleDocumentNodeSize(cell);
    if (cellSize === null) return [];
    ranges.push({
      from: cellPosition + 1,
      to: Math.max(cellPosition + 1, cellPosition + cellSize - 1),
    });
    cellPosition += cellSize;
  }
  return ranges;
}

function estimateLazyDocumentTableColumnCount(table: WorkDocumentNode): number {
  let maximum = 1;
  for (const row of table.content ?? []) {
    let columns = 0;
    for (const cell of row.content ?? []) {
      const colspan = Number(cell.attrs?.colspan);
      columns += Number.isSafeInteger(colspan) && colspan > 0 ? colspan : 1;
    }
    maximum = Math.max(maximum, columns);
  }
  return maximum;
}
