import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  documentInlineOffsets,
  elementForNode,
  nonNegativePixels,
  outerHeight,
  verticalBlockEnd,
  verticalBlockStart,
} from './work-document-pagination-dom';
import type {
  DocumentTableCellBoundary,
  DocumentTableCellFragmentMeasurement,
  DocumentTablePaginationBreak,
  DocumentTableRowFragmentPlan,
  MeasuredDocumentLayoutBlock,
} from './work-document-pagination-types';
import {
  documentTableRowCantSplit,
  documentTableRowRepeats,
} from './work-document-table-row';

interface MeasuredDocumentTableRow {
  node: ProseMirrorNode;
  element: HTMLElement;
  from: number;
  to: number;
  height: number;
}

const MAX_DOCUMENT_TABLE_ROW_FRAGMENTS = 256;

export function createDocumentTableRowFragmentPlan(
  cells: readonly DocumentTableCellFragmentMeasurement[],
  rowHeight: number,
  maximumFragmentedRowHeight: number,
): DocumentTableRowFragmentPlan[] {
  const height = Math.max(1, rowHeight);
  const atomic = (): DocumentTableRowFragmentPlan[] => [
    {
      height,
      cellRanges: cells.flatMap((cell) =>
        cell.to > cell.from ? [{ from: cell.from, to: cell.to }] : [],
      ),
    },
  ];
  if (
    !cells.length ||
    !Number.isFinite(maximumFragmentedRowHeight) ||
    height > maximumFragmentedRowHeight
  ) {
    return atomic();
  }

  const candidates = normalizedDocumentTableBreakCoordinates(
    cells.flatMap((cell) =>
      cell.boundaries
        .slice(1, -1)
        .map((boundary) => boundary.y)
        .filter((y) => y > 0.5 && y < height - 0.5),
    ),
  ).slice(0, MAX_DOCUMENT_TABLE_ROW_FRAGMENTS - 1);
  if (!candidates.length) return atomic();

  const coordinates = [0, ...candidates, height];
  const positions = coordinates.map((coordinate, coordinateIndex) =>
    cells.map((cell) => {
      if (coordinateIndex === 0) {
        return { position: cell.from, y: 0 };
      }
      if (coordinateIndex + 1 === coordinates.length) {
        return { position: cell.to, y: height };
      }
      return documentTableBoundaryAtOrBefore(cell, coordinate);
    }),
  );
  const fragments = coordinates.slice(0, -1).flatMap((coordinate, index) => {
    const nextCoordinate = coordinates[index + 1];
    const cellRanges = cells.flatMap((_cell, cellIndex) => {
      const from = positions[index]?.[cellIndex]?.position;
      const to = positions[index + 1]?.[cellIndex]?.position;
      return from !== undefined && to !== undefined && to > from
        ? [{ from, to }]
        : [];
    });
    if (!cellRanges.length) return [];
    const cellBreaks =
      index === 0
        ? undefined
        : cells.map((cell, cellIndex) => {
            const boundary = positions[index]?.[cellIndex] ?? {
              position: cell.from,
              y: 0,
            };
            return {
              cellIndex: cell.cellIndex,
              position: boundary.position,
              alignmentOffset: Math.max(0, coordinate - boundary.y),
              ...(cell.tableOffsetLeft === undefined
                ? {}
                : { tableOffsetLeft: cell.tableOffsetLeft }),
              ...(cell.outerWidth === undefined
                ? {}
                : { outerWidth: cell.outerWidth }),
              ...(cell.contentOffsetLeft === undefined
                ? {}
                : { contentOffsetLeft: cell.contentOffsetLeft }),
            };
          });
    return [
      {
        height: Math.max(1, nextCoordinate - coordinate),
        cellRanges,
        ...(cellBreaks ? { cellBreaks } : {}),
      },
    ];
  });
  return fragments.length > 1 ? fragments : atomic();
}

function normalizedDocumentTableBreakCoordinates(
  coordinates: readonly number[],
): number[] {
  const result: number[] = [];
  for (const coordinate of [...coordinates].sort(
    (left, right) => left - right,
  )) {
    if (!Number.isFinite(coordinate)) continue;
    const normalized = Math.round(coordinate * 4) / 4;
    if (result.length === 0 || normalized - (result.at(-1) as number) >= 1) {
      result.push(normalized);
    }
  }
  return result;
}

function documentTableBoundaryAtOrBefore(
  cell: DocumentTableCellFragmentMeasurement,
  coordinate: number,
): DocumentTableCellBoundary {
  let result = cell.boundaries[0] ?? { position: cell.from, y: 0 };
  for (const boundary of cell.boundaries) {
    if (boundary.y > coordinate + 0.5) break;
    result = boundary;
  }
  return result;
}

export function measureDocumentTableRows(
  editor: Editor,
  tableNode: ProseMirrorNode,
  tableElement: HTMLElement,
  tableId: string,
  tableFrom: number,
  maximumFragmentedRowHeight: number,
): MeasuredDocumentLayoutBlock[] {
  const rows: MeasuredDocumentTableRow[] = [];
  let complete = true;
  tableNode.forEach((row, offset) => {
    const from = tableFrom + offset + 1;
    const element = elementForNode(editor, from);
    if (
      row.type.name !== 'tableRow' ||
      !element ||
      element.tagName.toLowerCase() !== 'tr'
    ) {
      complete = false;
      return;
    }
    rows.push({
      node: row,
      element,
      from,
      to: from + row.nodeSize,
      height: Math.max(1, outerHeight(element)),
    });
  });
  if (!complete || rows.length !== tableNode.childCount || !rows.length) {
    return [];
  }

  let leadingHeaderRowCount = 0;
  for (const row of rows) {
    if (!documentTableRowRepeats(row.node)) break;
    leadingHeaderRowCount += 1;
  }
  const repeatHeaderCount =
    leadingHeaderRowCount < rows.length ? leadingHeaderRowCount : 0;

  const rowHeight = rows.reduce((sum, row) => sum + row.height, 0);
  const tableStyle = getComputedStyle(tableElement);
  const blockStart = verticalBlockStart(tableStyle);
  const blockEnd = verticalBlockEnd(tableStyle);
  const unassignedHeight = Math.max(
    0,
    outerHeight(tableElement) - rowHeight - blockStart - blockEnd,
  );
  const repeatHeaderHeight =
    repeatHeaderCount > 0
      ? rows
          .slice(0, repeatHeaderCount)
          .reduce((sum, row) => sum + row.height, 0)
      : 0;
  const tableBreak = documentTablePaginationBreak(
    tableElement,
    tableNode,
    tableId,
    rows,
    repeatHeaderCount,
    repeatHeaderHeight,
  );
  const { inlineOffsetLeft, inlineOffsetRight } = documentInlineOffsets(
    editor,
    tableElement,
  );
  const continuationRowHeight = Math.max(
    1,
    maximumFragmentedRowHeight - repeatHeaderHeight,
  );
  const rowPlans = rows.map((row, rowIndex) => {
    const cells = measureDocumentTableCellFragments(editor, row);
    const canSplit =
      rowIndex >= leadingHeaderRowCount && !documentTableRowCantSplit(row.node);
    return createDocumentTableRowFragmentPlan(
      cells,
      row.height,
      canSplit ? continuationRowHeight : 0,
    );
  });
  const flowCount = rowPlans.reduce(
    (count, fragments) => count + fragments.length,
    0,
  );
  const repeatHeaderFragmentCount =
    repeatHeaderCount > 0
      ? rowPlans
          .slice(0, repeatHeaderCount)
          .reduce((count, fragments) => count + fragments.length, 0)
      : 0;
  const result: MeasuredDocumentLayoutBlock[] = [];
  let flowIndex = 0;
  for (const [rowIndex, row] of rows.entries()) {
    const fragments = rowPlans[rowIndex] ?? [];
    for (const [fragmentIndex, fragment] of fragments.entries()) {
      const first = flowIndex === 0;
      const last = flowIndex + 1 === flowCount;
      const repeatsHeader =
        repeatHeaderFragmentCount > 0 && flowIndex >= repeatHeaderFragmentCount;
      result.push({
        block: {
          id:
            fragments.length > 1
              ? `${tableId}-row-${rowIndex}-fragment-${fragmentIndex}`
              : `${tableId}-row-${rowIndex}`,
          height:
            fragment.height +
            (first ? blockStart : 0) +
            (last ? blockEnd + unassignedHeight : 0),
          ...(flowCount > 1
            ? {
                flowId: tableId,
                flowIndex,
                flowCount,
                minimumFragmentsPerPage: 1,
              }
            : {}),
          ...(flowCount > 1 && repeatHeaderFragmentCount > 0
            ? {
                repeatHeaderCount: repeatHeaderFragmentCount,
                repeatHeaderHeight,
              }
            : {}),
        },
        element: tableElement,
        from: row.from,
        to: row.to,
        inlineOffsetLeft,
        inlineOffsetRight,
        observeResize: false,
        selectionRanges: fragment.cellRanges,
        tableBreak: {
          ...tableBreak,
          repeatedHeaderRowsHtml: repeatsHeader
            ? tableBreak.repeatedHeaderRowsHtml
            : [],
          repeatedHeaderOverlayHtml: repeatsHeader
            ? tableBreak.repeatedHeaderOverlayHtml
            : '',
          ...(fragment.cellBreaks ? { cellBreaks: fragment.cellBreaks } : {}),
        },
      });
      flowIndex += 1;
    }
  }
  return result;
}

function measureDocumentTableCellFragments(
  editor: Editor,
  row: MeasuredDocumentTableRow,
): DocumentTableCellFragmentMeasurement[] {
  const rowRect = row.element.getBoundingClientRect();
  const tableElement = row.element.closest('table');
  const tableRect = tableElement?.getBoundingClientRect();
  const scaleX =
    tableElement && tableElement.offsetWidth > 0 && tableRect
      ? tableRect.width / tableElement.offsetWidth
      : 1;
  const scaleY =
    row.element.offsetHeight > 0
      ? rowRect.height / row.element.offsetHeight
      : 1;
  const result: DocumentTableCellFragmentMeasurement[] = [];
  row.node.forEach((cell, cellOffset, cellIndex) => {
    const from = row.from + cellOffset + 2;
    const to = row.from + cellOffset + cell.nodeSize;
    const element = elementForNode(editor, from - 1);
    const boundaries: DocumentTableCellBoundary[] = [];
    cell.forEach((_block, blockOffset, blockIndex) => {
      const position = from + blockOffset;
      const blockElement = elementForNode(editor, position);
      if (!blockElement) return;
      const rect = blockElement.getBoundingClientRect();
      const style = getComputedStyle(blockElement);
      const y = Math.max(
        0,
        (rect.top - rowRect.top) / Math.max(scaleY, 0.01) -
          nonNegativePixels(style.marginTop),
      );
      if (blockIndex === 0) {
        boundaries.push({ position: from, y });
      } else {
        boundaries.push({ position, y });
      }
    });
    const cellStyle = element ? getComputedStyle(element) : null;
    const cellRect = element?.getBoundingClientRect();
    const contentStart =
      boundaries[0]?.y ??
      (cellRect
        ? (cellRect.top - rowRect.top) / Math.max(scaleY, 0.01) +
          nonNegativePixels(cellStyle?.borderTopWidth ?? '') +
          nonNegativePixels(cellStyle?.paddingTop ?? '')
        : 0);
    if (!boundaries.length)
      boundaries.push({ position: from, y: contentStart });
    boundaries[0] = { position: from, y: contentStart };
    const contentEnd = cellRect
      ? (cellRect.bottom - rowRect.top) / Math.max(scaleY, 0.01) -
        nonNegativePixels(cellStyle?.borderBottomWidth ?? '') -
        nonNegativePixels(cellStyle?.paddingBottom ?? '')
      : row.height;
    boundaries.push({
      position: to,
      y: Math.max(contentStart, Math.min(row.height, contentEnd)),
    });
    result.push({
      cellIndex,
      from,
      to,
      boundaries,
      ...(cellRect && tableRect
        ? {
            tableOffsetLeft: Math.max(
              0,
              (cellRect.left - tableRect.left) / Math.max(scaleX, 0.01),
            ),
            outerWidth: Math.max(
              1,
              element?.offsetWidth || cellRect.width / Math.max(scaleX, 0.01),
            ),
            contentOffsetLeft:
              nonNegativePixels(cellStyle?.borderLeftWidth ?? '') +
              nonNegativePixels(cellStyle?.paddingLeft ?? ''),
          }
        : {}),
    });
  });
  return result;
}

function documentTablePaginationBreak(
  tableElement: HTMLElement,
  tableNode: ProseMirrorNode,
  tableId: string,
  rows: readonly MeasuredDocumentTableRow[],
  repeatHeaderCount: number,
  repeatHeaderHeight: number,
): DocumentTablePaginationBreak {
  const table =
    tableElement.tagName.toLowerCase() === 'table'
      ? tableElement
      : Array.from(tableElement.children).find(
          (child) => child.tagName.toLowerCase() === 'table',
        );
  const colgroup = Array.from(table?.children ?? []).find(
    (child) => child.tagName.toLowerCase() === 'colgroup',
  );
  const tableRect = table?.getBoundingClientRect();
  const scaleX =
    table instanceof HTMLElement && table.offsetWidth > 0 && tableRect
      ? tableRect.width / table.offsetWidth
      : 1;
  const scaleY =
    table instanceof HTMLElement && table.offsetHeight > 0 && tableRect
      ? tableRect.height / table.offsetHeight
      : 1;
  const columnCount = logicalDocumentTableColumnCount(tableNode);
  const tableWidth =
    table instanceof HTMLElement
      ? table.offsetWidth || (tableRect?.width ?? 0) / Math.max(scaleX, 0.01)
      : tableElement.offsetWidth;
  const repeatedHeaderRows = rows.slice(0, repeatHeaderCount);
  const leadingCell = rows[0]?.element.querySelector<HTMLElement>(
    ':scope > th:first-child, :scope > td:first-child',
  );
  const leadingCellRect = leadingCell?.getBoundingClientRect();
  const leadingCellStyle = leadingCell ? getComputedStyle(leadingCell) : null;
  return {
    tableId,
    columnCount,
    colgroupHtml: colgroup instanceof HTMLElement ? colgroup.outerHTML : '',
    repeatHeaderHeight,
    tableWidth,
    leadingCellOffsetLeft:
      leadingCellRect && tableRect
        ? (leadingCellRect.left - tableRect.left) / Math.max(scaleX, 0.01) +
          nonNegativePixels(leadingCellStyle?.borderLeftWidth ?? '') +
          nonNegativePixels(leadingCellStyle?.paddingLeft ?? '')
        : 0,
    repeatedHeaderRowsHtml: repeatedHeaderRows.map((row) =>
      repeatedDocumentTableRowHtml(row.element),
    ),
    repeatedHeaderOverlayHtml: repeatedDocumentTableHeaderOverlayHtml(
      repeatedHeaderRows,
      tableRect,
      scaleX,
      scaleY,
      tableWidth,
      columnCount,
    ),
  };
}

function logicalDocumentTableColumnCount(table: ProseMirrorNode): number {
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

function repeatedDocumentTableRowHtml(row: HTMLElement): string {
  const clone = row.cloneNode(true) as HTMLElement;
  sanitizeRepeatedDocumentTableElement(clone);
  return clone.outerHTML;
}

function repeatedDocumentTableHeaderOverlayHtml(
  rows: readonly MeasuredDocumentTableRow[],
  tableRect: DOMRect | undefined,
  scaleX: number,
  scaleY: number,
  tableWidth: number,
  columnCount: number,
): string {
  const container = document.createElement('div');
  const originTop = rows[0]?.element.getBoundingClientRect().top ?? 0;
  let fallbackTop = 0;
  for (const row of rows) {
    const rowRect = row.element.getBoundingClientRect();
    const rowTop =
      rowRect.height > 0
        ? Math.max(0, (rowRect.top - originTop) / Math.max(scaleY, 0.01))
        : fallbackTop;
    const cells = Array.from(row.element.children).filter(
      (element): element is HTMLElement =>
        element instanceof HTMLElement &&
        ['td', 'th'].includes(element.tagName.toLowerCase()),
    );
    let fallbackColumn = 0;
    for (const source of cells) {
      const colspan = Math.max(
        1,
        Number.parseInt(source.getAttribute('colspan') ?? '1', 10) || 1,
      );
      const cellRect = source.getBoundingClientRect();
      const fallbackLeft =
        (fallbackColumn / Math.max(1, columnCount)) * tableWidth;
      const fallbackWidth = (colspan / Math.max(1, columnCount)) * tableWidth;
      const left =
        tableRect && cellRect.width > 0
          ? (cellRect.left - tableRect.left) / Math.max(scaleX, 0.01)
          : fallbackLeft;
      const width =
        cellRect.width > 0
          ? source.offsetWidth || cellRect.width / Math.max(scaleX, 0.01)
          : fallbackWidth;
      const height =
        cellRect.height > 0
          ? source.offsetHeight || cellRect.height / Math.max(scaleY, 0.01)
          : row.height;
      container.append(
        repeatedDocumentTableHeaderCellElement(
          source,
          left,
          rowTop,
          width,
          height,
        ),
      );
      fallbackColumn += colspan;
    }
    fallbackTop += row.height;
  }
  return container.innerHTML;
}

function repeatedDocumentTableHeaderCellElement(
  source: HTMLElement,
  left: number,
  top: number,
  width: number,
  height: number,
): HTMLDivElement {
  const sourceClone = source.cloneNode(true) as HTMLElement;
  sanitizeRepeatedDocumentTableElement(sourceClone);
  const cell = document.createElement('div');
  cell.className = 'work-document-table-repeated-header-cell';
  cell.dataset.sourceCell = source.tagName.toLowerCase();
  cell.innerHTML = sourceClone.innerHTML;
  const computed = getComputedStyle(source);
  for (const property of REPEATED_TABLE_HEADER_STYLE_PROPERTIES) {
    const value = computed.getPropertyValue(property);
    if (value) cell.style.setProperty(property, value);
  }
  cell.style.left = `${left}px`;
  cell.style.top = `${top}px`;
  cell.style.width = `${Math.max(1, width)}px`;
  cell.style.height = `${Math.max(1, height)}px`;
  return cell;
}

const REPEATED_TABLE_HEADER_STYLE_PROPERTIES = [
  'background-color',
  'border-top',
  'border-right',
  'border-bottom',
  'border-left',
  'color',
  'font-family',
  'font-size',
  'font-style',
  'font-weight',
  'letter-spacing',
  'line-height',
  'padding-top',
  'padding-right',
  'padding-bottom',
  'padding-left',
  'text-align',
  'white-space',
] as const;

function sanitizeRepeatedDocumentTableElement(clone: HTMLElement): void {
  clone.removeAttribute('contenteditable');
  clone.removeAttribute('data-node-view-content-react');
  clone.classList.remove('selectedCell');
  for (const descendant of clone.querySelectorAll<HTMLElement>('*')) {
    descendant.removeAttribute('contenteditable');
    descendant.removeAttribute('data-node-view-content-react');
    descendant.classList.remove('selectedCell');
    descendant.classList.remove('ProseMirror-selectednode');
  }
  clone
    .querySelectorAll(
      [
        '.work-document-auto-page-break',
        '.work-document-table-page-break',
        '.work-document-table-cell-page-break',
        '.work-document-table-repeated-header',
      ].join(','),
    )
    .forEach((pageBreak) => {
      pageBreak.remove();
    });
}
