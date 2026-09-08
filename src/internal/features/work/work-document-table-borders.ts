import type { CommandProps } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { NodeSelection } from '@tiptap/pm/state';
import {
  isInTable,
  selectedRect,
  TableMap,
  type Rect,
} from '@tiptap/pm/tables';

export type DocumentTableBorderStyle =
  | 'solid'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'none';

export type DocumentTableBorderEdge = 'top' | 'right' | 'bottom' | 'left';

export type DocumentTableBorderTarget =
  | 'all'
  | 'outside'
  | 'inside'
  | 'top'
  | 'right'
  | 'bottom'
  | 'left'
  | 'insideHorizontal'
  | 'insideVertical';

export interface DocumentTableBorder {
  color: string;
  style: DocumentTableBorderStyle;
  width: number;
}

export type DocumentTableCellBorders = Record<
  DocumentTableBorderEdge,
  DocumentTableBorder
>;

const BORDER_EDGES: readonly DocumentTableBorderEdge[] = [
  'top',
  'right',
  'bottom',
  'left',
];

export function uniformDocumentTableBorders(
  border: DocumentTableBorder,
): DocumentTableCellBorders {
  return {
    top: { ...border },
    right: { ...border },
    bottom: { ...border },
    left: { ...border },
  };
}

export function documentTableBordersFromAttributes(
  attributes: Record<string, unknown>,
  fallback: DocumentTableBorder,
): DocumentTableCellBorders {
  const candidate = attributes.borders;
  if (!candidate || typeof candidate !== 'object' || Array.isArray(candidate)) {
    return uniformDocumentTableBorders(fallback);
  }
  const record = candidate as Record<string, unknown>;
  return {
    top: normalizeDocumentTableBorder(record.top) ?? { ...fallback },
    right: normalizeDocumentTableBorder(record.right) ?? { ...fallback },
    bottom: normalizeDocumentTableBorder(record.bottom) ?? { ...fallback },
    left: normalizeDocumentTableBorder(record.left) ?? { ...fallback },
  };
}

export function documentTableBordersFromElement(
  element: HTMLElement,
  fallback: DocumentTableBorder,
): DocumentTableCellBorders {
  const uniform = documentTableBorderFromElement(element, fallback);
  return {
    top: documentTableEdgeFromElement(element, 'top', uniform),
    right: documentTableEdgeFromElement(element, 'right', uniform),
    bottom: documentTableEdgeFromElement(element, 'bottom', uniform),
    left: documentTableEdgeFromElement(element, 'left', uniform),
  };
}

export function renderDocumentTableBorders(
  borders: DocumentTableCellBorders,
): Record<string, string> {
  const normalized = documentTableBordersFromAttributes(
    { borders },
    borders.top,
  );
  const representative =
    uniformDocumentTableBorder(normalized) ?? normalized.top;
  const attributes: Record<string, string> = {
    'data-office-cell-border-color': representative.color,
    'data-office-cell-border-style': representative.style,
    'data-office-cell-border-width': String(representative.width),
  };
  for (const edge of BORDER_EDGES) {
    const border = normalized[edge];
    attributes[`data-office-cell-border-${edge}-color`] = border.color;
    attributes[`data-office-cell-border-${edge}-style`] = border.style;
    attributes[`data-office-cell-border-${edge}-width`] = String(border.width);
  }
  const uniform = uniformDocumentTableBorder(normalized);
  attributes.style = uniform
    ? `border: ${documentTableBorderCss(uniform)}`
    : BORDER_EDGES.map(
        (edge) => `border-${edge}: ${documentTableBorderCss(normalized[edge])}`,
      ).join('; ');
  return attributes;
}

export function uniformDocumentTableBorder(
  borders: DocumentTableCellBorders,
): DocumentTableBorder | null {
  const first = borders.top;
  return BORDER_EDGES.every((edge) =>
    sameDocumentTableBorder(borders[edge], first),
  )
    ? first
    : null;
}

/**
 * Word-style shared-edge conflict: thicker width wins, then style weight, then
 * the first argument for stability. `none` always loses to a painted edge.
 * Storage stays independent; callers project winners for CSS paint only.
 */
export function resolveDocumentTableSharedBorder(
  first: DocumentTableBorder,
  second: DocumentTableBorder,
): DocumentTableBorder {
  const a = normalizeResolvedBorder(first);
  const b = normalizeResolvedBorder(second);
  if (a.style === 'none' && b.style === 'none') return a;
  if (a.style === 'none') return b;
  if (b.style === 'none') return a;
  if (a.width !== b.width) return a.width > b.width ? a : b;
  const weightA = documentTableBorderStyleWeight(a.style);
  const weightB = documentTableBorderStyleWeight(b.style);
  if (weightA !== weightB) return weightA > weightB ? a : b;
  return a;
}

/**
 * Projects independent per-cell borders into a paint grid where each shared
 * edge uses the Word-style winner on both abutting sides.
 */
export function projectDocumentTableBordersForPaint(
  grid: ReadonlyArray<ReadonlyArray<DocumentTableCellBorders>>,
): DocumentTableCellBorders[][] {
  const projected = grid.map((row) =>
    row.map((cell) => ({
      top: { ...cell.top },
      right: { ...cell.right },
      bottom: { ...cell.bottom },
      left: { ...cell.left },
    })),
  );
  for (let rowIndex = 0; rowIndex < projected.length; rowIndex += 1) {
    const row = projected[rowIndex];
    if (!row) continue;
    for (let cellIndex = 0; cellIndex < row.length; cellIndex += 1) {
      const cell = row[cellIndex];
      if (!cell) continue;
      const rightNeighbor = row[cellIndex + 1];
      if (rightNeighbor) {
        const winner = resolveDocumentTableSharedBorder(
          cell.right,
          rightNeighbor.left,
        );
        cell.right = { ...winner };
        rightNeighbor.left = { ...winner };
      }
      const below = projected[rowIndex + 1]?.[cellIndex];
      if (below) {
        const winner = resolveDocumentTableSharedBorder(cell.bottom, below.top);
        cell.bottom = { ...winner };
        below.top = { ...winner };
      }
    }
  }
  return projected;
}

/**
 * Applies shared-edge winners to CSS paint on an HTML table (including
 * colspan/rowspan occupancy) without mutating `data-office-cell-border-*`
 * storage used for DOCX export. Irregular grids fail soft.
 */
export function applyDocumentTableSharedBorderPaint(
  table: HTMLTableElement,
): void {
  const occupancy = tableCellOccupancy(table);
  if (!occupancy) return;
  const fallback: DocumentTableBorder = {
    color: '#cfd5df',
    style: 'solid',
    width: 1,
  };
  const paint = new Map<HTMLTableCellElement, DocumentTableCellBorders>();
  for (const placement of occupancy.placements) {
    paint.set(
      placement.cell,
      documentTableBordersFromElement(placement.cell, fallback),
    );
  }
  const { grid, rowCount, columnCount } = occupancy;
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount - 1; columnIndex += 1) {
      const left = grid[rowIndex]?.[columnIndex];
      const right = grid[rowIndex]?.[columnIndex + 1];
      if (!left || !right || left === right) continue;
      const leftPlacement = occupancy.byCell.get(left);
      const rightPlacement = occupancy.byCell.get(right);
      if (
        !leftPlacement ||
        !rightPlacement ||
        leftPlacement.originColumn + leftPlacement.colSpan - 1 !==
          columnIndex ||
        rightPlacement.originColumn !== columnIndex + 1
      ) {
        continue;
      }
      const leftPaint = paint.get(left);
      const rightPaint = paint.get(right);
      if (!leftPaint || !rightPaint) continue;
      const winner = resolveDocumentTableSharedBorder(
        leftPaint.right,
        rightPaint.left,
      );
      leftPaint.right = { ...winner };
      rightPaint.left = { ...winner };
    }
  }
  for (let rowIndex = 0; rowIndex < rowCount - 1; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const above = grid[rowIndex]?.[columnIndex];
      const below = grid[rowIndex + 1]?.[columnIndex];
      if (!above || !below || above === below) continue;
      const abovePlacement = occupancy.byCell.get(above);
      const belowPlacement = occupancy.byCell.get(below);
      if (
        !abovePlacement ||
        !belowPlacement ||
        abovePlacement.originRow + abovePlacement.rowSpan - 1 !== rowIndex ||
        belowPlacement.originRow !== rowIndex + 1
      ) {
        continue;
      }
      const abovePaint = paint.get(above);
      const belowPaint = paint.get(below);
      if (!abovePaint || !belowPaint) continue;
      const winner = resolveDocumentTableSharedBorder(
        abovePaint.bottom,
        belowPaint.top,
      );
      abovePaint.bottom = { ...winner };
      belowPaint.top = { ...winner };
    }
  }
  for (const [cell, borders] of paint) {
    for (const edge of BORDER_EDGES) {
      cell.style.setProperty(
        `border-${edge}`,
        documentTableBorderCss(borders[edge]),
      );
    }
  }
}

interface TableCellPlacement {
  cell: HTMLTableCellElement;
  originRow: number;
  originColumn: number;
  rowSpan: number;
  colSpan: number;
}

interface TableCellOccupancy {
  byCell: Map<HTMLTableCellElement, TableCellPlacement>;
  grid: Array<Array<HTMLTableCellElement | null>>;
  placements: TableCellPlacement[];
  rowCount: number;
  columnCount: number;
}

function tableCellOccupancy(
  table: HTMLTableElement,
): TableCellOccupancy | null {
  const rows = Array.from(table.rows);
  if (!rows.length) return null;
  const grid: Array<Array<HTMLTableCellElement | null>> = [];
  const byCell = new Map<HTMLTableCellElement, TableCellPlacement>();
  const placements: TableCellPlacement[] = [];
  let columnCount = 0;

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    if (!row) return null;
    if (!grid[rowIndex]) grid[rowIndex] = [];
    const cells = Array.from(row.cells).filter(
      (cell): cell is HTMLTableCellElement =>
        cell instanceof HTMLTableCellElement,
    );
    let columnIndex = 0;
    for (const cell of cells) {
      const occupied = grid[rowIndex] ?? (grid[rowIndex] = []);
      while (occupied[columnIndex]) columnIndex += 1;
      const colSpan = Math.max(1, cell.colSpan || 1);
      const rowSpan = Math.max(1, cell.rowSpan || 1);
      if (
        !Number.isSafeInteger(colSpan) ||
        !Number.isSafeInteger(rowSpan) ||
        colSpan > 64 ||
        rowSpan > 64 ||
        byCell.has(cell)
      ) {
        return null;
      }
      const placement: TableCellPlacement = {
        cell,
        originRow: rowIndex,
        originColumn: columnIndex,
        rowSpan,
        colSpan,
      };
      byCell.set(cell, placement);
      placements.push(placement);
      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        const targetRow = rowIndex + rowOffset;
        if (!grid[targetRow]) grid[targetRow] = [];
        for (let columnOffset = 0; columnOffset < colSpan; columnOffset += 1) {
          const targetColumn = columnIndex + columnOffset;
          if (grid[targetRow]?.[targetColumn]) return null;
          grid[targetRow]![targetColumn] = cell;
        }
      }
      columnIndex += colSpan;
      columnCount = Math.max(columnCount, columnIndex);
    }
  }

  const rowCount = grid.length;
  if (!rowCount || !columnCount) return null;
  for (const row of grid) {
    while (row.length < columnCount) row.push(null);
    if (row.length !== columnCount) return null;
  }
  return { byCell, grid, placements, rowCount, columnCount };
}

function normalizeResolvedBorder(
  border: DocumentTableBorder,
): DocumentTableBorder {
  const style = border.style === 'none' ? 'none' : border.style;
  return {
    color: border.color,
    style,
    width: style === 'none' ? 0 : border.width,
  };
}

function documentTableBorderStyleWeight(
  style: DocumentTableBorderStyle,
): number {
  switch (style) {
    case 'none':
      return 0;
    case 'dotted':
      return 1;
    case 'dashed':
      return 2;
    case 'solid':
      return 3;
    case 'double':
      return 4;
    default:
      return 0;
  }
}

export function normalizeDocumentTableBorder(
  value: unknown,
): DocumentTableBorder | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const color = normalizeTableColor(String(record.color ?? ''));
  const style = normalizeDocumentTableBorderStyle(String(record.style ?? ''));
  const width = normalizeDocumentTableBorderWidth(record.width);
  if (!color || !style || width === null) return null;
  return {
    color,
    style,
    width: style === 'none' ? 0 : width,
  };
}

export function normalizeTableColor(
  value: string | null | undefined,
): string | null {
  const trimmed = value?.trim().toLowerCase();
  if (!trimmed) return null;
  if (/^#[0-9a-f]{3}$/.test(trimmed)) {
    return `#${trimmed
      .slice(1)
      .split('')
      .map((part) => `${part}${part}`)
      .join('')}`;
  }
  if (/^#[0-9a-f]{6}$/.test(trimmed)) return trimmed;
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(trimmed);
  if (!rgb) return null;
  const channels = rgb.slice(1, 4).map(Number);
  if (channels.some((channel) => channel < 0 || channel > 255)) return null;
  return `#${channels
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}

export function normalizeDocumentTableBorderStyle(
  value: string | null | undefined,
): DocumentTableBorderStyle | null {
  if (
    value === 'solid' ||
    value === 'dashed' ||
    value === 'dotted' ||
    value === 'double' ||
    value === 'none'
  ) {
    return value;
  }
  return null;
}

export function normalizeDocumentTableBorderWidth(
  value: unknown,
): number | null {
  const width =
    typeof value === 'number'
      ? value
      : Number.parseFloat(String(value ?? '').replace(/px$/i, ''));
  if (!Number.isFinite(width) || width < 0 || width > 6) return null;
  return Math.round(width * 2) / 2;
}

export function setSelectedDocumentTableBorders(
  { dispatch, state }: CommandProps,
  target: DocumentTableBorderTarget,
  requestedBorder: DocumentTableBorder,
): boolean {
  const border = normalizeDocumentTableBorder(requestedBorder);
  const context = selectedTableBorderContext(state);
  if (!border || !context) return false;
  const transaction = state.tr;
  const positions = context.map.cellsInRect(context.rectangle);
  for (const relativePosition of positions) {
    const position = context.tableStart + relativePosition;
    const cell = transaction.doc.nodeAt(position);
    if (!cell || !isTableCell(cell)) continue;
    const cellRectangle = context.map.findCell(relativePosition);
    const edges = documentTableBorderTargetEdges(
      target,
      cellRectangle,
      context.rectangle,
    );
    if (!edges.length) continue;
    const fallback = documentTableBorderFromAttributes(cell.attrs);
    const borders = documentTableBordersFromAttributes(cell.attrs, fallback);
    for (const edge of edges) borders[edge] = { ...border };
    const representative = uniformDocumentTableBorder(borders) ?? borders.top;
    transaction.setNodeMarkup(position, undefined, {
      ...cell.attrs,
      borderColor: representative.color,
      borderStyle: representative.style,
      borderWidth: representative.width,
      borders,
    });
  }
  if (dispatch && transaction.docChanged) {
    dispatch(transaction.scrollIntoView());
  }
  return true;
}

function selectedTableBorderContext(state: CommandProps['state']): {
  tableStart: number;
  map: TableMap;
  rectangle: Rect;
} | null {
  const selection = state.selection;
  if (
    selection instanceof NodeSelection &&
    selection.node.type.spec.tableRole === 'table'
  ) {
    const map = TableMap.get(selection.node);
    return {
      tableStart: selection.from + 1,
      map,
      rectangle: { left: 0, right: map.width, top: 0, bottom: map.height },
    };
  }
  if (!isInTable(state)) return null;
  const rectangle = selectedRect(state);
  return {
    tableStart: rectangle.tableStart,
    map: rectangle.map,
    rectangle,
  };
}

function documentTableBorderTargetEdges(
  target: DocumentTableBorderTarget,
  cell: Rect,
  selection: Rect,
): DocumentTableBorderEdge[] {
  if (target === 'all') return [...BORDER_EDGES];
  const edges: DocumentTableBorderEdge[] = [];
  const outside = target === 'outside';
  const inside = target === 'inside';
  if (
    (target === 'top' && cell.top === selection.top) ||
    (outside && cell.top === selection.top) ||
    ((inside || target === 'insideHorizontal') && cell.top > selection.top)
  ) {
    edges.push('top');
  }
  if (
    (target === 'right' && cell.right === selection.right) ||
    (outside && cell.right === selection.right) ||
    ((inside || target === 'insideVertical') && cell.right < selection.right)
  ) {
    edges.push('right');
  }
  if (
    (target === 'bottom' && cell.bottom === selection.bottom) ||
    (outside && cell.bottom === selection.bottom) ||
    ((inside || target === 'insideHorizontal') &&
      cell.bottom < selection.bottom)
  ) {
    edges.push('bottom');
  }
  if (
    (target === 'left' && cell.left === selection.left) ||
    (outside && cell.left === selection.left) ||
    ((inside || target === 'insideVertical') && cell.left > selection.left)
  ) {
    edges.push('left');
  }
  return edges;
}

function documentTableBorderFromAttributes(
  attributes: Record<string, unknown>,
): DocumentTableBorder {
  return {
    color:
      normalizeTableColor(String(attributes.borderColor ?? '')) ?? '#cfd5df',
    style:
      normalizeDocumentTableBorderStyle(String(attributes.borderStyle ?? '')) ??
      'solid',
    width: normalizeDocumentTableBorderWidth(attributes.borderWidth) ?? 1,
  };
}

function documentTableBorderFromElement(
  element: HTMLElement,
  fallback: DocumentTableBorder,
): DocumentTableBorder {
  const style =
    normalizeDocumentTableBorderStyle(
      element.dataset.officeCellBorderStyle || element.style.borderStyle,
    ) ?? fallback.style;
  return {
    color:
      normalizeTableColor(
        element.dataset.officeCellBorderColor || element.style.borderColor,
      ) ?? fallback.color,
    style,
    width:
      style === 'none'
        ? 0
        : (normalizeDocumentTableBorderWidth(
            element.dataset.officeCellBorderWidth || element.style.borderWidth,
          ) ?? fallback.width),
  };
}

function documentTableEdgeFromElement(
  element: HTMLElement,
  edge: DocumentTableBorderEdge,
  fallback: DocumentTableBorder,
): DocumentTableBorder {
  const datasetPrefix = `officeCellBorder${capitalize(edge)}`;
  const style =
    normalizeDocumentTableBorderStyle(
      element.dataset[`${datasetPrefix}Style`] ||
        element.style.getPropertyValue(`border-${edge}-style`),
    ) ?? fallback.style;
  return {
    color:
      normalizeTableColor(
        element.dataset[`${datasetPrefix}Color`] ||
          element.style.getPropertyValue(`border-${edge}-color`),
      ) ?? fallback.color,
    style,
    width:
      style === 'none'
        ? 0
        : (normalizeDocumentTableBorderWidth(
            element.dataset[`${datasetPrefix}Width`] ||
              element.style.getPropertyValue(`border-${edge}-width`),
          ) ?? fallback.width),
  };
}

function documentTableBorderCss(border: DocumentTableBorder): string {
  return border.style === 'none' || border.width === 0
    ? '0px none transparent'
    : `${border.width}px ${border.style} ${border.color}`;
}

function sameDocumentTableBorder(
  left: DocumentTableBorder,
  right: DocumentTableBorder,
): boolean {
  return (
    left.color === right.color &&
    left.style === right.style &&
    left.width === right.width
  );
}

function isTableCell(node: ProseMirrorNode): boolean {
  return (
    node.type.spec.tableRole === 'cell' ||
    node.type.spec.tableRole === 'header_cell'
  );
}

function capitalize(value: string): string {
  return `${value.slice(0, 1).toUpperCase()}${value.slice(1)}`;
}
