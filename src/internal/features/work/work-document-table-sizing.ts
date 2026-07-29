import { type CommandProps, Extension } from '@tiptap/core';
import { Table, TableView } from '@tiptap/extension-table';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Plugin, type EditorState } from '@tiptap/pm/state';
import { NodeSelection } from '@tiptap/pm/state';
import { isInTable, selectedRect, TableMap } from '@tiptap/pm/tables';
import {
  documentTableRowHeight,
  type DocumentTableRowHeightRule,
} from './work-document-table-row';

export type DocumentTableLayoutMode = 'window' | 'contents' | 'fixed';

export interface DocumentTableSizingState {
  columnWidth: number | null;
  rowHeight: number | null;
  rowHeightRule: DocumentTableRowHeightRule | null;
  layoutMode: DocumentTableLayoutMode;
  selectedColumnCount: number;
  selectedRowCount: number;
}

const DEFAULT_COLUMN_WIDTH = 120;
const DEFAULT_ROW_HEIGHT = 36;
const MIN_COLUMN_WIDTH = 25;
const MIN_ROW_HEIGHT = 12;
const MAX_TABLE_DIMENSION = 4_000;

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTableSizing: {
      setDocumentTableColumnWidth: (
        width: number,
        renderedColumnWidths?: readonly number[],
      ) => ReturnType;
      setDocumentTableRowHeight: (
        height: number | null,
        rule?: DocumentTableRowHeightRule,
      ) => ReturnType;
      setDocumentTableLayoutMode: (
        mode: DocumentTableLayoutMode,
        renderedTableWidth?: number,
      ) => ReturnType;
      distributeDocumentTableColumns: (
        renderedSelectionWidth?: number,
      ) => ReturnType;
      distributeDocumentTableRows: (
        renderedSelectionHeight?: number,
      ) => ReturnType;
    };
  }
}

class WorkDocumentTableView extends TableView {
  constructor(
    node: ProseMirrorNode,
    cellMinWidth: number,
    view?: ConstructorParameters<typeof TableView>[2],
  ) {
    super(node, cellMinWidth, view);
    this.syncLayout(node);
  }

  override update(node: ProseMirrorNode): boolean {
    const updated = super.update(node);
    if (updated) this.syncLayout(node);
    return updated;
  }

  private syncLayout(node: ProseMirrorNode): void {
    const mode = normalizeDocumentTableLayoutMode(node.attrs.layoutMode);
    this.table.dataset.officeTableLayout = mode;
    if (mode === 'window') {
      this.clearStaleColumnWidths();
      this.table.style.width = '100%';
    } else if (mode === 'contents') {
      this.clearStaleColumnWidths();
      this.table.style.width = '';
    }
  }

  private clearStaleColumnWidths(): void {
    for (const column of Array.from(this.colgroup.children)) {
      if (!(column instanceof HTMLElement)) continue;
      column.style.removeProperty('width');
      column.style.setProperty('min-width', `${this.cellMinWidth}px`);
    }
  }
}

export const DocumentTable = Table.extend({
  addAttributes() {
    return {
      ...(this.parent?.() ?? {}),
      layoutMode: {
        default: 'window',
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentTableLayoutMode(element.dataset.officeTableLayout),
        renderHTML: (attributes: Record<string, unknown>) => ({
          'data-office-table-layout': normalizeDocumentTableLayoutMode(
            attributes.layoutMode,
          ),
        }),
      },
    };
  },
}).configure({
  resizable: true,
  allowTableNodeSelection: true,
  View: WorkDocumentTableView,
});

export const DocumentTableSizing = Extension.create({
  name: 'documentTableSizing',

  addCommands() {
    return {
      setDocumentTableColumnWidth: (width, renderedColumnWidths) => (props) =>
        setSelectedColumnWidth(props, width, renderedColumnWidths),
      setDocumentTableRowHeight:
        (height, rule = 'atLeast') =>
        (props) =>
          setSelectedRowHeight(props, height, rule),
      setDocumentTableLayoutMode: (mode, renderedTableWidth) => (props) =>
        setTableLayoutMode(props, mode, renderedTableWidth),
      distributeDocumentTableColumns: (renderedSelectionWidth) => (props) =>
        distributeColumns(props, renderedSelectionWidth),
      distributeDocumentTableRows: (renderedSelectionHeight) => (props) =>
        distributeRows(props, renderedSelectionHeight),
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        appendTransaction: (transactions, _oldState, newState) => {
          if (!transactions.some((transaction) => transaction.docChanged)) {
            return null;
          }
          const transaction = newState.tr;
          newState.doc.descendants((node, position) => {
            if (node.type.spec.tableRole !== 'table') return true;
            if (
              normalizeDocumentTableLayoutMode(node.attrs.layoutMode) ===
                'window' &&
              tableHasExplicitColumnWidths(node)
            ) {
              transaction.setNodeMarkup(position, undefined, {
                ...node.attrs,
                layoutMode: 'fixed',
              });
            }
            return false;
          });
          return transaction.docChanged ? transaction : null;
        },
      }),
    ];
  },
});

export function documentTableSizing(
  state: EditorState,
): DocumentTableSizingState | null {
  const context = tableSizingContext(state);
  if (!context) return null;
  const columns = numberRange(context.left, context.right);
  const rows = numberRange(context.top, context.bottom);
  const widths = tableColumnWidths(context.table, context.map).filter(
    (_width, index) => columns.includes(index),
  );
  const rowSizes = tableRowSizes(context.table).filter((_size, index) =>
    rows.includes(index),
  );
  return {
    columnWidth: commonPositiveValue(widths),
    rowHeight: commonNullableValue(rowSizes.map(({ height }) => height)),
    rowHeightRule: commonNullableValue(rowSizes.map(({ rule }) => rule)),
    layoutMode: normalizeDocumentTableLayoutMode(
      context.table.attrs.layoutMode,
    ),
    selectedColumnCount: columns.length,
    selectedRowCount: rows.length,
  };
}

export function normalizeDocumentTableLayoutMode(
  value: unknown,
): DocumentTableLayoutMode {
  return value === 'contents' || value === 'fixed' ? value : 'window';
}

function setSelectedColumnWidth(
  { dispatch, state }: CommandProps,
  width: number,
  renderedColumnWidths?: readonly number[],
): boolean {
  const normalized = normalizedDimension(
    width,
    MIN_COLUMN_WIDTH,
    MAX_TABLE_DIMENSION,
  );
  const context = tableSizingContext(state);
  if (!context || normalized === null) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  const preservedWidths = normalizeRenderedDimensions(
    renderedColumnWidths,
    context.map.width,
    MIN_COLUMN_WIDTH,
  );
  if (preservedWidths) {
    for (let column = context.left; column < context.right; column += 1) {
      preservedWidths[column] = normalized;
    }
    setPhysicalColumnWidths(transaction, context, preservedWidths, 0);
  } else {
    setPhysicalColumnWidths(
      transaction,
      context,
      numberRange(context.left, context.right).map(() => normalized),
      context.left,
    );
  }
  setTableLayoutAttribute(transaction, context, 'fixed');
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setSelectedRowHeight(
  { dispatch, state }: CommandProps,
  height: number | null,
  rule: DocumentTableRowHeightRule,
): boolean {
  const normalized =
    height === null
      ? null
      : normalizedDimension(height, MIN_ROW_HEIGHT, MAX_TABLE_DIMENSION);
  const context = tableSizingContext(state);
  if (!context || (height !== null && normalized === null)) return false;
  if (rule !== 'atLeast' && rule !== 'exact') return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  setPhysicalRowHeights(
    transaction,
    context,
    numberRange(context.top, context.bottom),
    normalized,
    rule,
  );
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setTableLayoutMode(
  { dispatch, state }: CommandProps,
  requestedMode: DocumentTableLayoutMode,
  renderedTableWidth?: number,
): boolean {
  if (
    requestedMode !== 'window' &&
    requestedMode !== 'contents' &&
    requestedMode !== 'fixed'
  ) {
    return false;
  }
  const context = tableSizingContext(state);
  if (!context) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  setTableLayoutAttribute(transaction, context, requestedMode);
  if (requestedMode === 'fixed') {
    const existing = tableColumnWidths(context.table, context.map);
    const total = distributionTotal(
      existing,
      renderedTableWidth,
      DEFAULT_COLUMN_WIDTH,
    );
    setPhysicalColumnWidths(
      transaction,
      context,
      equalDimensions(total, context.map.width, MIN_COLUMN_WIDTH),
      0,
    );
  } else {
    clearPhysicalColumnWidths(transaction, context);
  }
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function distributeColumns(
  { dispatch, state }: CommandProps,
  renderedSelectionWidth?: number,
): boolean {
  const context = tableSizingContext(state);
  if (!context) return false;
  const selected = numberRange(context.left, context.right);
  const columns =
    selected.length > 1 ? selected : numberRange(0, context.map.width);
  if (columns.length < 2) return false;
  if (!dispatch) return true;
  const allWidths = tableColumnWidths(context.table, context.map);
  const existing = columns.map((column) => allWidths[column] ?? null);
  const total = distributionTotal(
    existing,
    renderedSelectionWidth,
    DEFAULT_COLUMN_WIDTH,
  );
  const transaction = state.tr;
  setPhysicalColumnWidths(
    transaction,
    context,
    equalDimensions(total, columns.length, MIN_COLUMN_WIDTH),
    columns[0] ?? 0,
  );
  setTableLayoutAttribute(transaction, context, 'fixed');
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function distributeRows(
  { dispatch, state }: CommandProps,
  renderedSelectionHeight?: number,
): boolean {
  const context = tableSizingContext(state);
  if (!context) return false;
  const selected = numberRange(context.top, context.bottom);
  const rows =
    selected.length > 1 ? selected : numberRange(0, context.map.height);
  if (rows.length < 2) return false;
  if (!dispatch) return true;
  const allSizes = tableRowSizes(context.table);
  const existing = rows.map((row) => allSizes[row]?.height ?? null);
  const total = distributionTotal(
    existing,
    renderedSelectionHeight,
    DEFAULT_ROW_HEIGHT,
  );
  const [height] = equalDimensions(total, rows.length, MIN_ROW_HEIGHT);
  const transaction = state.tr;
  setPhysicalRowHeights(
    transaction,
    context,
    rows,
    height ?? DEFAULT_ROW_HEIGHT,
    'atLeast',
  );
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

interface TableSizingContext {
  table: ProseMirrorNode;
  tableStart: number;
  map: TableMap;
  left: number;
  right: number;
  top: number;
  bottom: number;
}

function tableSizingContext(state: EditorState): TableSizingContext | null {
  if (
    state.selection instanceof NodeSelection &&
    state.selection.node.type.spec.tableRole === 'table'
  ) {
    const table = state.selection.node;
    const map = TableMap.get(table);
    return {
      table,
      tableStart: state.selection.from + 1,
      map,
      left: 0,
      right: map.width,
      top: 0,
      bottom: map.height,
    };
  }
  if (!isInTable(state)) return null;
  const rectangle = selectedRect(state);
  return {
    table: rectangle.table,
    tableStart: rectangle.tableStart,
    map: rectangle.map,
    left: rectangle.left,
    right: rectangle.right,
    top: rectangle.top,
    bottom: rectangle.bottom,
  };
}

function dispatchTableSizingTransaction(
  dispatch: NonNullable<CommandProps['dispatch']>,
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  state: EditorState,
): void {
  if (!transaction.docChanged) return;
  if (
    state.selection instanceof NodeSelection &&
    state.selection.node.type.spec.tableRole === 'table'
  ) {
    const tablePosition = context.tableStart - 1;
    if (
      transaction.doc.nodeAt(tablePosition)?.type.spec.tableRole === 'table'
    ) {
      transaction.setSelection(
        NodeSelection.create(transaction.doc, tablePosition),
      );
    }
  }
  dispatch(transaction.scrollIntoView());
}

function setPhysicalColumnWidths(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  widths: readonly number[],
  startColumn: number,
): void {
  const touched = new Set<string>();
  widths.forEach((width, widthIndex) => {
    const column = startColumn + widthIndex;
    if (column < 0 || column >= context.map.width) return;
    for (let row = 0; row < context.map.height; row += 1) {
      const relativePosition =
        context.map.map[row * context.map.width + column];
      if (relativePosition === undefined) continue;
      const cell = context.table.nodeAt(relativePosition);
      if (!cell) continue;
      const widthIndexInCell = column - context.map.colCount(relativePosition);
      const key = `${relativePosition}:${widthIndexInCell}`;
      if (touched.has(key)) continue;
      touched.add(key);
      const documentPosition = context.tableStart + relativePosition;
      const current = transaction.doc.nodeAt(documentPosition);
      if (!current) continue;
      const colwidth = Array.isArray(current.attrs.colwidth)
        ? [...current.attrs.colwidth]
        : Array.from({ length: Number(current.attrs.colspan ?? 1) }, () => 0);
      colwidth[widthIndexInCell] = width;
      transaction.setNodeMarkup(documentPosition, undefined, {
        ...current.attrs,
        colwidth,
      });
    }
  });
}

function clearPhysicalColumnWidths(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
): void {
  context.table.descendants((node, offset) => {
    if (!isTableCell(node)) return true;
    const position = context.tableStart + offset;
    const current = transaction.doc.nodeAt(position);
    if (current?.attrs.colwidth) {
      transaction.setNodeMarkup(position, undefined, {
        ...current.attrs,
        colwidth: null,
      });
    }
    return false;
  });
}

function setPhysicalRowHeights(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  rows: readonly number[],
  height: number | null,
  rule: DocumentTableRowHeightRule,
): void {
  const selectedRows = new Set(rows);
  context.table.forEach((_row, offset, index) => {
    if (!selectedRows.has(index)) return;
    const position = context.tableStart + offset;
    const current = transaction.doc.nodeAt(position);
    if (!current) return;
    transaction.setNodeMarkup(position, undefined, {
      ...current.attrs,
      rowHeight: height,
      rowHeightRule: height === null ? null : rule,
    });
  });
}

function setTableLayoutAttribute(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  layoutMode: DocumentTableLayoutMode,
): void {
  const position = context.tableStart - 1;
  const table = transaction.doc.nodeAt(position);
  if (!table) return;
  transaction.setNodeMarkup(position, undefined, {
    ...table.attrs,
    layoutMode,
  });
}

function tableColumnWidths(
  table: ProseMirrorNode,
  map: TableMap,
): Array<number | null> {
  return Array.from({ length: map.width }, (_unused, column) => {
    for (let row = 0; row < map.height; row += 1) {
      const position = map.map[row * map.width + column];
      if (position === undefined) continue;
      const cell = table.nodeAt(position);
      const index = column - map.colCount(position);
      const value = Array.isArray(cell?.attrs.colwidth)
        ? Number(cell.attrs.colwidth[index])
        : 0;
      if (Number.isFinite(value) && value > 0) return value;
    }
    return null;
  });
}

function tableRowSizes(table: ProseMirrorNode): Array<{
  height: number | null;
  rule: DocumentTableRowHeightRule | null;
}> {
  const sizes: Array<{
    height: number | null;
    rule: DocumentTableRowHeightRule | null;
  }> = [];
  table.forEach((row) => {
    sizes.push(documentTableRowHeight(row.attrs));
  });
  return sizes;
}

function tableHasExplicitColumnWidths(table: ProseMirrorNode): boolean {
  let found = false;
  table.descendants((node) => {
    if (!isTableCell(node)) return true;
    if (
      Array.isArray(node.attrs.colwidth) &&
      node.attrs.colwidth.some((width: unknown) => Number(width) > 0)
    ) {
      found = true;
    }
    return !found;
  });
  return found;
}

function isTableCell(node: ProseMirrorNode): boolean {
  return (
    node.type.spec.tableRole === 'cell' ||
    node.type.spec.tableRole === 'header_cell'
  );
}

function normalizedDimension(
  value: number,
  minimum: number,
  maximum: number,
): number | null {
  if (!Number.isFinite(value) || value < minimum || value > maximum)
    return null;
  return Math.round(value * 100) / 100;
}

function normalizeRenderedDimensions(
  values: readonly number[] | undefined,
  expectedCount: number,
  minimum: number,
): number[] | null {
  if (!values || values.length !== expectedCount) return null;
  const normalized = values.map((value) =>
    normalizedDimension(value, minimum, MAX_TABLE_DIMENSION),
  );
  return normalized.every((value) => value !== null)
    ? (normalized as number[])
    : null;
}

function numberRange(start: number, end: number): number[] {
  return Array.from(
    { length: Math.max(0, end - start) },
    (_unused, index) => start + index,
  );
}

function distributionTotal(
  existing: readonly (number | null)[],
  rendered: number | undefined,
  fallback: number,
): number {
  if (Number.isFinite(rendered) && Number(rendered) > 0) {
    return Number(rendered);
  }
  if (existing.every((value) => value !== null && value > 0)) {
    return existing.reduce<number>((sum, value) => sum + Number(value), 0);
  }
  return fallback * existing.length;
}

function equalDimensions(
  total: number,
  count: number,
  minimum: number,
): number[] {
  const size = Math.max(minimum, Math.round((total / count) * 100) / 100);
  return Array.from({ length: count }, () => size);
}

function commonPositiveValue(
  values: readonly (number | null)[],
): number | null {
  if (!values.length || values.some((value) => value === null || value <= 0)) {
    return null;
  }
  return values.every((value) => value === values[0]) ? values[0] : null;
}

function commonNullableValue<T>(values: readonly (T | null)[]): T | null {
  if (!values.length || values.some((value) => value === null)) return null;
  return values.every((value) => value === values[0]) ? values[0] : null;
}
