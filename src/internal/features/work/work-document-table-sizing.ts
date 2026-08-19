import { type CommandProps, Extension } from '@tiptap/core';
import { Table, TableView } from '@tiptap/extension-table';
import { closeHistory } from '@tiptap/pm/history';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Plugin, type EditorState } from '@tiptap/pm/state';
import { NodeSelection } from '@tiptap/pm/state';
import { isInTable, selectedRect, TableMap } from '@tiptap/pm/tables';
import {
  documentTableRowHeight,
  documentTableRowRepeats,
  type DocumentTableRowHeightRule,
} from './work-document-table-row';
import {
  applyDocumentTableGeometryToElement,
  documentTableGeometryForLayoutMode,
  documentTableGeometryFromElement,
  documentTableLayoutMode,
  normalizeDocumentTableAlignment,
  normalizeDocumentTableCellMargins,
  normalizeDocumentTableGeometry,
  normalizeDocumentTableProperties,
  renderDocumentTableGeometry,
  type DocumentTableAlignment,
  type DocumentTableCellMargins,
  type DocumentTableGeometry,
  type DocumentTableLayoutAlgorithm,
  type DocumentTableLayoutMode,
  type DocumentTablePreferredWidthType,
  type DocumentTableProperties,
} from './work-document-table-geometry';
import {
  MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
  MIN_DOCUMENT_TABLE_ROW_HEIGHT,
  normalizeDocumentTableDimension,
  normalizeDocumentTablePropertyChanges,
  type DocumentTablePropertyChanges,
} from './work-document-table-property-changes';
import {
  normalizeDocumentTableColumnPercent,
  normalizeDocumentTableColumnPercentages,
  type DocumentTableColumnWidthType,
} from './work-document-table-column-widths';

export {
  normalizeDocumentTableGeometry,
  normalizeDocumentTableLayoutMode,
} from './work-document-table-geometry';
export type {
  DocumentTableAlignment,
  DocumentTableCellMargins,
  DocumentTableGeometry,
  DocumentTableLayoutAlgorithm,
  DocumentTableLayoutMode,
  DocumentTablePreferredWidth,
  DocumentTablePreferredWidthType,
  DocumentTableProperties,
} from './work-document-table-geometry';
export type { DocumentTablePropertyChanges } from './work-document-table-property-changes';

export interface DocumentTableSizingState {
  columnWidth: number | null;
  columnWidthType: DocumentTableColumnWidthType;
  rowHeight: number | null;
  rowHeightRule: DocumentTableRowHeightRule | null;
  layoutMode: DocumentTableLayoutMode;
  layoutAlgorithm: DocumentTableLayoutAlgorithm;
  preferredWidthType: DocumentTablePreferredWidthType;
  preferredWidth: number | null;
  alignment: DocumentTableAlignment;
  indent: number;
  cellMargins: DocumentTableCellMargins;
  selectedColumnCount: number;
  selectedRowCount: number;
}

const DEFAULT_COLUMN_WIDTH = 120;
const DEFAULT_ROW_HEIGHT = 36;

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTableSizing: {
      setDocumentTableColumnWidth: (
        width: number,
        renderedColumnWidths?: readonly number[],
      ) => ReturnType;
      setDocumentTableColumnWidthPercent: (
        width: number,
        renderedColumnWidths?: readonly number[],
        renderedTableWidth?: number,
      ) => ReturnType;
      setDocumentTableRowHeight: (
        height: number | null,
        rule?: DocumentTableRowHeightRule,
      ) => ReturnType;
      setDocumentTableLayoutMode: (
        mode: DocumentTableLayoutMode,
        renderedTableWidth?: number,
      ) => ReturnType;
      setDocumentTableAlignment: (
        alignment: DocumentTableAlignment,
      ) => ReturnType;
      setDocumentTableProperties: (
        properties: DocumentTableProperties,
      ) => ReturnType;
      setDocumentTablePropertyChanges: (
        changes: DocumentTablePropertyChanges,
      ) => ReturnType;
      setDocumentTableCellMargins: (
        margins: DocumentTableCellMargins,
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
    this.syncGeometry(node);
  }

  override update(node: ProseMirrorNode): boolean {
    const updated = super.update(node);
    if (updated) this.syncGeometry(node);
    return updated;
  }

  private syncGeometry(node: ProseMirrorNode): void {
    const geometry = tableGeometry(node);
    applyDocumentTableGeometryToElement(this.table, geometry);
    if (node.attrs.officeImported) {
      this.table.dataset.officeTableImported = 'true';
    } else {
      delete this.table.dataset.officeTableImported;
    }
    if (!tableHasExplicitColumnWidths(node)) this.clearStaleColumnWidths();
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
      geometry: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentTableGeometryFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) =>
          renderDocumentTableGeometry(
            attributes.geometry,
            attributes.layoutMode,
          ),
      },
      layoutMode: {
        default: null,
        parseHTML: () => null,
        renderHTML: () => ({}),
      },
      officeImported: {
        default: false,
        parseHTML: (element: HTMLElement) =>
          element.dataset.officeTableImported === 'true',
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes.officeImported
            ? { 'data-office-table-imported': 'true' }
            : {},
      },
      virtualTableId: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          element.dataset.documentVirtualTableId || null,
        renderHTML: (attributes: Record<string, unknown>) =>
          typeof attributes.virtualTableId === 'string' &&
          attributes.virtualTableId
            ? { 'data-document-virtual-table-id': attributes.virtualTableId }
            : {},
      },
      virtualTableIndex: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          nonNegativeSafeInteger(element.dataset.documentVirtualTableIndex),
        renderHTML: (attributes: Record<string, unknown>) => {
          const index = nonNegativeSafeInteger(attributes.virtualTableIndex);
          return index === null
            ? {}
            : { 'data-document-virtual-table-index': String(index) };
        },
      },
      virtualTableCount: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          positiveSafeInteger(element.dataset.documentVirtualTableCount),
        renderHTML: (attributes: Record<string, unknown>) => {
          const count = positiveSafeInteger(attributes.virtualTableCount);
          return count === null
            ? {}
            : { 'data-document-virtual-table-count': String(count) };
        },
      },
    };
  },
}).configure({
  resizable: true,
  allowTableNodeSelection: true,
  View: WorkDocumentTableView,
});

function nonNegativeSafeInteger(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : null;
}

function positiveSafeInteger(value: unknown): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0 ? number : null;
}

export const DocumentTableSizing = Extension.create({
  name: 'documentTableSizing',

  addCommands() {
    return {
      setDocumentTableColumnWidth: (width, renderedColumnWidths) => (props) =>
        setSelectedColumnWidth(props, width, renderedColumnWidths),
      setDocumentTableColumnWidthPercent:
        (width, renderedColumnWidths, renderedTableWidth) => (props) =>
          setSelectedColumnWidthPercent(
            props,
            width,
            renderedColumnWidths,
            renderedTableWidth,
          ),
      setDocumentTableRowHeight:
        (height, rule = 'atLeast') =>
        (props) =>
          setSelectedRowHeight(props, height, rule),
      setDocumentTableLayoutMode: (mode, renderedTableWidth) => (props) =>
        setTableLayoutMode(props, mode, renderedTableWidth),
      setDocumentTableAlignment: (alignment) => (props) =>
        setTableAlignment(props, alignment),
      setDocumentTableProperties: (properties) => (props) =>
        setTableProperties(props, properties),
      setDocumentTablePropertyChanges: (changes) => (props) =>
        setTablePropertyChanges(props, changes),
      setDocumentTableCellMargins: (margins) => (props) =>
        setTableCellMargins(props, margins),
      distributeDocumentTableColumns: (renderedSelectionWidth) => (props) =>
        distributeColumns(props, renderedSelectionWidth),
      distributeDocumentTableRows: (renderedSelectionHeight) => (props) =>
        distributeRows(props, renderedSelectionHeight),
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        appendTransaction: (transactions, oldState, newState) => {
          if (!transactions.some((transaction) => transaction.docChanged)) {
            return null;
          }
          const previousTables = documentTables(oldState.doc);
          const currentTables = documentTables(newState.doc);
          if (previousTables.length !== currentTables.length) return null;
          const transaction = newState.tr;
          let tableIndex = 0;
          newState.doc.descendants((node, position) => {
            if (node.type.spec.tableRole !== 'table') return true;
            const previous = previousTables[tableIndex];
            tableIndex += 1;
            if (
              tableColumnWidthSignature(previous) !==
                tableColumnWidthSignature(node) &&
              tableColumnPercentageSignature(previous) ===
                tableColumnPercentageSignature(node) &&
              tableHasExplicitColumnWidths(node) &&
              tableGeometry(node).layout !== 'fixed'
            ) {
              transaction.setNodeMarkup(position, undefined, {
                ...node.attrs,
                geometry: fixedGeometryForTable(node),
                layoutMode: null,
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
  const percentages = tableColumnPercentages(context.table, context.map).filter(
    (_width, index) => columns.includes(index),
  );
  const columnPercentage = commonPositiveValue(percentages);
  const rowSizes = tableRowSizes(context.table).filter((_size, index) =>
    rows.includes(index),
  );
  const geometry = tableGeometry(context.table);
  return {
    columnWidth: columnPercentage ?? commonPositiveValue(widths),
    columnWidthType: columnPercentage === null ? 'pixels' : 'percent',
    rowHeight: commonNullableValue(rowSizes.map(({ height }) => height)),
    rowHeightRule: commonNullableValue(rowSizes.map(({ rule }) => rule)),
    layoutMode: documentTableLayoutMode(geometry),
    layoutAlgorithm: geometry.layout,
    preferredWidthType: geometry.width.type,
    preferredWidth: geometry.width.value,
    alignment: geometry.alignment,
    indent: geometry.indent,
    cellMargins: geometry.cellMargins,
    selectedColumnCount: columns.length,
    selectedRowCount: rows.length,
  };
}

function setSelectedColumnWidth(
  { dispatch, state }: CommandProps,
  width: number,
  renderedColumnWidths?: readonly number[],
): boolean {
  const normalized = normalizeDocumentTableDimension(
    width,
    MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
  );
  const context = tableSizingContext(state);
  if (!context || normalized === null) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  const preservedWidths = normalizeRenderedDimensions(
    renderedColumnWidths,
    context.map.width,
    MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
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
  clearPhysicalColumnPercentages(
    transaction,
    context,
    numberRange(context.left, context.right),
  );
  setFixedTableGeometryFromColumns(transaction, context);
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setSelectedColumnWidthPercent(
  { dispatch, state }: CommandProps,
  width: number,
  renderedColumnWidths?: readonly number[],
  renderedTableWidth?: number,
): boolean {
  const normalized = normalizeDocumentTableColumnPercent(width);
  const context = tableSizingContext(state);
  if (!context || normalized === null) return false;
  const fallbackWidths = percentageFallbackWidths(
    context,
    normalized,
    renderedColumnWidths,
    renderedTableWidth,
  );
  if (!fallbackWidths) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  setPhysicalColumnWidths(transaction, context, fallbackWidths, 0);
  const percentages = renderedPercentages(
    context,
    renderedColumnWidths,
    renderedTableWidth,
  );
  for (let column = context.left; column < context.right; column += 1) {
    percentages[column] = normalized;
  }
  setPhysicalColumnPercentages(transaction, context, percentages, 0);
  const geometry = tableGeometry(context.table);
  setTableGeometryAttribute(transaction, context, {
    ...geometry,
    layout: 'fixed',
    width:
      geometry.width.type === 'auto'
        ? { type: 'percent', value: 100 }
        : geometry.width,
  });
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
      : normalizeDocumentTableDimension(height, MIN_DOCUMENT_TABLE_ROW_HEIGHT);
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
  let preferredWidth = renderedTableWidth;
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
      equalDimensions(
        total,
        context.map.width,
        MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
      ),
      0,
    );
    preferredWidth = total;
  } else {
    clearPhysicalColumnWidths(transaction, context);
    clearPhysicalColumnPercentages(
      transaction,
      context,
      numberRange(0, context.map.width),
    );
  }
  setTableGeometryAttribute(
    transaction,
    context,
    documentTableGeometryForLayoutMode(
      requestedMode,
      tableGeometry(context.table),
      preferredWidth,
    ),
  );
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setTableAlignment(
  { dispatch, state }: CommandProps,
  requestedAlignment: DocumentTableAlignment,
): boolean {
  const alignment = normalizeDocumentTableAlignment(requestedAlignment);
  const context = tableSizingContext(state);
  if (!context || !alignment) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  setTableGeometryAttribute(transaction, context, {
    ...tableGeometry(context.table),
    alignment,
  });
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setTableProperties(
  { dispatch, state }: CommandProps,
  requestedProperties: DocumentTableProperties,
): boolean {
  const properties = normalizeDocumentTableProperties(requestedProperties);
  const context = tableSizingContext(state);
  if (!context || !properties) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  setTableGeometryAttribute(transaction, context, {
    ...tableGeometry(context.table),
    ...properties,
  });
  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setTablePropertyChanges(
  { dispatch, state }: CommandProps,
  requestedChanges: DocumentTablePropertyChanges,
): boolean {
  const changes = normalizeDocumentTablePropertyChanges(requestedChanges);
  const context = tableSizingContext(state);
  if (!context || !changes) return false;
  if (
    changes.row?.repeatHeader === true &&
    !canRepeatSelectedTableRows(context)
  ) {
    return false;
  }
  if (!dispatch) return true;

  const transaction = state.tr;
  closeHistory(transaction);
  if (changes.column) {
    if (changes.column.type === 'percent') {
      const fallbackWidths = percentageFallbackWidths(
        context,
        changes.column.width,
        changes.column.renderedColumnWidths,
        changes.column.renderedTableWidth,
      );
      if (!fallbackWidths) return false;
      setPhysicalColumnWidths(transaction, context, fallbackWidths, 0);
    } else {
      const preservedWidths = normalizeRenderedDimensions(
        changes.column.renderedColumnWidths,
        context.map.width,
        MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
      );
      if (preservedWidths) {
        for (let column = context.left; column < context.right; column += 1) {
          preservedWidths[column] = changes.column.width;
        }
        setPhysicalColumnWidths(transaction, context, preservedWidths, 0);
      } else {
        setPhysicalColumnWidths(
          transaction,
          context,
          numberRange(context.left, context.right).map(
            () => changes.column?.width ?? DEFAULT_COLUMN_WIDTH,
          ),
          context.left,
        );
      }
    }
    if (changes.column.type === 'percent') {
      const percentages = renderedPercentages(
        context,
        changes.column.renderedColumnWidths,
        changes.column.renderedTableWidth,
      );
      for (let column = context.left; column < context.right; column += 1) {
        percentages[column] = changes.column.width;
      }
      setPhysicalColumnPercentages(transaction, context, percentages, 0);
    } else {
      clearPhysicalColumnPercentages(
        transaction,
        context,
        numberRange(context.left, context.right),
      );
      setFixedTableGeometryFromColumns(transaction, context);
    }
  }

  if (changes.table) {
    const currentTable = transaction.doc.nodeAt(context.tableStart - 1);
    setTableGeometryAttribute(transaction, context, {
      ...tableGeometry(currentTable ?? context.table),
      ...changes.table,
    });
  }

  if (changes.row) {
    setPhysicalRowProperties(
      transaction,
      context,
      numberRange(context.top, context.bottom),
      changes.row,
    );
  }

  if (changes.cell) {
    setSelectedCellProperties(transaction, context, changes.cell);
  }

  dispatchTableSizingTransaction(dispatch, transaction, context, state);
  return true;
}

function setTableCellMargins(
  { dispatch, state }: CommandProps,
  requestedMargins: DocumentTableCellMargins,
): boolean {
  const cellMargins = normalizeDocumentTableCellMargins(requestedMargins);
  const context = tableSizingContext(state);
  if (!context || !cellMargins) return false;
  if (!dispatch) return true;
  const transaction = state.tr;
  setTableGeometryAttribute(transaction, context, {
    ...tableGeometry(context.table),
    cellMargins,
  });
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
    equalDimensions(total, columns.length, MIN_DOCUMENT_TABLE_COLUMN_WIDTH),
    columns[0] ?? 0,
  );
  clearPhysicalColumnPercentages(transaction, context, columns);
  setFixedTableGeometryFromColumns(transaction, context);
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
  const [height] = equalDimensions(
    total,
    rows.length,
    MIN_DOCUMENT_TABLE_ROW_HEIGHT,
  );
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

function setPhysicalColumnPercentages(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  percentages: readonly number[],
  startColumn: number,
): void {
  const touched = new Set<string>();
  percentages.forEach((percentage, widthIndex) => {
    const column = startColumn + widthIndex;
    if (column < 0 || column >= context.map.width) return;
    for (let row = 0; row < context.map.height; row += 1) {
      const relativePosition =
        context.map.map[row * context.map.width + column];
      if (relativePosition === undefined) continue;
      const widthIndexInCell = column - context.map.colCount(relativePosition);
      const key = `${relativePosition}:${widthIndexInCell}`;
      if (touched.has(key)) continue;
      touched.add(key);
      const documentPosition = context.tableStart + relativePosition;
      const current = transaction.doc.nodeAt(documentPosition);
      if (!current) continue;
      const values =
        normalizeDocumentTableColumnPercentages(
          current.attrs.columnWidthPercentages,
          Number(current.attrs.colspan ?? 1),
        ) ??
        Array.from(
          { length: Number(current.attrs.colspan ?? 1) },
          () => percentage,
        );
      values[widthIndexInCell] = percentage;
      transaction.setNodeMarkup(documentPosition, undefined, {
        ...current.attrs,
        columnWidthPercentages: values,
      });
    }
  });
}

function clearPhysicalColumnPercentages(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  columns: readonly number[],
): void {
  const touched = new Set<number>();
  for (const column of columns) {
    for (let row = 0; row < context.map.height; row += 1) {
      const relativePosition =
        context.map.map[row * context.map.width + column];
      if (relativePosition === undefined || touched.has(relativePosition))
        continue;
      touched.add(relativePosition);
      const position = context.tableStart + relativePosition;
      const current = transaction.doc.nodeAt(position);
      if (current?.attrs.columnWidthPercentages) {
        transaction.setNodeMarkup(position, undefined, {
          ...current.attrs,
          columnWidthPercentages: null,
        });
      }
    }
  }
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

function setPhysicalRowProperties(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  rows: readonly number[],
  properties: NonNullable<DocumentTablePropertyChanges['row']>,
): void {
  const selectedRows = new Set(rows);
  context.table.forEach((_row, offset, index) => {
    if (!selectedRows.has(index)) return;
    const position = context.tableStart + offset;
    const current = transaction.doc.nodeAt(position);
    if (!current) return;
    transaction.setNodeMarkup(position, undefined, {
      ...current.attrs,
      rowHeight: properties.height,
      rowHeightRule: properties.height === null ? null : properties.heightRule,
      cantSplit: properties.cantSplit,
      ...(properties.repeatHeader === undefined
        ? {}
        : { repeatHeader: properties.repeatHeader }),
    });
  });
}

function setSelectedCellProperties(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  properties: NonNullable<DocumentTablePropertyChanges['cell']>,
): void {
  const positions = context.map.cellsInRect({
    left: context.left,
    right: context.right,
    top: context.top,
    bottom: context.bottom,
  });
  for (const relativePosition of positions) {
    const position = context.tableStart + relativePosition;
    const current = transaction.doc.nodeAt(position);
    if (!current || !isTableCell(current)) continue;
    transaction.setNodeMarkup(position, undefined, {
      ...current.attrs,
      verticalAlign: properties.verticalAlign,
      margins: properties.margins,
    });
  }
}

function canRepeatSelectedTableRows(context: TableSizingContext): boolean {
  let leadingRepeatCount = 0;
  for (let index = 0; index < context.table.childCount; index += 1) {
    if (!documentTableRowRepeats(context.table.child(index))) break;
    leadingRepeatCount += 1;
  }
  return context.top <= leadingRepeatCount;
}

function setTableGeometryAttribute(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
  geometry: DocumentTableGeometry,
): void {
  const position = context.tableStart - 1;
  const table = transaction.doc.nodeAt(position);
  if (!table) return;
  transaction.setNodeMarkup(position, undefined, {
    ...table.attrs,
    geometry: normalizeDocumentTableGeometry(geometry),
    layoutMode: null,
  });
}

function setFixedTableGeometryFromColumns(
  transaction: CommandProps['tr'],
  context: TableSizingContext,
): void {
  const table = transaction.doc.nodeAt(context.tableStart - 1);
  if (!table || table.type.spec.tableRole !== 'table') return;
  setTableGeometryAttribute(transaction, context, fixedGeometryForTable(table));
}

function tableGeometry(table: ProseMirrorNode): DocumentTableGeometry {
  return normalizeDocumentTableGeometry(
    table.attrs.geometry,
    table.attrs.layoutMode,
  );
}

function fixedGeometryForTable(table: ProseMirrorNode): DocumentTableGeometry {
  const geometry = tableGeometry(table);
  const widths = tableColumnWidths(table, TableMap.get(table));
  const total = widths.every(
    (width): width is number => width !== null && width > 0,
  )
    ? widths.reduce((sum, width) => sum + width, 0)
    : null;
  return {
    ...geometry,
    layout: 'fixed',
    width:
      total === null
        ? geometry.width
        : {
            type: 'pixels',
            value: Math.round(total * 100) / 100,
          },
  };
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

function tableColumnPercentages(
  table: ProseMirrorNode,
  map: TableMap,
): Array<number | null> {
  return Array.from({ length: map.width }, (_unused, column) => {
    for (let row = 0; row < map.height; row += 1) {
      const position = map.map[row * map.width + column];
      if (position === undefined) continue;
      const cell = table.nodeAt(position);
      const index = column - map.colCount(position);
      const percentages = normalizeDocumentTableColumnPercentages(
        cell?.attrs.columnWidthPercentages,
        Number(cell?.attrs.colspan ?? 1),
      );
      if (percentages?.[index] !== undefined) return percentages[index] ?? null;
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

function documentTables(document: ProseMirrorNode): ProseMirrorNode[] {
  const tables: ProseMirrorNode[] = [];
  document.descendants((node) => {
    if (node.type.spec.tableRole !== 'table') return true;
    tables.push(node);
    return false;
  });
  return tables;
}

function tableColumnWidthSignature(table: ProseMirrorNode | undefined): string {
  if (!table) return '';
  const widths: unknown[] = [];
  table.descendants((node) => {
    if (!isTableCell(node)) return true;
    widths.push(node.attrs.colwidth ?? null);
    return false;
  });
  return JSON.stringify(widths);
}

function tableColumnPercentageSignature(
  table: ProseMirrorNode | undefined,
): string {
  if (!table) return '';
  const widths: unknown[] = [];
  table.descendants((node) => {
    if (!isTableCell(node)) return true;
    widths.push(node.attrs.columnWidthPercentages ?? null);
    return false;
  });
  return JSON.stringify(widths);
}

function renderedPercentages(
  context: TableSizingContext,
  renderedColumnWidths: readonly number[] | undefined,
  renderedTableWidth: number | undefined,
): number[] {
  const existing = tableColumnPercentages(context.table, context.map);
  const rendered = normalizeRenderedDimensions(
    renderedColumnWidths,
    context.map.width,
    1,
  );
  const physical = tableColumnWidths(context.table, context.map);
  const total =
    Number.isFinite(renderedTableWidth) && Number(renderedTableWidth) > 0
      ? Number(renderedTableWidth)
      : rendered
        ? rendered.reduce((sum, value) => sum + value, 0)
        : physical.every((value) => value !== null && value > 0)
          ? physical.reduce<number>((sum, value) => sum + Number(value), 0)
          : DEFAULT_COLUMN_WIDTH * context.map.width;
  return existing.map((value, index) => {
    if (value !== null) return value;
    const pixels = rendered?.[index] ?? physical[index] ?? DEFAULT_COLUMN_WIDTH;
    return Math.max(
      1,
      Math.min(100, Math.round((Number(pixels) / total) * 10_000) / 100),
    );
  });
}

function percentageFallbackWidths(
  context: TableSizingContext,
  selectedPercentage: number,
  renderedColumnWidths: readonly number[] | undefined,
  renderedTableWidth: number | undefined,
): number[] | null {
  const rendered = normalizeRenderedDimensions(
    renderedColumnWidths,
    context.map.width,
    1,
  );
  const physical = tableColumnWidths(context.table, context.map);
  const total =
    Number.isFinite(renderedTableWidth) && Number(renderedTableWidth) > 0
      ? Number(renderedTableWidth)
      : rendered
        ? rendered.reduce((sum, value) => sum + value, 0)
        : physical.every((value) => value !== null && value > 0)
          ? physical.reduce<number>((sum, value) => sum + Number(value), 0)
          : null;
  if (total === null) return null;
  const widths = (
    rendered ?? physical.map((value) => value ?? DEFAULT_COLUMN_WIDTH)
  ).map(Number);
  const selectedPixels = Math.max(
    MIN_DOCUMENT_TABLE_COLUMN_WIDTH,
    Math.round((selectedPercentage / 100) * total * 100) / 100,
  );
  for (let column = context.left; column < context.right; column += 1) {
    widths[column] = selectedPixels;
  }
  return widths;
}

function isTableCell(node: ProseMirrorNode): boolean {
  return (
    node.type.spec.tableRole === 'cell' ||
    node.type.spec.tableRole === 'header_cell'
  );
}

function normalizeRenderedDimensions(
  values: readonly number[] | undefined,
  expectedCount: number,
  minimum: number,
): number[] | null {
  if (!values || values.length !== expectedCount) return null;
  const normalized = values.map((value) =>
    normalizeDocumentTableDimension(value, minimum),
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
