import type { DocumentTableCellFormat } from '../work-document-table-cell-formatting';
import type {
  DocumentTableAlignment,
  DocumentTableCellMarginOverrides,
  DocumentTableCellMarginSide,
  DocumentTableCellMargins,
  DocumentTablePreferredWidthType,
  DocumentTableProperties,
} from '../work-document-table-geometry';
import type {
  DocumentTableRowHeightRule,
  DocumentTableRowOptions,
} from '../work-document-table-row';
import type {
  DocumentTablePropertyChanges,
  DocumentTableSizingState,
} from '../work-document-table-sizing';
import type { DocumentTableColumnWidthType } from '../work-document-table-column-widths';

export const PIXELS_PER_CENTIMETER = 96 / 2.54;

export type DocumentTablePropertiesTab = 'table' | 'row' | 'column' | 'cell';

export interface DocumentTablePropertiesDraft {
  table: {
    widthType: DocumentTablePreferredWidthType;
    width: string;
    alignment: DocumentTableAlignment;
    indent: string;
  };
  row: {
    heightEnabled: boolean;
    height: string;
    heightRule: DocumentTableRowHeightRule;
    cantSplit: boolean;
    repeatHeader: boolean;
  };
  column: {
    widthType: DocumentTableColumnWidthType;
    width: string;
  };
  cell: {
    verticalAlign: DocumentTableCellFormat['verticalAlign'];
    useTableMargins: boolean;
    margins: Record<DocumentTableCellMarginSide, string>;
  };
}

export interface DocumentTablePropertiesSource {
  sizing: DocumentTableSizingState;
  rowOptions: DocumentTableRowOptions;
  cellFormat: DocumentTableCellFormat;
  canRepeatHeader: boolean;
  renderedTableWidth?: number;
  renderedRowHeight?: number;
  renderedColumnWidth?: number;
  renderedColumnWidths?: readonly number[];
}

export interface DocumentTablePropertiesErrors {
  tableWidth: string | null;
  tableIndent: string | null;
  rowHeight: string | null;
  columnWidth: string | null;
  cellMargins: Partial<Record<DocumentTableCellMarginSide, string>>;
}

const FALLBACK_TABLE_WIDTH_CENTIMETERS = 15;
const FALLBACK_ROW_HEIGHT_CENTIMETERS = 0.95;
const FALLBACK_COLUMN_WIDTH_CENTIMETERS = 3.18;
const marginSides = [
  'top',
  'right',
  'bottom',
  'left',
] as const satisfies readonly DocumentTableCellMarginSide[];

export function createDocumentTablePropertiesDraft(
  source: DocumentTablePropertiesSource,
): DocumentTablePropertiesDraft {
  const widthType = source.sizing.preferredWidthType;
  const effectiveMargins = effectiveCellMargins(
    source.cellFormat,
    source.sizing.cellMargins,
  );
  return {
    table: {
      widthType,
      width:
        widthType === 'percent'
          ? formatNumber(source.sizing.preferredWidth ?? 100)
          : widthType === 'pixels'
            ? centimeters(
                source.sizing.preferredWidth ??
                  source.renderedTableWidth ??
                  FALLBACK_TABLE_WIDTH_CENTIMETERS * PIXELS_PER_CENTIMETER,
              )
            : '',
      alignment: source.sizing.alignment,
      indent: centimeters(source.sizing.indent),
    },
    row: {
      heightEnabled: source.sizing.rowHeight !== null,
      height: centimeters(
        source.sizing.rowHeight ??
          source.renderedRowHeight ??
          FALLBACK_ROW_HEIGHT_CENTIMETERS * PIXELS_PER_CENTIMETER,
      ),
      heightRule: source.sizing.rowHeightRule ?? 'atLeast',
      cantSplit: source.rowOptions.cantSplit,
      repeatHeader: source.rowOptions.repeatHeader,
    },
    column: {
      widthType: source.sizing.columnWidthType,
      width:
        source.sizing.columnWidthType === 'percent'
          ? formatNumber(source.sizing.columnWidth ?? 100)
          : centimeters(
              source.renderedColumnWidth ??
                source.sizing.columnWidth ??
                FALLBACK_COLUMN_WIDTH_CENTIMETERS * PIXELS_PER_CENTIMETER,
            ),
    },
    cell: {
      verticalAlign: source.cellFormat.verticalAlign,
      useTableMargins: source.cellFormat.margins === null,
      margins: marginDrafts(effectiveMargins),
    },
  };
}

export function documentTablePropertiesErrors(
  draft: DocumentTablePropertiesDraft,
): DocumentTablePropertiesErrors {
  const cellMargins: DocumentTablePropertiesErrors['cellMargins'] = {};
  if (!draft.cell.useTableMargins) {
    for (const side of marginSides) {
      if (!validNumber(draft.cell.margins[side], 0, 5)) {
        cellMargins[side] = '请输入 0 到 5 之间的厘米数。';
      }
    }
  }
  return {
    tableWidth: tableWidthError(draft),
    tableIndent: tableIndentError(draft),
    rowHeight: draft.row.heightEnabled
      ? dimensionError(draft.row.height)
      : null,
    columnWidth:
      draft.column.widthType === 'percent'
        ? percentageError(draft.column.width)
        : dimensionError(draft.column.width),
    cellMargins,
  };
}

export function hasDocumentTablePropertiesErrors(
  errors: DocumentTablePropertiesErrors,
): boolean {
  return Boolean(
    errors.tableWidth ||
      errors.tableIndent ||
      errors.rowHeight ||
      errors.columnWidth ||
      Object.keys(errors.cellMargins).length,
  );
}

export function draftForTableWidthType(
  current: DocumentTablePropertiesDraft,
  widthType: DocumentTablePreferredWidthType,
  renderedTableWidth: number | undefined,
): DocumentTablePropertiesDraft {
  if (widthType === current.table.widthType) return current;
  const width =
    widthType === 'auto'
      ? ''
      : widthType === 'percent'
        ? '100'
        : centimeters(
            renderedTableWidth ??
              FALLBACK_TABLE_WIDTH_CENTIMETERS * PIXELS_PER_CENTIMETER,
          );
  return { ...current, table: { ...current.table, widthType, width } };
}

export function draftForColumnWidthType(
  current: DocumentTablePropertiesDraft,
  widthType: DocumentTableColumnWidthType,
  source: Pick<
    DocumentTablePropertiesSource,
    'renderedColumnWidth' | 'renderedTableWidth' | 'sizing'
  >,
): DocumentTablePropertiesDraft {
  if (widthType === current.column.widthType) return current;
  const width =
    widthType === 'percent'
      ? formatNumber(
          source.sizing.columnWidthType === 'percent'
            ? (source.sizing.columnWidth ?? 100)
            : source.renderedColumnWidth && source.renderedTableWidth
              ? (source.renderedColumnWidth / source.renderedTableWidth) * 100
              : 100 / Math.max(1, source.sizing.selectedColumnCount),
        )
      : centimeters(
          source.renderedColumnWidth ??
            (source.sizing.columnWidthType === 'pixels'
              ? source.sizing.columnWidth
              : null) ??
            FALLBACK_COLUMN_WIDTH_CENTIMETERS * PIXELS_PER_CENTIMETER,
        );
  return { ...current, column: { widthType, width } };
}

export function documentTablePropertyChanges(
  initial: DocumentTablePropertiesDraft,
  current: DocumentTablePropertiesDraft,
  source: Pick<
    DocumentTablePropertiesSource,
    | 'canRepeatHeader'
    | 'cellFormat'
    | 'renderedColumnWidths'
    | 'renderedTableWidth'
    | 'sizing'
  >,
): DocumentTablePropertyChanges | null {
  const errors = documentTablePropertiesErrors(current);
  if (hasDocumentTablePropertiesErrors(errors)) return null;

  const changes: DocumentTablePropertyChanges = {};

  const tableChanged =
    initial.table.widthType !== current.table.widthType ||
    !sameDraftNumber(initial.table.width, current.table.width) ||
    initial.table.alignment !== current.table.alignment ||
    !sameDraftNumber(initial.table.indent, current.table.indent);
  if (tableChanged) {
    changes.table = tablePropertiesForChanges(initial, current, source.sizing);
  }

  const rowChanged =
    initial.row.heightEnabled !== current.row.heightEnabled ||
    (current.row.heightEnabled &&
      !sameDraftNumber(initial.row.height, current.row.height)) ||
    initial.row.heightRule !== current.row.heightRule ||
    initial.row.cantSplit !== current.row.cantSplit ||
    (source.canRepeatHeader &&
      initial.row.repeatHeader !== current.row.repeatHeader);
  if (rowChanged) {
    changes.row = {
      height: rowHeightForChanges(initial, current, source.sizing),
      heightRule: current.row.heightRule,
      cantSplit: current.row.cantSplit,
      ...(source.canRepeatHeader
        ? { repeatHeader: current.row.repeatHeader }
        : {}),
    };
  }

  if (
    initial.column.widthType !== current.column.widthType ||
    !sameDraftNumber(initial.column.width, current.column.width)
  ) {
    changes.column = {
      type: current.column.widthType,
      width:
        current.column.widthType === 'percent'
          ? rounded(Number(current.column.width))
          : rounded(Number(current.column.width) * PIXELS_PER_CENTIMETER),
      renderedColumnWidths: source.renderedColumnWidths,
      renderedTableWidth: source.renderedTableWidth,
    };
  }

  const marginsChanged =
    initial.cell.useTableMargins !== current.cell.useTableMargins ||
    (!current.cell.useTableMargins &&
      marginSides.some(
        (side) =>
          !sameDraftNumber(
            initial.cell.margins[side],
            current.cell.margins[side],
          ),
      ));
  if (
    initial.cell.verticalAlign !== current.cell.verticalAlign ||
    marginsChanged
  ) {
    changes.cell = {
      verticalAlign: current.cell.verticalAlign,
      margins: cellMarginsForChanges(initial, current, source),
    };
  }

  return Object.keys(changes).length ? changes : null;
}

function tablePropertiesForChanges(
  initial: DocumentTablePropertiesDraft,
  current: DocumentTablePropertiesDraft,
  sizing: DocumentTableSizingState,
): DocumentTableProperties {
  const widthType = current.table.widthType;
  const preserveWidth =
    initial.table.widthType === widthType &&
    sameDraftNumber(initial.table.width, current.table.width) &&
    sizing.preferredWidthType === widthType;
  const widthValue =
    widthType === 'auto'
      ? null
      : preserveWidth && sizing.preferredWidth !== null
        ? sizing.preferredWidth
        : widthType === 'percent'
          ? Number(current.table.width)
          : Number(current.table.width) * PIXELS_PER_CENTIMETER;
  return {
    width: {
      type: widthType,
      value: widthValue === null ? null : rounded(widthValue),
    } as DocumentTableProperties['width'],
    alignment: current.table.alignment,
    indent: sameDraftNumber(initial.table.indent, current.table.indent)
      ? sizing.indent
      : rounded(Number(current.table.indent) * PIXELS_PER_CENTIMETER),
  };
}

function rowHeightForChanges(
  initial: DocumentTablePropertiesDraft,
  current: DocumentTablePropertiesDraft,
  sizing: DocumentTableSizingState,
): number | null {
  if (!current.row.heightEnabled) return null;
  if (
    initial.row.heightEnabled &&
    sameDraftNumber(initial.row.height, current.row.height) &&
    sizing.rowHeight !== null
  ) {
    return sizing.rowHeight;
  }
  return rounded(Number(current.row.height) * PIXELS_PER_CENTIMETER);
}

function cellMarginsForChanges(
  initial: DocumentTablePropertiesDraft,
  current: DocumentTablePropertiesDraft,
  source: Pick<DocumentTablePropertiesSource, 'cellFormat' | 'sizing'>,
): DocumentTableCellMarginOverrides | null {
  if (current.cell.useTableMargins) return null;
  const margins: DocumentTableCellMarginOverrides = initial.cell.useTableMargins
    ? { ...source.sizing.cellMargins }
    : { ...(source.cellFormat.margins ?? {}) };
  for (const side of marginSides) {
    if (
      sameDraftNumber(initial.cell.margins[side], current.cell.margins[side])
    ) {
      continue;
    }
    margins[side] = rounded(
      Number(current.cell.margins[side]) * PIXELS_PER_CENTIMETER,
    );
  }
  return margins;
}

function effectiveCellMargins(
  format: DocumentTableCellFormat,
  tableMargins: DocumentTableCellMargins,
): DocumentTableCellMargins {
  return {
    top: format.margins?.top ?? tableMargins.top,
    right: format.margins?.right ?? tableMargins.right,
    bottom: format.margins?.bottom ?? tableMargins.bottom,
    left: format.margins?.left ?? tableMargins.left,
  };
}

function marginDrafts(
  margins: DocumentTableCellMargins,
): Record<DocumentTableCellMarginSide, string> {
  return {
    top: centimeters(margins.top),
    right: centimeters(margins.right),
    bottom: centimeters(margins.bottom),
    left: centimeters(margins.left),
  };
}

function tableWidthError(draft: DocumentTablePropertiesDraft): string | null {
  if (draft.table.widthType === 'auto') return null;
  if (draft.table.widthType === 'percent') {
    return validNumber(draft.table.width, 1, 100)
      ? null
      : '请输入 1 到 100 之间的百分比。';
  }
  return validNumber(draft.table.width, 0.5, 30)
    ? null
    : '请输入 0.5 到 30 之间的厘米数。';
}

function tableIndentError(draft: DocumentTablePropertiesDraft): string | null {
  if (draft.table.alignment !== 'left') return null;
  return validNumber(draft.table.indent, 0, 30)
    ? null
    : '请输入 0 到 30 之间的厘米数。';
}

function dimensionError(value: string): string | null {
  return validNumber(value, 0.5, 30) ? null : '请输入 0.5 到 30 之间的厘米数。';
}

function percentageError(value: string): string | null {
  return validNumber(value, 1, 100) ? null : '请输入 1 到 100 之间的百分比。';
}

function validNumber(value: string, minimum: number, maximum: number): boolean {
  if (!value.trim()) return false;
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= minimum && numeric <= maximum;
}

function sameDraftNumber(left: string, right: string): boolean {
  if (!left.trim() || !right.trim()) return left === right;
  return Number(left) === Number(right);
}

function centimeters(pixels: number): string {
  return formatNumber(pixels / PIXELS_PER_CENTIMETER);
}

function formatNumber(value: number): string {
  return Number.isFinite(value) ? value.toFixed(2).replace(/\.?0+$/, '') : '';
}

function rounded(value: number): number {
  return Math.round(value * 100) / 100;
}
