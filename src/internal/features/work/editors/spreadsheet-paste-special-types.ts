import type { Cell } from '@fortune-sheet/core';
import type { SpreadsheetResolvedCellBorders } from './spreadsheet-cell-border';
import type {
  SpreadsheetCellRange,
  SpreadsheetCellRangeInput,
} from './spreadsheet-cell-range';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetSheet,
} from '../work-types';

export const MAX_SPREADSHEET_PASTE_SPECIAL_CELLS = 50_000;
export const MAX_SPREADSHEET_ROWS = 1_048_576;
export const MAX_SPREADSHEET_COLUMNS = 16_384;

export const spreadsheetPasteContentOptions = [
  { value: 'all', label: '全部', richOnly: false },
  { value: 'formulas', label: '公式', richOnly: false },
  { value: 'values', label: '值', richOnly: false },
  { value: 'formats', label: '格式', richOnly: true },
  { value: 'comments', label: '批注', richOnly: true },
  { value: 'validation', label: '数据验证', richOnly: true },
  { value: 'all-except-borders', label: '边框除外的全部', richOnly: true },
  {
    value: 'formulas-and-number-formats',
    label: '公式和数字格式',
    richOnly: true,
  },
  {
    value: 'values-and-number-formats',
    label: '值和数字格式',
    richOnly: true,
  },
  { value: 'column-widths', label: '列宽', richOnly: true },
] as const;

export type SpreadsheetPasteContent =
  (typeof spreadsheetPasteContentOptions)[number]['value'];

export const spreadsheetPasteOperationOptions = [
  { value: 'none', label: '无' },
  { value: 'add', label: '加' },
  { value: 'subtract', label: '减' },
  { value: 'multiply', label: '乘' },
  { value: 'divide', label: '除' },
] as const;

export type SpreadsheetPasteOperation =
  (typeof spreadsheetPasteOperationOptions)[number]['value'];

export interface SpreadsheetPasteSpecialOptions {
  content: SpreadsheetPasteContent;
  operation: SpreadsheetPasteOperation;
  skipBlanks: boolean;
  transpose: boolean;
}

export interface SpreadsheetClipboardCell {
  cell: Cell | null;
  borders: SpreadsheetResolvedCellBorders;
  validation?: WorkSpreadsheetDataValidationItem;
  protection?: { locked: boolean; hidden: boolean };
  hyperlink?: NonNullable<WorkSpreadsheetSheet['hyperlink']>[string];
}

export interface SpreadsheetClipboardMerge {
  row: number;
  column: number;
  rowSpan: number;
  columnSpan: number;
}

export interface SpreadsheetClipboardSnapshot {
  version: 1;
  kind: 'rich' | 'text';
  plainText: string;
  sourceSheetId?: string;
  sourceRange: SpreadsheetCellRange;
  rowCount: number;
  columnCount: number;
  cells: SpreadsheetClipboardCell[][];
  columnWidths?: number[];
  merges: SpreadsheetClipboardMerge[];
  containsUnsupportedFormulaState: boolean;
}

export interface SpreadsheetPasteSpecialRequest {
  snapshot: SpreadsheetClipboardSnapshot;
  targetSheetId: string;
  targetSelection: SpreadsheetCellRangeInput;
  options: SpreadsheetPasteSpecialOptions;
}

export interface SpreadsheetPasteSpecialResult {
  content: WorkSpreadsheetContent;
  targetRange: SpreadsheetCellRange;
  firstCellValue: unknown;
}

export interface SpreadsheetPastePlan {
  targetRange: SpreadsheetCellRange;
  rowCount: number;
  columnCount: number;
}

export interface SpreadsheetCellWriter {
  get(row: number, column: number): Cell | null | undefined;
  set(row: number, column: number, cell: Cell | null): void;
  finish(sheet: WorkSpreadsheetSheet): WorkSpreadsheetSheet;
}

export interface SpreadsheetPasteSource {
  cell: SpreadsheetClipboardCell;
  sourceRow: number;
  sourceColumn: number;
}
