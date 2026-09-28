import type { Cell } from '@fortune-sheet/core';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetSheet,
} from '../work-types';
import type { SpreadsheetResolvedCellBorders } from './spreadsheet-cell-border';
import type {
  SpreadsheetCellRange,
  SpreadsheetCellRangeInput,
} from './spreadsheet-cell-range';

export const MAX_SPREADSHEET_PASTE_SPECIAL_CELLS = 50_000;
export const MAX_SPREADSHEET_ROWS = 1_048_576;
export const MAX_SPREADSHEET_COLUMNS = 16_384;

const PASTE_CONTENT_SPECS = [
  {
    value: 'all',
    labelKey: 'spreadsheet.pasteSpecial.content.all',
    richOnly: false,
  },
  {
    value: 'formulas',
    labelKey: 'spreadsheet.pasteSpecial.content.formulas',
    richOnly: false,
  },
  {
    value: 'values',
    labelKey: 'spreadsheet.pasteSpecial.content.values',
    richOnly: false,
  },
  {
    value: 'formats',
    labelKey: 'spreadsheet.pasteSpecial.content.formats',
    richOnly: true,
  },
  {
    value: 'comments',
    labelKey: 'spreadsheet.pasteSpecial.content.comments',
    richOnly: true,
  },
  {
    value: 'validation',
    labelKey: 'spreadsheet.pasteSpecial.content.validation',
    richOnly: true,
  },
  {
    value: 'all-except-borders',
    labelKey: 'spreadsheet.pasteSpecial.content.allExceptBorders',
    richOnly: true,
  },
  {
    value: 'formulas-and-number-formats',
    labelKey: 'spreadsheet.pasteSpecial.content.formulasAndNumberFormats',
    richOnly: true,
  },
  {
    value: 'values-and-number-formats',
    labelKey: 'spreadsheet.pasteSpecial.content.valuesAndNumberFormats',
    richOnly: true,
  },
  {
    value: 'column-widths',
    labelKey: 'spreadsheet.pasteSpecial.content.columnWidths',
    richOnly: true,
  },
] as const;

const PASTE_OPERATION_SPECS = [
  { value: 'none', labelKey: 'spreadsheet.pasteSpecial.op.none' },
  { value: 'add', labelKey: 'spreadsheet.pasteSpecial.op.add' },
  { value: 'subtract', labelKey: 'spreadsheet.pasteSpecial.op.subtract' },
  { value: 'multiply', labelKey: 'spreadsheet.pasteSpecial.op.multiply' },
  { value: 'divide', labelKey: 'spreadsheet.pasteSpecial.op.divide' },
] as const;

export type SpreadsheetPasteContent = (typeof PASTE_CONTENT_SPECS)[number]['value'];

export type SpreadsheetPasteOperation =
  (typeof PASTE_OPERATION_SPECS)[number]['value'];

export function spreadsheetPasteContentOptions(
  catalog: OfficeMessageCatalog = resolveOfficeMessages(),
): readonly {
  value: SpreadsheetPasteContent;
  label: string;
  richOnly: boolean;
}[] {
  return PASTE_CONTENT_SPECS.map((spec) => ({
    value: spec.value,
    label: officeMessage(catalog, spec.labelKey),
    richOnly: spec.richOnly,
  }));
}

export function spreadsheetPasteOperationOptions(
  catalog: OfficeMessageCatalog = resolveOfficeMessages(),
): readonly { value: SpreadsheetPasteOperation; label: string }[] {
  return PASTE_OPERATION_SPECS.map((spec) => ({
    value: spec.value,
    label: officeMessage(catalog, spec.labelKey),
  }));
}

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
