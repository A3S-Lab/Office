import type { Cell } from '@fortune-sheet/core';
import {
  formulaHasExternalReference,
  spreadsheetFormulaRangesForSelection,
} from '../work-spreadsheet-formulas';
import type { WorkSpreadsheetSheet } from '../work-types';
import { canMutateSpreadsheetCellRange } from './spreadsheet-cell-mutation-guard';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import {
  MAX_SPREADSHEET_COLUMNS,
  MAX_SPREADSHEET_ROWS,
} from './spreadsheet-paste-special-types';

export type SpreadsheetCopyFromAboveKind = 'formula' | 'value';

export type SpreadsheetCopyFromAboveValue = boolean | number | string | null;

export interface SpreadsheetCopyFromAbovePlan {
  formulaBarValue: SpreadsheetCopyFromAboveValue | '';
  kind: SpreadsheetCopyFromAboveKind;
  source: { row: number; column: number };
  targetRange: SpreadsheetCellRange;
  value: SpreadsheetCopyFromAboveValue;
}

export function planSpreadsheetCopyFromAbove(
  sheet: WorkSpreadsheetSheet | undefined,
  target: { row: number; column: number },
  kind: SpreadsheetCopyFromAboveKind,
): SpreadsheetCopyFromAbovePlan | null {
  if (
    !sheet ||
    !isSpreadsheetCopyFromAboveKind(kind) ||
    !spreadsheetCellCoordinateIsSupported(target.row, target.column) ||
    target.row === 0
  ) {
    return null;
  }

  const targetRange: SpreadsheetCellRange = {
    row: [target.row, target.row],
    column: [target.column, target.column],
  };
  if (!canMutateSpreadsheetCellRange(sheet, targetRange)) return null;

  const source = { row: target.row - 1, column: target.column };
  const sourceCell = spreadsheetCopyFromAboveCellAt(
    sheet,
    source.row,
    source.column,
  );
  const targetCell = spreadsheetCopyFromAboveCellAt(
    sheet,
    target.row,
    target.column,
  );
  const formula = spreadsheetCopyableFormula(
    sheet,
    sourceCell,
    source.row,
    source.column,
  );
  if (kind === 'formula' && sourceCell?.f && formula === null) return null;

  const copiesFormula = kind === 'formula' && formula !== null;
  const value = copiesFormula
    ? formula
    : spreadsheetCopyFromAboveCellValue(sourceCell);
  if (
    copiesFormula
      ? targetCell?.f === value
      : !targetCell?.f &&
        Object.is(spreadsheetCopyFromAboveCellValue(targetCell), value)
  ) {
    return null;
  }

  return {
    formulaBarValue: value ?? '',
    kind,
    source,
    targetRange,
    value,
  };
}

function spreadsheetCopyableFormula(
  sheet: WorkSpreadsheetSheet,
  cell: Cell | null | undefined,
  row: number,
  column: number,
): string | null {
  if (typeof cell?.f !== 'string' || !cell.f.trim()) return null;
  const formula = cell.f.trim();
  if (
    !formula.startsWith('=') ||
    formulaHasExternalReference(formula) ||
    cell.spl !== undefined ||
    spreadsheetFormulaRangesForSelection(sheet, {
      startRow: row,
      endRow: row,
      startColumn: column,
      endColumn: column,
    }).length > 0
  ) {
    return null;
  }
  return formula;
}

function spreadsheetCopyFromAboveCellValue(
  cell: Cell | null | undefined,
): SpreadsheetCopyFromAboveValue {
  if (cell?.v !== undefined && cell.v !== null) return cell.v;
  const inlineText = spreadsheetInlineCellText(cell);
  if (inlineText !== null) return inlineText;
  if (cell?.m !== undefined && cell.m !== null) return cell.m;
  return null;
}

function spreadsheetInlineCellText(
  cell: Cell | null | undefined,
): string | null {
  const format = cell?.ct as { s?: unknown; t?: unknown } | undefined;
  if (format?.t !== 'inlineStr' || !Array.isArray(format.s)) return null;
  return format.s
    .map((run) =>
      isRecord(run) && (typeof run.v === 'string' || typeof run.v === 'number')
        ? String(run.v)
        : '',
    )
    .join('');
}

function spreadsheetCopyFromAboveCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null | undefined {
  if (sheet.data !== undefined) return sheet.data[row]?.[column];
  return sheet.celldata?.find((entry) => entry.r === row && entry.c === column)
    ?.v;
}

function spreadsheetCellCoordinateIsSupported(
  row: number,
  column: number,
): boolean {
  return (
    Number.isSafeInteger(row) &&
    row >= 0 &&
    row < MAX_SPREADSHEET_ROWS &&
    Number.isSafeInteger(column) &&
    column >= 0 &&
    column < MAX_SPREADSHEET_COLUMNS
  );
}

function isSpreadsheetCopyFromAboveKind(
  value: unknown,
): value is SpreadsheetCopyFromAboveKind {
  return value === 'formula' || value === 'value';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
