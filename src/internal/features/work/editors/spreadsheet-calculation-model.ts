import {
  type Cell,
  genarate,
  type Op,
  type Sheet,
  update,
} from '@fortune-sheet/core';
import type {
  OfficeKernelSpreadsheetCalculatedCell,
  OfficeKernelSpreadsheetCoordinate,
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetSessionUpdate,
  OfficeKernelSpreadsheetValue,
} from '../../../kernel/office-kernel-protocol';
import {
  SPREADSHEET_KERNEL_MAX_PATCH_CELLS,
  type SpreadsheetKernelFallbackCell,
  type SpreadsheetKernelWorkbook,
  sameSpreadsheetKernelInputCell,
  spreadsheetKernelCellKey,
} from './spreadsheet-calculation-projection';
import type { SpreadsheetCalculationCommand } from './spreadsheet-command-controller';

export function spreadsheetCalculationTargets(
  workbook: SpreadsheetKernelWorkbook,
  command: SpreadsheetCalculationCommand,
): OfficeKernelSpreadsheetCoordinate[] | undefined {
  if (command.scope === 'workbook') return undefined;
  const rowStart = Math.min(
    command.range.row[0] ?? 0,
    command.range.row[1] ?? 0,
  );
  const rowEnd = Math.max(command.range.row[0] ?? 0, command.range.row[1] ?? 0);
  const columnStart = Math.min(
    command.range.column[0] ?? 0,
    command.range.column[1] ?? 0,
  );
  const columnEnd = Math.max(
    command.range.column[0] ?? 0,
    command.range.column[1] ?? 0,
  );
  const sheet = workbook.sheets.find(
    (candidate) => candidate.id === command.sheetId,
  );
  if (!sheet) return [];
  return sheet.cells.flatMap((cell) =>
    cell.formula &&
    cell.row >= rowStart &&
    cell.row <= rowEnd &&
    cell.column >= columnStart &&
    cell.column <= columnEnd
      ? [{ sheetId: sheet.id, row: cell.row, column: cell.column }]
      : [],
  );
}

export function spreadsheetCalculationSessionUpdate(
  previous: SpreadsheetKernelWorkbook | null,
  current: SpreadsheetKernelWorkbook,
  baseDocumentRevision: number,
  forceReplace = false,
): OfficeKernelSpreadsheetSessionUpdate {
  if (
    forceReplace ||
    !previous ||
    !sameSpreadsheetStructure(previous.sheets, current.sheets)
  ) {
    return { kind: 'replace', sheets: current.sheets };
  }
  const changes: Extract<
    OfficeKernelSpreadsheetSessionUpdate,
    { kind: 'patch' }
  >['changes'] = [];
  for (
    let sheetIndex = 0;
    sheetIndex < current.sheets.length;
    sheetIndex += 1
  ) {
    const currentSheet = current.sheets[sheetIndex];
    const previousSheet = previous.sheets[sheetIndex];
    if (!currentSheet || !previousSheet) {
      return { kind: 'replace', sheets: current.sheets };
    }
    const previousCells = new Map(
      previousSheet.cells.map((cell) => [spreadsheetKernelCellKey(cell), cell]),
    );
    for (const cell of currentSheet.cells) {
      const key = spreadsheetKernelCellKey(cell);
      const previousCell = previousCells.get(key);
      previousCells.delete(key);
      if (previousCell && sameSpreadsheetKernelInputCell(previousCell, cell)) {
        continue;
      }
      changes.push({
        kind: 'upsert',
        sheetId: currentSheet.id,
        row: cell.row,
        column: cell.column,
        formula: cell.formula,
        value: cell.value,
      });
      if (changes.length > SPREADSHEET_KERNEL_MAX_PATCH_CELLS) {
        return { kind: 'replace', sheets: current.sheets };
      }
    }
    for (const cell of previousCells.values()) {
      changes.push({
        kind: 'remove',
        sheetId: currentSheet.id,
        row: cell.row,
        column: cell.column,
      });
      if (changes.length > SPREADSHEET_KERNEL_MAX_PATCH_CELLS) {
        return { kind: 'replace', sheets: current.sheets };
      }
    }
  }
  return {
    kind: 'patch',
    baseDocumentRevision,
    changes,
  };
}

export function spreadsheetCalculationFallbackCells(
  workbook: SpreadsheetKernelWorkbook,
  command: SpreadsheetCalculationCommand,
  includeDataTables = true,
): SpreadsheetKernelFallbackCell[] {
  return workbook.fallbackCells.filter((cell) => {
    if (!includeDataTables && cell.type === 'data-table') return false;
    if (command.scope === 'workbook') return true;
    return (
      cell.sheetId === command.sheetId &&
      cell.row >=
        Math.min(command.range.row[0] ?? 0, command.range.row[1] ?? 0) &&
      cell.row <=
        Math.max(command.range.row[0] ?? 0, command.range.row[1] ?? 0) &&
      cell.column >=
        Math.min(command.range.column[0] ?? 0, command.range.column[1] ?? 0) &&
      cell.column <=
        Math.max(command.range.column[0] ?? 0, command.range.column[1] ?? 0)
    );
  });
}

export function spreadsheetCalculationOps(
  sheets: readonly Sheet[],
  calculatedCells: readonly OfficeKernelSpreadsheetCalculatedCell[],
): Op[] {
  const sheetsById = new Map(
    sheets.flatMap((sheet) => (sheet.id ? [[sheet.id, sheet] as const] : [])),
  );
  return calculatedCells.flatMap((calculated) => {
    const sheet = sheetsById.get(calculated.sheetId);
    const cell = sheet?.data?.[calculated.row]?.[calculated.column];
    if (!sheet?.id || !cell?.f) return [];
    const value = cellWithCalculatedValue(cell, calculated.value);
    if (sameCalculatedCell(cell, value)) return [];
    return [
      {
        id: sheet.id,
        op: 'replace',
        path: ['data', calculated.row, calculated.column],
        value,
      } satisfies Op,
    ];
  });
}

function sameSpreadsheetStructure(
  previous: readonly OfficeKernelSpreadsheetInputSheet[],
  current: readonly OfficeKernelSpreadsheetInputSheet[],
): boolean {
  return (
    previous.length === current.length &&
    previous.every(
      (sheet, index) =>
        sheet.id === current[index]?.id &&
        sheet.name === current[index]?.name &&
        JSON.stringify(sheet.tables ?? []) ===
          JSON.stringify(current[index]?.tables ?? []),
    )
  );
}

function cellWithCalculatedValue(
  cell: Cell,
  value: OfficeKernelSpreadsheetValue,
): Cell {
  const next = { ...cell };
  if (value.kind === 'blank') {
    delete next.v;
    delete next.m;
    return next;
  }
  if (value.kind === 'number') {
    next.v = value.value;
    next.m = formattedSpreadsheetNumber(value.value, cell.ct?.fa);
    next.ct = { fa: cell.ct?.fa ?? 'General', t: 'n' };
    return next;
  }
  if (value.kind === 'boolean') {
    next.v = value.value;
    next.m = value.value ? 'TRUE' : 'FALSE';
    next.ct = { fa: 'General', t: 'b' };
    return next;
  }
  if (value.kind === 'error') {
    next.v = value.value;
    next.m = value.value;
    next.ct = { fa: cell.ct?.fa ?? 'General', t: 'e' };
    return next;
  }
  next.v = value.value;
  next.m = value.value;
  next.ct = { fa: cell.ct?.fa ?? 'General', t: 's' };
  return next;
}

function formattedSpreadsheetNumber(value: number, format?: string): string {
  try {
    if (format && format !== 'General') return String(update(format, value));
    return String(genarate(value)?.[0] ?? value);
  } catch {
    return String(value);
  }
}

function sameCalculatedCell(left: Cell, right: Cell): boolean {
  return (
    left.v === right.v &&
    left.m === right.m &&
    JSON.stringify(left.ct) === JSON.stringify(right.ct)
  );
}
