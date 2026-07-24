import {
  genarate,
  update,
  type Cell,
  type Op,
  type Sheet,
} from '@fortune-sheet/core';
import type {
  OfficeKernelSpreadsheetCalculatedCell,
  OfficeKernelSpreadsheetCoordinate,
  OfficeKernelSpreadsheetInputCell,
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetSessionUpdate,
  OfficeKernelSpreadsheetValue,
} from '../../../kernel/office-kernel-protocol';
import {
  isOfficeKernelSpreadsheetError,
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
  OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES,
} from '../../../kernel/office-kernel-spreadsheet-protocol';
import { spreadsheetFormulaRangeForCell } from '../work-spreadsheet-formulas';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetFormulaRangeType,
  WorkSpreadsheetSheet,
} from '../work-types';
import type { SpreadsheetCalculationCommand } from './spreadsheet-command-controller';

const SPREADSHEET_KERNEL_MAX_CELLS = 100_000;
const SPREADSHEET_KERNEL_MAX_FORMULA_CHARACTERS = 8_192;
const SPREADSHEET_KERNEL_MAX_IDENTIFIER_BYTES = 256;
const SPREADSHEET_KERNEL_MAX_SHEETS = 1_024;
const SPREADSHEET_KERNEL_MAX_PATCH_CELLS = 10_000;
const textEncoder = new TextEncoder();

export interface SpreadsheetKernelWorkbook {
  fallbackCells: SpreadsheetKernelFallbackCell[];
  sheets: OfficeKernelSpreadsheetInputSheet[];
  sourceKey: string;
}

export interface SpreadsheetKernelFallbackCell
  extends OfficeKernelSpreadsheetCoordinate {
  type: WorkSpreadsheetFormulaRangeType;
}

export function createSpreadsheetKernelWorkbook(
  content: Pick<WorkSpreadsheetContent, 'sheets'>,
): SpreadsheetKernelWorkbook | null {
  const fallbackCells = content.sheets.flatMap((sheet) =>
    spreadsheetFormulaCoordinates(sheet).flatMap(({ row, column }) => {
      if (!sheet.id) return [];
      const range = spreadsheetFormulaRangeForCell(sheet, row, column);
      return range
        ? [
            {
              sheetId: sheet.id,
              row,
              column,
              type: range.type,
            },
          ]
        : [];
    }),
  );
  return compileSpreadsheetKernelWorkbook(content.sheets, fallbackCells);
}

export function spreadsheetCalculationSourceKey(
  sheets: readonly Sheet[],
): string | null {
  return compileSpreadsheetKernelWorkbook(sheets)?.sourceKey ?? null;
}

export function refreshSpreadsheetKernelWorkbook(
  sheets: readonly Sheet[],
  fallbackCells: readonly SpreadsheetKernelFallbackCell[],
): SpreadsheetKernelWorkbook | null {
  return compileSpreadsheetKernelWorkbook(sheets, [...fallbackCells]);
}

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
      previousSheet.cells.map((cell) => [spreadsheetCellKey(cell), cell]),
    );
    for (const cell of currentSheet.cells) {
      const key = spreadsheetCellKey(cell);
      const previousCell = previousCells.get(key);
      previousCells.delete(key);
      if (previousCell && sameSpreadsheetInputCell(previousCell, cell)) {
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

function compileSpreadsheetKernelWorkbook(
  sheets: readonly Sheet[],
  fallbackCells: SpreadsheetKernelFallbackCell[] = [],
): SpreadsheetKernelWorkbook | null {
  if (sheets.length > SPREADSHEET_KERNEL_MAX_SHEETS) return null;
  const fallbackKeys = new Set(
    fallbackCells.map((cell) => cellKey(cell.sheetId, cell.row, cell.column)),
  );
  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();
  const sourceSheets: OfficeKernelSpreadsheetInputSheet[] = [];
  let cellCount = 0;
  for (const sheet of sheets) {
    const id = sheet.id;
    const name = sheet.name;
    const normalizedName = name?.toLowerCase();
    if (
      !id?.trim() ||
      !name?.trim() ||
      !normalizedName ||
      utf8ByteLength(id) > SPREADSHEET_KERNEL_MAX_IDENTIFIER_BYTES ||
      utf8ByteLength(name) > SPREADSHEET_KERNEL_MAX_IDENTIFIER_BYTES ||
      sheetIds.has(id) ||
      sheetNames.has(normalizedName)
    ) {
      return null;
    }
    const cells = sparseSpreadsheetCells(
      sheet,
      SPREADSHEET_KERNEL_MAX_CELLS - cellCount,
    );
    if (!cells) return null;
    cellCount += cells.length;
    sheetIds.add(id);
    sheetNames.add(normalizedName);
    sourceSheets.push({ id, name, cells });
  }
  const compiled = sourceSheets.map((sheet) => ({
    ...sheet,
    cells: sheet.cells.map((cell) => {
      if (
        !cell.formula ||
        !fallbackKeys.has(cellKey(sheet.id, cell.row, cell.column))
      ) {
        return cell;
      }
      const { formula: _formula, ...cachedCell } = cell;
      return cachedCell;
    }),
  }));
  return {
    fallbackCells,
    sheets: compiled,
    sourceKey: JSON.stringify(
      sourceSheets.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        cells: sheet.cells.map((cell) =>
          cell.formula
            ? {
                row: cell.row,
                column: cell.column,
                formula: cell.formula,
              }
            : cell,
        ),
      })),
    ),
  };
}

function cellKey(sheetId: string, row: number, column: number): string {
  return `${sheetId}\u0000${row}\u0000${column}`;
}

function spreadsheetCellKey(
  cell: Pick<OfficeKernelSpreadsheetInputCell, 'row' | 'column'>,
): string {
  return `${cell.row}:${cell.column}`;
}

function sameSpreadsheetStructure(
  previous: readonly OfficeKernelSpreadsheetInputSheet[],
  current: readonly OfficeKernelSpreadsheetInputSheet[],
): boolean {
  return (
    previous.length === current.length &&
    previous.every(
      (sheet, index) =>
        sheet.id === current[index]?.id && sheet.name === current[index]?.name,
    )
  );
}

function sameSpreadsheetInputCell(
  previous: OfficeKernelSpreadsheetInputCell,
  current: OfficeKernelSpreadsheetInputCell,
): boolean {
  if (previous.formula && previous.formula === current.formula) {
    return true;
  }
  return (
    previous.formula === current.formula &&
    sameSpreadsheetValue(previous.value, current.value)
  );
}

function sameSpreadsheetValue(
  previous: OfficeKernelSpreadsheetValue,
  current: OfficeKernelSpreadsheetValue,
): boolean {
  if (previous.kind !== current.kind) return false;
  if (previous.kind === 'blank' || current.kind === 'blank') {
    return previous.kind === current.kind;
  }
  return previous.value === current.value;
}

function spreadsheetFormulaCoordinates(
  sheet: WorkSpreadsheetSheet,
): Array<{ row: number; column: number }> {
  if (sheet.data) {
    return sheet.data.flatMap((cells, row) =>
      cells.flatMap((cell, column) => (cell?.f ? [{ row, column }] : [])),
    );
  }
  return (sheet.celldata ?? []).flatMap((entry) =>
    entry.v?.f ? [{ row: entry.r, column: entry.c }] : [],
  );
}

function sparseSpreadsheetCells(
  sheet: Sheet,
  maximumCells: number,
): OfficeKernelSpreadsheetInputCell[] | null {
  const cells: OfficeKernelSpreadsheetInputCell[] = [];
  const coordinates = new Set<string>();
  const appendCell = (
    cell: Cell | null | undefined,
    row: number,
    column: number,
  ): boolean => {
    const input = spreadsheetInputCell(cell, row, column);
    if (input === undefined) return false;
    if (!input) return true;
    const coordinate = `${row}:${column}`;
    if (
      coordinates.has(coordinate) ||
      cells.length >= maximumCells ||
      cells.length >= SPREADSHEET_KERNEL_MAX_CELLS
    ) {
      return false;
    }
    coordinates.add(coordinate);
    cells.push(input);
    return true;
  };
  if (sheet.data) {
    if (sheet.data.length > OFFICE_KERNEL_SPREADSHEET_MAX_ROWS) return null;
    for (let row = 0; row < sheet.data.length; row += 1) {
      const values = sheet.data[row];
      if (!values) continue;
      if (values.length > OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS) return null;
      for (let column = 0; column < values.length; column += 1) {
        if (!appendCell(values[column], row, column)) return null;
      }
    }
  } else {
    for (const entry of sheet.celldata ?? []) {
      if (!appendCell(entry.v, entry.r, entry.c)) return null;
    }
  }
  return cells.sort(
    (left, right) => left.row - right.row || left.column - right.column,
  );
}

function spreadsheetInputCell(
  cell: Cell | null | undefined,
  row: number,
  column: number,
): OfficeKernelSpreadsheetInputCell | null | undefined {
  if (!cell || (!cell.f && cell.v === undefined && cell.m === undefined)) {
    return null;
  }
  if (
    !Number.isSafeInteger(row) ||
    row < 0 ||
    row >= OFFICE_KERNEL_SPREADSHEET_MAX_ROWS ||
    !Number.isSafeInteger(column) ||
    column < 0 ||
    column >= OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS ||
    (cell.f &&
      Array.from(cell.f).length > SPREADSHEET_KERNEL_MAX_FORMULA_CHARACTERS)
  ) {
    return undefined;
  }
  const value = spreadsheetKernelValue(cell);
  if (!value) return undefined;
  return {
    row,
    column,
    formula: cell.f,
    value,
  };
}

function spreadsheetKernelValue(
  cell: Cell,
): OfficeKernelSpreadsheetValue | null {
  const raw = cell.v ?? cell.m;
  if (cell.ct?.t === 'e') {
    const value = String(raw ?? '#VALUE!');
    return {
      kind: 'error',
      value: isOfficeKernelSpreadsheetError(value) ? value : '#VALUE!',
    };
  }
  if (raw === undefined || raw === null || raw === '') return { kind: 'blank' };
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? { kind: 'number', value: raw } : null;
  }
  if (typeof raw === 'boolean') return { kind: 'boolean', value: raw };
  const value = String(raw);
  return utf8ByteLength(value) <= OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES
    ? { kind: 'text', value }
    : null;
}

function utf8ByteLength(value: string): number {
  return textEncoder.encode(value).byteLength;
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
