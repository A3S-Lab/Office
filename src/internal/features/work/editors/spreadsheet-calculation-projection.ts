import type { Cell, Op, Sheet } from '@fortune-sheet/core';
import type {
  OfficeKernelSpreadsheetCoordinate,
  OfficeKernelSpreadsheetInputCell,
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetSessionCellChange,
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

const SPREADSHEET_KERNEL_MAX_CELLS = 100_000;
const SPREADSHEET_KERNEL_MAX_FORMULA_CHARACTERS = 8_192;
const SPREADSHEET_KERNEL_MAX_IDENTIFIER_BYTES = 256;
const SPREADSHEET_KERNEL_MAX_SHEETS = 1_024;
export const SPREADSHEET_KERNEL_MAX_PATCH_CELLS = 10_000;
const SPREADSHEET_FINGERPRINT_MASK = 0xffff_ffff_ffff_ffffn;
const SPREADSHEET_FINGERPRINT_OFFSET_A = 0xcbf2_9ce4_8422_2325n;
const SPREADSHEET_FINGERPRINT_OFFSET_B = 0x8422_2325_cbf2_9ce4n;
const SPREADSHEET_FINGERPRINT_PRIME = 0x0000_0100_0000_01b3n;
const textEncoder = new TextEncoder();

export interface SpreadsheetKernelWorkbook {
  fallbackCells: SpreadsheetKernelFallbackCell[];
  sheets: OfficeKernelSpreadsheetInputSheet[];
  sourceKey: string;
  sourceState: SpreadsheetKernelSourceState;
}

export interface SpreadsheetKernelFallbackCell
  extends OfficeKernelSpreadsheetCoordinate {
  type: WorkSpreadsheetFormulaRangeType;
}

export interface SpreadsheetKernelOperationProjection {
  changes: OfficeKernelSpreadsheetSessionCellChange[];
  sourceChanged: boolean;
  workbook: SpreadsheetKernelWorkbook;
}

interface SpreadsheetKernelSourceState {
  cellCount: number;
  fingerprintA: bigint;
  fingerprintB: bigint;
  sourceCells: Map<string, OfficeKernelSpreadsheetInputCell>;
  structureKey: string;
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

export function refreshSpreadsheetKernelWorkbook(
  sheets: readonly Sheet[],
  fallbackCells: readonly SpreadsheetKernelFallbackCell[],
): SpreadsheetKernelWorkbook | null {
  return compileSpreadsheetKernelWorkbook(sheets, [...fallbackCells]);
}

export function projectSpreadsheetKernelWorkbookOperations(
  previous: SpreadsheetKernelWorkbook,
  sheets: readonly WorkSpreadsheetSheet[],
  operations: readonly Op[],
): SpreadsheetKernelOperationProjection | null {
  if (operations.length > SPREADSHEET_KERNEL_MAX_PATCH_CELLS) return null;
  const structureKey = spreadsheetStructureKey(sheets);
  if (
    structureKey === null ||
    structureKey !== previous.sourceState.structureKey
  ) {
    return null;
  }
  const sheetsById = new Map(
    sheets.flatMap((sheet) => (sheet.id ? [[sheet.id, sheet] as const] : [])),
  );
  const coordinates = new Map<
    string,
    { sheetId: string; row: number; column: number }
  >();
  for (const operation of operations) {
    if (
      operation.op === 'insertRowCol' ||
      operation.op === 'deleteRowCol' ||
      operation.op === 'addSheet' ||
      operation.op === 'deleteSheet' ||
      operation.path.length === 0
    ) {
      return null;
    }
    if (operation.path[0] === 'name' || operation.path[0] === 'id') {
      return null;
    }
    if (operation.path[0] !== 'data') continue;
    const row = operation.path[1];
    const column = operation.path[2];
    if (
      !operation.id ||
      !sheetsById.has(operation.id) ||
      !isSpreadsheetCoordinateIndex(row, OFFICE_KERNEL_SPREADSHEET_MAX_ROWS) ||
      !isSpreadsheetCoordinateIndex(
        column,
        OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
      )
    ) {
      return null;
    }
    const coordinate = {
      sheetId: operation.id,
      row,
      column,
    };
    coordinates.set(
      sourceCellKey(coordinate.sheetId, coordinate.row, coordinate.column),
      coordinate,
    );
  }
  if (!coordinates.size) {
    return {
      changes: [],
      sourceChanged: false,
      workbook: previous,
    };
  }

  const nextSourceCells = new Map(previous.sourceState.sourceCells);
  const nextFallbackCells = new Map(
    previous.fallbackCells.map((cell) => [
      sourceCellKey(cell.sheetId, cell.row, cell.column),
      cell,
    ]),
  );
  const kernelChanges = new Map<
    string,
    OfficeKernelSpreadsheetCoordinate & {
      cell: OfficeKernelSpreadsheetInputCell | null;
    }
  >();
  const changes: OfficeKernelSpreadsheetSessionCellChange[] = [];
  let cellCount = previous.sourceState.cellCount;
  let fingerprintA = previous.sourceState.fingerprintA;
  let fingerprintB = previous.sourceState.fingerprintB;
  let sourceChanged = false;

  for (const [key, coordinate] of coordinates) {
    const sheet = sheetsById.get(coordinate.sheetId);
    if (!sheet) return null;
    const currentCell = spreadsheetInputCell(
      spreadsheetSourceCellAt(sheet, coordinate.row, coordinate.column),
      coordinate.row,
      coordinate.column,
    );
    if (currentCell === undefined) return null;
    const previousCell = previous.sourceState.sourceCells.get(key);
    if (
      (previousCell &&
        currentCell &&
        sameSpreadsheetKernelInputCell(previousCell, currentCell)) ||
      (!previousCell && !currentCell)
    ) {
      continue;
    }
    sourceChanged = true;
    if (previousCell) {
      const previousFingerprint = spreadsheetSourceCellFingerprint(
        coordinate.sheetId,
        previousCell,
      );
      fingerprintA = subtractFingerprint(fingerprintA, previousFingerprint[0]);
      fingerprintB = subtractFingerprint(fingerprintB, previousFingerprint[1]);
      nextSourceCells.delete(key);
      cellCount -= 1;
    }
    if (currentCell) {
      const currentFingerprint = spreadsheetSourceCellFingerprint(
        coordinate.sheetId,
        currentCell,
      );
      fingerprintA = addFingerprint(fingerprintA, currentFingerprint[0]);
      fingerprintB = addFingerprint(fingerprintB, currentFingerprint[1]);
      nextSourceCells.set(key, currentCell);
      cellCount += 1;
      if (cellCount > SPREADSHEET_KERNEL_MAX_CELLS) return null;
    }

    const fallbackRange = currentCell?.formula
      ? spreadsheetFormulaRangeForCell(sheet, coordinate.row, coordinate.column)
      : undefined;
    if (fallbackRange) {
      nextFallbackCells.set(key, {
        ...coordinate,
        type: fallbackRange.type,
      });
    } else {
      nextFallbackCells.delete(key);
    }
    const kernelCell = currentCell
      ? spreadsheetKernelInputCell(currentCell, Boolean(fallbackRange))
      : null;
    kernelChanges.set(key, { ...coordinate, cell: kernelCell });
    changes.push(
      kernelCell
        ? {
            kind: 'upsert',
            ...coordinate,
            formula: kernelCell.formula,
            value: kernelCell.value,
          }
        : {
            kind: 'remove',
            ...coordinate,
          },
    );
  }
  if (!sourceChanged) {
    return {
      changes: [],
      sourceChanged: false,
      workbook: previous,
    };
  }
  const fallbackCells = [...nextFallbackCells.values()].sort(
    compareSpreadsheetCoordinates,
  );
  const sourceState: SpreadsheetKernelSourceState = {
    cellCount,
    fingerprintA,
    fingerprintB,
    sourceCells: nextSourceCells,
    structureKey,
  };
  return {
    changes,
    sourceChanged: true,
    workbook: {
      fallbackCells,
      sheets: spreadsheetSheetsWithKernelChanges(
        previous.sheets,
        kernelChanges,
      ),
      sourceKey: spreadsheetSourceKey(sourceState, fallbackCells),
      sourceState,
    },
  };
}

export function spreadsheetOperationsMayChangeCalculation(
  operations: readonly Op[],
  workbook?: SpreadsheetKernelWorkbook | null,
): boolean {
  return operations.some((operation) => {
    if (
      operation.op === 'insertRowCol' ||
      operation.op === 'deleteRowCol' ||
      operation.op === 'addSheet' ||
      operation.op === 'deleteSheet' ||
      operation.path.length === 0 ||
      operation.path[0] === 'name' ||
      operation.path[0] === 'id'
    ) {
      return true;
    }
    if (operation.path[0] !== 'data') return false;
    if (!workbook)
      return operation.path.length <= 3 || isInputCellPath(operation);
    const row = operation.path[1];
    const column = operation.path[2];
    if (
      !operation.id ||
      !isSpreadsheetCoordinateIndex(row, OFFICE_KERNEL_SPREADSHEET_MAX_ROWS) ||
      !isSpreadsheetCoordinateIndex(
        column,
        OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
      )
    ) {
      return true;
    }
    const previous = workbook.sourceState.sourceCells.get(
      sourceCellKey(operation.id, row, column),
    );
    if (operation.path.length <= 3) {
      const current = spreadsheetInputCell(
        operation.op === 'remove'
          ? null
          : (operation.value as Cell | null | undefined),
        row,
        column,
      );
      if (current === undefined) return true;
      return !(
        (previous &&
          current &&
          sameSpreadsheetKernelInputCell(previous, current)) ||
        (!previous && !current)
      );
    }
    const attribute = operation.path[3];
    if (attribute === 'f') {
      const formula =
        operation.op === 'remove' || operation.value == null
          ? undefined
          : String(operation.value);
      return previous?.formula !== formula;
    }
    if (attribute === 'v' || attribute === 'm' || attribute === 'ct') {
      return !previous?.formula;
    }
    return false;
  });
}

function isInputCellPath(operation: Op): boolean {
  if (operation.path[0] !== 'data') return false;
  if (operation.path.length <= 3) return true;
  return (
    operation.path[3] === 'f' ||
    operation.path[3] === 'v' ||
    operation.path[3] === 'm' ||
    operation.path[3] === 'ct'
  );
}

function compileSpreadsheetKernelWorkbook(
  sheets: readonly Sheet[],
  fallbackCells: SpreadsheetKernelFallbackCell[] = [],
): SpreadsheetKernelWorkbook | null {
  if (sheets.length > SPREADSHEET_KERNEL_MAX_SHEETS) return null;
  const fallbackKeys = new Set(
    fallbackCells.map((cell) =>
      sourceCellKey(cell.sheetId, cell.row, cell.column),
    ),
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
        !fallbackKeys.has(sourceCellKey(sheet.id, cell.row, cell.column))
      ) {
        return cell;
      }
      const { formula: _formula, ...cachedCell } = cell;
      return cachedCell;
    }),
  }));
  const sourceState = spreadsheetKernelSourceState(sourceSheets);
  const sortedFallbackCells = [...fallbackCells].sort(
    compareSpreadsheetCoordinates,
  );
  return {
    fallbackCells: sortedFallbackCells,
    sheets: compiled,
    sourceKey: spreadsheetSourceKey(sourceState, sortedFallbackCells),
    sourceState,
  };
}

function sourceCellKey(sheetId: string, row: number, column: number): string {
  return `${sheetId}\u0000${row}\u0000${column}`;
}

function spreadsheetKernelSourceState(
  sheets: readonly OfficeKernelSpreadsheetInputSheet[],
): SpreadsheetKernelSourceState {
  const sourceCells = new Map<string, OfficeKernelSpreadsheetInputCell>();
  let fingerprintA = 0n;
  let fingerprintB = 0n;
  for (const sheet of sheets) {
    for (const cell of sheet.cells) {
      const key = sourceCellKey(sheet.id, cell.row, cell.column);
      const fingerprint = spreadsheetSourceCellFingerprint(sheet.id, cell);
      sourceCells.set(key, cell);
      fingerprintA = addFingerprint(fingerprintA, fingerprint[0]);
      fingerprintB = addFingerprint(fingerprintB, fingerprint[1]);
    }
  }
  return {
    cellCount: sourceCells.size,
    fingerprintA,
    fingerprintB,
    sourceCells,
    structureKey: JSON.stringify(
      sheets.map((sheet) => ({ id: sheet.id, name: sheet.name })),
    ),
  };
}

function spreadsheetStructureKey(sheets: readonly Sheet[]): string | null {
  if (sheets.length > SPREADSHEET_KERNEL_MAX_SHEETS) return null;
  const ids = new Set<string>();
  const names = new Set<string>();
  const structure: Array<{ id: string; name: string }> = [];
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
      ids.has(id) ||
      names.has(normalizedName)
    ) {
      return null;
    }
    ids.add(id);
    names.add(normalizedName);
    structure.push({ id, name });
  }
  return JSON.stringify(structure);
}

function spreadsheetSourceKey(
  sourceState: SpreadsheetKernelSourceState,
  fallbackCells: readonly SpreadsheetKernelFallbackCell[],
): string {
  let fallbackA = 0n;
  let fallbackB = 0n;
  for (const cell of fallbackCells) {
    const fingerprint = spreadsheetStringFingerprint(
      JSON.stringify([cell.sheetId, cell.row, cell.column, cell.type]),
    );
    fallbackA = addFingerprint(fallbackA, fingerprint[0]);
    fallbackB = addFingerprint(fallbackB, fingerprint[1]);
  }
  return [
    'v2',
    sourceState.structureKey,
    sourceState.cellCount,
    fingerprintHex(sourceState.fingerprintA),
    fingerprintHex(sourceState.fingerprintB),
    fallbackCells.length,
    fingerprintHex(fallbackA),
    fingerprintHex(fallbackB),
  ].join(':');
}

function spreadsheetSourceCellFingerprint(
  sheetId: string,
  cell: OfficeKernelSpreadsheetInputCell,
): readonly [bigint, bigint] {
  return spreadsheetStringFingerprint(
    JSON.stringify([
      sheetId,
      cell.row,
      cell.column,
      cell.formula
        ? ['formula', cell.formula]
        : ['value', cell.value.kind, spreadsheetValuePayload(cell.value)],
    ]),
  );
}

function spreadsheetStringFingerprint(
  value: string,
): readonly [bigint, bigint] {
  let fingerprintA = SPREADSHEET_FINGERPRINT_OFFSET_A;
  let fingerprintB = SPREADSHEET_FINGERPRINT_OFFSET_B;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    fingerprintA = fingerprintByte(fingerprintA, code & 0xff);
    fingerprintA = fingerprintByte(fingerprintA, code >>> 8);
    fingerprintB = fingerprintByte(fingerprintB, code >>> 8);
    fingerprintB = fingerprintByte(fingerprintB, code & 0xff);
  }
  return [fingerprintA, fingerprintB];
}

function fingerprintByte(fingerprint: bigint, value: number): bigint {
  return (
    ((fingerprint ^ BigInt(value)) * SPREADSHEET_FINGERPRINT_PRIME) &
    SPREADSHEET_FINGERPRINT_MASK
  );
}

function addFingerprint(left: bigint, right: bigint): bigint {
  return (left + right) & SPREADSHEET_FINGERPRINT_MASK;
}

function subtractFingerprint(left: bigint, right: bigint): bigint {
  return (left - right) & SPREADSHEET_FINGERPRINT_MASK;
}

function fingerprintHex(value: bigint): string {
  return value.toString(16).padStart(16, '0');
}

function spreadsheetValuePayload(
  value: OfficeKernelSpreadsheetValue,
): string | number | boolean | null {
  return value.kind === 'blank' ? null : value.value;
}

function spreadsheetSourceCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null | undefined {
  if (sheet.data) return sheet.data[row]?.[column];
  return sheet.celldata?.find((entry) => entry.r === row && entry.c === column)
    ?.v;
}

function spreadsheetKernelInputCell(
  cell: OfficeKernelSpreadsheetInputCell,
  fallback: boolean,
): OfficeKernelSpreadsheetInputCell {
  if (!fallback || !cell.formula) return cell;
  const { formula: _formula, ...cachedCell } = cell;
  return cachedCell;
}

function spreadsheetSheetsWithKernelChanges(
  sheets: readonly OfficeKernelSpreadsheetInputSheet[],
  changes: ReadonlyMap<
    string,
    OfficeKernelSpreadsheetCoordinate & {
      cell: OfficeKernelSpreadsheetInputCell | null;
    }
  >,
): OfficeKernelSpreadsheetInputSheet[] {
  const changesBySheet = new Map<
    string,
    Array<
      OfficeKernelSpreadsheetCoordinate & {
        cell: OfficeKernelSpreadsheetInputCell | null;
      }
    >
  >();
  for (const change of changes.values()) {
    const sheetChanges = changesBySheet.get(change.sheetId) ?? [];
    sheetChanges.push(change);
    changesBySheet.set(change.sheetId, sheetChanges);
  }
  return sheets.map((sheet) => {
    const sheetChanges = changesBySheet.get(sheet.id);
    if (!sheetChanges?.length) return sheet;
    const cells = new Map(
      sheet.cells.map((cell) => [spreadsheetKernelCellKey(cell), cell]),
    );
    for (const change of sheetChanges) {
      const key = spreadsheetKernelCellKey(change);
      if (change.cell) cells.set(key, change.cell);
      else cells.delete(key);
    }
    return {
      ...sheet,
      cells: [...cells.values()].sort(compareSpreadsheetInputCells),
    };
  });
}

function compareSpreadsheetCoordinates(
  left: OfficeKernelSpreadsheetCoordinate,
  right: OfficeKernelSpreadsheetCoordinate,
): number {
  return (
    left.sheetId.localeCompare(right.sheetId) ||
    left.row - right.row ||
    left.column - right.column
  );
}

function compareSpreadsheetInputCells(
  left: OfficeKernelSpreadsheetInputCell,
  right: OfficeKernelSpreadsheetInputCell,
): number {
  return left.row - right.row || left.column - right.column;
}

function isSpreadsheetCoordinateIndex(
  value: unknown,
  limit: number,
): value is number {
  return (
    typeof value === 'number' &&
    Number.isSafeInteger(value) &&
    value >= 0 &&
    value < limit
  );
}

export function spreadsheetKernelCellKey(
  cell: Pick<OfficeKernelSpreadsheetInputCell, 'row' | 'column'>,
): string {
  return `${cell.row}:${cell.column}`;
}

export function sameSpreadsheetKernelInputCell(
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
