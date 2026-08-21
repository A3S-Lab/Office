import type { Cell, CellMatrix, Op, Selection } from '@fortune-sheet/core';
import { sameSpreadsheetHistoryValue } from '../spreadsheet-history-value';
import { sparseArrayIndexes } from '../spreadsheet-sparse';
import {
  attachSpreadsheetShownCommentCells,
  spreadsheetMatrixProfile,
} from '../work-spreadsheet-matrix-profile';
import type { WorkSpreadsheetContent } from '../work-types';
import { officeFontFamilies } from './office-font-families';
import type { OfficeSelectOption } from './office-select';
import {
  MAXIMUM_INCREMENTAL_SPREADSHEET_OPERATIONS,
  projectSpreadsheetSheetsFromFortuneOperations,
} from './spreadsheet-operation-projection';
import { reconcileSpreadsheetTablesAfterFortune } from './spreadsheet-table-reconciliation';

const spreadsheetFontSizes = [
  9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72,
] as const;

export interface SpreadsheetSelectionSummary {
  average: number | null;
  nonEmptyCount: number;
  numericCount: number;
  sum: number | null;
}

export function spreadsheetSelectionReference(
  selection: Pick<Selection, 'row' | 'column'>,
): string {
  const rowStart = Math.min(
    selection.row[0] ?? 0,
    selection.row[1] ?? selection.row[0] ?? 0,
  );
  const rowEnd = Math.max(
    selection.row[0] ?? 0,
    selection.row[1] ?? selection.row[0] ?? 0,
  );
  const columnStart = Math.min(
    selection.column[0] ?? 0,
    selection.column[1] ?? selection.column[0] ?? 0,
  );
  const columnEnd = Math.max(
    selection.column[0] ?? 0,
    selection.column[1] ?? selection.column[0] ?? 0,
  );
  const start = `${spreadsheetColumnLabel(columnStart)}${rowStart + 1}`;
  const end = `${spreadsheetColumnLabel(columnEnd)}${rowEnd + 1}`;
  return start === end ? start : `${start}:${end}`;
}

export function spreadsheetSingleRange(
  selection: Pick<Selection, 'row' | 'column'>,
): {
  row: number[];
  column: number[];
} {
  return {
    row: finiteSpreadsheetSelectionAxis(selection.row),
    column: finiteSpreadsheetSelectionAxis(selection.column),
  };
}

export function spreadsheetCellAt(
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined,
  row: number | undefined,
  column: number | undefined,
): Cell | null {
  if (!sheet) return null;
  const safeRow = finiteSpreadsheetIndex(row, 0);
  const safeColumn = finiteSpreadsheetIndex(column, 0);
  return (
    sheet.data?.[safeRow]?.[safeColumn] ??
    sheet.celldata?.find(
      (entry) => entry.r === safeRow && entry.c === safeColumn,
    )?.v ??
    null
  );
}

export function spreadsheetSelectionSummary(
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined,
  selection: Pick<Selection, 'row' | 'column'>,
): SpreadsheetSelectionSummary {
  const summary: SpreadsheetSelectionSummary = {
    average: null,
    nonEmptyCount: 0,
    numericCount: 0,
    sum: null,
  };
  if (!sheet) return summary;
  const range = spreadsheetSingleRange(selection);
  let sum = 0;
  const collect = (cell: Cell | null | undefined) => {
    if (!cell || !spreadsheetCellHasContent(cell)) return;
    summary.nonEmptyCount += 1;
    if (typeof cell.v !== 'number' || !Number.isFinite(cell.v)) return;
    summary.numericCount += 1;
    sum += cell.v;
  };

  if (sheet.data?.length) {
    const rowStart = range.row[0] ?? 0;
    const rowEnd = Math.min(range.row[1] ?? rowStart, sheet.data.length - 1);
    const columnStart = range.column[0] ?? 0;
    const columnEnd = range.column[1] ?? columnStart;
    for (const rowIndex of sparseArrayIndexes(sheet.data)) {
      if (rowIndex < rowStart || rowIndex > rowEnd) continue;
      const row = sheet.data[rowIndex];
      if (!row) continue;
      for (const columnIndex of sparseArrayIndexes(row)) {
        if (columnIndex < columnStart || columnIndex > columnEnd) continue;
        collect(row[columnIndex]);
      }
    }
  } else {
    const rowStart = range.row[0] ?? 0;
    const rowEnd = range.row[1] ?? rowStart;
    const columnStart = range.column[0] ?? 0;
    const columnEnd = range.column[1] ?? columnStart;
    for (const entry of sheet.celldata ?? []) {
      if (
        entry.r < rowStart ||
        entry.r > rowEnd ||
        entry.c < columnStart ||
        entry.c > columnEnd
      ) {
        continue;
      }
      collect(entry.v);
    }
  }

  if (summary.numericCount > 0) {
    summary.sum = sum;
    summary.average = sum / summary.numericCount;
  }
  return summary;
}

export function spreadsheetFontSizeOptions(
  current: number | undefined,
): { value: string; label: string }[] {
  const values: number[] = [...spreadsheetFontSizes];
  if (current && !values.includes(current)) values.push(current);
  return values
    .sort((left, right) => left - right)
    .map((value) => ({ value: String(value), label: String(value) }));
}

export function spreadsheetFontFamilyOptions(
  current: string | undefined,
): OfficeSelectOption[] {
  const options: OfficeSelectOption[] = officeFontFamilies.map(
    ({ cssFamily, group, label, name }) => ({
      value: name,
      group,
      label,
      previewStyle: { fontFamily: cssFamily },
      searchText: `${name} ${label}`,
    }),
  );
  const normalizedCurrent = current?.trim();
  if (
    normalizedCurrent &&
    !options.some(({ value }) => value === normalizedCurrent)
  ) {
    options.push({
      value: normalizedCurrent,
      group: '文档字体',
      label: normalizedCurrent,
      previewStyle: { fontFamily: normalizedCurrent },
      searchText: normalizedCurrent,
    });
  }
  return options;
}

export function spreadsheetSheetsWithFiniteSelections(
  sheets: WorkSpreadsheetContent['sheets'],
): WorkSpreadsheetContent['sheets'] {
  return sheets.map((sheet) => ({
    ...sheet,
    luckysheet_select_save: (sheet.luckysheet_select_save?.length
      ? sheet.luckysheet_select_save
      : [undefined]
    ).map(finiteSpreadsheetSelection),
  }));
}

export function spreadsheetSheetsForFortune(
  sheets: WorkSpreadsheetContent['sheets'],
): WorkSpreadsheetContent['sheets'] {
  const startedAt = spreadsheetProjectionNow();
  let populatedCellCount = 0;
  const projectedSheets = sheets.map((sheet) => {
    const { celldata: sourceCellData, data, ...sourceMetadata } = sheet;
    const projected = structuredClone(sourceMetadata);
    const profile = spreadsheetMatrixProfile(data);
    if (data && profile?.fortuneReady) {
      populatedCellCount += profile.populatedCellCount;
      return { ...projected, data };
    }
    const fortuneProjection = spreadsheetMatrixForFortune(
      {
        ...projected,
        ...(data === undefined ? {} : { data }),
        ...(sourceCellData === undefined ? {} : { celldata: sourceCellData }),
      },
      projected.config?.merge,
    );
    populatedCellCount += fortuneProjection.populatedCellCount;
    return { ...projected, data: fortuneProjection.data };
  });
  recordSpreadsheetProjectionMeasure(startedAt, spreadsheetProjectionNow(), {
    populatedCellCount,
    sheetCount: sheets.length,
  });
  return projectedSheets;
}

function spreadsheetMatrixForFortune(
  sheet: WorkSpreadsheetContent['sheets'][number],
  merges: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['config']
  >['merge'],
): {
  data: CellMatrix;
  populatedCellCount: number;
} {
  const declaredRows = positiveSpreadsheetDimension(sheet.row);
  const declaredColumns = positiveSpreadsheetDimension(sheet.column);
  const hasDeclaredGrid = declaredRows > 0 && declaredColumns > 0;
  let rowCount = hasDeclaredGrid ? declaredRows : 60;
  let columnCount = hasDeclaredGrid ? declaredColumns : 26;
  let populatedCellCount = 0;
  const shownCommentCells: Array<{ c: number; r: number }> = [];
  const projected: CellMatrix = [];

  const setSourceCell = (rowIndex: number, columnIndex: number, cell: Cell) => {
    rowCount = Math.max(rowCount, rowIndex + 1);
    columnCount = Math.max(columnCount, columnIndex + 1);
    const row = projected[rowIndex] ?? [];
    projected[rowIndex] = row;
    row[columnIndex] = cloneSpreadsheetCellForFortune(cell);
    if (cell.ps?.isShow) {
      shownCommentCells.push({ c: columnIndex, r: rowIndex });
    }
    populatedCellCount += 1;
  };

  if (sheet.data !== undefined) {
    rowCount = Math.max(rowCount, sheet.data.length);
    for (const rowIndex of sparseArrayIndexes(sheet.data)) {
      const sourceRow = sheet.data[rowIndex];
      if (!sourceRow) continue;
      columnCount = Math.max(columnCount, sourceRow.length);
      for (const columnIndex of sparseArrayIndexes(sourceRow)) {
        const cell = sourceRow[columnIndex];
        if (cell != null) setSourceCell(rowIndex, columnIndex, cell);
      }
    }
  } else {
    for (const entry of sheet.celldata ?? []) {
      if (entry.v != null) setSourceCell(entry.r, entry.c, entry.v);
    }
  }

  for (const merge of Object.values(merges ?? {})) {
    rowCount = Math.max(rowCount, merge.r + merge.rs);
    columnCount = Math.max(columnCount, merge.c + merge.cs);
    for (let rowIndex = merge.r; rowIndex < merge.r + merge.rs; rowIndex += 1) {
      const row = projected[rowIndex] ?? [];
      projected[rowIndex] = row;
      for (
        let columnIndex = merge.c;
        columnIndex < merge.c + merge.cs;
        columnIndex += 1
      ) {
        row[columnIndex] = Object.freeze({
          ...(row[columnIndex] ?? {}),
          mc:
            rowIndex === merge.r && columnIndex === merge.c
              ? { ...merge }
              : { r: merge.r, c: merge.c },
        });
      }
    }
  }

  projected.length = rowCount;
  projected[0] ??= [];
  for (const rowIndex of sparseArrayIndexes(projected)) {
    const row = projected[rowIndex];
    if (!row) continue;
    row.length = Math.max(row.length, columnCount);
    Object.freeze(row);
  }
  attachSpreadsheetShownCommentCells(
    projected,
    Object.freeze(shownCommentCells.map((cell) => Object.freeze(cell))),
  );
  Object.freeze(projected);
  return { data: projected, populatedCellCount };
}

function cloneSpreadsheetCellForFortune(cell: Cell): Cell {
  const clone = { ...cell };
  if (cell.mc) clone.mc = { ...cell.mc };
  if (cell.ct) {
    clone.ct = {
      ...cell.ct,
      ...(cell.ct.s === undefined ? {} : { s: structuredClone(cell.ct.s) }),
    };
  }
  if (cell.ps) clone.ps = { ...cell.ps };
  if (cell.hl) clone.hl = { ...cell.hl };
  if (cell.spl && typeof cell.spl === 'object') {
    clone.spl = structuredClone(cell.spl);
  }
  Object.freeze(clone);
  return clone;
}

function spreadsheetProjectionNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function recordSpreadsheetProjectionMeasure(
  start: number,
  end: number,
  detail: Record<string, number>,
): void {
  try {
    globalThis.performance?.measure(
      'a3s-office.spreadsheet.fortune-projection',
      { detail, end, start },
    );
  } catch {
    // User Timing diagnostics must never affect the editor boundary.
  }
}

export function spreadsheetSheetsFromFortune(
  sheets: WorkSpreadsheetContent['sheets'],
  sourceSheets: WorkSpreadsheetContent['sheets'],
  operations: readonly Op[] = [],
): WorkSpreadsheetContent['sheets'] {
  const startedAt = spreadsheetProjectionNow();
  const incremental = projectSpreadsheetSheetsFromFortuneOperations(
    sheets,
    sourceSheets,
    operations,
  );
  if (incremental) {
    recordSpreadsheetControlledProjectionMeasure(
      startedAt,
      spreadsheetProjectionNow(),
      {
        affectedCellCount: incremental.affectedCellCount,
        incremental: true,
        operationCount: operations.length,
      },
    );
    return reconcileSpreadsheetTablesAfterFortune(
      incremental.sheets,
      sourceSheets,
      operations,
    );
  }

  const projected = sheets.map((sheet, index) => {
    const source =
      (sheet.id
        ? sourceSheets.find((candidate) => candidate.id === sheet.id)
        : undefined) ?? sourceSheets[index];
    const cells = spreadsheetPopulatedCellData(sheet);
    const { celldata: _cellData, data: _data, ...metadata } = sheet;
    const useMatrix = source
      ? source.data !== undefined
      : sheet.data !== undefined;
    if (!useMatrix) return { ...metadata, celldata: cells };

    const data: CellMatrix = [];
    data.length = Math.max(
      sheet.data?.length ?? 0,
      source?.data?.length ?? 0,
      positiveSpreadsheetDimension(sheet.row),
    );
    for (const { r: rowIndex, c: columnIndex, v: cell } of cells) {
      const row = data[rowIndex] ?? [];
      data[rowIndex] = row;
      row.length = Math.max(
        row.length,
        sheet.data?.[rowIndex]?.length ?? 0,
        source?.data?.[rowIndex]?.length ?? 0,
        positiveSpreadsheetDimension(sheet.column),
      );
      row[columnIndex] = cell;
    }
    return { ...metadata, data };
  });
  recordSpreadsheetControlledProjectionMeasure(
    startedAt,
    spreadsheetProjectionNow(),
    {
      affectedCellCount: 0,
      incremental: false,
      operationCount: operations.length,
    },
  );
  return reconcileSpreadsheetTablesAfterFortune(
    projected,
    sourceSheets,
    operations,
  );
}

function recordSpreadsheetControlledProjectionMeasure(
  start: number,
  end: number,
  detail: Record<string, boolean | number>,
): void {
  try {
    globalThis.performance?.measure(
      'a3s-office.spreadsheet.controlled-projection',
      { detail, end, start },
    );
  } catch {
    // User Timing diagnostics must never affect the editor boundary.
  }
}

export function finiteSpreadsheetSelection(
  selection: Selection | undefined,
): Selection {
  const row = finiteSpreadsheetSelectionAxis(selection?.row);
  const column = finiteSpreadsheetSelectionAxis(selection?.column);
  return {
    ...selection,
    row,
    column,
    row_focus: finiteSpreadsheetFocus(selection?.row_focus, row),
    column_focus: finiteSpreadsheetFocus(selection?.column_focus, column),
  };
}

export function sameSpreadsheetWorkbookState(
  changed: WorkSpreadsheetContent['sheets'],
  rendered: WorkSpreadsheetContent['sheets'],
): boolean {
  return (
    changed.length === rendered.length &&
    changed.every((sheet, index) =>
      sameSpreadsheetSheetState(sheet, rendered[index]),
    )
  );
}

/**
 * Answers the workbook equality check from Fortune's exact operation batch
 * without walking unrelated matrix cells. A null result requires the legacy
 * full comparison because the batch is missing or too broad to prove safely.
 */
export function sameSpreadsheetWorkbookStateAfterOperations(
  changed: WorkSpreadsheetContent['sheets'],
  rendered: WorkSpreadsheetContent['sheets'],
  operations: readonly Op[],
): boolean | null {
  if (
    operations.length === 0 ||
    operations.length > MAXIMUM_INCREMENTAL_SPREADSHEET_OPERATIONS
  ) {
    return null;
  }
  if (
    changed.length !== rendered.length ||
    !changed.every((sheet, index) =>
      sameSpreadsheetHistoryValue(
        spreadsheetSheetMetadata(sheet),
        spreadsheetSheetMetadata(rendered[index]),
      ),
    )
  ) {
    return false;
  }

  const changedById = spreadsheetSheetsById(changed);
  const renderedById = spreadsheetSheetsById(rendered);
  if (!changedById || !renderedById) return null;
  const coordinates = new Map<
    string,
    { column: number; row: number; sheetId: string }
  >();
  for (const operation of operations) {
    if (
      operation.op === 'insertRowCol' ||
      operation.op === 'deleteRowCol' ||
      operation.op === 'addSheet' ||
      operation.op === 'deleteSheet' ||
      operation.path.length === 0
    ) {
      return false;
    }
    if (
      !operation.id ||
      !changedById.has(operation.id) ||
      !renderedById.has(operation.id)
    ) {
      return null;
    }
    if (operation.path[0] !== 'data') continue;
    const row = operation.path[1];
    const column = operation.path[2];
    if (
      operation.path.length < 3 ||
      !isSpreadsheetCoordinate(row) ||
      !isSpreadsheetCoordinate(column)
    ) {
      return null;
    }
    coordinates.set(`${operation.id}:${row}:${column}`, {
      column,
      row,
      sheetId: operation.id,
    });
  }

  for (const { column, row, sheetId } of coordinates.values()) {
    if (
      !sameSpreadsheetHistoryValue(
        spreadsheetCellAt(changedById.get(sheetId), row, column),
        spreadsheetCellAt(renderedById.get(sheetId), row, column),
      )
    ) {
      return false;
    }
  }
  return true;
}

export function sameSpreadsheetHistoryContent(
  left: WorkSpreadsheetContent,
  right: WorkSpreadsheetContent,
): boolean {
  if (left === right) return true;
  const { sheets: leftSheets, ...leftMetadata } = left;
  const { sheets: rightSheets, ...rightMetadata } = right;
  return (
    sameSpreadsheetHistoryValue(leftMetadata, rightMetadata) &&
    leftSheets.length === rightSheets.length &&
    leftSheets.every((sheet, index) =>
      sameSpreadsheetSheetState(sheet, rightSheets[index]),
    )
  );
}

export function spreadsheetContentWithSelection(
  content: WorkSpreadsheetContent,
  sheetId: string,
  selection: Selection | null | undefined,
): WorkSpreadsheetContent {
  return spreadsheetContentWithSelections(
    content,
    sheetId,
    selection ? [selection] : [],
  );
}

export function spreadsheetContentWithSelections(
  content: WorkSpreadsheetContent,
  sheetId: string,
  selections: readonly Selection[],
): WorkSpreadsheetContent {
  if (!sheetId || !selections.length) return content;
  const nextSelections = selections.map(finiteSpreadsheetSelection);
  let changed = false;
  const sheets = content.sheets.map((sheet) => {
    if (sheet.id !== sheetId) return sheet;
    changed = true;
    return { ...sheet, luckysheet_select_save: nextSelections };
  });
  return changed ? { ...content, sheets } : content;
}

export function isSpreadsheetNativeTextUndoTarget(
  target: EventTarget | null,
): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (isSpreadsheetCellEditingTarget(target)) return true;
  if (target.closest('.fortune-container')) return false;
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    target.isContentEditable ||
    Boolean(target.closest('[contenteditable="true"]'))
  );
}

export function isSpreadsheetCellEditingTarget(
  target: EventTarget | null,
): boolean {
  if (!(target instanceof Element)) return false;
  const formulaInput = target.closest('.fortune-fx-input');
  if (formulaInput) return true;
  const cellInput = target.closest('.luckysheet-cell-input');
  if (!cellInput) return false;
  const inputBox = cellInput.closest<HTMLElement>('.luckysheet-input-box');
  if (!inputBox) return true;
  const zIndex = Number.parseInt(
    inputBox.style.zIndex || getComputedStyle(inputBox).zIndex,
    10,
  );
  return !Number.isFinite(zIndex) || zIndex >= 0;
}

export function spreadsheetFormulaBarSelectAllTarget(
  event: KeyboardEvent,
): HTMLElement | null {
  if (
    event.defaultPrevented ||
    event.repeat ||
    event.altKey ||
    event.shiftKey ||
    !(event.metaKey || event.ctrlKey) ||
    event.key.toLocaleLowerCase() !== 'a' ||
    !(event.target instanceof Element)
  ) {
    return null;
  }
  const formulaBar = event.target.closest('.fortune-fx-input');
  return formulaBar instanceof HTMLElement ? formulaBar : null;
}

export function selectSpreadsheetFormulaBarContents(
  event: KeyboardEvent,
): boolean {
  const formulaBar = spreadsheetFormulaBarSelectAllTarget(event);
  const selection = window.getSelection();
  if (!formulaBar || !selection) return false;
  const range = document.createRange();
  range.selectNodeContents(formulaBar);
  selection.removeAllRanges();
  selection.addRange(range);
  return true;
}

function spreadsheetColumnLabel(column: number): string {
  let value = Math.max(0, column) + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function spreadsheetCellHasContent(cell: Cell): boolean {
  if (typeof cell.f === 'string' && cell.f.trim()) return true;
  if (cell.v === null || cell.v === undefined) return false;
  return typeof cell.v !== 'string' || cell.v.trim().length > 0;
}

function finiteSpreadsheetSelectionAxis(axis: number[] | undefined): number[] {
  const first = finiteSpreadsheetIndex(axis?.[0], 0);
  const second = finiteSpreadsheetIndex(axis?.[1], first);
  return [Math.min(first, second), Math.max(first, second)];
}

function finiteSpreadsheetFocus(value: unknown, axis: number[]): number {
  const focus = finiteSpreadsheetIndex(value, axis[0] ?? 0);
  return Math.min(axis[1] ?? focus, Math.max(axis[0] ?? focus, focus));
}

function finiteSpreadsheetIndex(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(0, Math.trunc(value))
    : fallback;
}

function isSpreadsheetCoordinate(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function spreadsheetSheetsById(
  sheets: WorkSpreadsheetContent['sheets'],
): Map<string, WorkSpreadsheetContent['sheets'][number]> | null {
  const byId = new Map<string, WorkSpreadsheetContent['sheets'][number]>();
  for (const sheet of sheets) {
    if (!sheet.id || byId.has(sheet.id)) return null;
    byId.set(sheet.id, sheet);
  }
  return byId;
}

function spreadsheetSheetMetadata(
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined,
) {
  if (!sheet) return null;
  const {
    celldata: _cellData,
    calcChain: _calculationChain,
    data: _data,
    dynamicArray_compute: _dynamicArrayComputation,
    luckysheet_select_save: _selection,
    luckysheet_selection_range: _range,
    zoomRatio,
    ...content
  } = sheet;
  return {
    ...content,
    ...(zoomRatio === undefined || zoomRatio === 1 ? {} : { zoomRatio }),
  };
}

function sameSpreadsheetSheetState(
  left: WorkSpreadsheetContent['sheets'][number],
  right: WorkSpreadsheetContent['sheets'][number] | undefined,
): boolean {
  return Boolean(
    right &&
      sameSpreadsheetHistoryValue(
        spreadsheetSheetMetadata(left),
        spreadsheetSheetMetadata(right),
      ) &&
      sameSpreadsheetCellStorage(left, right),
  );
}

function sameSpreadsheetCellStorage(
  left: WorkSpreadsheetContent['sheets'][number],
  right: WorkSpreadsheetContent['sheets'][number],
): boolean {
  if (left.data !== undefined && right.data !== undefined) {
    const leftProfile = spreadsheetMatrixProfile(left.data);
    const rightProfile = spreadsheetMatrixProfile(right.data);
    if (
      leftProfile &&
      rightProfile &&
      leftProfile.historyRoot === rightProfile.historyRoot
    ) {
      return leftProfile.historyState === rightProfile.historyState;
    }
    return sameSpreadsheetDataMatrices(left.data, right.data);
  }
  if (left.data === undefined && right.data === undefined) {
    if (left.celldata === right.celldata) return true;
    return sameSpreadsheetHistoryValue(
      spreadsheetPopulatedCellData(left),
      spreadsheetPopulatedCellData(right),
    );
  }
  return sameSpreadsheetHistoryValue(
    spreadsheetPopulatedCellData(left),
    spreadsheetPopulatedCellData(right),
  );
}

function sameSpreadsheetDataMatrices(
  left: CellMatrix,
  right: CellMatrix,
): boolean {
  if (left === right) return true;
  for (const rowIndex of sparseArrayIndexes(left)) {
    if (!sameSpreadsheetDataRows(left[rowIndex], right[rowIndex])) return false;
  }
  for (const rowIndex of sparseArrayIndexes(right)) {
    if (
      !Object.hasOwn(left, rowIndex) &&
      !sameSpreadsheetDataRows(undefined, right[rowIndex])
    ) {
      return false;
    }
  }
  return true;
}

function sameSpreadsheetDataRows(
  left: CellMatrix[number] | undefined,
  right: CellMatrix[number] | undefined,
): boolean {
  if (left === right) return true;
  for (const columnIndex of sparseArrayIndexes(left ?? [])) {
    const leftCell = left?.[columnIndex];
    if (leftCell == null) continue;
    const rightCell = right?.[columnIndex];
    if (
      rightCell == null ||
      !sameSpreadsheetHistoryValue(leftCell, rightCell)
    ) {
      return false;
    }
  }
  for (const columnIndex of sparseArrayIndexes(right ?? [])) {
    if (right?.[columnIndex] != null && left?.[columnIndex] == null) {
      return false;
    }
  }
  return true;
}

function spreadsheetPopulatedCellData(
  sheet: WorkSpreadsheetContent['sheets'][number],
): NonNullable<WorkSpreadsheetContent['sheets'][number]['celldata']> {
  const cells: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['celldata']
  > = [];
  if (sheet.data !== undefined) {
    for (const rowIndex of sparseArrayIndexes(sheet.data)) {
      const row = sheet.data[rowIndex];
      if (!row) continue;
      for (const columnIndex of sparseArrayIndexes(row)) {
        const cell = row[columnIndex];
        if (cell == null) continue;
        cells.push({ r: rowIndex, c: columnIndex, v: cell });
      }
    }
  } else {
    for (const cell of sheet.celldata ?? []) {
      if (cell.v != null) cells.push(cell);
    }
  }
  return cells.sort((left, right) => left.r - right.r || left.c - right.c);
}

function positiveSpreadsheetDimension(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? Math.floor(value)
    : 0;
}
