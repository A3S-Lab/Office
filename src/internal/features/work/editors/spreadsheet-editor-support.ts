import type { Cell, Selection } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import { officeFontFamilies } from './office-font-families';
import type { OfficeSelectOption } from './office-select';

const spreadsheetFontSizes = [
  9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72,
] as const;
const spreadsheetFormulaHistoryIgnoredKeys = new Set(['ct', 'm', 'v']);

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
    for (let rowIndex = rowStart; rowIndex <= rowEnd; rowIndex += 1) {
      const row = sheet.data[rowIndex];
      if (!row) continue;
      const populatedColumnEnd = Math.min(columnEnd, row.length - 1);
      for (
        let columnIndex = columnStart;
        columnIndex <= populatedColumnEnd;
        columnIndex += 1
      ) {
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
  return structuredClone(sheets).map((sheet) => {
    for (const merge of Object.values(sheet.config?.merge ?? {})) {
      for (
        let rowIndex = merge.r;
        rowIndex < merge.r + merge.rs;
        rowIndex += 1
      ) {
        const row = sheet.data?.[rowIndex];
        if (!row) continue;
        for (
          let columnIndex = merge.c;
          columnIndex < merge.c + merge.cs;
          columnIndex += 1
        ) {
          row[columnIndex] = {
            ...(row[columnIndex] ?? {}),
            mc:
              rowIndex === merge.r && columnIndex === merge.c
                ? { ...merge }
                : { r: merge.r, c: merge.c },
          };
        }
      }
    }
    return {
      ...sheet,
      celldata: sheet.data
        ? sheet.data.flatMap((row, rowIndex) =>
            row.flatMap((cell, columnIndex) =>
              cell == null ? [] : [{ r: rowIndex, c: columnIndex, v: cell }],
            ),
          )
        : (sheet.celldata ?? []),
    };
  });
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
    JSON.stringify(changed.map(spreadsheetSheetWithoutTransientSelection)) ===
    JSON.stringify(rendered.map(spreadsheetSheetWithoutTransientSelection))
  );
}

export function sameSpreadsheetHistoryContent(
  left: WorkSpreadsheetContent,
  right: WorkSpreadsheetContent,
): boolean {
  return sameSpreadsheetHistoryValue(
    {
      ...left,
      sheets: left.sheets.map(spreadsheetSheetWithoutTransientSelection),
    },
    {
      ...right,
      sheets: right.sheets.map(spreadsheetSheetWithoutTransientSelection),
    },
  );
}

export function spreadsheetContentWithSelection(
  content: WorkSpreadsheetContent,
  sheetId: string,
  selection: Selection | null | undefined,
): WorkSpreadsheetContent {
  if (!sheetId || !selection) return content;
  const nextSelection = finiteSpreadsheetSelection(selection);
  let changed = false;
  const sheets = content.sheets.map((sheet) => {
    if (sheet.id !== sheetId) return sheet;
    changed = true;
    return { ...sheet, luckysheet_select_save: [nextSelection] };
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

function spreadsheetSheetWithoutTransientSelection(
  sheet: WorkSpreadsheetContent['sheets'][number],
) {
  const {
    celldata: _cellData,
    luckysheet_select_save: _selection,
    luckysheet_selection_range: _range,
    ...content
  } = sheet;
  return content;
}

function sameSpreadsheetHistoryValue(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (left === null || right === null || typeof left !== typeof right)
    return false;
  const leftFormulaCell = spreadsheetFormulaCell(left);
  const rightFormulaCell = spreadsheetFormulaCell(right);
  if (leftFormulaCell || rightFormulaCell) {
    return Boolean(
      leftFormulaCell &&
        rightFormulaCell &&
        sameSpreadsheetFormulaHistoryCell(leftFormulaCell, rightFormulaCell),
    );
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (
      !Array.isArray(left) ||
      !Array.isArray(right) ||
      left.length !== right.length
    ) {
      return false;
    }
    return left.every((value, index) =>
      sameSpreadsheetHistoryValue(value, right[index]),
    );
  }
  if (typeof left !== 'object' || typeof right !== 'object') return false;
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).filter(
    (key) => leftRecord[key] !== undefined,
  );
  const rightKeys = Object.keys(rightRecord).filter(
    (key) => rightRecord[key] !== undefined,
  );
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every((key) =>
      sameSpreadsheetHistoryValue(leftRecord[key], rightRecord[key]),
    )
  );
}

function spreadsheetFormulaCell(value: unknown): Cell | null {
  return value &&
    typeof value === 'object' &&
    typeof (value as Cell).f === 'string'
    ? (value as Cell)
    : null;
}

function sameSpreadsheetFormulaHistoryCell(left: Cell, right: Cell): boolean {
  const leftRecord = left as unknown as Record<string, unknown>;
  const rightRecord = right as unknown as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).filter(
    (key) =>
      !spreadsheetFormulaHistoryIgnoredKeys.has(key) &&
      leftRecord[key] !== undefined,
  );
  const rightKeys = Object.keys(rightRecord).filter(
    (key) =>
      !spreadsheetFormulaHistoryIgnoredKeys.has(key) &&
      rightRecord[key] !== undefined,
  );
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every((key) =>
      sameSpreadsheetHistoryValue(leftRecord[key], rightRecord[key]),
    ) &&
    sameSpreadsheetHistoryValue(
      spreadsheetFormulaHistoryCellType(left.ct),
      spreadsheetFormulaHistoryCellType(right.ct),
    )
  );
}

function spreadsheetFormulaHistoryCellType(
  cellType: Cell['ct'],
): Omit<NonNullable<Cell['ct']>, 't'> | undefined {
  if (!cellType) return undefined;
  const { fa, t: _type, ...retainedCellType } = cellType;
  const normalized = {
    ...retainedCellType,
    ...(fa && fa !== 'General' ? { fa } : {}),
  };
  return Object.keys(normalized).length ? normalized : undefined;
}
