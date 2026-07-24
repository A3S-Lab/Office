import type { Cell, Selection } from '@fortune-sheet/core';
import type { KeyboardEvent as ReactKeyboardEvent } from 'react';
import type { WorkSpreadsheetContent } from '../work-types';

const spreadsheetFontSizes = [
  9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72,
] as const;

export function spreadsheetSelectionReference(selection: Selection): string {
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

export function spreadsheetFontSizeOptions(
  current: number | undefined,
): { value: string; label: string }[] {
  const values: number[] = [...spreadsheetFontSizes];
  if (current && !values.includes(current)) values.push(current);
  return values
    .sort((left, right) => left - right)
    .map((value) => ({ value: String(value), label: String(value) }));
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

export function spreadsheetFormulaInitializationKey(
  content: WorkSpreadsheetContent,
): string {
  return content.sheets
    .flatMap((sheet) =>
      (sheet.data ?? []).flatMap((row, rowIndex) =>
        row.flatMap((cell, columnIndex) =>
          cell?.f
            ? [`${sheet.id ?? sheet.name}:${rowIndex}:${columnIndex}:${cell.f}`]
            : [],
        ),
      ),
    )
    .join('|');
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

export function isSpreadsheetNativeTextUndoTarget(
  target: EventTarget | null,
): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (target.closest('.fortune-container')) return false;
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    target.isContentEditable ||
    Boolean(target.closest('[contenteditable="true"]'))
  );
}

export function spreadsheetFormulaBarSelectAllTarget(
  event: ReactKeyboardEvent<HTMLElement>,
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
