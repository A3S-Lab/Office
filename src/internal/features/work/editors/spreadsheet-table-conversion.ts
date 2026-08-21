import type { Cell } from '@fortune-sheet/core';
import { sparseArrayEntries, sparseArrayIndexes } from '../spreadsheet-sparse';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
} from '../work-types';
import { setSpreadsheetCellBorders } from './spreadsheet-cell-border';
import { spreadsheetCellRangeArea } from './spreadsheet-cell-range';
import { MAX_SPREADSHEET_TABLE_CELLS } from './spreadsheet-table-limits';
import {
  createSpreadsheetTableRenderResolver,
  spreadsheetTableStylePalette,
} from './spreadsheet-table-style';

const MAX_TABLE_CONVERSION_DENSE_FOOTPRINT = 1_000_000;

export function canMaterializeSpreadsheetTableAppearance(
  table: WorkSpreadsheetTable,
): boolean {
  return (
    table.style.family === 'none' ||
    spreadsheetCellRangeArea(table.range) <= MAX_SPREADSHEET_TABLE_CELLS
  );
}

/**
 * Convert to Range removes ListObject semantics, so the render-only style must
 * become native cell formatting to preserve the appearance users confirmed.
 */
export function materializeSpreadsheetTableAppearance(
  content: WorkSpreadsheetContent,
  sheetId: string,
  table: WorkSpreadsheetTable,
): WorkSpreadsheetContent | null {
  if (!canMaterializeSpreadsheetTableAppearance(table)) return null;
  if (table.style.family === 'none') return content;
  const sheetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = content.sheets[sheetIndex];
  const palette = spreadsheetTableStylePalette(table.style);
  if (!sheet || !palette) return null;
  const resolve = createSpreadsheetTableRenderResolver([table]);
  const styledSheet = shouldMaterializeSpreadsheetTableAsSparseCells(
    sheet,
    table,
  )
    ? materializeSpreadsheetTableSparseCells(sheet, table, resolve)
    : materializeSpreadsheetTableMatrix(sheet, table, resolve);
  const sheets = [...content.sheets];
  sheets[sheetIndex] = styledSheet;
  return setSpreadsheetCellBorders(
    { ...content, sheets },
    sheetId,
    table.range,
    { target: 'all', color: palette.border, style: 'thin' },
  );
}

function shouldMaterializeSpreadsheetTableAsSparseCells(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
): boolean {
  if (sheet.data === undefined) return true;
  const [startRow, endRow] = table.range.row;
  const requiredRowLength = table.range.column[1] + 1;
  let footprint = 0;
  let presentTableRows = 0;
  for (const [rowIndex, row] of sparseArrayEntries(sheet.data)) {
    if (rowIndex >= startRow && rowIndex <= endRow) {
      presentTableRows += 1;
      footprint += Math.max(row.length, requiredRowLength);
    } else {
      footprint += row.length;
    }
    if (footprint > MAX_TABLE_CONVERSION_DENSE_FOOTPRINT) return true;
  }
  footprint += (endRow - startRow + 1 - presentTableRows) * requiredRowLength;
  return footprint > MAX_TABLE_CONVERSION_DENSE_FOOTPRINT;
}

function materializeSpreadsheetTableMatrix(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  resolve: ReturnType<typeof createSpreadsheetTableRenderResolver>,
): WorkSpreadsheetSheet {
  const source = sheet.data ?? [];
  const data: NonNullable<WorkSpreadsheetSheet['data']> = [];
  data.length = Math.max(source.length, table.range.row[1] + 1);
  for (const [row, value] of sparseArrayEntries(source)) data[row] = value;
  for (let row = table.range.row[0]; row <= table.range.row[1]; row += 1) {
    const sourceRow = source[row];
    const values: NonNullable<(typeof data)[number]> = [];
    values.length = Math.max(sourceRow?.length ?? 0, table.range.column[1] + 1);
    for (const column of sparseArrayIndexes(sourceRow)) {
      values[column] = sourceRow?.[column];
    }
    for (
      let column = table.range.column[0];
      column <= table.range.column[1];
      column += 1
    ) {
      const style = resolve(row, column);
      if (style) {
        values[column] = spreadsheetCellWithTableAppearance(
          values[column],
          style,
        );
      }
    }
    data[row] = values;
  }
  return { ...sheet, data };
}

function materializeSpreadsheetTableSparseCells(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  resolve: ReturnType<typeof createSpreadsheetTableRenderResolver>,
): WorkSpreadsheetSheet {
  const celldata: NonNullable<WorkSpreadsheetSheet['celldata']> = [];
  if (sheet.data !== undefined) {
    for (const [row, values] of sparseArrayEntries(sheet.data)) {
      for (const column of sparseArrayIndexes(values)) {
        const cell = values[column];
        if (cell != null) celldata.push({ c: column, r: row, v: cell });
      }
    }
  } else {
    celldata.push(...(sheet.celldata ?? []));
  }
  const indexes = new Map(
    celldata.map((entry, index) => [`${entry.r}:${entry.c}`, index]),
  );
  for (let row = table.range.row[0]; row <= table.range.row[1]; row += 1) {
    for (
      let column = table.range.column[0];
      column <= table.range.column[1];
      column += 1
    ) {
      const style = resolve(row, column);
      if (!style) continue;
      const key = `${row}:${column}`;
      const index = indexes.get(key);
      const cell = spreadsheetCellWithTableAppearance(
        index === undefined ? undefined : celldata[index]?.v,
        style,
      );
      if (index === undefined) {
        indexes.set(key, celldata.length);
        celldata.push({ c: column, r: row, v: cell });
      } else {
        celldata[index] = { c: column, r: row, v: cell };
      }
    }
  }
  celldata.sort((left, right) => left.r - right.r || left.c - right.c);
  const { data: _data, ...metadata } = sheet;
  return { ...metadata, celldata };
}

function spreadsheetCellWithTableAppearance(
  cell: Cell | null | undefined,
  style: NonNullable<
    ReturnType<ReturnType<typeof createSpreadsheetTableRenderResolver>>
  >,
): Cell {
  return {
    ...(cell ?? {}),
    bg: style.background,
    fc: style.textColor,
    ...(style.bold ? { bl: 1 } : {}),
  };
}
