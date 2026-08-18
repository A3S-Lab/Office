import type { Cell } from '@fortune-sheet/core';
import { sparseArrayEntries } from '../spreadsheet-sparse';
import type { WorkSpreadsheetContent } from '../work-types';
import { spreadsheetSelectionReference } from './spreadsheet-editor-support';

export interface SpreadsheetFindMatch {
  sheetId: string;
  row: number;
  column: number;
  reference: string;
  text: string;
}

interface IndexedSpreadsheetCell {
  row: number;
  column: number;
  cell: Cell;
}

export function spreadsheetFindMatches(
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined,
  query: string,
): SpreadsheetFindMatch[] {
  if (!sheet || !query) return [];
  const normalizedQuery = normalizeSpreadsheetFindText(query);
  if (!normalizedQuery) return [];

  return indexedSpreadsheetCells(sheet)
    .filter(({ row, column, cell }) => {
      const reference = spreadsheetSelectionReference({
        row: [row, row],
        column: [column, column],
      });
      const sourceFormula = sheet.formulaMetadata?.sourceFormulas?.[reference];
      return spreadsheetCellSearchValues(cell, sourceFormula).some((value) =>
        normalizeSpreadsheetFindText(value).includes(normalizedQuery),
      );
    })
    .map(({ row, column, cell }) => ({
      sheetId: sheet.id ?? '',
      row,
      column,
      reference: spreadsheetSelectionReference({
        row: [row, row],
        column: [column, column],
      }),
      text: spreadsheetCellDisplayText(cell),
    }));
}

function indexedSpreadsheetCells(
  sheet: WorkSpreadsheetContent['sheets'][number],
): IndexedSpreadsheetCell[] {
  const indexed = new Map<string, IndexedSpreadsheetCell>();
  for (const [row, cells] of sparseArrayEntries(sheet.data)) {
    for (const [column, cell] of sparseArrayEntries(cells)) {
      if (!cell) continue;
      indexed.set(`${row}:${column}`, { row, column, cell });
    }
  }
  for (const entry of sheet.celldata ?? []) {
    if (
      !Number.isSafeInteger(entry.r) ||
      !Number.isSafeInteger(entry.c) ||
      entry.r < 0 ||
      entry.c < 0 ||
      !entry.v
    ) {
      continue;
    }
    const key = `${entry.r}:${entry.c}`;
    if (!indexed.has(key)) {
      indexed.set(key, { row: entry.r, column: entry.c, cell: entry.v });
    }
  }
  return [...indexed.values()].sort(
    (left, right) => left.row - right.row || left.column - right.column,
  );
}

function spreadsheetCellSearchValues(
  cell: Cell,
  sourceFormula: string | undefined,
): string[] {
  const values = [cell.m, cell.v, cell.f, sourceFormula]
    .map(spreadsheetFindValueText)
    .filter((value): value is string => value !== null);
  return [...new Set(values)];
}

function spreadsheetCellDisplayText(cell: Cell): string {
  for (const value of [cell.m, cell.v, cell.f]) {
    const text = spreadsheetFindValueText(value);
    if (text) return text;
  }
  return '';
}

function spreadsheetFindValueText(value: unknown): string | null {
  if (typeof value === 'string') return value;
  if (typeof value === 'number' && Number.isFinite(value)) return String(value);
  if (typeof value === 'boolean' || typeof value === 'bigint') {
    return String(value);
  }
  return null;
}

function normalizeSpreadsheetFindText(value: string): string {
  return value.normalize('NFKC').toLocaleLowerCase();
}
