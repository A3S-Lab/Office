import type { Cell } from '@fortune-sheet/core';
import {
  editableSpreadsheetFormula,
  formulaHasExternalReference,
  formulaHasStructuredReference,
  spreadsheetFormulaFunctions,
} from '../work-spreadsheet-formulas';
import { sparseArrayEntries } from '../spreadsheet-sparse';
import type {
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
  WorkSpreadsheetTableColumn,
} from '../work-types';

/** Keep calculated-column metadata small enough to cross every editor boundary. */
export const MAX_SPREADSHEET_TABLE_CALCULATED_FORMULA_LENGTH = 8_192;

/**
 * Functions that can reach outside the deterministic browser calculation
 * boundary. A calculated-column rule containing one of these functions is
 * retained as an ordinary cell formula, but is never propagated implicitly.
 */
const UNSAFE_SPREADSHEET_TABLE_CALCULATED_FUNCTIONS = new Set([
  'CALL',
  'DDE',
  'DDE.REQUEST',
  'DDE.POKE',
  'EXEC',
  'HYPERLINK',
  'INDIRECT',
  'OFFSET',
  'REGISTER.ID',
  'RTD',
  'WEBSERVICE',
]);

/**
 * Normalizes the editable form of a table-column formula. Native table parts
 * omit the leading `=`, while cell formulas and the calculation kernel keep
 * it. Unsafe or external formulas are deliberately not treated as automatic
 * fill rules.
 */
export function normalizeSpreadsheetTableCalculatedFormula(
  value: unknown,
): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  const formula = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
  if (
    formula.startsWith('==') ||
    formula.length > MAX_SPREADSHEET_TABLE_CALCULATED_FORMULA_LENGTH ||
    /[\u0000-\u001f\u007f]/u.test(formula) ||
    formulaHasExternalReference(formula) ||
    !formulaHasStructuredReference(formula) ||
    spreadsheetFormulaFunctions(formula).some((name) =>
      UNSAFE_SPREADSHEET_TABLE_CALCULATED_FUNCTIONS.has(name),
    )
  ) {
    return undefined;
  }
  const editable = editableSpreadsheetFormula(formula).trim();
  return editable || undefined;
}

/**
 * Finds a shared current-row formula already authored in a table column.
 * Ordinary one-off A1 formulas are intentionally ignored so inserting a row
 * cannot unexpectedly overwrite a user's manual exception.
 */
export function inferredSpreadsheetTableCalculatedFormula(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  columnOffset: number,
): string | undefined {
  const startRow = table.range.row[0] + Number(table.headerRow);
  const endRow = table.range.row[1] - Number(table.totalsRow);
  const column = table.range.column[0] + columnOffset;
  const formulas: string[] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    const formula = normalizeSpreadsheetTableCalculatedFormula(
      spreadsheetTableCellAt(sheet, row, column)?.f,
    );
    if (!formula || !hasCurrentRowReference(formula)) continue;
    formulas.push(formula);
  }
  const first = formulas[0];
  return first && formulas.every((formula) => formula === first)
    ? first
    : undefined;
}

/**
 * Reconciles declared calculated-column formulas with formulas authored in the
 * controlled cell matrix. A conflicting column is made manual instead of
 * guessing which formula should be copied into a new row.
 */
export function reconcileSpreadsheetTableCalculatedColumns(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
): WorkSpreadsheetTableColumn[] {
  return table.columns.map((column, offset) => {
    const declared = normalizeSpreadsheetTableCalculatedFormula(
      column.calculatedFormula,
    );
    const observed = spreadsheetTableCurrentRowFormulas(sheet, table, offset);
    const uniqueObserved = new Set(observed);
    if (uniqueObserved.size > 1) return stripCalculatedFormula(column);
    const inferred = uniqueObserved.values().next().value as string | undefined;
    if (declared && inferred && declared !== inferred) {
      // Two different current-row formulas mean that this is no longer a
      // single calculated column. Failing closed keeps a manual exception
      // from being silently overwritten on a later row insertion.
      return stripCalculatedFormula(column);
    }
    if (declared) return { ...column, calculatedFormula: declared };
    return inferred
      ? { ...column, calculatedFormula: inferred }
      : stripCalculatedFormula(column);
  });
}

/** Returns true only for a formula safe to copy into a new table row. */
export function isSpreadsheetTableCalculatedFormula(
  value: unknown,
): value is string {
  const formula = normalizeSpreadsheetTableCalculatedFormula(value);
  return Boolean(formula && hasCurrentRowReference(formula));
}

/**
 * Applies calculated-column formulas to newly inserted body rows. The helper
 * deliberately edits only empty cells; copied values and authored exceptions
 * remain authoritative. Both Fortune's dense matrix and sparse `celldata`
 * storage are updated without materializing unrelated rows or columns.
 */
export function fillSpreadsheetTableCalculatedColumns(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  rows: readonly number[],
): WorkSpreadsheetSheet {
  const targetRows = Array.from(
    new Set(
      rows.filter(
        (row) =>
          Number.isSafeInteger(row) &&
          row >= table.range.row[0] + Number(table.headerRow) &&
          row <= table.range.row[1] - Number(table.totalsRow),
      ),
    ),
  ).sort((left, right) => left - right);
  if (!targetRows.length) return sheet;

  const formulas = table.columns.map((column) =>
    isSpreadsheetTableCalculatedFormula(column.calculatedFormula)
      ? normalizeSpreadsheetTableCalculatedFormula(column.calculatedFormula)
      : undefined,
  );
  if (!formulas.some(Boolean)) return sheet;

  if (sheet.data !== undefined) {
    const data = sheet.data.slice();
    const mutableRows = new Map<number, NonNullable<typeof data>[number]>();
    let changed = false;
    for (const rowIndex of targetRows) {
      const sourceRow = data[rowIndex];
      let row = mutableRows.get(rowIndex);
      for (let offset = 0; offset < formulas.length; offset += 1) {
        const formula = formulas[offset];
        if (!formula) continue;
        const column = table.range.column[0] + offset;
        const current = sourceRow?.[column] ?? null;
        if (!spreadsheetTableCellIsEmpty(current)) continue;
        row ??= cloneSpreadsheetTableRow(sourceRow);
        mutableRows.set(rowIndex, row);
        const styleSource =
          current ??
          nearestSpreadsheetTableCell(sheet, table, rowIndex, column);
        row[column] = spreadsheetCellWithCalculatedFormula(
          styleSource,
          formula,
        );
        changed = true;
      }
      if (row) data[rowIndex] = row;
    }
    return changed ? { ...sheet, data } : sheet;
  }

  const entries = [...(sheet.celldata ?? [])];
  const byCoordinate = new Map(
    entries.map((entry, index) => [`${entry.r}:${entry.c}`, index]),
  );
  let changed = false;
  for (const row of targetRows) {
    for (let offset = 0; offset < formulas.length; offset += 1) {
      const formula = formulas[offset];
      if (!formula) continue;
      const column = table.range.column[0] + offset;
      const key = `${row}:${column}`;
      const index = byCoordinate.get(key);
      const current = index === undefined ? null : entries[index]?.v;
      if (!spreadsheetTableCellIsEmpty(current)) continue;
      const styleSource =
        current ?? nearestSpreadsheetTableCell(sheet, table, row, column);
      const entry = {
        r: row,
        c: column,
        v: spreadsheetCellWithCalculatedFormula(styleSource, formula),
      };
      if (index === undefined) {
        byCoordinate.set(key, entries.length);
        entries.push(entry);
      } else {
        entries[index] = entry;
      }
      changed = true;
    }
  }
  if (!changed) return sheet;
  entries.sort((left, right) => left.r - right.r || left.c - right.c);
  return { ...sheet, celldata: entries };
}

export function spreadsheetTableCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  return (
    sheet.data?.[row]?.[column] ??
    sheet.celldata?.find((entry) => entry.r === row && entry.c === column)?.v ??
    null
  );
}

export function spreadsheetTableCellIsEmpty(
  cell: Cell | null | undefined,
): boolean {
  return (
    !cell?.f &&
    (cell?.v === undefined || cell.v === null || cell.v === '') &&
    (cell?.m === undefined || cell.m === null || cell.m === '')
  );
}

/** Copy visual cell metadata without carrying a stale cached formula result. */
export function spreadsheetCellWithCalculatedFormula(
  source: Cell | null | undefined,
  formula: string,
): Cell {
  const { f: _formula, m: _display, v: _value, ...presentation } = source ?? {};
  return { ...presentation, f: formula };
}

function stripCalculatedFormula(
  column: WorkSpreadsheetTableColumn,
): WorkSpreadsheetTableColumn {
  if (column.calculatedFormula === undefined) return column;
  const { calculatedFormula: _formula, ...withoutFormula } = column;
  return withoutFormula;
}

function hasCurrentRowReference(formula: string): boolean {
  const source = formula.replace(/"(?:[^"]|"")*"/g, '""');
  return /\[@/i.test(source) || /\[#This Row\]/i.test(source);
}

/** Iterate formula-bearing cells without forcing a dense worksheet matrix. */
export function spreadsheetTableFormulaCells(
  sheet: WorkSpreadsheetSheet,
): Array<{ cell: Cell; column: number; row: number }> {
  const cells: Array<{ cell: Cell; column: number; row: number }> = [];
  for (const [row, values] of sparseArrayEntries(sheet.data)) {
    for (const [column, cell] of sparseArrayEntries(values)) {
      if (cell?.f) cells.push({ cell, column, row });
    }
  }
  for (const entry of sheet.celldata ?? []) {
    if (entry.v?.f)
      cells.push({ cell: entry.v, column: entry.c, row: entry.r });
  }
  return cells;
}

function spreadsheetTableCurrentRowFormulas(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  columnOffset: number,
): string[] {
  const startRow = table.range.row[0] + Number(table.headerRow);
  const endRow = table.range.row[1] - Number(table.totalsRow);
  const column = table.range.column[0] + columnOffset;
  const formulas: string[] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    const formula = normalizeSpreadsheetTableCalculatedFormula(
      spreadsheetTableCellAt(sheet, row, column)?.f,
    );
    if (formula && hasCurrentRowReference(formula)) formulas.push(formula);
  }
  return formulas;
}

function cloneSpreadsheetTableRow(
  source: NonNullable<WorkSpreadsheetSheet['data']>[number] | undefined,
): NonNullable<WorkSpreadsheetSheet['data']>[number] {
  const row: NonNullable<WorkSpreadsheetSheet['data']>[number] = [];
  if (!source) return row;
  row.length = source.length;
  for (const column of sparseArrayEntries(source).map(([index]) => index)) {
    row[column] = source[column];
  }
  return row;
}

function nearestSpreadsheetTableCell(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  row: number,
  column: number,
): Cell | null {
  const start = table.range.row[0] + Number(table.headerRow);
  const end = table.range.row[1] - Number(table.totalsRow);
  for (let distance = 1; distance <= end - start; distance += 1) {
    const before = row - distance;
    if (before >= start) {
      const cell = spreadsheetTableCellAt(sheet, before, column);
      if (cell) return cell;
    }
    const after = row + distance;
    if (after <= end) {
      const cell = spreadsheetTableCellAt(sheet, after, column);
      if (cell) return cell;
    }
  }
  return null;
}
