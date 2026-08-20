import type { Cell } from '@fortune-sheet/core';
import { spreadsheetCellAddress } from '../work-spreadsheet-formulas';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
} from './spreadsheet-cell-range';

const spreadsheetMaximumRows = 1_048_576;
const spreadsheetMaximumColumns = 16_384;

export const spreadsheetAutoSumMaximumTargets = 1_000;

export type SpreadsheetAutoSumFunction =
  | 'sum'
  | 'average'
  | 'count'
  | 'max'
  | 'min';

export interface SpreadsheetAutoSumWrite {
  range: SpreadsheetCellRange;
  values: string[][];
}

export interface SpreadsheetAutoSumPlan {
  formulaBarValue: string;
  function: SpreadsheetAutoSumFunction;
  selection: SpreadsheetCellRange & {
    row_focus: number;
    column_focus: number;
  };
  targetCellCount: number;
  writes: SpreadsheetAutoSumWrite[];
}

interface SpreadsheetAutoSumTarget {
  column: number;
  formula: string;
  row: number;
}

interface SpreadsheetAutoSumCandidate {
  blocked: boolean;
  orientation: 'column' | 'row';
  targets: SpreadsheetAutoSumTarget[];
}

interface SpreadsheetAutoSumSourceRange {
  endColumn: number;
  endRow: number;
  startColumn: number;
  startRow: number;
}

export function planSpreadsheetAutoSum(
  sheet: WorkSpreadsheetSheet | undefined,
  selection: SpreadsheetCellRangeInput,
  functionName: SpreadsheetAutoSumFunction,
): SpreadsheetAutoSumPlan | null {
  const range = normalizeSpreadsheetCellRange(selection);
  const nativeFunction = spreadsheetAutoSumNativeFunction(functionName);
  if (
    !sheet ||
    !range ||
    !nativeFunction ||
    !spreadsheetRangeIsSupported(range)
  ) {
    return null;
  }

  const cellAt = createSpreadsheetCellReader(sheet);
  let targets: SpreadsheetAutoSumTarget[];
  let writes: SpreadsheetAutoSumWrite[];
  if (spreadsheetRangeIsSingleCell(range)) {
    const target = spreadsheetSingleAutoSumTarget(
      cellAt,
      range.row[0],
      range.column[0],
      nativeFunction,
    );
    if (!target) return null;
    targets = [target];
    writes = compactSpreadsheetAutoSumTargets(targets, 'row');
  } else {
    const rowCandidate = spreadsheetTotalsRowCandidate(
      cellAt,
      range,
      nativeFunction,
    );
    const columnCandidate = spreadsheetTotalsColumnCandidate(
      cellAt,
      range,
      nativeFunction,
    );
    const targetCoordinates = new Set<string>();
    const candidates = [rowCandidate, columnCandidate]
      .filter((candidate) => !candidate.blocked && candidate.targets.length > 0)
      .map((candidate) => ({
        ...candidate,
        targets: candidate.targets.filter((target) => {
          const coordinate = `${target.row}:${target.column}`;
          if (targetCoordinates.has(coordinate)) return false;
          targetCoordinates.add(coordinate);
          return true;
        }),
      }))
      .filter((candidate) => candidate.targets.length > 0);
    if (candidates.length === 0) return null;
    targets = candidates.flatMap((candidate) => candidate.targets);
    writes = candidates.flatMap((candidate) =>
      compactSpreadsheetAutoSumTargets(
        candidate.targets,
        candidate.orientation,
      ),
    );
  }

  if (
    targets.length === 0 ||
    targets.length > spreadsheetAutoSumMaximumTargets
  ) {
    return null;
  }

  const orderedTargets = [...targets].sort(
    (left, right) => left.row - right.row || left.column - right.column,
  );
  const firstTarget = orderedTargets[0];
  const firstRow = Math.min(...orderedTargets.map((target) => target.row));
  const lastRow = Math.max(...orderedTargets.map((target) => target.row));
  const firstColumn = Math.min(
    ...orderedTargets.map((target) => target.column),
  );
  const lastColumn = Math.max(...orderedTargets.map((target) => target.column));
  return {
    formulaBarValue: firstTarget.formula,
    function: functionName,
    selection: {
      row: [firstRow, lastRow],
      column: [firstColumn, lastColumn],
      row_focus: firstTarget.row,
      column_focus: firstTarget.column,
    },
    targetCellCount: targets.length,
    writes,
  };
}

function spreadsheetSingleAutoSumTarget(
  cellAt: (row: number, column: number) => Cell | null | undefined,
  row: number,
  column: number,
  nativeFunction: string,
): SpreadsheetAutoSumTarget | null {
  if (spreadsheetCellHasContent(cellAt(row, column))) return null;
  const source =
    spreadsheetContiguousSourceAbove(cellAt, row, column, 0) ??
    spreadsheetContiguousSourceLeft(cellAt, row, column, 0);
  return source
    ? spreadsheetAutoSumTarget(row, column, source, nativeFunction)
    : null;
}

function spreadsheetTotalsRowCandidate(
  cellAt: (row: number, column: number) => Cell | null | undefined,
  range: SpreadsheetCellRange,
  nativeFunction: string,
): SpreadsheetAutoSumCandidate {
  const targets: SpreadsheetAutoSumTarget[] = [];
  const targetRow = range.row[1];
  if (targetRow <= range.row[0]) {
    return { blocked: false, orientation: 'row', targets };
  }
  for (let column = range.column[0]; column <= range.column[1]; column += 1) {
    const source = spreadsheetContiguousSourceAbove(
      cellAt,
      targetRow,
      column,
      range.row[0],
    );
    if (!source) continue;
    if (spreadsheetCellHasContent(cellAt(targetRow, column))) {
      return { blocked: true, orientation: 'row', targets: [] };
    }
    targets.push(
      spreadsheetAutoSumTarget(targetRow, column, source, nativeFunction),
    );
    if (targets.length > spreadsheetAutoSumMaximumTargets) {
      return { blocked: true, orientation: 'row', targets: [] };
    }
  }
  return { blocked: false, orientation: 'row', targets };
}

function spreadsheetTotalsColumnCandidate(
  cellAt: (row: number, column: number) => Cell | null | undefined,
  range: SpreadsheetCellRange,
  nativeFunction: string,
): SpreadsheetAutoSumCandidate {
  const targets: SpreadsheetAutoSumTarget[] = [];
  const targetColumn = range.column[1];
  if (targetColumn <= range.column[0]) {
    return { blocked: false, orientation: 'column', targets };
  }
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    const source = spreadsheetContiguousSourceLeft(
      cellAt,
      row,
      targetColumn,
      range.column[0],
    );
    if (!source) continue;
    if (spreadsheetCellHasContent(cellAt(row, targetColumn))) {
      return { blocked: true, orientation: 'column', targets: [] };
    }
    targets.push(
      spreadsheetAutoSumTarget(row, targetColumn, source, nativeFunction),
    );
    if (targets.length > spreadsheetAutoSumMaximumTargets) {
      return { blocked: true, orientation: 'column', targets: [] };
    }
  }
  return { blocked: false, orientation: 'column', targets };
}

function spreadsheetContiguousSourceAbove(
  cellAt: (row: number, column: number) => Cell | null | undefined,
  targetRow: number,
  column: number,
  minimumRow: number,
): SpreadsheetAutoSumSourceRange | null {
  const endRow = targetRow - 1;
  if (
    endRow < minimumRow ||
    !spreadsheetCellCanParticipate(cellAt(endRow, column))
  ) {
    return null;
  }
  let startRow = endRow;
  while (
    startRow > minimumRow &&
    spreadsheetCellCanParticipate(cellAt(startRow - 1, column))
  ) {
    startRow -= 1;
  }
  return {
    startRow,
    endRow,
    startColumn: column,
    endColumn: column,
  };
}

function spreadsheetContiguousSourceLeft(
  cellAt: (row: number, column: number) => Cell | null | undefined,
  row: number,
  targetColumn: number,
  minimumColumn: number,
): SpreadsheetAutoSumSourceRange | null {
  const endColumn = targetColumn - 1;
  if (
    endColumn < minimumColumn ||
    !spreadsheetCellCanParticipate(cellAt(row, endColumn))
  ) {
    return null;
  }
  let startColumn = endColumn;
  while (
    startColumn > minimumColumn &&
    spreadsheetCellCanParticipate(cellAt(row, startColumn - 1))
  ) {
    startColumn -= 1;
  }
  return {
    startRow: row,
    endRow: row,
    startColumn,
    endColumn,
  };
}

function spreadsheetAutoSumTarget(
  row: number,
  column: number,
  source: SpreadsheetAutoSumSourceRange,
  nativeFunction: string,
): SpreadsheetAutoSumTarget {
  const start = spreadsheetCellAddress(source.startRow, source.startColumn);
  const end = spreadsheetCellAddress(source.endRow, source.endColumn);
  return {
    row,
    column,
    formula: `=${nativeFunction}(${start}:${end})`,
  };
}

function compactSpreadsheetAutoSumTargets(
  targets: readonly SpreadsheetAutoSumTarget[],
  orientation: 'column' | 'row',
): SpreadsheetAutoSumWrite[] {
  const ordered = [...targets].sort(
    orientation === 'row'
      ? (left, right) => left.row - right.row || left.column - right.column
      : (left, right) => left.column - right.column || left.row - right.row,
  );
  const writes: SpreadsheetAutoSumWrite[] = [];
  for (const target of ordered) {
    const previous = writes.at(-1);
    if (
      previous &&
      orientation === 'row' &&
      previous.range.row[0] === target.row &&
      previous.range.column[1] + 1 === target.column
    ) {
      previous.range.column[1] = target.column;
      previous.values[0]?.push(target.formula);
      continue;
    }
    if (
      previous &&
      orientation === 'column' &&
      previous.range.column[0] === target.column &&
      previous.range.row[1] + 1 === target.row
    ) {
      previous.range.row[1] = target.row;
      previous.values.push([target.formula]);
      continue;
    }
    writes.push({
      range: {
        row: [target.row, target.row],
        column: [target.column, target.column],
      },
      values: [[target.formula]],
    });
  }
  return writes;
}

function createSpreadsheetCellReader(
  sheet: WorkSpreadsheetSheet,
): (row: number, column: number) => Cell | null | undefined {
  let sparseRows: Map<number, Map<number, Cell | null>> | null = null;
  const sparseCell = (row: number, column: number) => {
    if (!sparseRows) {
      sparseRows = new Map();
      for (const entry of sheet.celldata ?? []) {
        let sparseRow = sparseRows.get(entry.r);
        if (!sparseRow) {
          sparseRow = new Map<number, Cell | null>();
          sparseRows.set(entry.r, sparseRow);
        }
        sparseRow.set(entry.c, entry.v);
      }
    }
    return sparseRows.get(row)?.get(column);
  };
  return (row, column) =>
    sheet.data?.[row]?.[column] ?? sparseCell(row, column);
}

function spreadsheetCellCanParticipate(cell: Cell | null | undefined): boolean {
  return Boolean(
    (typeof cell?.f === 'string' && cell.f.trim()) ||
      (typeof cell?.v === 'number' && Number.isFinite(cell.v)),
  );
}

function spreadsheetCellHasContent(cell: Cell | null | undefined): boolean {
  if (!cell) return false;
  if (typeof cell.f === 'string' && cell.f.trim()) return true;
  if (cell.v === null || cell.v === undefined) return false;
  return typeof cell.v !== 'string' || cell.v.trim().length > 0;
}

function spreadsheetRangeIsSingleCell(range: SpreadsheetCellRange): boolean {
  return range.row[0] === range.row[1] && range.column[0] === range.column[1];
}

function spreadsheetRangeIsSupported(range: SpreadsheetCellRange): boolean {
  return (
    range.row[0] >= 0 &&
    range.row[1] < spreadsheetMaximumRows &&
    range.column[0] >= 0 &&
    range.column[1] < spreadsheetMaximumColumns &&
    [...range.row, ...range.column].every(Number.isSafeInteger)
  );
}

function spreadsheetAutoSumNativeFunction(
  functionName: SpreadsheetAutoSumFunction,
): string | null {
  switch (functionName) {
    case 'sum':
      return 'SUM';
    case 'average':
      return 'AVERAGE';
    case 'count':
      return 'COUNT';
    case 'max':
      return 'MAX';
    case 'min':
      return 'MIN';
    default:
      return null;
  }
}
