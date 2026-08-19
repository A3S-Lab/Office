import type { CellMatrix, Sheet } from '@fortune-sheet/core';

export interface SpreadsheetGridSize {
  rowCount: number;
  columnCount: number;
}

export function sparseArrayIndexes(
  values: readonly unknown[] | undefined,
): number[] {
  if (!values) return [];
  return Object.keys(values).flatMap((key) => {
    const index = Number(key);
    return Number.isSafeInteger(index) && index >= 0 && index < values.length
      ? [index]
      : [];
  });
}

export function sparseArrayEntries<T>(
  values: readonly T[] | undefined,
): Array<[number, T]> {
  if (!values) return [];
  return sparseArrayIndexes(values).flatMap((index) => {
    const value = values[index];
    return value === undefined ? [] : [[index, value]];
  });
}

export function sparseMatrixColumnCount(
  matrix: CellMatrix | undefined,
): number {
  let maximum = 0;
  for (const [, row] of sparseArrayEntries(matrix)) {
    maximum = Math.max(maximum, row.length);
  }
  return maximum;
}

export function spreadsheetGridSize(
  sheet: Pick<Sheet, 'column' | 'data' | 'row'> | undefined,
): SpreadsheetGridSize | null {
  if (!sheet) return null;
  return {
    rowCount: Math.max(sheet.row ?? 0, sheet.data?.length ?? 0),
    columnCount: Math.max(
      sheet.column ?? 0,
      sparseMatrixColumnCount(sheet.data),
    ),
  };
}

export function cloneSparseMatrix(source: CellMatrix | undefined): CellMatrix {
  const clone: CellMatrix = [];
  if (!source) return clone;
  clone.length = source.length;
  for (const [rowIndex, sourceRow] of sparseArrayEntries(source)) {
    const row = [] as CellMatrix[number];
    row.length = sourceRow.length;
    for (const columnIndex of sparseArrayIndexes(sourceRow)) {
      row[columnIndex] = sourceRow[columnIndex];
    }
    clone[rowIndex] = row;
  }
  return clone;
}
