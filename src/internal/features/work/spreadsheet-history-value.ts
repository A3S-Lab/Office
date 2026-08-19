import type { Cell } from '@fortune-sheet/core';

const spreadsheetFormulaHistoryIgnoredKeys = new Set(['ct', 'm', 'v']);

export function sameSpreadsheetHistoryValue(
  left: unknown,
  right: unknown,
): boolean {
  if (left === right) return true;
  if (left === null || right === null || typeof left !== typeof right) {
    return false;
  }
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
    const leftKeys = Object.keys(left).filter(
      (key) => left[Number(key)] !== undefined,
    );
    const rightKeys = Object.keys(right).filter(
      (key) => right[Number(key)] !== undefined,
    );
    return (
      leftKeys.length === rightKeys.length &&
      leftKeys.every(
        (key) =>
          Object.hasOwn(right, key) &&
          sameSpreadsheetHistoryValue(left[Number(key)], right[Number(key)]),
      )
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
