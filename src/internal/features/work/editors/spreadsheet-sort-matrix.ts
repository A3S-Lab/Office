import type { Cell } from '@fortune-sheet/core';
import { translateSpreadsheetFormula } from './spreadsheet-paste-special-cell';
import {
  createSpreadsheetSortDirectAppearanceRows,
  spreadsheetSortCellMatchesAppearance,
  type SpreadsheetSortCellAppearance,
  type SpreadsheetSortAppearancePosition,
  type SpreadsheetSortAppearanceRows,
  type SpreadsheetSortAppearanceTarget,
} from './spreadsheet-sort-appearance';
import { spreadsheetSortCustomListMatchKey } from './spreadsheet-sort-custom-list';
import {
  spreadsheetSortAppearanceRowsMatch,
  spreadsheetSortError,
  type SpreadsheetSortDirection,
  type SpreadsheetSortKey,
  type SpreadsheetSortRequest,
  type SpreadsheetSortResult,
  validateSpreadsheetSortRequest,
} from './spreadsheet-sort';

interface SpreadsheetSortItem {
  appearance: readonly (SpreadsheetSortCellAppearance | undefined)[];
  cells: readonly (Cell | null)[];
  index: number;
}

type SpreadsheetSortCompiledKey =
  | {
      kind: 'appearance';
      offset: number;
      position: SpreadsheetSortAppearancePosition;
      target: SpreadsheetSortAppearanceTarget;
    }
  | {
      customRanks: ReadonlyMap<string, number>;
      kind: 'custom-list';
      offset: number;
    }
  | {
      direction: SpreadsheetSortDirection;
      kind: 'value';
      offset: number;
    };

export function sortSpreadsheetMatrix(
  source: readonly (readonly (Cell | null)[])[],
  request: SpreadsheetSortRequest,
  appearanceRows?: SpreadsheetSortAppearanceRows,
): SpreadsheetSortResult {
  const validation = validateSpreadsheetSortRequest(request);
  if (!validation.ok) return validation;
  const { range, hasHeader, keys, orientation } = validation.request;
  const width = range.column[1] - range.column[0] + 1;
  const height = range.row[1] - range.row[0] + 1;
  if (source.length !== height || source.some((row) => row.length !== width)) {
    return spreadsheetSortError('invalid-matrix');
  }
  if (source.some((row) => row.some((cell) => Boolean(cell?.hl)))) {
    return spreadsheetSortError('unsupported-linked-cell');
  }

  const resolvedAppearanceRows =
    appearanceRows ?? createSpreadsheetSortDirectAppearanceRows(source);
  if (
    keys.some((key) => key.sortOn !== undefined) &&
    !spreadsheetSortAppearanceRowsMatch(source, resolvedAppearanceRows)
  ) {
    return spreadsheetSortError('invalid-appearance');
  }

  const vertical = orientation === 'top-to-bottom';
  const header = vertical && hasHeader ? source[0] : undefined;
  const items = vertical
    ? spreadsheetSortRowItems(source, resolvedAppearanceRows, hasHeader)
    : spreadsheetSortColumnItems(source, resolvedAppearanceRows, width);
  const firstKeyIndex = vertical ? range.column[0] : range.row[0];
  const compiledKeys = keys.map((key) =>
    compileSpreadsheetSortKey(key, firstKeyIndex),
  );
  if (
    compiledKeys.some(
      (key) =>
        key.kind === 'appearance' &&
        !items.some((item) =>
          spreadsheetSortCellMatchesAppearance(
            item.appearance[key.offset],
            key.target,
          ),
        ),
    )
  ) {
    return spreadsheetSortError('invalid-appearance');
  }

  const sorted = [...items].sort((left, right) => {
    for (const key of compiledKeys) {
      const order = compareSpreadsheetSortCells(
        left.cells[key.offset] ?? null,
        right.cells[key.offset] ?? null,
        left.appearance[key.offset],
        right.appearance[key.offset],
        key,
      );
      if (order) return order;
    }
    return left.index - right.index;
  });

  const translatedItems: (Cell | null)[][] = [];
  for (const [targetIndex, item] of sorted.entries()) {
    const delta = targetIndex - item.index;
    const translated = translateSpreadsheetSortItem(
      item.cells,
      vertical ? delta : 0,
      vertical ? 0 : delta,
    );
    if (!translated) {
      return spreadsheetSortError('formula-reference-out-of-range');
    }
    translatedItems.push(translated);
  }

  if (vertical) {
    return {
      ok: true,
      rows: header
        ? [header as (Cell | null)[], ...translatedItems]
        : translatedItems,
    };
  }
  return {
    ok: true,
    rows: Array.from({ length: height }, (_, rowOffset) =>
      translatedItems.map((column) => column[rowOffset] ?? null),
    ),
  };
}

function spreadsheetSortRowItems(
  source: readonly (readonly (Cell | null)[])[],
  appearances: SpreadsheetSortAppearanceRows,
  hasHeader: boolean,
): SpreadsheetSortItem[] {
  const offset = hasHeader ? 1 : 0;
  return source.slice(offset).map((cells, index) => ({
    appearance: appearances[index + offset] ?? [],
    cells,
    index,
  }));
}

function spreadsheetSortColumnItems(
  source: readonly (readonly (Cell | null)[])[],
  appearances: SpreadsheetSortAppearanceRows,
  width: number,
): SpreadsheetSortItem[] {
  return Array.from({ length: width }, (_, columnOffset) => ({
    appearance: source.map(
      (_, rowOffset) => appearances[rowOffset]?.[columnOffset],
    ),
    cells: source.map((row) => row[columnOffset] ?? null),
    index: columnOffset,
  }));
}

function compareSpreadsheetSortCells(
  left: Cell | null,
  right: Cell | null,
  leftAppearance: SpreadsheetSortAppearanceRows[number][number] | undefined,
  rightAppearance: SpreadsheetSortAppearanceRows[number][number] | undefined,
  key: SpreadsheetSortCompiledKey,
): number {
  if (key.kind === 'appearance') {
    const leftMatches = spreadsheetSortCellMatchesAppearance(
      leftAppearance,
      key.target,
    );
    const rightMatches = spreadsheetSortCellMatchesAppearance(
      rightAppearance,
      key.target,
    );
    if (leftMatches === rightMatches) return 0;
    const matchedFirst = key.position === 'first';
    return leftMatches === matchedFirst ? -1 : 1;
  }
  const leftValue = spreadsheetSortValue(left);
  const rightValue = spreadsheetSortValue(right);
  if (leftValue === null) return rightValue === null ? 0 : 1;
  if (rightValue === null) return -1;
  if (key.kind === 'custom-list') {
    const leftRank = key.customRanks.get(
      spreadsheetSortCustomListMatchKey(leftValue),
    );
    const rightRank = key.customRanks.get(
      spreadsheetSortCustomListMatchKey(rightValue),
    );
    if (leftRank !== undefined || rightRank !== undefined) {
      if (leftRank === undefined) return 1;
      if (rightRank === undefined) return -1;
      return leftRank - rightRank;
    }
  }
  const order =
    typeof leftValue === 'number' && typeof rightValue === 'number'
      ? leftValue - rightValue
      : spreadsheetSortCollator.compare(String(leftValue), String(rightValue));
  return key.kind === 'value' && key.direction === 'descending'
    ? -order
    : order;
}

function compileSpreadsheetSortKey(
  key: SpreadsheetSortKey,
  firstKeyIndex: number,
): SpreadsheetSortCompiledKey {
  if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
    return {
      kind: 'appearance',
      offset: key.index - firstKeyIndex,
      position: key.position,
      target: { kind: key.sortOn, color: key.color },
    };
  }
  if (key.sortOn === 'icon') {
    return {
      kind: 'appearance',
      offset: key.index - firstKeyIndex,
      position: key.position,
      target: { kind: 'icon', icon: key.icon },
    };
  }
  if (key.customList !== undefined) {
    return {
      kind: 'custom-list',
      offset: key.index - firstKeyIndex,
      customRanks: new Map(
        key.customList.map((entry, index) => [
          spreadsheetSortCustomListMatchKey(entry),
          index,
        ]),
      ),
    };
  }
  return {
    kind: 'value',
    offset: key.index - firstKeyIndex,
    direction: key.direction ?? 'ascending',
  };
}

function spreadsheetSortValue(cell: Cell | null): number | string | null {
  const value = cell?.v ?? cell?.m;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (typeof value === 'boolean') return value ? 1 : 0;
  return null;
}

function translateSpreadsheetSortItem(
  cells: readonly (Cell | null)[],
  rowOffset: number,
  columnOffset: number,
): (Cell | null)[] | null {
  if (rowOffset === 0 && columnOffset === 0) {
    return cells as (Cell | null)[];
  }
  const translated: (Cell | null)[] = [];
  for (const cell of cells) {
    if (!cell?.f) {
      translated.push(cell);
      continue;
    }
    const formula = translateSpreadsheetFormula(
      cell.f,
      rowOffset,
      columnOffset,
    );
    if (formula === null) return null;
    translated.push(formula === cell.f ? cell : { ...cell, f: formula });
  }
  return translated;
}

const spreadsheetSortCollator = new Intl.Collator('zh-CN', {
  numeric: true,
  sensitivity: 'base',
});
