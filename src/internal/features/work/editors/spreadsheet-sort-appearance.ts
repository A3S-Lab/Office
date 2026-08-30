import type { Cell, CellMatrix } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import {
  parseSpreadsheetConditionalColor,
  type SpreadsheetConditionalRgbColor,
} from '../work-spreadsheet-conditional-colors';
import { spreadsheetConditionalFormatStyles } from '../work-spreadsheet-conditional-format';
import {
  isSpreadsheetConditionalIconSetName,
  spreadsheetConditionalIconAppearance,
  spreadsheetConditionalIconSetCount,
  type SpreadsheetConditionalIconSetName,
} from '../work-spreadsheet-conditional-icons';
import type { WorkSpreadsheetSheet } from '../work-types';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import { spreadsheetCellFillFormat } from './spreadsheet-cell-fill-format';

export type SpreadsheetSortAppearanceKind =
  | 'cell-color'
  | 'font-color'
  | 'icon';

export type SpreadsheetSortAppearancePosition = 'top' | 'bottom';

export interface SpreadsheetSortIconTarget {
  iconSet: SpreadsheetConditionalIconSetName;
  index: number;
}

export type SpreadsheetSortAppearanceTarget =
  | { color: string | null; kind: 'cell-color' }
  | { color: string | null; kind: 'font-color' }
  | { icon: SpreadsheetSortIconTarget; kind: 'icon' };

export interface SpreadsheetSortCellAppearance {
  /** Undefined means the cell has a native pattern or gradient without one color. */
  cellColor?: string | null;
  /** Null represents the automatic/default font color. */
  fontColor: string | null;
  icon: SpreadsheetSortIconTarget | null;
}

export type SpreadsheetSortAppearanceRows =
  readonly (readonly SpreadsheetSortCellAppearance[])[];

export interface SpreadsheetSortAppearanceColumn {
  cellColors: readonly (string | null)[];
  column: number;
  fontColors: readonly (string | null)[];
  icons: readonly SpreadsheetSortIconTarget[];
}

export function createSpreadsheetSortAppearanceRows(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  rows: readonly (readonly (Cell | null)[])[],
): SpreadsheetSortAppearanceRows {
  const data = liveSpreadsheetSortMatrix(sheet, range, rows);
  const conditional = spreadsheetConditionalFormatStyles({ ...sheet, data });
  return rows.map((row, rowOffset) =>
    row.map((cell, columnOffset) => {
      const direct = spreadsheetSortDirectCellAppearance(cell);
      const style = conditional.get(
        `${range.row[0] + rowOffset}_${range.column[0] + columnOffset}`,
      );
      const conditionalCellColor = normalizeSpreadsheetSortColor(
        style?.cellColor,
      );
      const conditionalFontColor = normalizeSpreadsheetSortColor(
        style?.textColor,
      );
      return {
        cellColor:
          conditionalCellColor !== undefined
            ? conditionalCellColor
            : direct.cellColor,
        fontColor:
          conditionalFontColor !== undefined
            ? conditionalFontColor
            : direct.fontColor,
        icon: normalizeSpreadsheetSortIcon(style?.icon) ?? null,
      };
    }),
  );
}

export function createSpreadsheetSortDirectAppearanceRows(
  rows: readonly (readonly (Cell | null)[])[],
): SpreadsheetSortAppearanceRows {
  return rows.map((row) => row.map(spreadsheetSortDirectCellAppearance));
}

export function spreadsheetSortAppearanceColumns(
  rows: SpreadsheetSortAppearanceRows,
  range: SpreadsheetCellRange,
  hasHeader: boolean,
): readonly SpreadsheetSortAppearanceColumn[] {
  const width = range.column[1] - range.column[0] + 1;
  const sourceRows = rows.slice(hasHeader ? 1 : 0);
  return Array.from({ length: width }, (_, offset) => {
    const cellColors: Array<string | null> = [];
    const fontColors: Array<string | null> = [];
    const icons: SpreadsheetSortIconTarget[] = [];
    const seenCellColors = new Set<string>();
    const seenFontColors = new Set<string>();
    const seenIcons = new Set<string>();
    for (const row of sourceRows) {
      const appearance = row[offset];
      if (!appearance) continue;
      if (appearance.cellColor !== undefined) {
        appendUniqueColor(cellColors, seenCellColors, appearance.cellColor);
      }
      appendUniqueColor(fontColors, seenFontColors, appearance.fontColor);
      if (appearance.icon) {
        const key = spreadsheetSortIconTargetKey(appearance.icon);
        if (!seenIcons.has(key)) {
          seenIcons.add(key);
          icons.push({ ...appearance.icon });
        }
      }
    }
    return {
      column: range.column[0] + offset,
      cellColors,
      fontColors,
      icons,
    };
  });
}

export function normalizeSpreadsheetSortColor(
  value: unknown,
): string | null | undefined {
  if (value === null) return null;
  if (typeof value !== 'string') return undefined;
  const source = value.trim();
  const shorthand = /^#([0-9a-f]{3})$/iu.exec(source);
  const parsed = parseSpreadsheetConditionalColor(
    shorthand
      ? `#${[...shorthand[1]].map((character) => character.repeat(2)).join('')}`
      : source,
  );
  return parsed ? spreadsheetSortHexColor(parsed) : undefined;
}

export function normalizeSpreadsheetSortIcon(
  value: unknown,
): SpreadsheetSortIconTarget | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const source = value as Partial<SpreadsheetSortIconTarget>;
  if (
    !isSpreadsheetConditionalIconSetName(source.iconSet) ||
    !Number.isSafeInteger(source.index) ||
    source.index! < 0 ||
    source.index! >= spreadsheetConditionalIconSetCount(source.iconSet)
  ) {
    return null;
  }
  return { iconSet: source.iconSet, index: source.index! };
}

export function spreadsheetSortAppearanceTargetValue(
  target: SpreadsheetSortAppearanceTarget,
): string {
  if (target.kind === 'icon') {
    return `icon:${target.icon.iconSet}:${target.icon.index}`;
  }
  const value =
    target.color ?? (target.kind === 'cell-color' ? 'none' : 'automatic');
  return `${target.kind}:${value}`;
}

export function parseSpreadsheetSortAppearanceTargetValue(
  value: string,
): SpreadsheetSortAppearanceTarget | null {
  if (value.startsWith('cell-color:')) {
    const source = value.slice('cell-color:'.length);
    if (source === 'none') return { kind: 'cell-color', color: null };
    const color = normalizeSpreadsheetSortColor(source);
    return typeof color === 'string' ? { kind: 'cell-color', color } : null;
  }
  if (value.startsWith('font-color:')) {
    const source = value.slice('font-color:'.length);
    if (source === 'automatic') return { kind: 'font-color', color: null };
    const color = normalizeSpreadsheetSortColor(source);
    return typeof color === 'string' ? { kind: 'font-color', color } : null;
  }
  const icon = /^icon:([^:]+):(\d+)$/u.exec(value);
  if (!icon) return null;
  const normalized = normalizeSpreadsheetSortIcon({
    iconSet: icon[1],
    index: Number(icon[2]),
  });
  return normalized ? { kind: 'icon', icon: normalized } : null;
}

export function spreadsheetSortAppearanceTargetLabel(
  target: SpreadsheetSortAppearanceTarget,
): string {
  if (target.kind === 'cell-color') {
    return target.color
      ? `单元格颜色 ${target.color.toUpperCase()}`
      : '无单元格颜色';
  }
  if (target.kind === 'font-color') {
    return target.color
      ? `字体颜色 ${target.color.toUpperCase()}`
      : '自动字体颜色';
  }
  const count = spreadsheetConditionalIconSetCount(target.icon.iconSet);
  const appearance = spreadsheetConditionalIconAppearance({
    ...target.icon,
    count,
    showValue: true,
  });
  return `${appearance.glyph} ${appearance.label}`;
}

export function spreadsheetSortAppearanceTargets(
  column: SpreadsheetSortAppearanceColumn | undefined,
  kind: SpreadsheetSortAppearanceKind,
): readonly SpreadsheetSortAppearanceTarget[] {
  if (!column) return [];
  if (kind === 'cell-color') {
    return column.cellColors.map((color) => ({ kind, color }));
  }
  if (kind === 'font-color') {
    return column.fontColors.map((color) => ({ kind, color }));
  }
  return column.icons.map((icon) => ({ kind, icon: { ...icon } }));
}

export function spreadsheetSortAppearanceTargetsEqual(
  left: SpreadsheetSortAppearanceTarget,
  right: SpreadsheetSortAppearanceTarget,
): boolean {
  return (
    spreadsheetSortAppearanceTargetValue(left) ===
    spreadsheetSortAppearanceTargetValue(right)
  );
}

export function spreadsheetSortCellMatchesAppearance(
  appearance: SpreadsheetSortCellAppearance | undefined,
  target: SpreadsheetSortAppearanceTarget,
): boolean {
  if (!appearance) return false;
  if (target.kind === 'cell-color') {
    return appearance.cellColor === target.color;
  }
  if (target.kind === 'font-color') {
    return appearance.fontColor === target.color;
  }
  return Boolean(
    appearance.icon &&
      spreadsheetSortIconTargetKey(appearance.icon) ===
        spreadsheetSortIconTargetKey(target.icon),
  );
}

function spreadsheetSortDirectCellAppearance(
  cell: Cell | null,
): SpreadsheetSortCellAppearance {
  const fill = spreadsheetCellFillFormat(cell);
  const fontColor = normalizeSpreadsheetSortColor(cell?.fc);
  return {
    cellColor:
      fill.kind === 'solid'
        ? fill.color
        : fill.kind === 'none'
          ? null
          : undefined,
    fontColor: typeof fontColor === 'string' ? fontColor : null,
    icon: null,
  };
}

function liveSpreadsheetSortMatrix(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  rows: readonly (readonly (Cell | null)[])[],
): CellMatrix {
  const data = cloneSparseMatrix(sheet.data);
  for (const entry of sheet.celldata ?? []) {
    const row = data[entry.r] ? [...data[entry.r]!] : [];
    if (row[entry.c] === undefined) row[entry.c] = entry.v;
    data[entry.r] = row;
  }
  for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
    const absoluteRow = range.row[0] + rowOffset;
    const target = data[absoluteRow] ? [...data[absoluteRow]!] : [];
    const source = rows[rowOffset] ?? [];
    for (
      let columnOffset = 0;
      columnOffset < source.length;
      columnOffset += 1
    ) {
      target[range.column[0] + columnOffset] = source[columnOffset] ?? null;
    }
    data[absoluteRow] = target;
  }
  return data;
}

function appendUniqueColor(
  target: Array<string | null>,
  seen: Set<string>,
  color: string | null,
): void {
  const key = color ?? 'none';
  if (seen.has(key)) return;
  seen.add(key);
  target.push(color);
}

function spreadsheetSortIconTargetKey(target: SpreadsheetSortIconTarget) {
  return `${target.iconSet}:${target.index}`;
}

function spreadsheetSortHexColor(
  color: SpreadsheetConditionalRgbColor,
): string {
  return `#${[color.red, color.green, color.blue]
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}
