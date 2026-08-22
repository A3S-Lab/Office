import type { Cell } from '@fortune-sheet/core';
import {
  type XlsxNativeFillCellKey,
  xlsxNativeFillCellKeys,
} from '../work-xlsx-native-fill';
import type {
  SpreadsheetClipboardCell,
  SpreadsheetPasteContent,
  SpreadsheetPasteOperation,
} from './spreadsheet-paste-special-types';

export const spreadsheetPasteCellInvalid = Symbol(
  'spreadsheet-paste-cell-invalid',
);

export type SpreadsheetPasteCellResult =
  | Cell
  | null
  | typeof spreadsheetPasteCellInvalid;

type MutableCell = Cell & { hi?: number } & Partial<
    Record<XlsxNativeFillCellKey, unknown>
  >;
type UnknownRecord = Record<string, unknown>;

const spreadsheetCellContentKeys = [
  'v',
  'm',
  'f',
  'qp',
  'spl',
] as const satisfies readonly (keyof Cell)[];

const spreadsheetCellFormatKeys = [
  'bg',
  'bl',
  'it',
  'ff',
  'fs',
  'fc',
  'ht',
  'vt',
  'tb',
  'cl',
  'un',
  'tr',
  'rt',
  ...xlsxNativeFillCellKeys,
] as const satisfies readonly (keyof MutableCell)[];

export function pasteSpreadsheetSpecialCell({
  source,
  destination,
  content,
  operation,
  rowOffset,
  columnOffset,
}: {
  source: SpreadsheetClipboardCell;
  destination: Cell | null | undefined;
  content: SpreadsheetPasteContent;
  operation: SpreadsheetPasteOperation;
  rowOffset: number;
  columnOffset: number;
}): SpreadsheetPasteCellResult {
  let next: MutableCell;
  switch (content) {
    case 'all':
    case 'all-except-borders':
      next = cloneCellWithoutMerge(source.cell) ?? {};
      if (!translateCellFormula(next, rowOffset, columnOffset)) {
        return spreadsheetPasteCellInvalid;
      }
      applyClipboardProtection(next, source);
      break;
    case 'formulas':
      next = cloneCellWithoutMerge(destination) ?? {};
      if (
        !replaceCellContent(next, source.cell, true, rowOffset, columnOffset)
      ) {
        return spreadsheetPasteCellInvalid;
      }
      break;
    case 'values':
      next = cloneCellWithoutMerge(destination) ?? {};
      replaceCellContent(next, source.cell, false, rowOffset, columnOffset);
      break;
    case 'formats':
      next = cloneCellWithoutMerge(destination) ?? {};
      replaceCellFormatting(next, source.cell);
      applyClipboardProtection(next, source);
      break;
    case 'comments':
      next = cloneCellWithoutMerge(destination) ?? {};
      if (source.cell?.ps) next.ps = structuredClone(source.cell.ps);
      else delete next.ps;
      break;
    case 'formulas-and-number-formats':
      next = cloneCellWithoutMerge(destination) ?? {};
      if (
        !replaceCellContent(next, source.cell, true, rowOffset, columnOffset)
      ) {
        return spreadsheetPasteCellInvalid;
      }
      replaceNumberFormat(next, source.cell);
      break;
    case 'values-and-number-formats':
      next = cloneCellWithoutMerge(destination) ?? {};
      replaceCellContent(next, source.cell, false, rowOffset, columnOffset);
      replaceNumberFormat(next, source.cell);
      break;
    case 'validation':
    case 'column-widths':
      return destination ? structuredClone(destination) : null;
  }

  if (operation !== 'none') {
    const sourceValue = numericSpreadsheetCellValue(source.cell);
    const destinationValue = numericSpreadsheetCellValue(destination);
    if (
      sourceValue === null ||
      (destinationValue === null && spreadsheetCellHasContent(destination))
    ) {
      return spreadsheetPasteCellInvalid;
    }
    const value = spreadsheetPasteOperationValue(
      destinationValue ?? 0,
      sourceValue,
      operation,
    );
    if (value === null) return spreadsheetPasteCellInvalid;
    clearCellContent(next);
    next.v = value;
    next.m = String(value);
  }

  removeUndefinedProperties(next);
  return Object.keys(next).length ? next : null;
}

export function spreadsheetClipboardCellIsBlank(
  source: SpreadsheetClipboardCell,
  content: SpreadsheetPasteContent,
): boolean {
  switch (content) {
    case 'comments':
      return !source.cell?.ps;
    case 'validation':
      return source.validation === undefined;
    case 'formats':
      return (
        !spreadsheetCellHasFormatting(source.cell) &&
        !source.protection &&
        Object.keys(source.borders).length === 0
      );
    case 'column-widths':
      return false;
    case 'all':
    case 'all-except-borders':
      return (
        !source.cell &&
        source.validation === undefined &&
        source.hyperlink === undefined &&
        !source.protection &&
        Object.keys(source.borders).length === 0
      );
    case 'formulas':
    case 'values':
    case 'formulas-and-number-formats':
    case 'values-and-number-formats':
      return !spreadsheetCellHasContent(source.cell);
  }
}

export function spreadsheetPasteContentSupportsOperation(
  content: SpreadsheetPasteContent,
): boolean {
  return !['formats', 'comments', 'validation', 'column-widths'].includes(
    content,
  );
}

export function numericSpreadsheetCellValue(
  cell: Cell | null | undefined,
): number | null {
  if (typeof cell?.v === 'number' && Number.isFinite(cell.v)) return cell.v;
  const value = typeof cell?.v === 'string' ? cell.v.trim() : '';
  if (value && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value)) {
    const number = Number(value);
    if (Number.isFinite(number)) return number;
  }
  return null;
}

export function spreadsheetCellHasContent(
  cell: Cell | null | undefined,
): boolean {
  return Boolean(
    cell &&
      (cell.v !== undefined ||
        cell.m !== undefined ||
        Boolean(cell.f) ||
        cell.qp !== undefined ||
        cell.spl !== undefined ||
        (Array.isArray(cell.ct?.s) && cell.ct.s.length > 0)),
  );
}

export function translateSpreadsheetFormula(
  formula: string,
  rowOffset: number,
  columnOffset: number,
): string | null {
  if (!Number.isSafeInteger(rowOffset) || !Number.isSafeInteger(columnOffset)) {
    return null;
  }
  let result = '';
  let segmentStart = 0;
  let quoted = false;
  for (let index = 0; index < formula.length; index += 1) {
    if (formula[index] !== '"') continue;
    if (quoted && formula[index + 1] === '"') {
      index += 1;
      continue;
    }
    if (!quoted) {
      const translated = translateFormulaSegment(
        formula.slice(segmentStart, index),
        rowOffset,
        columnOffset,
      );
      if (translated === null) return null;
      result += translated;
      segmentStart = index;
      quoted = true;
    } else {
      result += formula.slice(segmentStart, index + 1);
      segmentStart = index + 1;
      quoted = false;
    }
  }
  if (quoted) return result + formula.slice(segmentStart);
  const translated = translateFormulaSegment(
    formula.slice(segmentStart),
    rowOffset,
    columnOffset,
  );
  return translated === null ? null : result + translated;
}

function replaceCellContent(
  target: MutableCell,
  source: Cell | null | undefined,
  formulas: boolean,
  rowOffset: number,
  columnOffset: number,
): boolean {
  clearCellContent(target);
  if (formulas && source?.f) {
    const translated = translateSpreadsheetFormula(
      source.f,
      rowOffset,
      columnOffset,
    );
    if (translated === null) return false;
    target.f = translated;
    if (source.v !== undefined) target.v = cloneValue(source.v);
    return true;
  }
  const value = spreadsheetLiteralCellValue(source);
  if (value !== undefined) target.v = cloneValue(value);
  return true;
}

function clearCellContent(cell: MutableCell): void {
  for (const key of spreadsheetCellContentKeys) delete cell[key];
  if (cell.ct && 's' in cell.ct) {
    const type = { ...cell.ct };
    delete type.s;
    if (Object.keys(type).length) cell.ct = type;
    else delete cell.ct;
  }
}

function replaceCellFormatting(
  target: MutableCell,
  source: Cell | null | undefined,
): void {
  const sourceCell = source as MutableCell | null | undefined;
  for (const key of spreadsheetCellFormatKeys) {
    if (sourceCell?.[key] === undefined) delete target[key];
    else target[key] = cloneValue(sourceCell[key]) as never;
  }
  const inline = Array.isArray(target.ct?.s)
    ? structuredClone(target.ct.s)
    : undefined;
  const format = source?.ct ? structuredClone(source.ct) : undefined;
  if (format) delete format.s;
  const nextType = {
    ...(format ?? {}),
    ...(inline ? { s: inline } : {}),
  };
  if (Object.keys(nextType).length) target.ct = nextType;
  else delete target.ct;
}

function replaceNumberFormat(
  target: MutableCell,
  source: Cell | null | undefined,
): void {
  const inline = Array.isArray(target.ct?.s)
    ? structuredClone(target.ct.s)
    : undefined;
  const nextType: NonNullable<Cell['ct']> = {
    ...(inline ? { s: inline } : {}),
  };
  if (source?.ct?.fa !== undefined) nextType.fa = source.ct.fa;
  if (source?.ct?.t !== undefined) nextType.t = source.ct.t;
  if (Object.keys(nextType).length) target.ct = nextType;
  else delete target.ct;
}

function applyClipboardProtection(
  target: MutableCell,
  source: SpreadsheetClipboardCell,
): void {
  delete target.lo;
  delete target.hi;
  if (!source.protection) return;
  target.lo = source.protection.locked ? 1 : 0;
  if (source.protection.hidden) target.hi = 1;
}

function translateCellFormula(
  cell: MutableCell,
  rowOffset: number,
  columnOffset: number,
): boolean {
  if (!cell.f) return true;
  const translated = translateSpreadsheetFormula(
    cell.f,
    rowOffset,
    columnOffset,
  );
  if (translated === null) return false;
  cell.f = translated;
  return true;
}

function spreadsheetLiteralCellValue(
  source: Cell | null | undefined,
): Cell['v'] | undefined {
  if (source?.v !== undefined) return source.v;
  if (Array.isArray(source?.ct?.s)) {
    return source.ct.s
      .map((segment: unknown) => {
        if (!isRecord(segment)) return '';
        const value = segment.v;
        return value === undefined || value === null ? '' : String(value);
      })
      .join('');
  }
  return source?.m;
}

function spreadsheetCellHasFormatting(cell: Cell | null | undefined): boolean {
  if (!cell) return false;
  const formattedCell = cell as MutableCell;
  return Boolean(
    spreadsheetCellFormatKeys.some((key) => formattedCell[key] !== undefined) ||
      cell.lo !== undefined ||
      (cell as MutableCell).hi !== undefined ||
      cell.ct?.fa !== undefined ||
      cell.ct?.t !== undefined,
  );
}

function spreadsheetPasteOperationValue(
  destination: number,
  source: number,
  operation: Exclude<SpreadsheetPasteOperation, 'none'>,
): number | null {
  if (operation === 'divide' && source === 0) return null;
  const value = {
    add: destination + source,
    subtract: destination - source,
    multiply: destination * source,
    divide: destination / source,
  }[operation];
  return Number.isFinite(value) ? value : null;
}

function translateFormulaSegment(
  segment: string,
  rowOffset: number,
  columnOffset: number,
): string | null {
  let valid = true;
  const sheetPrefix = "(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!";
  let next = segment.replace(
    new RegExp(
      String.raw`(^|[^A-Za-z0-9_.])(${sheetPrefix})?(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})(?![A-Za-z0-9_(])`,
      'g',
    ),
    (
      match,
      prefix: string,
      sheet: string | undefined,
      absoluteColumn: string,
      column: string,
      absoluteRow: string,
      row: string,
    ) => {
      const translatedColumn = absoluteColumn
        ? spreadsheetColumnIndex(column)
        : spreadsheetColumnIndex(column) + columnOffset;
      const translatedRow = absoluteRow ? Number(row) : Number(row) + rowOffset;
      if (
        translatedColumn < 0 ||
        translatedColumn >= 16_384 ||
        translatedRow < 1 ||
        translatedRow > 1_048_576
      ) {
        valid = false;
        return match;
      }
      return `${prefix}${sheet ?? ''}${absoluteColumn}${spreadsheetColumnLabel(translatedColumn)}${absoluteRow}${translatedRow}`;
    },
  );
  next = next.replace(
    new RegExp(
      String.raw`(^|[^A-Za-z0-9_.])(${sheetPrefix})?(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_])`,
      'g',
    ),
    (
      match,
      prefix: string,
      sheet: string | undefined,
      leftAbsolute: string,
      left: string,
      rightAbsolute: string,
      right: string,
    ) => {
      const leftColumn = leftAbsolute
        ? spreadsheetColumnIndex(left)
        : spreadsheetColumnIndex(left) + columnOffset;
      const rightColumn = rightAbsolute
        ? spreadsheetColumnIndex(right)
        : spreadsheetColumnIndex(right) + columnOffset;
      if (
        leftColumn < 0 ||
        rightColumn < 0 ||
        leftColumn >= 16_384 ||
        rightColumn >= 16_384
      ) {
        valid = false;
        return match;
      }
      return `${prefix}${sheet ?? ''}${leftAbsolute}${spreadsheetColumnLabel(leftColumn)}:${rightAbsolute}${spreadsheetColumnLabel(rightColumn)}`;
    },
  );
  next = next.replace(
    new RegExp(
      String.raw`(^|[^A-Za-z0-9_.])(${sheetPrefix})?(\$?)([1-9]\d{0,6}):(\$?)([1-9]\d{0,6})(?![A-Za-z0-9_])`,
      'g',
    ),
    (
      match,
      prefix: string,
      sheet: string | undefined,
      topAbsolute: string,
      top: string,
      bottomAbsolute: string,
      bottom: string,
    ) => {
      const topRow = topAbsolute ? Number(top) : Number(top) + rowOffset;
      const bottomRow = bottomAbsolute
        ? Number(bottom)
        : Number(bottom) + rowOffset;
      if (
        topRow < 1 ||
        bottomRow < 1 ||
        topRow > 1_048_576 ||
        bottomRow > 1_048_576
      ) {
        valid = false;
        return match;
      }
      return `${prefix}${sheet ?? ''}${topAbsolute}${topRow}:${bottomAbsolute}${bottomRow}`;
    },
  );
  return valid ? next : null;
}

function spreadsheetColumnIndex(value: string): number {
  let index = 0;
  for (const character of value.toUpperCase()) {
    index = index * 26 + character.charCodeAt(0) - 64;
  }
  return index - 1;
}

function spreadsheetColumnLabel(index: number): string {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function cloneCellWithoutMerge(
  cell: Cell | null | undefined,
): MutableCell | null {
  if (!cell) return null;
  const clone = structuredClone(cell) as MutableCell;
  delete clone.mc;
  return clone;
}

function removeUndefinedProperties(cell: MutableCell): void {
  const record = cell as UnknownRecord;
  for (const key of Object.keys(record)) {
    if (record[key] === undefined) delete record[key];
  }
}

function cloneValue<T>(value: T): T {
  return value && typeof value === 'object' ? structuredClone(value) : value;
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
