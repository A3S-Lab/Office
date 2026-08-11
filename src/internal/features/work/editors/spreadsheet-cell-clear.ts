import type { Cell, Selection, Sheet } from '@fortune-sheet/core';

export type SpreadsheetCellClearMode =
  | 'all'
  | 'formats'
  | 'contents'
  | 'comments'
  | 'hyperlinks';

type SpreadsheetCellClearRange = {
  row: [number, number];
  column: [number, number];
};

type ClearableCell = Cell & { hi?: number };
type UnknownRecord = Record<string, unknown>;

const spreadsheetCellFormatProperties = [
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
  'ct',
  'lo',
  'hi',
] as const satisfies readonly (keyof ClearableCell)[];

export function clearSpreadsheetSheetSelection(
  sheet: Sheet,
  selection: Pick<Selection, 'row' | 'column'>,
  mode: SpreadsheetCellClearMode,
): Sheet {
  const range = normalizeClearRange(selection);
  const data = sheet.data?.map((row, rowIndex) => {
    if (rowIndex < range.row[0] || rowIndex > range.row[1]) return row;
    return row.map((cell, columnIndex) =>
      clearSpreadsheetCellAt(cell, rowIndex, columnIndex, range, mode),
    );
  });
  const celldata = data
    ? undefined
    : sheet.celldata?.flatMap((entry) => {
        const cell = clearSpreadsheetCellAt(
          entry.v,
          entry.r,
          entry.c,
          range,
          mode,
        );
        return cell ? [{ ...entry, v: cell }] : [];
      });
  let next: Sheet = {
    ...sheet,
    ...(data ? { data } : {}),
    ...(celldata ? { celldata } : {}),
  };
  if (data && next.celldata) delete next.celldata;

  if (mode === 'formats' || mode === 'all') {
    next = clearSpreadsheetSheetFormats(next, range);
  }
  if (mode === 'hyperlinks' || mode === 'all') {
    next = clearSpreadsheetSheetHyperlinks(next, range);
  }
  return next;
}

function clearSpreadsheetCellAt(
  cell: Cell | null,
  row: number,
  column: number,
  range: SpreadsheetCellClearRange,
  mode: SpreadsheetCellClearMode,
): Cell | null {
  if (!cell || !clearRangeContains(range, row, column)) return cell;
  if (mode === 'all') {
    return cell.mc ? { mc: { ...cell.mc } } : null;
  }

  const next: ClearableCell = { ...cell };
  if (mode === 'contents') clearSpreadsheetCellContents(next);
  if (mode === 'formats') clearSpreadsheetCellFormats(next);
  if (mode === 'comments') delete next.ps;
  if (mode === 'hyperlinks') delete next.hl;
  removeUndefinedCellProperties(next);
  return Object.keys(next).length ? next : null;
}

function clearSpreadsheetCellContents(cell: ClearableCell): void {
  delete cell.v;
  delete cell.m;
  delete cell.f;
  delete cell.spl;
  delete cell.qp;
  if (cell.ct && 's' in cell.ct) {
    const cellType = { ...cell.ct };
    delete cellType.s;
    cell.ct = Object.keys(cellType).length ? cellType : undefined;
  }
}

function clearSpreadsheetCellFormats(cell: ClearableCell): void {
  const inlineText = spreadsheetInlineCellText(cell);
  for (const property of spreadsheetCellFormatProperties) {
    delete cell[property];
  }
  if (cell.v !== undefined) {
    cell.m = spreadsheetGeneralDisplayValue(cell.v);
  } else if (!cell.f && inlineText !== null) {
    cell.v = inlineText;
    cell.m = inlineText;
  }
}

function spreadsheetInlineCellText(cell: Cell): string | null {
  const segments = cell.ct?.s;
  if (!Array.isArray(segments)) return null;
  return segments
    .map((segment) => {
      if (!segment || typeof segment !== 'object') return '';
      const value = (segment as UnknownRecord).v;
      return value === null || value === undefined ? '' : String(value);
    })
    .join('');
}

function spreadsheetGeneralDisplayValue(value: Cell['v']): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return value === undefined ? '' : String(value);
}

function removeUndefinedCellProperties(cell: ClearableCell): void {
  const record = cell as UnknownRecord;
  for (const key of Object.keys(record)) {
    if (record[key] === undefined) delete record[key];
  }
}

function clearSpreadsheetSheetFormats(
  sheet: Sheet,
  range: SpreadsheetCellClearRange,
): Sheet {
  const config = sheet.config ? { ...sheet.config } : undefined;
  if (config?.borderInfo) {
    const borderInfo = clearSpreadsheetBorderFormats(
      config.borderInfo as unknown[],
      range,
    );
    if (borderInfo.length) {
      config.borderInfo = borderInfo as NonNullable<
        Sheet['config']
      >['borderInfo'];
    } else {
      delete config.borderInfo;
    }
  }

  const conditionalFormats = clearSpreadsheetConditionalFormats(
    sheet.luckysheet_conditionformat_save,
    range,
  );
  const alternateFormats = clearSpreadsheetAlternateFormats(
    sheet.luckysheet_alternateformat_save,
    range,
  );
  const next: Sheet = {
    ...sheet,
    ...(config ? { config } : {}),
    ...(conditionalFormats.length
      ? { luckysheet_conditionformat_save: conditionalFormats }
      : {}),
    ...(alternateFormats.length
      ? { luckysheet_alternateformat_save: alternateFormats }
      : {}),
  };
  if (!conditionalFormats.length) delete next.luckysheet_conditionformat_save;
  if (!alternateFormats.length) delete next.luckysheet_alternateformat_save;
  return next;
}

function clearSpreadsheetBorderFormats(
  formats: unknown[],
  range: SpreadsheetCellClearRange,
): unknown[] {
  return formats.flatMap((format) => {
    if (!isRecord(format)) return [format];
    if (format.rangeType === 'cell') {
      const value = isRecord(format.value) ? format.value : null;
      const row = finiteIndex(value?.row_index);
      const column = finiteIndex(value?.col_index);
      return row !== null &&
        column !== null &&
        clearRangeContains(range, row, column)
        ? []
        : [format];
    }
    if (format.rangeType !== 'range' || !Array.isArray(format.range)) {
      return [format];
    }
    const remaining = format.range.flatMap((candidate) => {
      const parsed = parseClearRange(candidate);
      if (!parsed) return [candidate];
      if (format.borderType === 'border-slash') {
        return clearRangesIntersect(parsed, range) ? [] : [parsed];
      }
      return subtractSpreadsheetRange(parsed, range);
    });
    return remaining.length ? [{ ...format, range: remaining }] : [];
  });
}

function clearSpreadsheetConditionalFormats(
  formats: Sheet['luckysheet_conditionformat_save'],
  range: SpreadsheetCellClearRange,
): unknown[] {
  if (!Array.isArray(formats)) return [];
  return formats.flatMap((format) => {
    if (!isRecord(format) || !Array.isArray(format.cellrange)) return [format];
    const remaining = format.cellrange.flatMap((candidate) => {
      const parsed = parseClearRange(candidate);
      return parsed ? subtractSpreadsheetRange(parsed, range) : [candidate];
    });
    return remaining.length ? [{ ...format, cellrange: remaining }] : [];
  });
}

function clearSpreadsheetAlternateFormats(
  formats: Sheet['luckysheet_alternateformat_save'],
  range: SpreadsheetCellClearRange,
): unknown[] {
  if (!Array.isArray(formats)) return [];
  return formats.flatMap((format) => {
    if (!isRecord(format)) return [format];
    const source = parseClearRange(format.cellrange);
    if (!source) return [format];
    return subtractSpreadsheetRange(source, range).map((cellrange) => ({
      ...format,
      cellrange,
    }));
  });
}

function clearSpreadsheetSheetHyperlinks(
  sheet: Sheet,
  range: SpreadsheetCellClearRange,
): Sheet {
  if (!sheet.hyperlink) return sheet;
  const hyperlink = Object.fromEntries(
    Object.entries(sheet.hyperlink).filter(([key]) => {
      const position = spreadsheetHyperlinkPosition(key);
      return (
        !position || !clearRangeContains(range, position.row, position.column)
      );
    }),
  );
  const next: Sheet = { ...sheet, hyperlink };
  if (!Object.keys(hyperlink).length) delete next.hyperlink;
  return next;
}

function spreadsheetHyperlinkPosition(
  value: string,
): { row: number; column: number } | null {
  const match = /^(\d+)_(\d+)$/.exec(value);
  if (!match) return null;
  return { row: Number(match[1]), column: Number(match[2]) };
}

function subtractSpreadsheetRange(
  source: SpreadsheetCellClearRange,
  removed: SpreadsheetCellClearRange,
): SpreadsheetCellClearRange[] {
  const top = Math.max(source.row[0], removed.row[0]);
  const bottom = Math.min(source.row[1], removed.row[1]);
  const left = Math.max(source.column[0], removed.column[0]);
  const right = Math.min(source.column[1], removed.column[1]);
  if (top > bottom || left > right) return [source];

  const remaining: SpreadsheetCellClearRange[] = [];
  if (source.row[0] < top) {
    remaining.push({
      row: [source.row[0], top - 1],
      column: [...source.column],
    });
  }
  if (bottom < source.row[1]) {
    remaining.push({
      row: [bottom + 1, source.row[1]],
      column: [...source.column],
    });
  }
  if (source.column[0] < left) {
    remaining.push({
      row: [top, bottom],
      column: [source.column[0], left - 1],
    });
  }
  if (right < source.column[1]) {
    remaining.push({
      row: [top, bottom],
      column: [right + 1, source.column[1]],
    });
  }
  return remaining;
}

function clearRangesIntersect(
  left: SpreadsheetCellClearRange,
  right: SpreadsheetCellClearRange,
): boolean {
  return !(
    left.row[1] < right.row[0] ||
    right.row[1] < left.row[0] ||
    left.column[1] < right.column[0] ||
    right.column[1] < left.column[0]
  );
}

function clearRangeContains(
  range: SpreadsheetCellClearRange,
  row: number,
  column: number,
): boolean {
  return (
    row >= range.row[0] &&
    row <= range.row[1] &&
    column >= range.column[0] &&
    column <= range.column[1]
  );
}

function normalizeClearRange(
  range: Pick<Selection, 'row' | 'column'>,
): SpreadsheetCellClearRange {
  return {
    row: normalizeClearAxis(range.row),
    column: normalizeClearAxis(range.column),
  };
}

function parseClearRange(value: unknown): SpreadsheetCellClearRange | null {
  if (!isRecord(value)) return null;
  const row = parseClearAxis(value.row);
  const column = parseClearAxis(value.column);
  return row && column ? { row, column } : null;
}

function normalizeClearAxis(value: number[]): [number, number] {
  const first = finiteIndex(value[0]) ?? 0;
  const second = finiteIndex(value[1]) ?? first;
  return [Math.min(first, second), Math.max(first, second)];
}

function parseClearAxis(value: unknown): [number, number] | null {
  if (!Array.isArray(value)) return null;
  const first = finiteIndex(value[0]);
  const second = finiteIndex(value[1]);
  if (first === null || second === null) return null;
  return [Math.min(first, second), Math.max(first, second)];
}

function finiteIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0
    ? Math.trunc(value)
    : null;
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
