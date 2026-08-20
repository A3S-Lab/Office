import { normalizeSheetProtectionAuthority } from '../work-spreadsheet-protection';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { setSpreadsheetCellBorders } from './spreadsheet-cell-border';
import {
  type SpreadsheetCellRange,
  subtractSpreadsheetCellRange,
} from './spreadsheet-cell-range';
import {
  pasteContentCopiesHyperlinks,
  pasteContentCopiesProtection,
  pasteContentCopiesValidation,
} from './spreadsheet-paste-special-mode';
import { spreadsheetPasteSourceAt } from './spreadsheet-paste-special-plan';
import type {
  SpreadsheetCellWriter,
  SpreadsheetClipboardSnapshot,
  SpreadsheetPastePlan,
  SpreadsheetPasteSpecialOptions,
} from './spreadsheet-paste-special-types';

type UnknownRecord = Record<string, unknown>;
type NativeMerge = { r: number; c: number; rs: number; cs: number };

export function applySpreadsheetPasteSheetMetadata(
  sheet: WorkSpreadsheetSheet,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  processed: Array<{ row: number; column: number }>,
  options: SpreadsheetPasteSpecialOptions,
): WorkSpreadsheetSheet {
  let next = sheet;
  const processedRanges = compactSpreadsheetCoordinates(processed);
  if (pasteContentCopiesValidation(options.content)) {
    next = applySpreadsheetPasteValidation(
      next,
      snapshot,
      plan,
      processed,
      processedRanges,
      options.transpose,
    );
  }
  if (pasteContentCopiesHyperlinks(options.content)) {
    next = applySpreadsheetPasteHyperlinks(
      next,
      snapshot,
      plan,
      processed,
      options.transpose,
    );
  }
  if (pasteContentCopiesProtection(options.content)) {
    next = clearSpreadsheetPasteProtectionRanges(next, processedRanges);
  }
  return next;
}

export function applySpreadsheetPasteBorders(
  content: WorkSpreadsheetContent,
  sheetId: string,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  processed: Array<{ row: number; column: number }>,
  transpose: boolean,
): WorkSpreadsheetContent {
  let next = content;
  for (const range of compactSpreadsheetCoordinates(processed)) {
    next =
      setSpreadsheetCellBorders(next, sheetId, range, {
        target: 'none',
        color: '#000000',
        style: 'thin',
      }) ?? next;
  }
  const sheetIndex = next.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = next.sheets[sheetIndex];
  if (!sheet) return next;

  const records = processed.flatMap(({ row, column }) => {
    const source = spreadsheetPasteSourceAt(
      snapshot,
      plan,
      row,
      column,
      transpose,
    ).cell;
    const value: UnknownRecord = { row_index: row, col_index: column };
    const sides: Array<
      [keyof typeof source.borders, 'l' | 'r' | 't' | 'b' | 's']
    > = [
      ['left', 'l'],
      ['right', 'r'],
      ['top', 't'],
      ['bottom', 'b'],
      ['diagonal', 's'],
    ];
    for (const [side, key] of sides) {
      const line = source.borders[side];
      if (line) value[key] = structuredClone(line);
    }
    return Object.keys(value).length > 2 ? [{ rangeType: 'cell', value }] : [];
  });
  const borderInfo = [
    ...(Array.isArray(sheet.config?.borderInfo)
      ? (sheet.config.borderInfo as unknown[])
      : []),
    ...records,
  ];
  const sheets = [...next.sheets];
  sheets[sheetIndex] = {
    ...sheet,
    config: {
      ...sheet.config,
      borderInfo: borderInfo as NonNullable<
        WorkSpreadsheetSheet['config']
      >['borderInfo'],
    },
  };
  return { ...next, sheets };
}

export function applySpreadsheetPasteMerges(
  writer: SpreadsheetCellWriter,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  transpose: boolean,
): Record<string, NativeMerge> {
  const merges: Record<string, NativeMerge> = {};
  if (!snapshot.merges.length) return merges;

  for (
    let blockRow = plan.targetRange.row[0];
    blockRow <= plan.targetRange.row[1];
    blockRow += plan.rowCount
  ) {
    for (
      let blockColumn = plan.targetRange.column[0];
      blockColumn <= plan.targetRange.column[1];
      blockColumn += plan.columnCount
    ) {
      for (const source of snapshot.merges) {
        const row = blockRow + (transpose ? source.column : source.row);
        const column = blockColumn + (transpose ? source.row : source.column);
        const rowSpan = transpose ? source.columnSpan : source.rowSpan;
        const columnSpan = transpose ? source.rowSpan : source.columnSpan;
        const merge = { r: row, c: column, rs: rowSpan, cs: columnSpan };
        merges[`${row}_${column}`] = merge;
        for (let targetRow = row; targetRow < row + rowSpan; targetRow += 1) {
          for (
            let targetColumn = column;
            targetColumn < column + columnSpan;
            targetColumn += 1
          ) {
            const cell = structuredClone(
              writer.get(targetRow, targetColumn) ?? {},
            );
            cell.mc =
              targetRow === row && targetColumn === column
                ? { r: row, c: column, rs: rowSpan, cs: columnSpan }
                : { r: row, c: column };
            writer.set(targetRow, targetColumn, cell);
          }
        }
      }
    }
  }
  return merges;
}

export function withSpreadsheetPasteMerges(
  sheet: WorkSpreadsheetSheet,
  merges: Record<string, NativeMerge>,
): WorkSpreadsheetSheet {
  if (!Object.keys(merges).length) return sheet;
  return {
    ...sheet,
    config: {
      ...sheet.config,
      merge: { ...(sheet.config?.merge ?? {}), ...merges },
    },
  };
}

export function pasteSpreadsheetColumnWidths(
  sheet: WorkSpreadsheetSheet,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
): WorkSpreadsheetSheet {
  const widths = snapshot.columnWidths ?? [];
  const columnlen = { ...(sheet.config?.columnlen ?? {}) };
  const customWidth = { ...(sheet.config?.customWidth ?? {}) };
  for (
    let column = plan.targetRange.column[0];
    column <= plan.targetRange.column[1];
    column += 1
  ) {
    const width = widths[(column - plan.targetRange.column[0]) % widths.length];
    if (width === undefined) continue;
    columnlen[column] = width;
    customWidth[column] = 1;
  }
  return {
    ...sheet,
    config: { ...sheet.config, columnlen, customWidth },
  };
}

export function createSpreadsheetCellWriter(
  source: WorkSpreadsheetSheet,
): SpreadsheetCellWriter {
  if (source.data !== undefined) {
    const data = source.data.slice();
    const mutableRows = new Set<number>();
    const mutableRow = (row: number) => {
      if (!mutableRows.has(row)) {
        data[row] = data[row]?.slice() ?? [];
        mutableRows.add(row);
      }
      return data[row]!;
    };
    return {
      get: (row, column) => data[row]?.[column],
      set: (row, column, cell) => {
        const values = mutableRow(row);
        if (cell) values[column] = cell;
        else delete values[column];
      },
      finish: (sheet) => ({
        ...sheet,
        row: Math.max(sheet.row ?? 0, data.length),
        column: Math.max(
          sheet.column ?? 0,
          ...[...mutableRows].map((row) => data[row]?.length ?? 0),
        ),
        data,
      }),
    };
  }

  const entries = new Map(
    (source.celldata ?? []).map((entry) => [
      `${entry.r}:${entry.c}`,
      { ...entry },
    ]),
  );
  return {
    get: (row, column) => entries.get(`${row}:${column}`)?.v,
    set: (row, column, cell) => {
      const key = `${row}:${column}`;
      if (cell) entries.set(key, { r: row, c: column, v: cell });
      else entries.delete(key);
    },
    finish: (sheet) => ({
      ...sheet,
      celldata: [...entries.values()].sort(
        (left, right) => left.r - right.r || left.c - right.c,
      ),
    }),
  };
}

export function withSpreadsheetPasteSelection(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): WorkSpreadsheetSheet {
  return {
    ...sheet,
    row: Math.max(sheet.row ?? 0, range.row[1] + 1),
    column: Math.max(sheet.column ?? 0, range.column[1] + 1),
    luckysheet_select_save: [
      {
        row: [...range.row],
        column: [...range.column],
        row_focus: range.row[0],
        column_focus: range.column[0],
      },
    ],
  };
}

function applySpreadsheetPasteValidation(
  sheet: WorkSpreadsheetSheet,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  processed: Array<{ row: number; column: number }>,
  processedRanges: SpreadsheetCellRange[],
  transpose: boolean,
): WorkSpreadsheetSheet {
  const dataVerification = { ...(sheet.dataVerification ?? {}) };
  for (const { row, column } of processed) {
    delete dataVerification[`${row}_${column}`];
    const source = spreadsheetPasteSourceAt(
      snapshot,
      plan,
      row,
      column,
      transpose,
    ).cell;
    if (source.validation) {
      dataVerification[`${row}_${column}`] = {
        ...structuredClone(source.validation),
        rangeTxt: spreadsheetCellReference(row, column),
      };
    }
  }

  let ranges = [...(sheet.dataValidationRanges ?? [])];
  for (const processedRange of processedRanges) {
    ranges = ranges.flatMap((item) => {
      const remaining = item.ranges.flatMap((range) =>
        subtractSpreadsheetCellRange(range, processedRange),
      );
      return remaining.length ? [{ ...item, ranges: remaining }] : [];
    });
  }
  return {
    ...sheet,
    dataVerification: Object.keys(dataVerification).length
      ? dataVerification
      : undefined,
    dataValidationRanges: ranges.length ? ranges : undefined,
  };
}

function applySpreadsheetPasteHyperlinks(
  sheet: WorkSpreadsheetSheet,
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  processed: Array<{ row: number; column: number }>,
  transpose: boolean,
): WorkSpreadsheetSheet {
  const hyperlink = { ...(sheet.hyperlink ?? {}) };
  for (const { row, column } of processed) {
    const key = `${row}_${column}`;
    delete hyperlink[key];
    const source = spreadsheetPasteSourceAt(
      snapshot,
      plan,
      row,
      column,
      transpose,
    ).cell;
    if (source.hyperlink) hyperlink[key] = structuredClone(source.hyperlink);
  }
  return {
    ...sheet,
    hyperlink: Object.keys(hyperlink).length ? hyperlink : undefined,
  };
}

function clearSpreadsheetPasteProtectionRanges(
  sheet: WorkSpreadsheetSheet,
  processedRanges: SpreadsheetCellRange[],
): WorkSpreadsheetSheet {
  if (!processedRanges.length) return sheet;
  const authority = normalizeSheetProtectionAuthority(sheet.config?.authority);
  let ranges = authority.cellProtectionRanges;
  for (const processedRange of processedRanges) {
    ranges = ranges.flatMap((item) =>
      subtractSpreadsheetCellRange(item.range, processedRange).map((range) => ({
        ...item,
        range,
      })),
    );
  }
  authority.cellProtectionRanges = ranges;
  return { ...sheet, config: { ...sheet.config, authority } };
}

function compactSpreadsheetCoordinates(
  coordinates: Array<{ row: number; column: number }>,
): SpreadsheetCellRange[] {
  if (!coordinates.length) return [];
  const columnsByRow = new Map<number, number[]>();
  for (const { row, column } of coordinates) {
    const columns = columnsByRow.get(row) ?? [];
    columnsByRow.set(row, columns);
    columns.push(column);
  }

  const completed: SpreadsheetCellRange[] = [];
  let active = new Map<string, SpreadsheetCellRange>();
  for (const row of [...columnsByRow.keys()].sort(
    (left, right) => left - right,
  )) {
    const columns = [...new Set(columnsByRow.get(row))].sort(
      (left, right) => left - right,
    );
    const runs: Array<[number, number]> = [];
    for (const column of columns) {
      const current = runs.at(-1);
      if (current && column === current[1] + 1) current[1] = column;
      else runs.push([column, column]);
    }
    const next = new Map<string, SpreadsheetCellRange>();
    for (const column of runs) {
      const signature = `${column[0]}:${column[1]}`;
      const previous = active.get(signature);
      if (previous && previous.row[1] === row - 1) {
        previous.row[1] = row;
        next.set(signature, previous);
      } else {
        next.set(signature, { row: [row, row], column });
      }
    }
    for (const [signature, range] of active) {
      if (!next.has(signature)) completed.push(range);
    }
    active = next;
  }
  completed.push(...active.values());
  return completed;
}

function spreadsheetCellReference(row: number, column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return `${label}${row + 1}`;
}
