import { describe, expect, test } from '@rstest/core';
import {
  MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS,
  setSpreadsheetCellBorders,
  spreadsheetRenderableDiagonalBorders,
} from '../src/internal/features/work/editors/spreadsheet-cell-border';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import {
  collectXlsxCellBorders,
  fortuneBorderInfoFromXlsxCells,
} from '../src/internal/features/work/work-xlsx-cell-borders';
import type { Sheet } from '@fortune-sheet/core';

describe('spreadsheet cell border model', () => {
  test('authors Fortune-native range borders without mutating the workbook', () => {
    const source = workbook();

    const next = setSpreadsheetCellBorders(
      source,
      'sheet-1',
      { row: [3, 1], column: [4, 2] },
      { target: 'all', color: '#2463EB', style: 'medium' },
    );

    expect(next).not.toBeNull();
    expect(next).not.toBe(source);
    expect(next?.sheets[0]).not.toBe(source.sheets[0]);
    expect(source.sheets[0]?.config).toBeUndefined();
    expect(next?.sheets[0]?.config?.borderInfo).toEqual([
      {
        rangeType: 'range',
        borderType: 'border-all',
        color: '#2463eb',
        style: '8',
        range: [{ row: [1, 3], column: [2, 4] }],
      },
    ]);
  });

  test('subtracts replaced ranges while preserving malformed vendor records', () => {
    const malformed = {
      rangeType: 'range',
      borderType: 'border-all',
      color: 'vendor-color',
      range: [{ row: ['vendor'], column: [0, 4] }],
    };
    const source = workbook([
      malformed,
      {
        rangeType: 'range',
        borderType: 'border-all',
        color: '#111111',
        style: '1',
        range: [{ row: [0, 4], column: [0, 4] }],
      },
      {
        rangeType: 'range',
        borderType: 'border-top',
        color: '#778899',
        style: '3',
        range: [{ row: [1, 3], column: [1, 3] }],
      },
    ]);

    const next = setSpreadsheetCellBorders(
      source,
      'sheet-1',
      { row: [1, 3], column: [1, 3] },
      { target: 'all', color: '#224466', style: 'thin' },
    );

    expect(next?.sheets[0]?.config?.borderInfo).toEqual([
      malformed,
      {
        rangeType: 'range',
        borderType: 'border-all',
        color: '#111111',
        style: '1',
        range: [
          { row: [0, 0], column: [0, 4] },
          { row: [4, 4], column: [0, 4] },
          { row: [1, 3], column: [0, 0] },
          { row: [1, 3], column: [4, 4] },
        ],
      },
      {
        rangeType: 'range',
        borderType: 'border-all',
        color: '#224466',
        style: '1',
        range: [{ row: [1, 3], column: [1, 3] }],
      },
    ]);
  });

  test('writes a final no-border record and removes superseded cell borders', () => {
    const malformedCell = { rangeType: 'cell', value: { row_index: 'x' } };
    const source = workbook([
      malformedCell,
      {
        rangeType: 'cell',
        value: {
          row_index: 1,
          col_index: 1,
          t: { color: '#111111', style: '1' },
        },
      },
      {
        rangeType: 'range',
        borderType: 'border-outside',
        color: '#111111',
        style: '1',
        range: [{ row: [0, 2], column: [0, 2] }],
      },
      {
        rangeType: 'range',
        borderType: 'border-slash',
        color: '#222222',
        style: '3',
        range: [
          {
            row: [1, 1],
            column: [1, 1],
            row_focus: 1,
            column_focus: 1,
          },
        ],
      },
    ]);

    const next = setSpreadsheetCellBorders(
      source,
      'sheet-1',
      { row: [1, 1], column: [1, 1] },
      { target: 'none', color: '#000000', style: 'thin' },
    );

    expect(next?.sheets[0]?.config?.borderInfo).toEqual([
      malformedCell,
      {
        rangeType: 'range',
        borderType: 'border-outside',
        color: '#111111',
        style: '1',
        range: [
          { row: [0, 0], column: [0, 2] },
          { row: [2, 2], column: [0, 2] },
          { row: [1, 1], column: [0, 0] },
          { row: [1, 1], column: [2, 2] },
        ],
      },
      {
        rangeType: 'range',
        borderType: 'border-none',
        color: '#000000',
        style: '1',
        range: [{ row: [1, 1], column: [1, 1] }],
      },
    ]);
  });

  test('bounds repeated applications to one managed record per target', () => {
    let current = workbook();
    for (let index = 0; index < 100; index += 1) {
      current =
        setSpreadsheetCellBorders(
          current,
          'sheet-1',
          { row: [0, 8], column: [0, 5] },
          {
            target: 'outside',
            color: index % 2 ? '#112233' : '#445566',
            style: index % 2 ? 'dashed' : 'thick',
          },
        ) ?? current;
    }

    expect(current.sheets[0]?.config?.borderInfo).toHaveLength(1);
    expect(current.sheets[0]?.config?.borderInfo?.[0]).toMatchObject({
      borderType: 'border-outside',
      color: '#112233',
      style: '4',
    });
  });

  test('authors independent diagonal-down and diagonal-up borders per cell', () => {
    const source = workbook();
    const withDown = setSpreadsheetCellBorders(
      source,
      'sheet-1',
      { row: [1, 2], column: [3, 4] },
      { target: 'diagonalDown', color: '#b42318', style: 'dotted' },
    );

    expect(withDown?.sheets[0]?.config?.borderInfo).toEqual([
      diagonal(1, 3, 'down'),
      diagonal(1, 4, 'down'),
      diagonal(2, 3, 'down'),
      diagonal(2, 4, 'down'),
    ]);

    const withBoth = setSpreadsheetCellBorders(
      withDown ?? source,
      'sheet-1',
      { row: [1, 2], column: [3, 4] },
      { target: 'diagonalUp', color: '#2463eb', style: 'thick' },
    );
    expect(withBoth?.sheets[0]?.config?.borderInfo).toEqual([
      diagonal(1, 3, 'both', '#2463eb', '13'),
      diagonal(1, 4, 'both', '#2463eb', '13'),
      diagonal(2, 3, 'both', '#2463eb', '13'),
      diagonal(2, 4, 'both', '#2463eb', '13'),
    ]);

    for (const target of ['diagonalDown', 'diagonalUp'] as const) {
      expect(
        setSpreadsheetCellBorders(
          source,
          'sheet-1',
          {
            row: [0, MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS],
            column: [0, 0],
          },
          { target, color: '#b42318', style: 'dotted' },
        ),
      ).toBeNull();
    }
  });

  test('indexes existing diagonal directions once for the maximum selection', () => {
    let recordReads = 0;
    const records = Array.from({ length: 128 }, (_, index) => ({
      rangeType: 'cell',
      value: {
        row_index: 10_000 + index,
        col_index: index,
        s: { color: '#123456', style: '1' },
      },
    }));
    const borderInfo = new Proxy(records, {
      get(target, property, receiver) {
        if (typeof property === 'string' && /^\d+$/.test(property)) {
          recordReads += 1;
        }
        return Reflect.get(target, property, receiver);
      },
    });

    const next = setSpreadsheetCellBorders(
      workbook(borderInfo),
      'sheet-1',
      { row: [0, 63], column: [0, 63] },
      { target: 'diagonalUp', color: '#2463eb', style: 'thin' },
    );

    expect(next).not.toBeNull();
    expect(next?.sheets[0]?.config?.borderInfo).toHaveLength(
      MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS + records.length,
    );
    expect(recordReads).toBeLessThan(1_000);
  });

  test('honors later native slash and no-border ranges when indexing renderable diagonals', () => {
    const [sheet] = workbook([
      diagonal(0, 0, 'both'),
      diagonal(0, 1, 'both'),
      diagonal(0, 2, 'both'),
      {
        rangeType: 'range',
        borderType: 'border-slash',
        color: '#778899',
        style: '8',
        range: [{ row: [0, 0], column: [0, 0] }],
      },
      {
        rangeType: 'range',
        borderType: 'border-none',
        color: '#000000',
        style: '1',
        range: [{ row: [0, 0], column: [1, 2] }],
      },
      diagonal(0, 2, 'up', '#2463eb', '13'),
    ]).sheets;
    expect(sheet).toBeDefined();
    if (!sheet) return;

    expect([...spreadsheetRenderableDiagonalBorders(sheet)]).toEqual([
      ['0_0', { up: { color: '#778899', style: '8' } }],
      ['0_2', { up: { color: '#2463eb', style: '13' } }],
    ]);
  });

  test('preserves all OOXML diagonal directions and treats legacy Fortune slashes as diagonal-down', () => {
    const line = { color: '#123456', style: 'thin' as const };
    const borderInfo = fortuneBorderInfoFromXlsxCells([
      {
        row: 0,
        column: 0,
        border: { diagonal: line, diagonalDown: true, diagonalUp: false },
      },
      {
        row: 0,
        column: 1,
        border: { diagonal: line, diagonalDown: false, diagonalUp: true },
      },
      {
        row: 0,
        column: 2,
        border: { diagonal: line, diagonalDown: true, diagonalUp: true },
      },
    ]);
    const sheet = {
      id: 'sheet-1',
      name: 'Sheet 1',
      data: [[{ v: 'Down' }, { v: 'Up' }, { v: 'Both' }]],
      config: { borderInfo },
    } as Sheet;

    expect([...collectXlsxCellBorders(sheet).entries()]).toEqual([
      ['0_0', { diagonal: line, diagonalDown: true, diagonalUp: false }],
      ['0_1', { diagonal: line, diagonalDown: false, diagonalUp: true }],
      ['0_2', { diagonal: line, diagonalDown: true, diagonalUp: true }],
    ]);

    const legacy = {
      ...sheet,
      data: [[{ v: 'Legacy' }]],
      config: {
        borderInfo: [
          {
            rangeType: 'range',
            borderType: 'border-slash',
            color: '#123456',
            style: '1',
            range: [
              {
                row: [0, 0],
                column: [0, 0],
                row_focus: 0,
                column_focus: 0,
              },
            ],
          },
        ],
      },
    } as Sheet;
    expect(collectXlsxCellBorders(legacy).get('0_0')).toEqual({
      diagonal: line,
      diagonalDown: true,
      diagonalUp: false,
    });

    const ordered = {
      ...sheet,
      data: [[{ v: 'Crossed' }]],
      config: {
        borderInfo: [
          diagonal(0, 0, 'up'),
          {
            rangeType: 'range',
            borderType: 'border-slash',
            color: '#778899',
            style: '8',
            range: [{ row: [0, 0], column: [0, 0] }],
          },
        ],
      },
    } as Sheet;
    expect(collectXlsxCellBorders(ordered).get('0_0')).toEqual({
      diagonal: { color: '#778899', style: 'medium' },
      diagonalDown: true,
      diagonalUp: true,
    });
  });

  test('rejects unknown sheets, malformed ranges, colors, and styles', () => {
    const source = workbook();
    expect(
      setSpreadsheetCellBorders(
        source,
        'missing',
        { row: [0, 0], column: [0, 0] },
        { target: 'all', color: '#000000', style: 'thin' },
      ),
    ).toBeNull();
    expect(
      setSpreadsheetCellBorders(
        source,
        'sheet-1',
        { row: [Number.NaN, 0], column: [0, 0] },
        { target: 'all', color: '#000000', style: 'thin' },
      ),
    ).toBeNull();
    expect(
      setSpreadsheetCellBorders(
        source,
        'sheet-1',
        { row: [0, 0], column: [0, 0] },
        { target: 'all', color: 'black', style: 'thin' },
      ),
    ).toBeNull();
    expect(
      setSpreadsheetCellBorders(
        source,
        'sheet-1',
        { row: [0, 0], column: [0, 0] },
        {
          target: 'all',
          color: '#000000',
          style: 'double' as never,
        },
      ),
    ).toBeNull();
  });
});

function workbook(borderInfo?: unknown[]): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        ...(borderInfo ? { config: { borderInfo } } : {}),
      },
    ],
  };
}

function diagonal(
  row: number,
  column: number,
  direction: 'both' | 'down' | 'up',
  color = '#b42318',
  style = '3',
) {
  const line = { color, style };
  return {
    rangeType: 'cell',
    value: {
      row_index: row,
      col_index: column,
      ...(direction === 'down' || direction === 'both' ? { s: line } : {}),
      a3sDiagonal: {
        up: direction === 'up' || direction === 'both',
        down: direction === 'down' || direction === 'both',
        line,
      },
    },
  };
}
