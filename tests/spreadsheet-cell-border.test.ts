import { describe, expect, test } from '@rstest/core';
import {
  MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS,
  setSpreadsheetCellBorders,
} from '../src/internal/features/work/editors/spreadsheet-cell-border';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

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

  test('emits native diagonal records per cell and rejects explosive ranges', () => {
    const source = workbook();
    const next = setSpreadsheetCellBorders(
      source,
      'sheet-1',
      { row: [1, 2], column: [3, 4] },
      { target: 'diagonal', color: '#b42318', style: 'dotted' },
    );

    expect(next?.sheets[0]?.config?.borderInfo).toEqual([
      diagonal(1, 3),
      diagonal(1, 4),
      diagonal(2, 3),
      diagonal(2, 4),
    ]);

    expect(
      setSpreadsheetCellBorders(
        source,
        'sheet-1',
        {
          row: [0, MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS],
          column: [0, 0],
        },
        { target: 'diagonal', color: '#b42318', style: 'dotted' },
      ),
    ).toBeNull();
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

function diagonal(row: number, column: number) {
  return {
    rangeType: 'range',
    borderType: 'border-slash',
    color: '#b42318',
    style: '3',
    range: [
      {
        row: [row, row],
        column: [column, column],
        row_focus: row,
        column_focus: column,
      },
    ],
  };
}
