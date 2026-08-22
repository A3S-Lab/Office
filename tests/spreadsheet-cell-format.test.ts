import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetCellFormat,
  canApplySpreadsheetCellFormat,
  MAX_SPREADSHEET_CELL_FORMAT_CELLS,
} from '../src/internal/features/work/editors/spreadsheet-cell-format';
import { setSpreadsheetCellBordersPerCell } from '../src/internal/features/work/editors/spreadsheet-cell-border-per-cell';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet cell format', () => {
  test('applies mixed format families to matrix cells without losing content', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 2,
          column: 3,
          data: [
            [
              {
                v: 12.5,
                m: '12.5',
                f: '=25/2',
                ct: { fa: 'General', t: 'n' },
                ps: {
                  left: null,
                  top: null,
                  width: null,
                  height: null,
                  value: 'Keep comment',
                  isShow: false,
                },
              },
              {
                v: 'A3S',
                m: 'A3S',
                ct: { fa: '@', t: 's' },
                hl: { r: 0, c: 1, id: 'https://a3s.dev' },
              },
              { v: true, m: 'TRUE', ct: { fa: 'General', t: 'b' } },
            ],
          ],
        },
      ],
    } satisfies WorkSpreadsheetContent;

    const next = applySpreadsheetCellFormat(content, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 2] },
      patch: {
        numberFormat: '"Value "General',
        horizontalAlignment: 'center',
        verticalAlignment: 'middle',
        wrapText: true,
        rotation: -30,
        fontFamily: 'Arial',
        fontSize: 12,
        fontColor: '#2463eb',
        bold: true,
        italic: true,
        underline: 'doubleAccounting',
        strike: true,
        fillColor: '#fff2cc',
        borders: [
          { target: 'top', color: '#172033', style: 'medium' },
          { target: 'diagonalDown', color: '#d84b4f', style: 'dashed' },
          { target: 'diagonalUp', color: '#d84b4f', style: 'dashed' },
        ],
        locked: false,
        hidden: true,
      },
    });

    expect(next).not.toBeNull();
    expect(content.sheets[0]?.data?.[0]?.[0]).toMatchObject({
      m: '12.5',
      ct: { fa: 'General', t: 'n' },
    });
    expect(next?.sheets[0]?.data?.[0]?.map((cell) => cell?.ct?.t)).toEqual([
      'n',
      's',
      'b',
    ]);
    expect(next?.sheets[0]?.data?.[0]?.[0]).toMatchObject({
      v: 12.5,
      f: '=25/2',
      ht: '0',
      vt: 0,
      tb: '2',
      rt: 120,
      ff: 'Arial',
      fs: 12,
      fc: '#2463eb',
      bl: 1,
      it: 1,
      un: 4,
      cl: 1,
      bg: '#fff2cc',
      ps: expect.objectContaining({ value: 'Keep comment' }),
    });
    expect(next?.sheets[0]?.data?.[0]?.[0]?.m).toBeUndefined();
    expect(next?.sheets[0]?.data?.[0]?.[1]?.hl).toEqual({
      r: 0,
      c: 1,
      id: 'https://a3s.dev',
    });
    expect(next?.sheets[0]?.config?.borderInfo).toHaveLength(4);
    expect(next?.sheets[0]?.config?.borderInfo?.[0]).toMatchObject({
      rangeType: 'range',
      borderType: 'border-none',
    });
    expect(
      (
        next?.sheets[0]?.config?.authority as {
          cellProtectionRanges?: unknown[];
        }
      )?.cellProtectionRanges,
    ).toEqual([
      {
        range: { row: [0, 0], column: [0, 2] },
        locked: false,
        hidden: true,
      },
    ]);
  });

  test('switches numeric and stacked orientations without stale Fortune fields', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          data: [[{ v: 'A3S', rt: 45, tr: '1' }]],
        },
      ],
    } satisfies WorkSpreadsheetContent;

    const stacked = applySpreadsheetCellFormat(content, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 0] },
      patch: { textOrientation: 'vertical' },
    });
    expect(stacked?.sheets[0]?.data?.[0]?.[0]).toMatchObject({
      tr: '3',
      v: 'A3S',
    });
    expect(stacked?.sheets[0]?.data?.[0]?.[0]?.rt).toBeUndefined();

    const rotated = stacked
      ? applySpreadsheetCellFormat(stacked, {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [0, 0] },
          patch: { rotation: -30 },
        })
      : null;
    expect(rotated?.sheets[0]?.data?.[0]?.[0]).toMatchObject({ rt: 120 });
    expect(rotated?.sheets[0]?.data?.[0]?.[0]?.tr).toBeUndefined();

    expect(
      canApplySpreadsheetCellFormat(content, {
        sheetId: 'sheet-1',
        range: { row: [0, 0], column: [0, 0] },
        patch: { rotation: 0, textOrientation: 'horizontal' },
      }),
    ).toBe(false);
  });

  test('preserves celldata storage and materializes only formatted blank cells', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 4,
          column: 4,
          celldata: [
            { r: 1, c: 1, v: { v: 'Anchor', bg: '#ff0000' } },
            { r: 3, c: 3, v: { v: 'Outside', fc: '#123456' } },
          ],
        },
      ],
    } satisfies WorkSpreadsheetContent;

    const next = applySpreadsheetCellFormat(content, {
      sheetId: 'sheet-1',
      range: { row: [1, 1], column: [1, 2] },
      patch: { fillColor: '#d9ead3', bold: true },
    });

    expect(next).not.toBeNull();
    expect(next?.sheets[0]).not.toHaveProperty('data');
    expect(next?.sheets[0]?.celldata).toEqual([
      {
        r: 1,
        c: 1,
        v: { v: 'Anchor', bg: '#d9ead3', bl: 1 },
      },
      { r: 1, c: 2, v: { bg: '#d9ead3', bl: 1 } },
      { r: 3, c: 3, v: { v: 'Outside', fc: '#123456' } },
    ]);
    expect(content.sheets[0]?.celldata).toHaveLength(2);
  });

  test('does not materialize a blank cell when removing a property', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 2,
          column: 2,
          celldata: [{ r: 0, c: 0, v: { v: 'Keep' } }],
        },
      ],
    } satisfies WorkSpreadsheetContent;

    const next = applySpreadsheetCellFormat(content, {
      sheetId: 'sheet-1',
      range: { row: [1, 1], column: [1, 1] },
      patch: { fillColor: null, horizontalAlignment: 'general' },
    });

    expect(next?.sheets[0]?.celldata).toEqual(content.sheets[0]?.celldata);
    expect(next?.sheets[0]).not.toHaveProperty('data');
  });

  test('keeps untouched mixed protection values while compacting the target', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 2,
          column: 2,
          data: [
            [{ v: 'A' }, { v: 'B' }],
            [{ v: 'C' }, { v: 'D' }],
          ],
          config: {
            authority: {
              cellProtectionRanges: [
                {
                  range: { row: [0, 0], column: [0, 1] },
                  locked: true,
                  hidden: true,
                },
                {
                  range: { row: [1, 1], column: [0, 1] },
                  locked: true,
                  hidden: false,
                },
              ],
            },
          },
        },
      ],
    } satisfies WorkSpreadsheetContent;

    const next = applySpreadsheetCellFormat(content, {
      sheetId: 'sheet-1',
      range: { row: [0, 1], column: [0, 1] },
      patch: { locked: false },
    });
    const authority = next?.sheets[0]?.config?.authority as {
      cellProtectionRanges?: unknown[];
    };

    expect(authority.cellProtectionRanges).toEqual([
      {
        range: { row: [0, 0], column: [0, 1] },
        locked: false,
        hidden: true,
      },
      {
        range: { row: [1, 1], column: [0, 1] },
        locked: false,
        hidden: false,
      },
    ]);
  });

  test('rejects oversized, invalid, and diagonal-border requests', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sheet 1', row: 20_000, column: 2 }],
    } satisfies WorkSpreadsheetContent;
    const oversized = {
      sheetId: 'sheet-1',
      range: {
        row: [0, MAX_SPREADSHEET_CELL_FORMAT_CELLS],
        column: [0, 0],
      },
      patch: { bold: true },
    } as const;
    const diagonal = {
      sheetId: 'sheet-1',
      range: { row: [0, 4_096], column: [0, 0] },
      patch: {
        borders: [{ target: 'diagonalDown', color: '#172033', style: 'thin' }],
      },
    } as const;

    expect(canApplySpreadsheetCellFormat(content, oversized)).toBe(false);
    expect(applySpreadsheetCellFormat(content, oversized)).toBeNull();
    expect(canApplySpreadsheetCellFormat(content, diagonal)).toBe(false);
    expect(
      setSpreadsheetCellBordersPerCell(
        content,
        'sheet-1',
        diagonal.range,
        diagonal.patch.borders,
      ),
    ).toBeNull();
    expect(
      canApplySpreadsheetCellFormat(content, {
        sheetId: 'sheet-1',
        range: { row: [0, 0], column: [0, 0] },
        patch: { fontSize: 0 },
      }),
    ).toBe(false);
  });
});
