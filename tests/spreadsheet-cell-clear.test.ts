import type { Cell, Sheet } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import { clearSpreadsheetSheetSelection } from '../src/internal/features/work/editors/spreadsheet-cell-clear';

const selectedCell = { row: [1, 1], column: [1, 1] };

describe('spreadsheet cell clear', () => {
  test('clears contents while retaining formatting, comments, hyperlinks, and merges', () => {
    const sheet = clearSpreadsheetSheetSelection(
      richSheet(),
      selectedCell,
      'contents',
    );

    expect(sheet.data?.[1]?.[1]).toEqual({
      bg: '#fff2cc',
      bl: 1,
      ct: { fa: '@', t: 'inlineStr' },
      fc: '#0070c0',
      ff: 'Arial',
      fs: 14,
      hi: 1,
      hl: { r: 1, c: 1, id: 'sheet-1' },
      lo: 0,
      mc: { r: 1, c: 1, rs: 1, cs: 1 },
      ps: {
        left: null,
        top: null,
        width: null,
        height: null,
        value: 'Keep this comment',
        isShow: false,
      },
      un: 1,
    });
    expect(sheet.hyperlink).toEqual({
      '0_0': { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
      '1_1': { linkType: 'webpage', linkAddress: 'https://wps.com' },
    });
  });

  test('clears direct, border, conditional, and alternating formats without changing content', () => {
    const sheet = clearSpreadsheetSheetSelection(
      richSheet(),
      selectedCell,
      'formats',
    );

    expect(sheet.data?.[1]?.[1]).toEqual({
      f: '=A1',
      hl: { r: 1, c: 1, id: 'sheet-1' },
      m: '0.25',
      mc: { r: 1, c: 1, rs: 1, cs: 1 },
      ps: {
        left: null,
        top: null,
        width: null,
        height: null,
        value: 'Keep this comment',
        isShow: false,
      },
      qp: 1,
      spl: { anchor: 'A1' },
      v: 0.25,
    });
    expect(sheet.config?.borderInfo).toEqual([
      expect.objectContaining({
        range: [
          { row: [0, 0], column: [0, 2] },
          { row: [2, 2], column: [0, 2] },
          { row: [1, 1], column: [0, 0] },
          { row: [1, 1], column: [2, 2] },
        ],
      }),
    ]);
    expect(sheet.luckysheet_conditionformat_save).toEqual([
      expect.objectContaining({
        cellrange: [
          { row: [0, 0], column: [0, 2] },
          { row: [2, 2], column: [0, 2] },
          { row: [1, 1], column: [0, 0] },
          { row: [1, 1], column: [2, 2] },
        ],
      }),
    ]);
    expect(sheet.luckysheet_alternateformat_save).toEqual([
      expect.objectContaining({ cellrange: { row: [0, 0], column: [0, 2] } }),
      expect.objectContaining({ cellrange: { row: [2, 2], column: [0, 2] } }),
      expect.objectContaining({ cellrange: { row: [1, 1], column: [0, 0] } }),
      expect.objectContaining({ cellrange: { row: [1, 1], column: [2, 2] } }),
    ]);
  });

  test('clears comments and hyperlinks independently', () => {
    const withoutComments = clearSpreadsheetSheetSelection(
      richSheet(),
      selectedCell,
      'comments',
    );
    expect(withoutComments.data?.[1]?.[1]?.ps).toBeUndefined();
    expect(withoutComments.data?.[1]?.[1]?.v).toBe(0.25);
    expect(withoutComments.data?.[1]?.[1]?.hl).toEqual({
      r: 1,
      c: 1,
      id: 'sheet-1',
    });

    const withoutHyperlinks = clearSpreadsheetSheetSelection(
      richSheet(),
      selectedCell,
      'hyperlinks',
    );
    expect(withoutHyperlinks.data?.[1]?.[1]?.hl).toBeUndefined();
    expect(withoutHyperlinks.data?.[1]?.[1]?.ps?.value).toBe(
      'Keep this comment',
    );
    expect(withoutHyperlinks.data?.[1]?.[1]?.v).toBe(0.25);
    expect(withoutHyperlinks.hyperlink).toEqual({
      '0_0': { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
    });
  });

  test('clears all cell state while retaining the merge geometry', () => {
    const sheet = clearSpreadsheetSheetSelection(
      richSheet(),
      selectedCell,
      'all',
    );

    expect(sheet.data?.[1]?.[1]).toEqual({
      mc: { r: 1, c: 1, rs: 1, cs: 1 },
    });
    expect(sheet.hyperlink).toEqual({
      '0_0': { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
    });
    expect(sheet.config?.borderInfo).toHaveLength(1);
    expect(sheet.luckysheet_conditionformat_save).toHaveLength(1);
  });

  test('retains vendor formatting ranges it cannot safely interpret', () => {
    const source = richSheet();
    const border = source.config?.borderInfo?.[0] as unknown as {
      range: unknown[];
    };
    border.range.push('vendor-border-range');
    const conditional = source
      .luckysheet_conditionformat_save?.[0] as unknown as {
      cellrange: unknown[];
    };
    conditional.cellrange.push('vendor-condition-range');

    const sheet = clearSpreadsheetSheetSelection(
      source,
      selectedCell,
      'formats',
    );

    expect(
      (sheet.config?.borderInfo?.[0] as unknown as { range: unknown[] }).range,
    ).toContain('vendor-border-range');
    expect(
      (
        sheet.luckysheet_conditionformat_save?.[0] as unknown as {
          cellrange: unknown[];
        }
      ).cellrange,
    ).toContain('vendor-condition-range');
  });
});

function richSheet(): Sheet {
  const cell = {
    v: 0.25,
    m: '25%',
    f: '=A1',
    spl: { anchor: 'A1' },
    qp: 1,
    ct: {
      fa: '@',
      t: 'inlineStr',
      s: [{ v: 'A3S' }, { v: ' Office' }],
    },
    bg: '#fff2cc',
    bl: 1,
    ff: 'Arial',
    fs: 14,
    fc: '#0070c0',
    un: 1,
    lo: 0,
    hi: 1,
    mc: { r: 1, c: 1, rs: 1, cs: 1 },
    ps: {
      left: null,
      top: null,
      width: null,
      height: null,
      value: 'Keep this comment',
      isShow: false,
    },
    hl: { r: 1, c: 1, id: 'sheet-1' },
  } satisfies Cell & { hi?: number };
  return {
    id: 'sheet-1',
    name: 'Sheet 1',
    data: [
      [{ v: 'outside' }, null, null],
      [null, cell, null],
      [null, null, null],
    ],
    config: {
      merge: { '1_1': { r: 1, c: 1, rs: 1, cs: 1 } },
      borderInfo: [
        {
          rangeType: 'range',
          borderType: 'border-all',
          color: '#000000',
          range: [{ row: [0, 2], column: [0, 2] }],
        },
      ],
    },
    hyperlink: {
      '0_0': { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
      '1_1': { linkType: 'webpage', linkAddress: 'https://wps.com' },
    },
    luckysheet_conditionformat_save: [
      {
        type: 'default',
        conditionName: 'greaterThan',
        conditionValue: ['0'],
        format: { textColor: '#9c0006', cellColor: '#ffc7ce' },
        cellrange: [{ row: [0, 2], column: [0, 2] }],
      },
    ],
    luckysheet_alternateformat_save: [
      {
        format: { head: { fc: '#ffffff' } },
        cellrange: { row: [0, 2], column: [0, 2] },
      },
    ],
  };
}
