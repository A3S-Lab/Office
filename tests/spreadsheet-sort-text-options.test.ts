import type { Cell } from '@fortune-sheet/core';
import { describe, expect, rstest, test } from '@rstest/core';
import { validateSpreadsheetSortRequest } from '../src/internal/features/work/editors/spreadsheet-sort';
import { createSpreadsheetSortTextComparator } from '../src/internal/features/work/editors/spreadsheet-sort-collation';
import { sortSpreadsheetMatrix } from '../src/internal/features/work/editors/spreadsheet-sort-matrix';

describe('spreadsheet text sort options', () => {
  test('switches Chinese text between pinyin and stroke order', () => {
    const rows = ['赵', '阿', '丁', '安', '王'].map((name) => [cell(name)]);
    const pinyin = sortSpreadsheetMatrix(rows, {
      sheetId: 'sheet-1',
      range: { row: [0, 4], column: [0, 0] },
      orientation: 'top-to-bottom',
      hasHeader: false,
      caseSensitive: false,
      textMethod: 'pinyin',
      keys: [{ index: 0, direction: 'ascending' }],
    });
    const stroke = sortSpreadsheetMatrix(rows, {
      sheetId: 'sheet-1',
      range: { row: [0, 4], column: [0, 0] },
      orientation: 'top-to-bottom',
      hasHeader: false,
      caseSensitive: false,
      textMethod: 'stroke',
      keys: [{ index: 0, direction: 'ascending' }],
    });

    expect(pinyin.ok && pinyin.rows.map((row) => row[0]?.v)).toEqual([
      '阿',
      '安',
      '丁',
      '王',
      '赵',
    ]);
    expect(stroke.ok && stroke.rows.map((row) => row[0]?.v)).toEqual([
      '丁',
      '王',
      '安',
      '阿',
      '赵',
    ]);
  });

  test('accepts compatible runtime collation aliases', () => {
    const resolvedOptions = Intl.Collator.prototype.resolvedOptions;
    const resolvedOptionsSpy = rstest
      .spyOn(Intl.Collator.prototype, 'resolvedOptions')
      .mockImplementation(function (this: Intl.Collator) {
        return {
          ...resolvedOptions.call(this),
          collation: 'default',
        };
      });

    try {
      for (const [textMethod, expected] of [
        ['pinyin', ['阿', '安', '丁', '王', '赵']],
        ['stroke', ['丁', '王', '安', '阿', '赵']],
      ] as const) {
        const compare = createSpreadsheetSortTextComparator({
          caseSensitive: false,
          textMethod,
        });

        expect(compare).not.toBeNull();
        expect(
          ['赵', '阿', '丁', '安', '王'].sort(compare ?? undefined),
        ).toEqual(expected);
      }
    } finally {
      resolvedOptionsSpy.mockRestore();
    }
  });

  test('puts lowercase before uppercase only when case sensitivity is enabled', () => {
    const columns = [['A', 'a', 'B', 'b'].map(cell)];
    const insensitive = sortSpreadsheetMatrix(columns, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 3] },
      orientation: 'left-to-right',
      hasHeader: false,
      caseSensitive: false,
      textMethod: 'pinyin',
      keys: [{ index: 0, direction: 'ascending' }],
    });
    const sensitive = sortSpreadsheetMatrix(columns, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 3] },
      orientation: 'left-to-right',
      hasHeader: false,
      caseSensitive: true,
      textMethod: 'pinyin',
      keys: [{ index: 0, direction: 'ascending' }],
    });

    expect(
      insensitive.ok && insensitive.rows[0]?.map((item) => item?.v),
    ).toEqual(['A', 'a', 'B', 'b']);
    expect(sensitive.ok && sensitive.rows[0]?.map((item) => item?.v)).toEqual([
      'a',
      'A',
      'b',
      'B',
    ]);
  });

  test('uses lexical order for numeric text while preserving numeric cell order', () => {
    const text = sortSpreadsheetMatrix(
      ['K2', 'K11', 'K100', 'K1'].map((value) => [cell(value)]),
      {
        sheetId: 'sheet-1',
        range: { row: [0, 3], column: [0, 0] },
        orientation: 'top-to-bottom',
        hasHeader: false,
        caseSensitive: false,
        textMethod: 'pinyin',
        keys: [{ index: 0, direction: 'ascending' }],
      },
    );
    const numbers = sortSpreadsheetMatrix(
      [20, 11, 100, 1].map((value) => [cell(value)]),
      {
        sheetId: 'sheet-1',
        range: { row: [0, 3], column: [0, 0] },
        orientation: 'top-to-bottom',
        hasHeader: false,
        caseSensitive: false,
        textMethod: 'stroke',
        keys: [{ index: 0, direction: 'ascending' }],
      },
    );

    expect(text.ok && text.rows.map((row) => row[0]?.v)).toEqual([
      'K1',
      'K100',
      'K11',
      'K2',
    ]);
    expect(numbers.ok && numbers.rows.map((row) => row[0]?.v)).toEqual([
      1, 11, 20, 100,
    ]);
  });

  test('rejects malformed text options before reading the matrix', () => {
    const base = {
      sheetId: 'sheet-1',
      range: {
        row: [0, 1] as [number, number],
        column: [0, 0] as [number, number],
      },
      orientation: 'top-to-bottom' as const,
      hasHeader: false,
      keys: [{ index: 0, direction: 'ascending' as const }],
    };

    expect(
      validateSpreadsheetSortRequest({
        ...base,
        caseSensitive: 'yes',
        textMethod: 'pinyin',
      } as never),
    ).toMatchObject({ ok: false, code: 'invalid-case-sensitivity' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        caseSensitive: false,
        textMethod: 'radical',
      } as never),
    ).toMatchObject({ ok: false, code: 'invalid-text-method' });
  });

  test('normalizes omitted text options for existing command callers', () => {
    expect(
      validateSpreadsheetSortRequest({
        sheetId: 'sheet-1',
        range: { row: [0, 1], column: [0, 0] },
        orientation: 'top-to-bottom',
        hasHeader: false,
        keys: [{ index: 0, direction: 'ascending' }],
      }),
    ).toMatchObject({
      ok: true,
      request: { caseSensitive: false, textMethod: 'pinyin' },
    });
  });
});

function cell(value: Cell['v']): Cell {
  return { v: value };
}
