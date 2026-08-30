import { describe, expect, test } from '@rstest/core';
import type { Cell } from '@fortune-sheet/core';
import {
  createSpreadsheetSortDialogSource,
  MAX_SPREADSHEET_SORT_CELLS,
  MAX_SPREADSHEET_SORT_KEYS,
  sortSpreadsheetRows,
  spreadsheetSortFailureMessage,
  validateSpreadsheetSortRequest,
} from '../src/internal/features/work/editors/spreadsheet-sort';

describe('spreadsheet custom sort', () => {
  test('keeps the header and applies stable ordered keys to complete rows', () => {
    const header = [cell('Team'), cell('Score'), cell('Name')];
    const alphaZoe = [cell('Alpha', { bl: 1 }), cell(90), cell('Zoe')];
    const betaAmy = [cell('Beta'), cell(80), cell('Amy')];
    const alphaAmy = [cell('Alpha'), cell(90), cell('Amy', { it: 1 })];
    const betaZoe = [cell('Beta'), cell(95), cell('Zoe')];

    const result = sortSpreadsheetRows(
      [header, alphaZoe, betaAmy, alphaAmy, betaZoe],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 4], column: [0, 2] },
        hasHeader: true,
        keys: [
          { column: 0, direction: 'ascending' },
          { column: 1, direction: 'descending' },
          { column: 2, direction: 'ascending' },
        ],
      },
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows.map((row) => row.map((item) => item?.v))).toEqual([
      ['Team', 'Score', 'Name'],
      ['Alpha', 90, 'Amy'],
      ['Alpha', 90, 'Zoe'],
      ['Beta', 95, 'Zoe'],
      ['Beta', 80, 'Amy'],
    ]);
    expect(result.rows[0]).toBe(header);
    expect(result.rows[1]?.[2]).toBe(alphaAmy[2]);
    expect(result.rows[2]?.[0]).toBe(alphaZoe[0]);
    expect(result.rows[1]?.[2]).toMatchObject({ it: 1 });
    expect(result.rows[2]?.[0]).toMatchObject({ bl: 1 });
  });

  test('sorts numbers naturally, leaves blanks last in either direction, and preserves ties', () => {
    const blank = [cell('blank'), null];
    const tenFirst = [cell('ten-first'), cell(10)];
    const two = [cell('two'), cell(2)];
    const tenSecond = [cell('ten-second'), cell(10)];
    const rows = [blank, tenFirst, two, tenSecond];

    const ascending = sortSpreadsheetRows(rows, {
      sheetId: 'sheet-1',
      range: { row: [0, 3], column: [0, 1] },
      hasHeader: false,
      keys: [{ column: 1, direction: 'ascending' }],
    });
    const descending = sortSpreadsheetRows(rows, {
      sheetId: 'sheet-1',
      range: { row: [0, 3], column: [0, 1] },
      hasHeader: false,
      keys: [{ column: 1, direction: 'descending' }],
    });

    expect(ascending.ok && ascending.rows.map((row) => row[0]?.v)).toEqual([
      'two',
      'ten-first',
      'ten-second',
      'blank',
    ]);
    expect(descending.ok && descending.rows.map((row) => row[0]?.v)).toEqual([
      'ten-first',
      'ten-second',
      'two',
      'blank',
    ]);
  });

  test('moves formulas with rows and translates only relative references', () => {
    const result = sortSpreadsheetRows(
      [
        [cell('Name'), cell('Amount'), cell('Calculated')],
        [cell('Beta'), cell(20), cell(40, { f: '=B2*2+$B$2' })],
        [cell('Alpha'), cell(10), cell(20, { f: '=B3*2+$B$2' })],
      ],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 2], column: [0, 2] },
        hasHeader: true,
        keys: [{ column: 0, direction: 'ascending' }],
      },
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows[1]?.map((item) => item?.v)).toEqual(['Alpha', 10, 20]);
    expect(result.rows[1]?.[2]?.f).toBe('=B2*2+$B$2');
    expect(result.rows[2]?.[2]?.f).toBe('=B3*2+$B$2');
  });

  test('rejects a sort atomically when a translated formula would leave the sheet', () => {
    const result = sortSpreadsheetRows(
      [
        [cell('Zed'), cell(1, { f: '=A1' })],
        [cell('Alpha'), cell(2, { f: '=A1' })],
      ],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 1], column: [0, 1] },
        hasHeader: false,
        keys: [{ column: 0, direction: 'ascending' }],
      },
    );

    expect(result).toMatchObject({
      ok: false,
      code: 'formula-reference-out-of-range',
    });
  });

  test('rejects coordinate-linked cells before reordering any row', () => {
    const result = sortSpreadsheetRows(
      [
        [cell('Zed'), cell('link', { hl: { r: 0, c: 1, id: 'sheet-1' } })],
        [cell('Alpha'), cell('plain')],
      ],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 1], column: [0, 1] },
        hasHeader: false,
        keys: [{ column: 0, direction: 'ascending' }],
      },
    );

    expect(result).toMatchObject({
      ok: false,
      code: 'unsupported-linked-cell',
    });
  });

  test('derives bounded header labels and starts from the active column', () => {
    const source = createSpreadsheetSortDialogSource(
      'sheet-1',
      'Sales',
      {
        range: { row: [0, 2], column: [2, 4] },
        activeColumn: 3,
      },
      [
        [cell('Region'), cell('Revenue'), cell('Owner')],
        [cell('East'), cell(120), cell('Ada')],
        [cell('West'), cell(90), cell('Lin')],
      ],
    );

    expect(source).toMatchObject({
      sheetId: 'sheet-1',
      sheetName: 'Sales',
      rangeReference: 'C1:E3',
      columns: [
        { column: 2, label: 'C（Region）' },
        { column: 3, label: 'D（Revenue）' },
        { column: 4, label: 'E（Owner）' },
      ],
      value: {
        hasHeader: true,
        keys: [{ column: 3, direction: 'ascending' }],
      },
    });
  });

  test('rejects duplicate, out-of-range, oversized, and malformed sort requests', () => {
    const base = {
      sheetId: 'sheet-1',
      range: {
        row: [0, 4] as [number, number],
        column: [0, 2] as [number, number],
      },
      hasHeader: true,
    };
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [
          { column: 0, direction: 'ascending' as const },
          { column: 0, direction: 'descending' as const },
        ],
      }),
    ).toMatchObject({ ok: false, code: 'duplicate-key' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [{ column: 9, direction: 'ascending' }],
      }),
    ).toMatchObject({ ok: false, code: 'column-out-of-range' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: Array.from(
          { length: MAX_SPREADSHEET_SORT_KEYS + 1 },
          (_, column) => ({
            column,
            direction: 'ascending' as const,
          }),
        ),
        range: {
          row: [0, 4],
          column: [0, MAX_SPREADSHEET_SORT_KEYS],
        },
      }),
    ).toMatchObject({ ok: false, code: 'too-many-keys' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [{ column: 0, direction: 'ascending' }],
        range: { row: [0, MAX_SPREADSHEET_SORT_CELLS], column: [0, 1] },
      }),
    ).toMatchObject({ ok: false, code: 'range-too-large' });

    const malformedRows = sortSpreadsheetRows([[cell('A')], [cell('B')]], {
      ...base,
      range: { row: [0, 1], column: [0, 1] },
      hasHeader: false,
      keys: [{ column: 0, direction: 'ascending' }],
    });
    expect(malformedRows).toMatchObject({ ok: false, code: 'invalid-matrix' });
    expect(spreadsheetSortFailureMessage(malformedRows)).toContain(
      '已完整读取的矩形区域',
    );
  });
});

function cell(value: Cell['v'], format: Partial<Cell> = {}): Cell {
  return { v: value, ...format };
}
