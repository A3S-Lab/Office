import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  createSpreadsheetSortDialogSource,
  createSpreadsheetSortRangePlan,
  MAX_SPREADSHEET_SORT_CELLS,
  MAX_SPREADSHEET_SORT_KEYS,
  spreadsheetSortFailureMessage,
  spreadsheetSortRowsFromSheet,
  spreadsheetSortRowsMatchRange,
  validateSpreadsheetSortRequest,
} from '../src/internal/features/work/editors/spreadsheet-sort';
import { sortSpreadsheetMatrix } from '../src/internal/features/work/editors/spreadsheet-sort-matrix';

describe('spreadsheet custom sort', () => {
  test('plans a bounded WPS expansion from dense and sparse current regions', () => {
    const dense = createSpreadsheetSortRangePlan(
      {
        id: 'dense',
        name: 'Dense',
        data: [
          [{ v: 'Title' }],
          [],
          [{ v: 'Name' }, { v: 'Owner' }, { v: 'Score' }],
          [{ v: 'Alpha' }, { v: 'Ada' }, { v: 90 }],
          [{ v: 'Beta' }, { v: 'Lin' }, { v: 80 }],
        ],
      },
      { row: [3, 3], column: [1, 1] },
    );

    expect(dense).toEqual({
      selectedRange: { row: [3, 3], column: [1, 1] },
      expandedRange: { row: [2, 4], column: [0, 2] },
    });

    const sparse = createSpreadsheetSortRangePlan(
      {
        id: 'sparse',
        name: 'Sparse',
        row: 1_048_576,
        column: 16_384,
        celldata: [
          { r: 8, c: 4, v: { v: 'Name' } },
          { r: 8, c: 5, v: { v: 'Score' } },
          { r: 9, c: 4, v: { v: 'Alpha' } },
          { r: 9, c: 5, v: { v: 90 } },
          { r: 1_048_575, c: 16_383, v: { v: 'Tail' } },
        ],
      },
      { row: [9, 9], column: [5, 5] },
    );

    expect(sparse).toEqual({
      selectedRange: { row: [9, 9], column: [5, 5] },
      expandedRange: { row: [8, 9], column: [4, 5] },
    });
  });

  test('does not warn when the selection already owns its current region', () => {
    expect(
      createSpreadsheetSortRangePlan(
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          data: [
            [{ v: 'Name' }, { v: 'Score' }],
            [{ v: 'Alpha' }, { v: 90 }],
            [{ v: 'Beta' }, { v: 80 }],
          ],
        },
        { row: [0, 2], column: [0, 1] },
      ),
    ).toEqual({
      selectedRange: { row: [0, 2], column: [0, 1] },
      expandedRange: null,
    });
  });

  test('materializes bounded controlled rows when a filtered native view omits them', () => {
    const range = { row: [1, 2], column: [1, 2] } as const;
    const rows = spreadsheetSortRowsFromSheet(
      {
        id: 'filtered',
        name: 'Filtered',
        data: [
          [{ v: 'Outside' }],
          [null, { v: 'Beta' }],
          [null, null, { v: 90 }],
        ],
        celldata: [{ r: 1, c: 2, v: { v: 80 } }],
      },
      range,
    );

    expect(rows).toEqual([
      [{ v: 'Beta' }, { v: 80 }],
      [null, { v: 90 }],
    ]);
    expect(spreadsheetSortRowsMatchRange(rows ?? [], range)).toBe(true);
    expect(spreadsheetSortRowsMatchRange([rows?.[0] ?? []], range)).toBe(false);
  });

  test('expands selections to the exact AutoFilter or table-owned sortable range', () => {
    const autoFilter = createSpreadsheetSortRangePlan(
      {
        id: 'filtered',
        name: 'Filtered',
        data: [
          [{ v: 'Name' }, { v: 'Score' }],
          [{ v: 'Beta' }, { v: 80 }],
          [{ v: 'Alpha' }, { v: 90 }],
        ],
        filter_select: { row: [0, 2], column: [0, 1] },
      },
      { row: [1, 1], column: [1, 1] },
    );
    const table = createSpreadsheetSortRangePlan(
      {
        id: 'table',
        name: 'Table',
        data: [
          [{ v: 'Name' }, { v: 'Score' }],
          [{ v: 'Beta' }, { v: 80 }],
          [{ v: 'Alpha' }, { v: 90 }],
          [{ v: 'Total' }, { v: 170 }],
        ],
        tables: [
          {
            id: 'table-1',
            name: 'Table1',
            range: { row: [0, 3], column: [0, 1] },
            columns: [{ name: 'Name' }, { name: 'Score' }],
            filters: [],
            headerRow: true,
            totalsRow: true,
            style: { family: 'medium', number: 2 },
            showFirstColumn: false,
            showLastColumn: false,
            showRowStripes: true,
            showColumnStripes: false,
          },
        ],
      },
      { row: [2, 2], column: [0, 0] },
    );

    expect(autoFilter).toEqual({
      selectedRange: { row: [1, 1], column: [1, 1] },
      expandedRange: { row: [0, 2], column: [0, 1] },
    });
    expect(table).toEqual({
      selectedRange: { row: [2, 2], column: [0, 0] },
      expandedRange: { row: [0, 2], column: [0, 1] },
    });
  });

  test('keeps the header and applies stable ordered keys to complete rows', () => {
    const header = [cell('Team'), cell('Score'), cell('Name')];
    const alphaZoe = [cell('Alpha', { bl: 1 }), cell(90), cell('Zoe')];
    const betaAmy = [cell('Beta'), cell(80), cell('Amy')];
    const alphaAmy = [cell('Alpha'), cell(90), cell('Amy', { it: 1 })];
    const betaZoe = [cell('Beta'), cell(95), cell('Zoe')];

    const result = sortSpreadsheetMatrix(
      [header, alphaZoe, betaAmy, alphaAmy, betaZoe],
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 4], column: [0, 2] },
        hasHeader: true,
        keys: [
          { index: 0, direction: 'ascending' },
          { index: 1, direction: 'descending' },
          { index: 2, direction: 'ascending' },
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

    const ascending = sortSpreadsheetMatrix(rows, {
      sheetId: 'sheet-1',
      orientation: 'top-to-bottom',
      range: { row: [0, 3], column: [0, 1] },
      hasHeader: false,
      keys: [{ index: 1, direction: 'ascending' }],
    });
    const descending = sortSpreadsheetMatrix(rows, {
      sheetId: 'sheet-1',
      orientation: 'top-to-bottom',
      range: { row: [0, 3], column: [0, 1] },
      hasHeader: false,
      keys: [{ index: 1, direction: 'descending' }],
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

  test('sorts listed values first, falls back naturally, and composes with later keys', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Task'), cell('Status')],
        [cell('Beta'), cell('进行中')],
        [cell('Alpha'), cell('已完成')],
        [cell('Zeta'), cell('有风险')],
        [cell('Delta'), cell('待确认')],
        [cell('Epsilon'), null],
        [cell('Gamma'), cell('有风险')],
      ],
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 6], column: [0, 1] },
        hasHeader: true,
        keys: [
          {
            index: 1,
            customList: ['有风险', '进行中', '正常', '已完成'],
          },
          { index: 0, direction: 'ascending' },
        ],
      },
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows.map((row) => row.map((item) => item?.v))).toEqual([
      ['Task', 'Status'],
      ['Gamma', '有风险'],
      ['Zeta', '有风险'],
      ['Beta', '进行中'],
      ['Alpha', '已完成'],
      ['Delta', '待确认'],
      ['Epsilon', undefined],
    ]);
  });

  test('sorts matching appearances top or bottom before later stable keys', () => {
    const rows = [
      [cell('Task'), cell('Status')],
      [cell('Blank red'), cell(null, { bg: '#ff0000' })],
      [cell('Zulu red'), cell('Blocked', { bg: '#f00' })],
      [cell('Alpha blue'), cell('Ready', { bg: '#0000ff' })],
      [cell('Beta plain'), cell('Ready')],
      [cell('Omega red'), cell('Blocked', { bg: '#ff0000' })],
    ];
    const result = sortSpreadsheetMatrix(rows, {
      sheetId: 'sheet-1',
      orientation: 'top-to-bottom',
      range: { row: [0, 5], column: [0, 1] },
      hasHeader: true,
      keys: [
        {
          index: 1,
          sortOn: 'cell-color',
          color: '#FF0000',
          position: 'first',
        },
        { index: 0, direction: 'ascending' },
      ],
    });

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows.map((row) => row[0]?.v)).toEqual([
      'Task',
      'Blank red',
      'Omega red',
      'Zulu red',
      'Alpha blue',
      'Beta plain',
    ]);

    const automaticFontLast = sortSpreadsheetMatrix(rows, {
      sheetId: 'sheet-1',
      orientation: 'top-to-bottom',
      range: { row: [0, 5], column: [0, 1] },
      hasHeader: true,
      keys: [
        {
          index: 1,
          sortOn: 'font-color',
          color: null,
          position: 'last',
        },
        { index: 0, direction: 'ascending' },
      ],
    });
    expect(
      automaticFontLast.ok && automaticFontLast.rows.map((row) => row[0]?.v),
    ).toEqual([
      'Task',
      'Alpha blue',
      'Beta plain',
      'Blank red',
      'Omega red',
      'Zulu red',
    ]);
  });

  test('sorts a computed conditional icon target from a supplied live snapshot', () => {
    const rows = [
      [cell('Task'), cell('Score')],
      [cell('Low'), cell(10)],
      [cell('High B'), cell(30)],
      [cell('Middle'), cell(20)],
      [cell('High A'), cell(40)],
    ];
    const appearances = rows.map((row, rowIndex) =>
      row.map((_, columnIndex) => ({
        cellColor: null,
        fontColor: null,
        icon:
          rowIndex === 0 || columnIndex === 0
            ? null
            : {
                iconSet: '3TrafficLights1' as const,
                index: rowIndex === 1 ? 0 : rowIndex === 3 ? 1 : 2,
              },
      })),
    );
    const result = sortSpreadsheetMatrix(
      rows,
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 4], column: [0, 1] },
        hasHeader: true,
        keys: [
          {
            index: 1,
            sortOn: 'icon',
            icon: { iconSet: '3TrafficLights1', index: 2 },
            position: 'first',
          },
          { index: 0, direction: 'ascending' },
        ],
      },
      appearances,
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows.map((row) => row[0]?.v)).toEqual([
      'Task',
      'High A',
      'High B',
      'Low',
      'Middle',
    ]);
  });

  test('rejects an appearance target missing from the live snapshot', () => {
    const rows = [
      [cell('Task'), cell('Status')],
      [cell('Alpha'), cell('Ready')],
      [cell('Beta'), cell('Blocked')],
    ];
    const appearances = rows.map((row) =>
      row.map(() => ({ cellColor: null, fontColor: null, icon: null })),
    );

    expect(
      sortSpreadsheetMatrix(
        rows,
        {
          sheetId: 'sheet-1',
          orientation: 'top-to-bottom',
          range: { row: [0, 2], column: [0, 1] },
          hasHeader: true,
          keys: [
            {
              index: 1,
              sortOn: 'cell-color',
              color: '#fce8e6',
              position: 'first',
            },
          ],
        },
        appearances,
      ),
    ).toMatchObject({ ok: false, code: 'invalid-appearance' });
  });

  test('composes distinct appearance priorities on the same column', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Task'), cell('Status')],
        [cell('Blue'), cell('Ready', { bg: '#4472c4' })],
        [cell('Plain'), cell('Waiting')],
        [cell('Yellow'), cell('Review', { bg: '#fff2cc' })],
        [cell('Red B'), cell('Blocked', { bg: '#fce8e6' })],
        [cell('Red A'), cell('Blocked', { bg: '#fce8e6' })],
      ],
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 5], column: [0, 1] },
        hasHeader: true,
        keys: [
          {
            index: 1,
            sortOn: 'cell-color',
            color: '#fce8e6',
            position: 'first',
          },
          {
            index: 1,
            sortOn: 'cell-color',
            color: '#fff2cc',
            position: 'first',
          },
          {
            index: 1,
            sortOn: 'cell-color',
            color: '#4472c4',
            position: 'last',
          },
          { index: 0, direction: 'ascending' },
        ],
      },
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows.map((row) => row[0]?.v)).toEqual([
      'Task',
      'Red A',
      'Red B',
      'Yellow',
      'Plain',
      'Blue',
    ]);
  });

  test('moves formulas with rows and translates only relative references', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Name'), cell('Amount'), cell('Calculated')],
        [cell('Beta'), cell(20), cell(40, { f: '=B2*2+$B$2' })],
        [cell('Alpha'), cell(10), cell(20, { f: '=B3*2+$B$2' })],
      ],
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 2], column: [0, 2] },
        hasHeader: true,
        keys: [{ index: 0, direction: 'ascending' }],
      },
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows[1]?.map((item) => item?.v)).toEqual(['Alpha', 10, 20]);
    expect(result.rows[1]?.[2]?.f).toBe('=B2*2+$B$2');
    expect(result.rows[2]?.[2]?.f).toBe('=B3*2+$B$2');
  });

  test('rejects a sort atomically when a translated formula would leave the sheet', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Zed'), cell(1, { f: '=A1' })],
        [cell('Alpha'), cell(2, { f: '=A1' })],
      ],
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 1], column: [0, 1] },
        hasHeader: false,
        keys: [{ index: 0, direction: 'ascending' }],
      },
    );

    expect(result).toMatchObject({
      ok: false,
      code: 'formula-reference-out-of-range',
    });
  });

  test('rejects coordinate-linked cells before reordering any row', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Zed'), cell('link', { hl: { r: 0, c: 1, id: 'sheet-1' } })],
        [cell('Alpha'), cell('plain')],
      ],
      {
        sheetId: 'sheet-1',
        orientation: 'top-to-bottom',
        range: { row: [0, 1], column: [0, 1] },
        hasHeader: false,
        keys: [{ index: 0, direction: 'ascending' }],
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
        activeRow: 1,
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
        { index: 2, label: 'C（Region）' },
        { index: 3, label: 'D（Revenue）' },
        { index: 4, label: 'E（Owner）' },
      ],
      value: {
        orientation: 'top-to-bottom',
        hasHeader: true,
        keys: [{ index: 3, direction: 'ascending' }],
      },
    });
  });

  test('does not infer headers from formatted values or formula results', () => {
    const formatted = createSpreadsheetSortDialogSource(
      'sheet-1',
      'Sales',
      {
        range: { row: [3, 6], column: [5, 5] },
        activeColumn: 5,
        activeRow: 3,
      },
      [
        [{ v: 0.67, m: '67%' }],
        [{ v: 0.47, m: '47%' }],
        [{ v: 0.87, m: '87%' }],
        [{ v: 1, m: '100%' }],
      ],
    );
    const formulas = createSpreadsheetSortDialogSource(
      'sheet-1',
      'Sales',
      {
        range: { row: [3, 6], column: [5, 5] },
        activeColumn: 5,
        activeRow: 3,
      },
      [
        [{ f: '=SUM(C4:E4)/3', v: 0.67, m: '67%' }],
        [{ f: '=AVERAGE(C5:E5)', v: 0.47, m: '47%' }],
        [{ f: '=AVERAGE(C6:E6)', v: 0.87, m: '87%' }],
        [{ f: '=AVERAGE(C7:E7)', v: 1, m: '100%' }],
      ],
    );

    expect(formatted?.value.hasHeader).toBe(false);
    expect(formatted?.columns).toEqual([{ index: 5, label: 'F' }]);
    expect(formulas?.value.hasHeader).toBe(false);
    expect(formulas?.columns).toEqual([{ index: 5, label: 'F' }]);
  });

  test('locks structural headers even when value heuristics would not detect them', () => {
    const source = createSpreadsheetSortDialogSource(
      'sheet-1',
      'Filtered',
      {
        range: { row: [0, 2], column: [0, 0] },
        activeColumn: 0,
        activeRow: 1,
        scope: { kind: 'auto-filter', hasHeader: true },
      },
      [[cell(10)], [cell(5)], [cell(7)]],
    );

    expect(source).toMatchObject({
      scope: { kind: 'auto-filter', hasHeader: true },
      value: { hasHeader: true, orientation: 'top-to-bottom' },
    });
  });

  test('rejects duplicate, out-of-range, oversized, and malformed sort requests', () => {
    const base = {
      sheetId: 'sheet-1',
      range: {
        row: [0, 4] as [number, number],
        column: [0, 2] as [number, number],
      },
      orientation: 'top-to-bottom' as const,
      hasHeader: true,
    };
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [
          { index: 0, direction: 'ascending' as const },
          { index: 0, direction: 'descending' as const },
        ],
      }),
    ).toMatchObject({ ok: false, code: 'duplicate-key' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        scope: {
          kind: 'table',
          tableId: '',
          hasHeader: true,
        },
        keys: [{ index: 0, direction: 'ascending' }],
      }),
    ).toMatchObject({ ok: false, code: 'invalid-scope' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [
          {
            index: 0,
            sortOn: 'cell-color',
            color: '#fce8e6',
            position: 'first',
          },
          {
            index: 0,
            sortOn: 'cell-color',
            color: '#fff2cc',
            position: 'first',
          },
        ],
      }),
    ).toMatchObject({ ok: true });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [
          {
            index: 0,
            sortOn: 'cell-color',
            color: '#fce8e6',
            position: 'first',
          },
          {
            index: 0,
            sortOn: 'cell-color',
            color: '#fce8e6',
            position: 'last',
          },
        ],
      }),
    ).toMatchObject({ ok: false, code: 'duplicate-key' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [{ index: 9, direction: 'ascending' }],
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
        keys: [{ index: 0, direction: 'ascending' }],
        range: { row: [0, MAX_SPREADSHEET_SORT_CELLS], column: [0, 1] },
      }),
    ).toMatchObject({ ok: false, code: 'range-too-large' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [{ index: 0, customList: ['High', ' high ', 'Low'] }],
      }),
    ).toMatchObject({ ok: false, code: 'invalid-custom-list' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [
          {
            index: 0,
            sortOn: 'cell-color',
            color: 'not-a-color',
            position: 'first',
          },
        ],
      }),
    ).toMatchObject({ ok: false, code: 'invalid-appearance' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [
          {
            index: 0,
            sortOn: 'icon',
            icon: { iconSet: '3TrafficLights1', index: 9 },
            position: 'first',
          },
        ],
      }),
    ).toMatchObject({ ok: false, code: 'invalid-appearance' });

    const malformedRows = sortSpreadsheetMatrix([[cell('A')], [cell('B')]], {
      ...base,
      range: { row: [0, 1], column: [0, 1] },
      hasHeader: false,
      keys: [{ index: 0, direction: 'ascending' }],
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
