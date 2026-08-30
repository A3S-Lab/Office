import { describe, expect, test } from '@rstest/core';
import {
  canApplySpreadsheetTableStructureChange,
  reconcileSpreadsheetTablesAfterFortune,
} from '../src/internal/features/work/editors/spreadsheet-table-reconciliation';
import type {
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
} from '../src/internal/features/work/work-types';

describe('spreadsheet table reconciliation', () => {
  test('canonicalizes an edited table header in the controlled model', () => {
    const source = tableSheet();
    const changed = structuredClone(source);
    const header = changed.data?.[0];
    if (!header) throw new Error('Expected table fixture header data.');
    header[1] = { m: 'Region', v: 'Region' };

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'replace',
          path: ['data', 0, 1],
          value: changed.data?.[0]?.[1],
        },
      ],
    );

    expect(reconciled[0]?.tables?.[0]?.columns).toEqual([
      { name: 'Region' },
      { name: 'Region2' },
      { name: 'Status' },
    ]);
    expect(reconciled[0]?.data?.[0]?.[1]).toMatchObject({
      m: 'Region2',
      v: 'Region2',
    });
    expect(source.tables?.[0]?.columns[1]?.name).toBe('Amount');
  });

  test('expands a table for an inserted column and keeps filters aligned', () => {
    const source = tableSheet();
    const changed = structuredClone(source);
    for (const row of changed.data ?? []) row?.splice(1, 0, null);

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'insertRowCol',
          path: [],
          value: {
            count: 1,
            direction: 'rightbottom',
            id: source.id,
            index: 0,
            type: 'column',
          },
        },
      ],
    );

    expect(reconciled[0]?.tables?.[0]).toMatchObject({
      range: { row: [0, 2], column: [0, 3] },
      columns: [
        { name: 'Region' },
        { name: 'Column2' },
        { name: 'Amount' },
        { name: 'Status' },
      ],
      filters: [
        {
          column: 2,
          criteria: { type: 'greater-than', value: '10' },
        },
      ],
    });
    expect(reconciled[0]?.data?.[0]?.[1]).toMatchObject({
      m: 'Column2',
      v: 'Column2',
    });
  });

  test('allows safe body and column deletion but blocks table destruction', () => {
    const sheet = tableSheet();

    expect(
      canApplySpreadsheetTableStructureChange(sheet, {
        axis: 'row',
        kind: 'delete',
        start: 1,
        end: 1,
      }),
    ).toBe(true);
    expect(
      canApplySpreadsheetTableStructureChange(sheet, {
        axis: 'column',
        kind: 'delete',
        start: 1,
        end: 1,
      }),
    ).toBe(true);
    expect(
      canApplySpreadsheetTableStructureChange(sheet, {
        axis: 'row',
        kind: 'delete',
        start: 0,
        end: 0,
      }),
    ).toBe(false);
    expect(
      canApplySpreadsheetTableStructureChange(sheet, {
        axis: 'row',
        kind: 'delete',
        start: 1,
        end: 2,
      }),
    ).toBe(false);
    expect(
      canApplySpreadsheetTableStructureChange(sheet, {
        axis: 'column',
        kind: 'delete',
        start: 0,
        end: 2,
      }),
    ).toBe(false);
  });

  test('shrinks columns and removes or shifts their filter criteria', () => {
    const source = tableSheet({
      filters: [
        { column: 1, criteria: { type: 'greater-than', value: '10' } },
        { column: 2, criteria: { type: 'equals', value: 'Ready' } },
      ],
    });
    const changed = structuredClone(source);
    for (const row of changed.data ?? []) row?.splice(1, 1);

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'deleteRowCol',
          path: [],
          value: { end: 1, id: source.id, start: 1, type: 'column' },
        },
      ],
    );

    expect(reconciled[0]?.tables?.[0]).toMatchObject({
      range: { row: [0, 2], column: [0, 1] },
      columns: [{ name: 'Region' }, { name: 'Status' }],
      filters: [{ column: 1, criteria: { type: 'equals', value: 'Ready' } }],
    });
  });

  test('fills an inserted table body row from the calculated-column rule', () => {
    const source = calculatedTableSheet();
    const changed = structuredClone(source);
    changed.data = [
      changed.data?.[0] ?? [],
      [{ v: 'New row' }, { v: 3 }, undefined],
      changed.data?.[1] ?? [],
      changed.data?.[2] ?? [],
    ];

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'insertRowCol',
          path: [],
          value: {
            count: 1,
            direction: 'lefttop',
            id: source.id,
            index: 1,
            type: 'row',
          },
        },
      ],
    );

    expect(reconciled[0]?.tables?.[0]?.range).toEqual({
      row: [0, 3],
      column: [0, 2],
    });
    expect(reconciled[0]?.tables?.[0]?.columns[2]).toEqual({
      name: 'Amount',
      calculatedFormula: '=[@Units]*2',
    });
    expect(reconciled[0]?.data?.[1]?.[2]).toMatchObject({
      f: '=[@Units]*2',
    });
    expect(reconciled[0]?.data?.[1]?.[2]?.v).toBeUndefined();
  });

  test('does not fill rows inserted outside a table', () => {
    const source = calculatedTableSheet();
    const changed = structuredClone(source);
    changed.data = [
      ...(changed.data ?? []),
      [{ v: 'Outside' }, undefined, undefined],
    ];

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'insertRowCol',
          path: [],
          value: {
            count: 1,
            direction: 'lefttop',
            id: source.id,
            index: 5,
            type: 'row',
          },
        },
      ],
    );

    expect(reconciled[0]?.tables?.[0]?.range).toEqual({
      row: [0, 2],
      column: [0, 2],
    });
    expect(reconciled[0]?.data?.[5]?.[2]).toBeUndefined();
  });

  test('fills every row in a multi-row table insertion', () => {
    const source = calculatedTableSheet();
    const changed = structuredClone(source);
    changed.data = [
      changed.data?.[0] ?? [],
      [{ v: 'One' }, { v: 1 }, undefined],
      [{ v: 'Two' }, { v: 2 }, undefined],
      changed.data?.[1] ?? [],
      changed.data?.[2] ?? [],
    ];

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'insertRowCol',
          path: [],
          value: {
            count: 2,
            direction: 'lefttop',
            id: source.id,
            index: 1,
            type: 'row',
          },
        },
      ],
    );

    expect(reconciled[0]?.data?.[1]?.[2]?.f).toBe('=[@Units]*2');
    expect(reconciled[0]?.data?.[2]?.[2]?.f).toBe('=[@Units]*2');
  });

  test('keeps a manual value in a newly inserted calculated cell', () => {
    const source = calculatedTableSheet();
    const changed = structuredClone(source);
    changed.data = [
      changed.data?.[0] ?? [],
      [{ v: 'Exception' }, { v: 3 }, { v: 99, m: '99' }],
      changed.data?.[1] ?? [],
      changed.data?.[2] ?? [],
    ];

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'insertRowCol',
          path: [],
          value: {
            count: 1,
            direction: 'lefttop',
            id: source.id,
            index: 1,
            type: 'row',
          },
        },
      ],
    );

    expect(reconciled[0]?.data?.[1]?.[2]).toEqual({ v: 99, m: '99' });
    expect(reconciled[0]?.tables?.[0]?.columns[2]?.calculatedFormula).toBe(
      '=[@Units]*2',
    );
  });

  test('fails closed when current-row formulas conflict', () => {
    const source = calculatedTableSheet();
    const changed = structuredClone(source);
    const firstBody = changed.data?.[1];
    if (!firstBody) throw new Error('Expected calculated table body.');
    firstBody[2] = { f: '=[@Units]*3' };

    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [source],
      [
        {
          id: source.id,
          op: 'replace',
          path: ['data', 1, 2],
          value: firstBody[2],
        },
      ],
    );

    expect(reconciled[0]?.tables?.[0]?.columns[2]).toEqual({ name: 'Amount' });
  });

  test('fills sparse celldata without materializing an unrelated matrix', () => {
    const source = calculatedTableSheet();
    const sparse: WorkSpreadsheetSheet = {
      ...source,
      data: undefined,
      celldata: [
        { r: 0, c: 0, v: { v: 'Region' } },
        { r: 0, c: 1, v: { v: 'Units' } },
        { r: 0, c: 2, v: { v: 'Amount' } },
        { r: 1, c: 0, v: { v: 'East' } },
        { r: 1, c: 1, v: { v: 10 } },
        { r: 1, c: 2, v: { f: '=[@Units]*2', v: 20 } },
        { r: 2, c: 0, v: { v: 'West' } },
        { r: 2, c: 1, v: { v: 12 } },
        { r: 2, c: 2, v: { f: '=[@Units]*2', v: 24 } },
      ],
    };
    const changed = {
      ...sparse,
      celldata: [
        ...(sparse.celldata ?? []).map((entry) =>
          entry.r === 2 ? { ...entry, r: 3 } : entry,
        ),
        { r: 2, c: 0, v: { v: 'New' } },
        { r: 2, c: 1, v: { v: 4 } },
      ],
    };
    const reconciled = reconcileSpreadsheetTablesAfterFortune(
      [changed],
      [sparse],
      [
        {
          id: source.id,
          op: 'insertRowCol',
          path: [],
          value: {
            count: 1,
            direction: 'rightbottom',
            id: source.id,
            index: 1,
            type: 'row',
          },
        },
      ],
    );

    expect(reconciled[0]?.data).toBeUndefined();
    expect(
      reconciled[0]?.celldata?.find((entry) => entry.r === 2 && entry.c === 2)
        ?.v?.f,
    ).toBe('=[@Units]*2');
  });
});

function tableSheet(
  patch: Partial<WorkSpreadsheetTable> = {},
): WorkSpreadsheetSheet {
  const table: WorkSpreadsheetTable = {
    id: 'table-1',
    name: 'Table1',
    range: { row: [0, 2], column: [0, 2] },
    columns: [{ name: 'Region' }, { name: 'Amount' }, { name: 'Status' }],
    filters: [
      {
        column: 1,
        criteria: { type: 'greater-than', value: '10' },
      },
    ],
    headerRow: true,
    totalsRow: false,
    style: { family: 'medium', number: 2 },
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
    ...patch,
  };
  return {
    id: 'sheet-1',
    name: 'Sales',
    row: 20,
    column: 8,
    data: [
      [{ v: 'Region' }, { v: 'Amount' }, { v: 'Status' }],
      [{ v: 'East' }, { v: 10 }, { v: 'Ready' }],
      [{ v: 'West' }, { v: 12 }, { v: 'Blocked' }],
    ],
    tables: [table],
  };
}

function calculatedTableSheet(): WorkSpreadsheetSheet {
  const sheet = tableSheet({
    columns: [
      { name: 'Region' },
      { name: 'Units' },
      { name: 'Amount', calculatedFormula: '=[@Units]*2' },
    ],
  });
  sheet.data?.[0]?.splice(
    0,
    3,
    { v: 'Region' },
    { v: 'Units' },
    { v: 'Amount' },
  );
  sheet.data?.[1]?.splice(0, 3, { v: 'East' }, { v: 10 }, undefined);
  sheet.data?.[2]?.splice(0, 3, { v: 'West' }, { v: 12 }, undefined);
  sheet.data?.[1]?.splice(2, 1, { f: '=[@Units]*2', v: 20 });
  sheet.data?.[2]?.splice(2, 1, { f: '=[@Units]*2', v: 24 });
  return sheet;
}
