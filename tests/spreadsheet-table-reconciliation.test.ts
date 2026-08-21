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
