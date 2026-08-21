import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetTable,
  convertSpreadsheetTableToRange,
  createSpreadsheetTableDialogSource,
  spreadsheetTableAtCell,
  updateSpreadsheetTable,
  validateSpreadsheetTableRequest,
} from '../src/internal/features/work/editors/spreadsheet-table';
import {
  createSpreadsheetTableRenderResolver,
  spreadsheetTableStylePalette,
} from '../src/internal/features/work/editors/spreadsheet-table-style';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../src/internal/features/work/work-types';

describe('spreadsheet tables', () => {
  test('creates one native table from the current region without styling every cell', () => {
    const content = tableContent();
    const bodyRow = content.sheets[0]?.data?.[1];
    const source = createSpreadsheetTableDialogSource(content, {
      sheetId: 'sheet-1',
      selection: {
        row: [1, 1],
        column: [1, 1],
        row_focus: 1,
        column_focus: 1,
      },
    });

    expect(source).toMatchObject({
      name: 'Table1',
      rangeReference: 'A1:C3',
      value: { headerRow: true, rangeReference: 'A1:C3' },
    });
    const next = source
      ? applySpreadsheetTable(content, {
          sheetId: source.sheetId,
          name: source.name,
          range: source.range,
          headerRow: true,
        })
      : null;

    expect(next?.sheets[0]?.tables).toEqual([
      expect.objectContaining({
        id: expect.any(String),
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        columns: [{ name: 'Region' }, { name: 'Column2' }, { name: 'Region2' }],
        headerRow: true,
        totalsRow: false,
        style: { family: 'medium', number: 2 },
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      }),
    ]);
    expect(next?.sheets[0]?.data?.[0]?.map((cell) => cell?.v)).toEqual([
      'Region',
      'Column2',
      'Region2',
    ]);
    expect(next?.sheets[0]?.data?.[1]).toBe(bodyRow);
    expect(next?.sheets[0]?.data?.[1]?.[0]).not.toHaveProperty('bg');
  });

  test('opens a one-row current region as a headerless table', () => {
    const content = tableContent();
    const sheet = firstSheet(content);
    sheet.data = [[{ v: 'East' }, { v: 12 }, { v: 'Ready' }]];
    const source = createSpreadsheetTableDialogSource(content, {
      sheetId: 'sheet-1',
      selection: {
        row: [0, 0],
        column: [1, 1],
        row_focus: 0,
        column_focus: 1,
      },
    });

    expect(source).toMatchObject({
      rangeReference: 'A1:C1',
      value: { headerRow: false, rangeReference: 'A1:C1' },
    });
    const next = source
      ? applySpreadsheetTable(content, {
          sheetId: source.sheetId,
          name: source.name,
          range: source.range,
          headerRow: source.value.headerRow,
        })
      : null;
    expect(next?.sheets[0]?.tables?.[0]).toMatchObject({
      columns: [{ name: 'Column1' }, { name: 'Column2' }, { name: 'Column3' }],
      headerRow: false,
      range: { row: [0, 0], column: [0, 2] },
    });
    expect(next?.sheets[0]?.data?.[0]?.[0]?.v).toBe('East');
  });

  test('keeps a maximum-row sparse sheet sparse when a bounded table is added', () => {
    const content = tableContent();
    const sheet = content.sheets[0];
    if (!sheet?.data) throw new Error('Expected table fixture data.');
    sheet.row = 1_048_576;
    sheet.column = 16_384;
    sheet.data.length = 1_048_576;

    const next = applySpreadsheetTable(content, {
      sheetId: 'sheet-1',
      name: 'SparseTable',
      range: { row: [0, 2], column: [0, 2] },
      headerRow: true,
    });

    expect(next?.sheets[0]?.data).toHaveLength(1_048_576);
    expect(Object.keys(next?.sheets[0]?.data ?? [])).toEqual(['0', '1', '2']);
  });

  test('uses materialized bounds when worksheet dimensions are omitted', () => {
    const content = tableContent();
    const sheet = firstSheet(content);
    delete sheet.row;
    delete sheet.column;

    expect(
      validateSpreadsheetTableRequest(content, {
        sheetId: 'sheet-1',
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        headerRow: true,
      }),
    ).toMatchObject({ ok: true });
  });

  test('keeps canonical duplicate headers within the 255-character limit', () => {
    const content = tableContent();
    const header = firstSheet(content).data?.[0];
    if (!header) throw new Error('Expected table fixture header data.');
    const maximumName = 'x'.repeat(255);
    header[0] = { v: maximumName };
    header[1] = { v: maximumName };

    const next = applySpreadsheetTable(content, {
      sheetId: 'sheet-1',
      name: 'Table1',
      range: { row: [0, 2], column: [0, 2] },
      headerRow: true,
    });

    expect(next?.sheets[0]?.tables?.[0]?.columns).toEqual([
      { name: maximumName },
      { name: `${'x'.repeat(254)}2` },
      { name: 'Region' },
    ]);
  });

  test('counts and truncates table headers by Unicode code point', () => {
    const content = tableContent();
    const header = firstSheet(content).data?.[0];
    if (!header) throw new Error('Expected table fixture header data.');
    const maximumName = '🙂'.repeat(255);
    header[0] = { v: maximumName };
    header[1] = { v: maximumName };

    const next = applySpreadsheetTable(content, {
      sheetId: 'sheet-1',
      name: 'Table1',
      range: { row: [0, 2], column: [0, 2] },
      headerRow: true,
    });
    const columns = next?.sheets[0]?.tables?.[0]?.columns ?? [];

    expect(columns[0]?.name).toBe(maximumName);
    expect(columns[1]?.name).toBe(`${'🙂'.repeat(254)}2`);
    expect(Array.from(columns[1]?.name ?? '')).toHaveLength(255);
  });

  test('rejects overlaps, merges, filters, protection, and pivot sheets', () => {
    const overlapping = tableContent();
    const first = applySpreadsheetTable(overlapping, {
      sheetId: 'sheet-1',
      name: 'Table1',
      range: { row: [0, 2], column: [0, 2] },
      headerRow: true,
    });
    expect(
      first &&
        validateSpreadsheetTableRequest(first, {
          sheetId: 'sheet-1',
          name: 'Table2',
          range: { row: [2, 4], column: [2, 3] },
          headerRow: true,
        }),
    ).toMatchObject({ ok: false, code: 'table-overlap' });

    const merged = tableContent();
    firstSheet(merged).config = {
      merge: { '1_1': { r: 1, c: 1, rs: 1, cs: 2 } },
    };
    expect(
      validateSpreadsheetTableRequest(merged, {
        sheetId: 'sheet-1',
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        headerRow: true,
      }),
    ).toMatchObject({ ok: false, code: 'merged-range' });

    const filtered = tableContent();
    firstSheet(filtered).filter_select = {
      row: [0, 2],
      column: [0, 2],
    };
    expect(
      validateSpreadsheetTableRequest(filtered, {
        sheetId: 'sheet-1',
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        headerRow: true,
      }),
    ).toMatchObject({ ok: false, code: 'auto-filter-overlap' });

    const protectedContent = tableContent();
    firstSheet(protectedContent).config = { authority: { sheet: 1 } };
    expect(
      validateSpreadsheetTableRequest(protectedContent, {
        sheetId: 'sheet-1',
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        headerRow: true,
      }),
    ).toMatchObject({ ok: false, code: 'protected-range' });

    const pivot = tableContent();
    firstSheet(pivot).isPivotTable = true;
    expect(
      validateSpreadsheetTableRequest(pivot, {
        sheetId: 'sheet-1',
        name: 'Table1',
        range: { row: [0, 2], column: [0, 2] },
        headerRow: true,
      }),
    ).toMatchObject({ ok: false, code: 'pivot-table' });
  });

  test('updates table design metadata and converts safely back to cell values', () => {
    const created = applySpreadsheetTable(tableContent(), {
      sheetId: 'sheet-1',
      name: 'Table1',
      range: { row: [0, 2], column: [0, 2] },
      headerRow: true,
    });
    const table = created?.sheets[0]?.tables?.[0];
    if (!created || !table) throw new Error('Expected a created table.');

    const updated = updateSpreadsheetTable(created, 'sheet-1', table.id, {
      name: 'Sales_2026',
      style: { family: 'dark', number: 7 },
      showFirstColumn: true,
      showLastColumn: true,
      showRowStripes: false,
      showColumnStripes: true,
    });
    expect(updated?.sheets[0]?.tables?.[0]).toMatchObject({
      name: 'Sales_2026',
      style: { family: 'dark', number: 7 },
      showFirstColumn: true,
      showLastColumn: true,
      showRowStripes: false,
      showColumnStripes: true,
    });
    expect(
      updated && spreadsheetTableAtCell(updated.sheets[0], 2, 2)?.name,
    ).toBe('Sales_2026');

    const converted = updated
      ? convertSpreadsheetTableToRange(updated, 'sheet-1', table.id)
      : null;
    expect(converted?.sheets[0]?.tables).toBeUndefined();
    expect(converted?.sheets[0]?.data?.[1]?.[1]?.v).toBe(10);
    const palette = spreadsheetTableStylePalette({
      family: 'dark',
      number: 7,
    });
    const updatedTable = updated?.sheets[0]?.tables?.[0];
    if (!updatedTable) throw new Error('Expected an updated table.');
    const expectedBody = createSpreadsheetTableRenderResolver([updatedTable])(
      1,
      1,
    );
    expect(converted?.sheets[0]?.data?.[0]?.[0]).toMatchObject({
      bg: palette?.header,
      bl: 1,
      fc: palette?.headerText,
    });
    expect(converted?.sheets[0]?.data?.[1]?.[1]).toMatchObject({
      bg: expectedBody?.background,
      fc: palette?.text,
    });
    expect(converted?.sheets[0]?.config?.borderInfo).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          borderType: 'border-all',
          rangeType: 'range',
        }),
      ]),
    );
    expect(updated.sheets[0]?.data?.[1]?.[1]?.bg).toBeUndefined();
  });

  test('refuses to materialize an oversized imported table during conversion', () => {
    const content = tableContent();
    const sheet = firstSheet(content);
    sheet.row = 100_002;
    sheet.tables = [
      {
        id: 'large-table',
        name: 'LargeTable',
        range: { row: [0, 100_000], column: [0, 0] },
        columns: [{ name: 'Region' }],
        filters: [],
        headerRow: true,
        totalsRow: false,
        style: { family: 'medium', number: 2 },
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    ];

    expect(
      convertSpreadsheetTableToRange(content, 'sheet-1', 'large-table'),
    ).toBeNull();
  });

  test('refuses conversion while a structured reference still targets the table', () => {
    const created = applySpreadsheetTable(tableContent(), {
      sheetId: 'sheet-1',
      name: 'Table1',
      range: { row: [0, 2], column: [0, 2] },
      headerRow: true,
    });
    const table = created?.sheets[0]?.tables?.[0];
    if (!created || !table) throw new Error('Expected a created table.');
    const data = firstSheet(created).data;
    if (!data) throw new Error('Expected table fixture data.');
    data[3] = [{ f: '=SUM(Table1[Region])', v: 0 }];

    expect(convertSpreadsheetTableToRange(created, 'sheet-1', table.id)).toBe(
      null,
    );
  });
});

function tableContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sales',
        row: 20,
        column: 8,
        data: [
          [{ v: 'Region' }, { v: '' }, { v: 'Region' }],
          [{ v: 'East' }, { v: 10 }, { v: 'Ready' }],
          [{ v: 'West' }, { v: 12 }, { v: 'Blocked' }],
        ],
      },
    ],
  };
}

function firstSheet(content: WorkSpreadsheetContent): WorkSpreadsheetSheet {
  const sheet = content.sheets[0];
  if (!sheet) throw new Error('Expected a spreadsheet fixture sheet.');
  return sheet;
}
