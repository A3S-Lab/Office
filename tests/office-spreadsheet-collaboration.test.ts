import { readFileSync } from 'node:fs';
import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeSpreadsheetCollaborationBinding,
  initializeOfficeSpreadsheetCollaboration,
  readOfficeSpreadsheetCollaboration,
  type SpreadsheetContent,
} from '../src/core';
import {
  NATIVE_SPREADSHEET_BATCH_CELLS_BASE64,
  NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
  NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
  NATIVE_SPREADSHEET_SET_CELL_BASE64,
} from './fixtures/native-spreadsheet-cell-updates';
import { spreadsheetCollaborationFixture as fixture } from './fixtures/spreadsheet-collaboration';

const BROWSER_SPREADSHEET_FIXTURE_BASE64 = readFileSync(
  'tests/fixtures/browser-spreadsheet-collaboration-update.base64',
  'utf8',
).trim();

test('initializes sparse typed Spreadsheet roots without a workbook blob', () => {
  const session = spreadsheetSession('spreadsheet-typed');
  const expected = fixtureWithoutTransientViewState();

  expect(initializeOfficeSpreadsheetCollaboration(session, fixture())).toEqual({
    initialized: true,
    content: expected,
  });
  expect(readOfficeSpreadsheetCollaboration(session)).toEqual(expected);
  expect(
    session.document.getMap(session.rootName('spreadsheet.sheets')).size,
  ).toBe(2);
  expect(
    session.document.share.has(session.rootName('spreadsheet.content')),
  ).toBe(false);
  const sheets = session.document.getMap(
    session.rootName('spreadsheet.sheets'),
  );
  const input = sheets.get('sheet-input') as Y.Map<unknown>;
  expect(input).toBeInstanceOf(Y.Map);
  expect(input.has('data')).toBe(false);
  expect(input.has('celldata')).toBe(false);
  expect((input.get('cellPresence') as Y.Map<unknown>).size).toBe(4);
});

test('rejects duplicate identities before bootstrap metadata is written', () => {
  const session = spreadsheetSession('spreadsheet-duplicate');
  const content = fixture();

  expect(() =>
    initializeOfficeSpreadsheetCollaboration(session, {
      ...content,
      sheets: [content.sheets[0], { ...content.sheets[1], id: 'sheet-input' }],
    }),
  ).toThrow(/unique sheet ID/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
});

test('bounds dense matrices while retaining sparse Excel coordinates', () => {
  const denseSession = spreadsheetSession('spreadsheet-dense-bound');
  expect(() =>
    initializeOfficeSpreadsheetCollaboration(denseSession, {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-large-dense',
          name: 'Large dense',
          data: Array.from({ length: 1_001 }, () => Array(1_000).fill(null)),
        },
      ],
    }),
  ).toThrow(/materialized dense cells/);
  expect(
    denseSession.document.getMap(denseSession.rootName('metadata')).size,
  ).toBe(0);

  const sparseSession = spreadsheetSession('spreadsheet-sparse-bound');
  initializeOfficeSpreadsheetCollaboration(sparseSession, {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-large-sparse',
        name: 'Large sparse',
        row: 1_048_576,
        column: 16_384,
        celldata: [{ r: 1_048_575, c: 16_383, v: { v: 'edge' } }],
      },
    ],
  });
  expect(
    readOfficeSpreadsheetCollaboration(sparseSession).sheets[0].celldata,
  ).toEqual([{ r: 1_048_575, c: 16_383, v: { v: 'edge' } }]);
});

test('patches sparse data matrices without treating holes as cells', () => {
  const session = spreadsheetSession('spreadsheet-sparse-data-patch');
  const data: NonNullable<SpreadsheetContent['sheets'][number]['data']> = [];
  data.length = 8;
  data[0] = [{ v: 'Anchor', m: 'Anchor' }];
  data[7] = [];
  data[7].length = 4;
  data[7][3] = { v: 'Tail', m: 'Tail' };
  const initial: SpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [{ id: 'sheet-sparse-data', name: 'Sparse data', data }],
  };
  initializeOfficeSpreadsheetCollaboration(session, initial);

  const nextData: NonNullable<SpreadsheetContent['sheets'][number]['data']> =
    [];
  nextData.length = data.length;
  nextData[0] = data[0];
  nextData[7] = [];
  nextData[7].length = 4;
  nextData[7][3] = { v: 'Updated tail', m: 'Updated tail' };
  const binding = createOfficeSpreadsheetCollaborationBinding(session);

  expect(
    binding.replace(initial, {
      ...initial,
      sheets: [{ ...initial.sheets[0], data: nextData }],
    }),
  ).toBe(true);
  expect(binding.content().sheets[0]?.data?.[7]?.[3]).toEqual({
    v: 'Updated tail',
    m: 'Updated tail',
  });
  const sheets = session.document.getMap(
    session.rootName('spreadsheet.sheets'),
  );
  const sheet = sheets.get('sheet-sparse-data') as Y.Map<unknown>;
  expect((sheet.get('cellPresence') as Y.Map<unknown>).size).toBe(2);
  binding.destroy();
});

test('rejects malformed Spreadsheet roots without writes during read', () => {
  const session = spreadsheetSession('spreadsheet-pure-read');
  initializeOfficeSpreadsheetCollaboration(session, fixture());
  const sheets = session.document.getMap(
    session.rootName('spreadsheet.sheets'),
  );
  (sheets.get('sheet-input') as Y.Map<unknown>).delete('cellPresence');
  let transactions = 0;
  const countTransaction = () => {
    transactions += 1;
  };
  session.document.on('afterTransaction', countTransaction);

  expect(() => readOfficeSpreadsheetCollaboration(session)).toThrow(
    /sheet cell presence is invalid/,
  );
  expect(transactions).toBe(0);
});

test('merges formula, style, note, config, and metadata edits by object field', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-field-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 2, (cell) => ({
      ...cell,
      f: '=SUM(A2:B2)*2',
      v: 60,
      m: '60',
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateInputSheet(secondBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        rowlen: { ...sheet.config?.rowlen, '2': 36 },
      },
      formulaMetadata: {
        ...sheet.formulaMetadata,
        sourceFormulas: {
          ...sheet.formulaMetadata?.sourceFormulas,
          D2: '=C2*2',
        },
      },
      data: sheet.data?.map((row, rowIndex) =>
        row.map((cell, column) =>
          rowIndex === 1 && column === 2 && cell
            ? {
                ...cell,
                bg: '#FEF3C7',
                ps: {
                  left: null,
                  top: null,
                  width: 120,
                  height: 80,
                  value: 'Remote note',
                  isShow: false,
                },
              }
            : cell,
        ),
      ),
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].data?.[1]?.[2]).toMatchObject({
    f: '=SUM(A2:B2)*2',
    v: 60,
    m: '60',
    bg: '#FEF3C7',
    ps: { value: 'Remote note' },
  });
  expect(converged.sheets[0].config?.rowlen).toMatchObject({
    '1': 28,
    '2': 36,
  });
  expect(converged.sheets[0].formulaMetadata?.sourceFormulas).toMatchObject({
    C2: '=SUM(A2:B2)',
    D2: '=C2*2',
  });
});

test('preserves directional diagonal borders across unrelated collaboration edits', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-diagonal-border',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();
  const line = { color: '#2463eb', style: '10' };

  firstBinding.replace(
    firstBefore,
    updateInputSheet(firstBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        borderInfo: [
          ...(sheet.config?.borderInfo ?? []),
          {
            rangeType: 'cell',
            value: {
              row_index: 1,
              col_index: 1,
              s: line,
              a3sDiagonal: { down: true, line, up: true },
            },
          },
        ],
      },
    })),
  );
  secondBinding.replace(
    stale,
    updateInputCell(stale, 1, 0, (cell) => ({
      ...cell,
      v: 12,
      m: '12',
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].config?.borderInfo).toContainEqual({
    rangeType: 'cell',
    value: {
      row_index: 1,
      col_index: 1,
      s: line,
      a3sDiagonal: { down: true, line, up: true },
    },
  });
  expect(converged.sheets[0].data?.[1]?.[0]).toMatchObject({ v: 12 });
});

test('syncs native table records and merges independent design edits', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-table-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const before = firstBinding.content();
  const table = {
    id: 'table-inputs',
    name: 'InputTable',
    range: {
      row: [1, 2] as [number, number],
      column: [0, 2] as [number, number],
    },
    columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }],
    filters: [],
    headerRow: true,
    totalsRow: false,
    style: { family: 'medium' as const, number: 2 },
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
  };

  expect(
    firstBinding.replace(before, {
      ...before,
      sheets: before.sheets.map((sheet) =>
        sheet.id === 'sheet-input' ? { ...sheet, tables: [table] } : sheet,
      ),
    }),
  ).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  expect(secondBinding.content().sheets[0]?.tables).toEqual([table]);

  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();
  firstBinding.replace(firstBefore, {
    ...firstBefore,
    sheets: firstBefore.sheets.map((sheet) =>
      sheet.id === 'sheet-input'
        ? {
            ...sheet,
            tables: sheet.tables?.map((candidate) =>
              candidate.id === table.id
                ? { ...candidate, name: 'Inputs_2026' }
                : candidate,
            ),
          }
        : sheet,
    ),
  });
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    sheets: secondBefore.sheets.map((sheet) =>
      sheet.id === 'sheet-input'
        ? {
            ...sheet,
            tables: sheet.tables?.map((candidate) =>
              candidate.id === table.id
                ? { ...candidate, showColumnStripes: true }
                : candidate,
            ),
          }
        : sheet,
    ),
  });
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0]?.tables?.[0]).toMatchObject({
    id: table.id,
    name: 'Inputs_2026',
    showColumnStripes: true,
  });
});

test('syncs validated calculated-column formulas across Yjs clients', () => {
  const content = spreadsheetTableFilterContent([
    { type: 'blanks' },
    { type: 'non-blanks' },
  ]);
  const table = content.sheets[0]?.tables?.[0];
  if (!table) throw new Error('Expected a table fixture.');
  table.columns[1] = {
    ...table.columns[1],
    calculatedFormula: '=[@Column 1]*2',
  };

  const firstDocument = new Y.Doc();
  const first = spreadsheetSession(
    'spreadsheet-calculated-column-sync',
    firstDocument,
  );
  initializeOfficeSpreadsheetCollaboration(first, content);

  const secondDocument = new Y.Doc();
  Y.applyUpdate(secondDocument, Y.encodeStateAsUpdate(firstDocument));
  const second = spreadsheetSession(
    'spreadsheet-calculated-column-sync',
    secondDocument,
  );

  expect(readOfficeSpreadsheetCollaboration(second).sheets[0]?.tables).toEqual([
    table,
  ]);
});

test('syncs validated totals-row metadata across Yjs clients', () => {
  const content = spreadsheetTableFilterContent([
    { type: 'blanks' },
    { type: 'non-blanks' },
  ]);
  const table = content.sheets[0]?.tables?.[0];
  if (!table) throw new Error('Expected a table fixture.');
  table.totalsRow = true;
  table.range = { row: [0, 3], column: [0, 1] };
  table.columns = [
    { name: 'Column 1', totalsLabel: 'Total' },
    { name: 'Column 2', totalsFunction: 'sum' },
  ];

  const firstDocument = new Y.Doc();
  const first = spreadsheetSession(
    'spreadsheet-totals-row-sync',
    firstDocument,
  );
  initializeOfficeSpreadsheetCollaboration(first, content);
  const secondDocument = new Y.Doc();
  Y.applyUpdate(secondDocument, Y.encodeStateAsUpdate(firstDocument));
  const second = spreadsheetSession(
    'spreadsheet-totals-row-sync',
    secondDocument,
  );
  expect(readOfficeSpreadsheetCollaboration(second).sheets[0]?.tables).toEqual([
    table,
  ]);
});

test('accepts every closed native table filter criterion', () => {
  const criteria = [
    { type: 'values', values: ['Open', 'Closed'], includeBlanks: true },
    { type: 'equals', value: 'Exact' },
    { type: 'not-equals', value: 'Excluded' },
    { type: 'contains', value: 'middle' },
    { type: 'does-not-contain', value: 'blocked' },
    { type: 'begins-with', value: 'prefix' },
    { type: 'does-not-begin-with', value: 'excluded-prefix' },
    { type: 'ends-with', value: 'suffix' },
    { type: 'does-not-end-with', value: 'excluded-suffix' },
    { type: 'greater-than', value: '10' },
    { type: 'greater-than-or-equal', value: '20' },
    { type: 'less-than', value: '90' },
    { type: 'less-than-or-equal', value: '80' },
    { type: 'between', lower: '30', upper: '70' },
    { type: 'not-between', lower: '40', upper: '60' },
    { type: 'blanks' },
    { type: 'non-blanks' },
    { type: 'top', count: 10 },
    { type: 'top-percent', percent: 20 },
    { type: 'bottom', count: 5 },
    { type: 'bottom-percent', percent: 15 },
    { type: 'dynamic', kind: 'this-month' },
    {
      type: 'compound',
      conjunction: 'or',
      conditions: [
        { type: 'contains', value: 'Risk' },
        { type: 'greater-than', value: '100' },
      ],
    },
  ];
  const session = spreadsheetSession('spreadsheet-table-filter-criteria');

  initializeOfficeSpreadsheetCollaboration(
    session,
    spreadsheetTableFilterContent(criteria),
  );

  expect(
    readOfficeSpreadsheetCollaboration(session).sheets[0]?.tables?.[0]?.filters,
  ).toEqual(
    criteria.map((criterion, column) => ({ column, criteria: criterion })),
  );
});

test('rejects malformed native table filters before collaboration writes', () => {
  const oversizedAggregate = Array.from({ length: 33 }, (_, index) =>
    `${index}:`.padEnd(32_767, 'x'),
  );
  const cases: Array<{
    criteria: unknown;
    expected: RegExp;
    name: string;
  }> = [
    {
      name: 'unknown criterion',
      criteria: { type: 'future-filter', value: 'Open' },
      expected: /supported filter criteria/,
    },
    {
      name: 'unknown field',
      criteria: { type: 'blanks', executable: true },
      expected: /without unknown fields/,
    },
    {
      name: 'empty value set',
      criteria: { type: 'values', values: [], includeBlanks: false },
      expected: /at least one value/,
    },
    {
      name: 'too many values',
      criteria: {
        type: 'values',
        values: Array.from({ length: 10_001 }, (_, index) => String(index)),
        includeBlanks: false,
      },
      expected: /at most 10,000 values/,
    },
    {
      name: 'duplicate value',
      criteria: {
        type: 'values',
        values: ['Open', 'Open'],
        includeBlanks: false,
      },
      expected: /unique filter values/,
    },
    {
      name: 'empty comparison',
      criteria: { type: 'equals', value: '' },
      expected: /1 to 32,767 characters/,
    },
    {
      name: 'oversized comparison',
      criteria: { type: 'contains', value: 'x'.repeat(32_768) },
      expected: /1 to 32,767 characters/,
    },
    {
      name: 'XML-forbidden comparison',
      criteria: { type: 'ends-with', value: 'unsafe\u0000value' },
      expected: /XML-compatible filter text/,
    },
    {
      name: 'compound without exactly two conditions',
      criteria: {
        type: 'compound',
        conjunction: 'and',
        conditions: [{ type: 'equals', value: 'Open' }],
      },
      expected: /exactly two custom filter conditions/,
    },
    {
      name: 'compound with an unknown conjunction',
      criteria: {
        type: 'compound',
        conjunction: 'xor',
        conditions: [
          { type: 'equals', value: 'Open' },
          { type: 'equals', value: 'Closed' },
        ],
      },
      expected: /an 'and' or 'or' conjunction/,
    },
    {
      name: 'recursive compound condition',
      criteria: {
        type: 'compound',
        conjunction: 'or',
        conditions: [
          { type: 'equals', value: 'Open' },
          {
            type: 'compound',
            conjunction: 'and',
            conditions: [
              { type: 'equals', value: 'Open' },
              { type: 'equals', value: 'Closed' },
            ],
          },
        ],
      },
      expected: /supported custom filter condition/,
    },
    {
      name: 'top count outside contract',
      criteria: { type: 'top', count: 501 },
      expected: /1 through 500/,
    },
    {
      name: 'percentage outside contract',
      criteria: { type: 'bottom-percent', percent: 101 },
      expected: /1 through 100/,
    },
    {
      name: 'unknown dynamic family',
      criteria: { type: 'dynamic', kind: 'current-century' },
      expected: /supported dynamic filter/,
    },
    {
      name: 'aggregate text budget',
      criteria: {
        type: 'values',
        values: oversizedAggregate,
        includeBlanks: false,
      },
      expected: /1,048,576 bytes/,
    },
  ];

  for (const [index, candidate] of cases.entries()) {
    const session = spreadsheetSession(`spreadsheet-table-filter-${index}`);
    expect(
      () =>
        initializeOfficeSpreadsheetCollaboration(
          session,
          spreadsheetTableFilterContent([candidate.criteria]),
        ),
      candidate.name,
    ).toThrow(candidate.expected);
    expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
  }
});

test('rejects malformed native table metadata before collaboration writes', () => {
  const cases: Array<{
    expected: RegExp;
    mutate: (table: Record<string, unknown>) => void;
    name: string;
  }> = [
    {
      name: 'A1-like table name',
      mutate: (table) => {
        table.name = 'A1';
      },
      expected: /valid table name/,
    },
    {
      name: 'filter without header',
      mutate: (table) => {
        table.headerRow = false;
      },
      expected: /filters to require an enabled header row/,
    },
    {
      name: 'style flags without style',
      mutate: (table) => {
        table.style = { family: 'none' };
      },
      expected: /style flags to require a built-in style/,
    },
    {
      name: 'unknown style field',
      mutate: (table) => {
        table.style = { family: 'none', number: 2 };
      },
      expected: /without unknown fields/,
    },
    {
      name: 'unknown column field',
      mutate: (table) => {
        const columns = table.columns as Array<Record<string, unknown>>;
        if (columns[0]) columns[0].formula = '=1';
      },
      expected: /without unknown fields/,
    },
    {
      name: 'dangerous calculated-column formula',
      mutate: (table) => {
        const columns = table.columns as Array<Record<string, unknown>>;
        if (columns[0]) {
          columns[0].calculatedFormula = '=INDIRECT([@Column 1])';
        }
      },
      expected: /bounded structured calculated-column formula/,
    },
    {
      name: 'external calculated-column formula',
      mutate: (table) => {
        const columns = table.columns as Array<Record<string, unknown>>;
        if (columns[0]) {
          columns[0].calculatedFormula = '=SUM([Book.xlsx]Sheet1![@Column 1])';
        }
      },
      expected: /bounded structured calculated-column formula/,
    },
    {
      name: 'unknown totals function',
      mutate: (table) => {
        const columns = table.columns as Array<Record<string, unknown>>;
        if (columns[0]) columns[0].totalsFunction = 'median';
      },
      expected: /supported totals-row function/,
    },
    {
      name: 'custom totals function without formula',
      mutate: (table) => {
        const columns = table.columns as Array<Record<string, unknown>>;
        if (columns[0]) columns[0].totalsFunction = 'custom';
      },
      expected: /custom totals-row functions to include a formula/,
    },
  ];

  for (const [index, candidate] of cases.entries()) {
    const content = spreadsheetTableFilterContent([{ type: 'blanks' }]);
    const table = content.sheets[0]?.tables?.[0] as unknown as Record<
      string,
      unknown
    >;
    candidate.mutate(table);
    const session = spreadsheetSession(`spreadsheet-table-metadata-${index}`);
    expect(
      () => initializeOfficeSpreadsheetCollaboration(session, content),
      candidate.name,
    ).toThrow(candidate.expected);
    expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
  }
});

test('accepts Rust-compatible Unicode table and column name limits', () => {
  const content = spreadsheetTableFilterContent([
    { type: 'blanks' },
    { type: 'non-blanks' },
  ]);
  const table = content.sheets[0]?.tables?.[0];
  if (!table) throw new Error('Expected a table fixture.');
  table.name = '𐐀'.repeat(255);
  table.columns[0] = { name: '🙂'.repeat(255) };
  table.columns[1] = { name: 'join\u200Der' };
  const session = spreadsheetSession('spreadsheet-table-unicode-contract');

  initializeOfficeSpreadsheetCollaboration(session, content);

  expect(
    readOfficeSpreadsheetCollaboration(session).sheets[0]?.tables?.[0],
  ).toMatchObject({
    name: table.name,
    columns: table.columns,
  });
});

test('converges concurrent first writes to one blank cell by nested field', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-blank-cell-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 2, 1, () => ({
      f: '=A2+B2',
      v: 30,
      m: '30',
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateInputCell(secondBefore, 2, 1, () => ({
      bg: '#DBEAFE',
      ct: { fa: '$0.00', t: 'n' },
      ps: {
        left: null,
        top: null,
        width: 140,
        height: 80,
        value: 'Concurrent note',
        isShow: false,
      },
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].data?.[2]?.[1]).toMatchObject({
    f: '=A2+B2',
    v: 30,
    m: '30',
    bg: '#DBEAFE',
    ct: { fa: '$0.00', t: 'n' },
    ps: { value: 'Concurrent note' },
  });
});

test('merges different fields in the same OOXML merge definition', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-merge-config-convergence',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputSheet(firstBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        merge: {
          ...sheet.config?.merge,
          '0_0': { ...sheet.config?.merge?.['0_0'], rs: 2 } as {
            r: number;
            c: number;
            rs: number;
            cs: number;
          },
        },
      },
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateInputSheet(secondBefore, (sheet) => ({
      ...sheet,
      config: {
        ...sheet.config,
        merge: {
          ...sheet.config?.merge,
          '0_0': { ...sheet.config?.merge?.['0_0'], cs: 3 } as {
            r: number;
            c: number;
            rs: number;
            cs: number;
          },
        },
      },
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.content().sheets[0].config?.merge?.['0_0']).toEqual({
    r: 0,
    c: 0,
    rs: 2,
    cs: 3,
  });
  expect(secondBinding.content()).toEqual(firstBinding.content());
});

test('preserves remotely added cells when applying an unrelated stale snapshot', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-stale-cell',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 2, 0, () => ({ v: 40, m: '40' })),
  );
  exchangeUpdates(firstDocument, secondDocument);
  secondBinding.replace(
    stale,
    updateInputCell(stale, 1, 0, (cell) => ({ ...cell, v: 11, m: '11' })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.sheets[0].data?.[2]?.[0]).toMatchObject({ v: 40 });
  expect(converged.sheets[0].data?.[1]?.[0]).toMatchObject({ v: 11 });
});

test('rejects deleting a cell that was concurrently edited', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-delete-edit',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 0, (cell) => ({ ...cell, v: 99, m: '99' })),
  );
  applyMissingUpdate(firstDocument, secondDocument);

  expect(() =>
    secondBinding.replace(
      stale,
      updateInputCell(stale, 1, 0, () => null),
    ),
  ).toThrow(/changed concurrently/);
});

test('rejects stale conflicts before writing any part of the transaction', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-atomic-conflict',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const stale = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 0, (cell) => ({ ...cell, v: 99, m: '99' })),
  );
  applyMissingUpdate(firstDocument, secondDocument);
  const updateBeforeFailure = Y.encodeStateVector(secondDocument);
  expect(() =>
    secondBinding.replace(
      stale,
      updateInputCell(
        updateInputCell(stale, 1, 0, () => null),
        1,
        1,
        (cell) => ({ ...cell, v: 25, m: '25' }),
      ),
    ),
  ).toThrow(/changed concurrently/);

  expect(Y.encodeStateVector(secondDocument)).toEqual(updateBeforeFailure);
  expect(secondBinding.content().sheets[0].data?.[1]?.[0]?.v).toBe(99);
  expect(secondBinding.content().sheets[0].data?.[1]?.[1]?.v).toBe(20);
});

test('rejects concurrent reuse of the same sheet ID for different sheets', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-sheet-id-claim',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(firstBefore, {
    ...firstBefore,
    sheets: [...firstBefore.sheets, addedSheet('sheet-collision', 'Ada sheet')],
  });
  secondBinding.replace(secondBefore, {
    ...secondBefore,
    sheets: [
      ...secondBefore.sheets,
      addedSheet('sheet-collision', 'Grace sheet'),
    ],
  });
  exchangeUpdates(firstDocument, secondDocument);

  expect(() => firstBinding.content()).toThrow(/concurrently assigned/);
  expect(() => secondBinding.content()).toThrow(/concurrently assigned/);
});

test('detects concurrent independent Spreadsheet bootstrap', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = spreadsheetSession('spreadsheet-bootstrap-race', firstDocument);
  const second = spreadsheetSession(
    'spreadsheet-bootstrap-race',
    secondDocument,
  );
  initializeOfficeSpreadsheetCollaboration(first, fixture());
  const secondFixture = fixture();
  initializeOfficeSpreadsheetCollaboration(second, {
    ...secondFixture,
    calculation: {
      ...(secondFixture.calculation ?? {
        mode: 'automatic',
        fullCalculationOnLoad: false,
        forceFullCalculation: false,
        iterativeCalculation: false,
        maximumIterations: 100,
        maximumChange: 0.001,
        fullPrecision: true,
      }),
      maximumIterations: 200,
    },
  });

  exchangeUpdates(firstDocument, secondDocument);

  expect(() => first.metadata()).toThrow(/Multiple clients initialized/);
  expect(() => second.metadata()).toThrow(/Multiple clients initialized/);
});

test('keeps Spreadsheet undo local to one client', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'spreadsheet-local-undo',
  );
  const firstBinding = createOfficeSpreadsheetCollaborationBinding(first);
  const secondBinding = createOfficeSpreadsheetCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  firstBinding.replace(
    firstBefore,
    updateInputCell(firstBefore, 1, 0, (cell) => ({ ...cell, v: 12, m: '12' })),
  );
  exchangeUpdates(firstDocument, secondDocument);
  const secondBefore = secondBinding.content();
  secondBinding.stopCapturing();
  secondBinding.replace(
    secondBefore,
    updateInputCell(secondBefore, 1, 1, (cell) => ({
      ...cell,
      v: 25,
      m: '25',
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  expect(firstBinding.content().sheets[0].data?.[1]?.[0]?.v).toBe(10);
  expect(firstBinding.content().sheets[0].data?.[1]?.[1]?.v).toBe(25);
  expect(secondBinding.content()).toEqual(firstBinding.content());
});

test('rejects Spreadsheet mutation outside edit mode', () => {
  const document = new Y.Doc();
  const writable = spreadsheetSession('spreadsheet-view', document);
  initializeOfficeSpreadsheetCollaboration(writable, fixture());
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'spreadsheet-view',
    document,
    kind: 'spreadsheet',
    mode: 'view',
  });
  const binding = createOfficeSpreadsheetCollaborationBinding(readOnly);
  const before = binding.content();

  expect(() =>
    binding.replace(
      before,
      updateInputCell(before, 1, 0, () => ({ v: 1 })),
    ),
  ).toThrow(/cannot modify canonical content/);
  expect(() => binding.undo()).toThrow(/cannot modify canonical content/);
});

test('applies native Spreadsheet cell updates in Yjs across reordered delivery', () => {
  const orderedDocument = new Y.Doc();
  const reorderedDocument = new Y.Doc();
  for (const document of [orderedDocument, reorderedDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_SPREADSHEET_FIXTURE_BASE64));
    const binding = createOfficeSpreadsheetCollaborationBinding(
      spreadsheetSession('fixture-spreadsheet', document),
    );
    const before = binding.content();
    binding.replace(before, {
      ...before,
      sheets: before.sheets.map((sheet) => {
        if (sheet.id !== 'sheet-data') return sheet;
        const data = (sheet.data ?? []).map((row) => [...row]);
        const cell = data[1]?.[0];
        if (!cell) throw new Error('Expected the browser fixture data cell.');
        data[1][0] = {
          ...cell,
          bg: '#DBEAFE',
          ps: {
            left: null,
            top: null,
            width: 140,
            height: 80,
            value: 'Browser note',
            isShow: false,
          },
        };
        return { ...sheet, data };
      }),
    });
  }

  for (const encoded of [
    NATIVE_SPREADSHEET_BATCH_CELLS_BASE64,
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
    NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
    NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
  ]) {
    Y.applyUpdate(orderedDocument, decodeBase64(encoded));
  }
  for (const encoded of [
    NATIVE_SPREADSHEET_DELETE_CELL_BASE64,
    NATIVE_SPREADSHEET_CREATE_CELL_BASE64,
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
    NATIVE_SPREADSHEET_BATCH_CELLS_BASE64,
    NATIVE_SPREADSHEET_SET_CELL_BASE64,
    NATIVE_SPREADSHEET_BATCH_CELLS_BASE64,
  ]) {
    Y.applyUpdate(reorderedDocument, decodeBase64(encoded));
  }

  const contents = [orderedDocument, reorderedDocument].map((document) =>
    readOfficeSpreadsheetCollaboration(
      spreadsheetSession('fixture-spreadsheet', document),
    ),
  );
  expect(contents[0]).toEqual(contents[1]);
  const dataSheet = contents[0]?.sheets.find(({ id }) => id === 'sheet-data');
  expect(dataSheet?.data?.[1]?.[0]).toMatchObject({
    v: 12,
    m: '12',
    f: '=6*2',
    bl: 1,
    bg: '#DBEAFE',
    ct: { fa: '0.00', t: 'n' },
    ps: { value: 'Browser note' },
  });
  expect(dataSheet?.data?.[0]?.[0]).toBeNull();
  expect(dataSheet?.data?.[3]?.[4]).toEqual({
    v: 'Batched',
    m: 'Batched',
  });
  const emptySheet = contents[0]?.sheets.find(({ id }) => id === 'sheet-empty');
  expect(emptySheet?.data).toBeUndefined();
  expect(emptySheet?.celldata).toEqual([
    {
      r: 100,
      c: 5,
      v: {
        v: 'sparse native',
        m: 'sparse native',
        ps: { value: 'Agent note', isShow: false },
      },
    },
  ]);
  const sparseSheet = contents[0]?.sheets.find(
    ({ id }) => id === 'sheet-sparse',
  );
  expect(sparseSheet?.celldata).toEqual([]);
});

function fixtureWithoutTransientViewState() {
  const content = fixture();
  return {
    ...content,
    sheets: content.sheets.map(
      ({
        status: _status,
        zoomRatio: _zoomRatio,
        luckysheet_select_save: _selection,
        ...sheet
      }) => sheet,
    ),
  };
}

function spreadsheetSession(artifactId: string, document = new Y.Doc()) {
  return createOfficeCollaborationSession({
    artifactId,
    document,
    kind: 'spreadsheet',
  });
}

function connectedPair(artifactId: string) {
  const firstDocument = new Y.Doc();
  const first = spreadsheetSession(artifactId, firstDocument);
  initializeOfficeSpreadsheetCollaboration(first, fixture());
  const secondDocument = new Y.Doc();
  Y.applyUpdate(secondDocument, Y.encodeStateAsUpdate(firstDocument));
  const second = spreadsheetSession(artifactId, secondDocument);
  return { first, firstDocument, second, secondDocument };
}

function updateInputCell(
  content: ReturnType<typeof fixtureWithoutTransientViewState>,
  row: number,
  column: number,
  update: (
    cell: NonNullable<
      NonNullable<(typeof content.sheets)[number]['data']>[number][number]
    >,
  ) => NonNullable<(typeof content.sheets)[number]['data']>[number][number],
) {
  return updateInputSheet(content, (sheet) => {
    const data = (sheet.data ?? []).map((values) => [...values]);
    while (data.length <= row) data.push([]);
    while (data[row].length <= column) data[row].push(null);
    data[row][column] = update(data[row][column] ?? {});
    return { ...sheet, data };
  });
}

function updateInputSheet(
  content: ReturnType<typeof fixtureWithoutTransientViewState>,
  update: (
    sheet: (typeof content.sheets)[number],
  ) => (typeof content.sheets)[number],
) {
  return {
    ...content,
    sheets: content.sheets.map((sheet) =>
      sheet.id === 'sheet-input' ? update(sheet) : sheet,
    ),
  };
}

function addedSheet(id: string, name: string) {
  return {
    id,
    name,
    row: 1,
    column: 1,
    data: [[{ v: name, m: name }]],
  };
}

function spreadsheetTableFilterContent(
  criteria: readonly unknown[],
): SpreadsheetContent {
  const width = criteria.length;
  const content = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-table-filters',
        name: 'Table filters',
        data: [Array(width).fill(null), Array(width).fill(null)],
        tables: [
          {
            id: 'table-filters',
            name: 'FilterTable',
            range: { row: [0, 1], column: [0, width - 1] },
            columns: Array.from({ length: width }, (_, index) => ({
              name: `Column ${index + 1}`,
            })),
            filters: criteria.map((criterion, column) => ({
              column,
              criteria: criterion,
            })),
            headerRow: true,
            totalsRow: false,
            style: { family: 'medium', number: 2 },
            showFirstColumn: false,
            showLastColumn: false,
            showRowStripes: true,
            showColumnStripes: false,
          },
        ],
      },
    ],
  };
  return content as unknown as SpreadsheetContent;
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}

function applyMissingUpdate(source: Y.Doc, target: Y.Doc): void {
  Y.applyUpdate(
    target,
    Y.encodeStateAsUpdate(source, Y.encodeStateVector(target)),
    'test-network',
  );
}

function decodeBase64(value: string): Uint8Array {
  return Uint8Array.from(Buffer.from(value, 'base64'));
}
