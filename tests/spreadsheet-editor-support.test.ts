import { expect, test } from '@rstest/core';
import { produce } from 'immer';
import {
  sameSpreadsheetHistoryContent,
  sameSpreadsheetWorkbookState,
  sameSpreadsheetWorkbookStateAfterOperations,
  spreadsheetSheetsForFortune,
  spreadsheetSheetsFromFortune,
} from '../src/internal/features/work/editors/spreadsheet-editor-support';
import { stageSpreadsheetRichTextPaste } from '../src/internal/features/work/editors/spreadsheet-rich-text-paste';
import {
  freezeImportedSpreadsheetCell,
  registerImportedSpreadsheetMatrix,
  SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY,
  spreadsheetMatrixProfile,
} from '../src/internal/features/work/work-spreadsheet-matrix-profile';
import { spreadsheetProtectionKey } from '../src/internal/features/work/work-spreadsheet-protection';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('projects sparse workbook cells without cloning logical empty ranges', () => {
  const data = sparseMatrix();
  const projected = spreadsheetSheetsForFortune([
    {
      id: 'sheet-1',
      name: 'Sparse',
      row: 1_000_000,
      column: 16_384,
      data,
      config: {
        merge: {
          '0_0': { r: 0, c: 0, rs: 1, cs: 2 },
        },
      },
    },
  ]);

  const fortuneData = projected[0]?.data;
  expect(projected[0]?.celldata).toBeUndefined();
  expect(fortuneData).toHaveLength(1_000_000);
  expect(Object.keys(fortuneData ?? [])).toEqual(['0', '999999']);
  expect(fortuneData?.[0]?.[0]).toEqual({
    v: 'Anchor',
    m: 'Anchor',
    mc: { r: 0, c: 0, rs: 1, cs: 2 },
  });
  expect(fortuneData?.[0]?.[1]).toEqual({ mc: { r: 0, c: 0 } });
  expect(fortuneData?.[999_999]?.[16_383]).toEqual({
    v: 'Tail',
    m: 'Tail',
  });
  expect(fortuneData).not.toBe(data);
  expect(fortuneData?.[0]).not.toBe(data[0]);
  expect(fortuneData?.[0]?.[0]).not.toBe(data[0]?.[0]);
  expect(Object.isFrozen(fortuneData)).toBe(true);
  expect(Object.isFrozen(fortuneData?.[0])).toBe(true);
  expect(Object.isFrozen(fortuneData?.[0]?.[0])).toBe(true);
  expect(Object.keys(data)).toEqual(['0', '999999']);
  expect(data[0]?.[0]).not.toHaveProperty('mc');
});

test('projects celldata as an independent sparse Fortune matrix', () => {
  const sourceCell = {
    f: '=A1*2',
    v: 4,
    m: '4',
    ct: { fa: '0.00', t: 'n' },
  };
  const projected = spreadsheetSheetsForFortune([
    {
      id: 'sheet-celldata',
      name: 'Cell data',
      row: 100_000,
      column: 10,
      celldata: [{ r: 99_999, c: 9, v: sourceCell }],
    },
  ]);

  const data = projected[0]?.data;
  expect(projected[0]?.celldata).toBeUndefined();
  expect(data).toHaveLength(100_000);
  expect(Object.keys(data ?? [])).toEqual(['0', '99999']);
  expect(data?.[0]).toHaveLength(10);
  expect(data?.[99_999]?.[9]).toEqual(sourceCell);
  expect(data?.[99_999]?.[9]).not.toBe(sourceCell);
  expect(data?.[99_999]?.[9]?.ct).not.toBe(sourceCell.ct);
});

test('reuses an authenticated immutable import matrix at the Fortune boundary', () => {
  const sourceCell = freezeImportedSpreadsheetCell({
    v: 'Imported',
    m: 'Imported',
  });
  const sourceRow = [sourceCell];
  const data = [sourceRow];
  data.length = 100_000;
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 10,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 100_000,
    shownCommentCells: [],
  });

  const projected = spreadsheetSheetsForFortune([
    {
      id: 'sheet-imported',
      name: 'Imported',
      row: 100_000,
      column: 10,
      data,
    },
  ]);
  const projectedData = projected[0]?.data;
  expect(projectedData).toBe(data);
  expect(projectedData?.[0]).toBe(sourceRow);
  expect(projectedData?.[0]?.[0]).toBe(sourceCell);
  expect(Object.isFrozen(data)).toBe(true);
  expect(Object.isFrozen(sourceRow)).toBe(true);
  expect(Object.isFrozen(sourceCell)).toBe(true);
  expect(
    (
      data as typeof data & {
        __a3sShownCommentCells?: readonly unknown[];
      }
    )[SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY],
  ).toEqual([]);
  expect(Object.keys(data)).not.toContain(
    SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY,
  );

  const edited = produce(projected, (draft) => {
    const cell = draft[0]?.data?.[0]?.[0];
    if (cell) cell.v = 'Edited';
  });
  expect(edited[0]?.data?.[0]?.[0]?.v).toBe('Edited');
  expect(data[0]?.[0]?.v).toBe('Imported');
  expect(
    (
      edited[0]?.data as
        | (NonNullable<(typeof edited)[number]['data']> & {
            __a3sShownCommentCells?: readonly unknown[];
          })
        | undefined
    )?.[SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY],
  ).toBeUndefined();
});

test('reconciles authenticated cell operations with row-level copy-on-write', () => {
  const visibleComment = {
    height: null,
    isShow: true,
    left: null,
    top: null,
    value: 'Visible',
    width: null,
  };
  const sourceCell = freezeImportedSpreadsheetCell({
    f: '=1+1',
    lo: 0,
    m: '2',
    ps: visibleComment,
    v: 2,
  });
  const untouchedCell = freezeImportedSpreadsheetCell({
    m: 'Untouched',
    v: 'Untouched',
  });
  const sourceRow = [sourceCell];
  const untouchedRow = [untouchedCell];
  const data = [sourceRow, untouchedRow];
  data.length = 100_000;
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 10,
    formulaCells: [{ column: 0, row: 0 }],
    fortuneReady: true,
    populatedCellCount: 2,
    protectionCellKey: '0_0:0:',
    rowCount: 100_000,
    shownCommentCells: [{ c: 0, r: 0 }],
  });
  const source = {
    column: 10,
    data,
    id: 'sheet-imported',
    name: 'Imported',
    row: 100_000,
  };
  const changed = produce([source], (draft) => {
    const cell = draft[0]?.data?.[0]?.[0] as
      | (NonNullable<(typeof draft)[0]>['data'][number][number] & {
          hi?: number;
        })
      | undefined;
    if (!cell) throw new Error('Expected the imported formula cell.');
    delete cell.f;
    delete cell.ps;
    cell.hi = 1;
    cell.lo = 1;
    cell.m = '3';
    cell.v = 3;
  });

  const controlled = spreadsheetSheetsFromFortune(
    changed,
    [source],
    [
      {
        id: 'sheet-imported',
        op: 'replace',
        path: ['data', 0, 0],
        value: changed[0]?.data?.[0]?.[0],
      },
    ],
  );
  const controlledData = controlled[0]?.data;
  const profile = spreadsheetMatrixProfile(controlledData);

  expect(controlledData).not.toBe(data);
  expect(controlledData?.[0]).not.toBe(sourceRow);
  expect(controlledData?.[1]).toBe(untouchedRow);
  expect(controlledData?.[0]?.[0]).not.toBe(changed[0]?.data?.[0]?.[0]);
  expect(controlledData?.[0]?.[0]).toMatchObject({ hi: 1, lo: 1, v: 3 });
  expect(profile).toMatchObject({
    columnCount: 10,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 2,
    protectionCellKey: '0_0:1:1',
    rowCount: 100_000,
    shownCommentCells: [],
  });
  expect(profile?.historyRoot).toBe(
    spreadsheetMatrixProfile(data)?.historyRoot,
  );
  expect(profile?.historyState).not.toBe(
    spreadsheetMatrixProfile(data)?.historyState,
  );
  expect(
    sameSpreadsheetHistoryContent(
      { sheets: [source], type: 'spreadsheet' },
      { sheets: controlled, type: 'spreadsheet' },
    ),
  ).toBe(false);
  expect(spreadsheetSheetsForFortune(controlled)[0]?.data).toBe(controlledData);
});

test('reuses incremental history state for formula result-only changes', () => {
  const data = [[freezeImportedSpreadsheetCell({ f: '=1+1', m: '2', v: 2 })]];
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 1,
    formulaCells: [{ column: 0, row: 0 }],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 1,
    shownCommentCells: [],
  });
  const source = { data, id: 'sheet-formula', name: 'Formula' };
  const changed = produce([source], (draft) => {
    const cell = draft[0]?.data?.[0]?.[0];
    if (!cell) throw new Error('Expected the formula cell.');
    cell.ct = { fa: 'General', t: 'n' };
    cell.m = '999';
    cell.v = 999;
  });
  const controlled = spreadsheetSheetsFromFortune(
    changed,
    [source],
    [
      {
        id: 'sheet-formula',
        op: 'replace',
        path: ['data', 0, 0],
        value: changed[0]?.data?.[0]?.[0],
      },
    ],
  );

  expect(spreadsheetMatrixProfile(controlled[0]?.data)?.historyState).toBe(
    spreadsheetMatrixProfile(data)?.historyState,
  );
  expect(
    sameSpreadsheetHistoryContent(
      { sheets: [source], type: 'spreadsheet' },
      { sheets: controlled, type: 'spreadsheet' },
    ),
  ).toBe(true);
});

test('preserves consecutive controlled edits through incremental operations', () => {
  const sourceCell = freezeImportedSpreadsheetCell({ m: 'First', v: 'First' });
  const data = [[sourceCell]];
  data.length = 100_000;
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 10,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 100_000,
    shownCommentCells: [],
  });
  const source = {
    column: 10,
    data,
    id: 'sheet-imported',
    name: 'Imported',
    row: 100_000,
  };
  const firstChanged = produce([source], (draft) => {
    const row = [] as NonNullable<(typeof source)['data']>[number];
    row.length = 10;
    row[9] = { m: 'Second', v: 'Second' };
    if (draft[0]?.data) draft[0].data[99_999] = row;
  });
  const firstControlled = spreadsheetSheetsFromFortune(
    firstChanged,
    [source],
    [
      {
        id: 'sheet-imported',
        op: 'add',
        path: ['data', 99_999, 9],
        value: firstChanged[0]?.data?.[99_999]?.[9],
      },
    ],
  );
  const secondChanged = produce(firstControlled, (draft) => {
    const cell = draft[0]?.data?.[0]?.[0];
    if (!cell) throw new Error('Expected the first controlled cell.');
    cell.m = 'Updated';
    cell.v = 'Updated';
  });
  const secondControlled = spreadsheetSheetsFromFortune(
    secondChanged,
    firstControlled,
    [
      {
        id: 'sheet-imported',
        op: 'replace',
        path: ['data', 0, 0],
        value: secondChanged[0]?.data?.[0]?.[0],
      },
    ],
  );

  expect(firstControlled[0]?.data?.[99_999]?.[9]?.v).toBe('Second');
  expect(secondControlled[0]?.data?.[0]?.[0]?.v).toBe('Updated');
  expect(secondControlled[0]?.data?.[99_999]?.[9]?.v).toBe('Second');
  expect(spreadsheetMatrixProfile(secondControlled[0]?.data)).toMatchObject({
    fortuneReady: true,
    populatedCellCount: 2,
  });
});

test('reconstructs formula-bar rich text through the unregistered matrix fallback', () => {
  const colorOrigin = {
    baseColor: '#4472c4',
    index: 4,
    kind: 'theme',
    renderedColor: '#4472c4',
  } as const;
  const source = {
    data: [
      [
        {
          ct: {
            fa: 'General',
            t: 'inlineStr',
            s: [
              {
                a3sXlsxColorOrigin: colorOrigin,
                bl: 1,
                fc: '#4472c4',
                v: 'Native ',
              },
              { it: 1, v: 'rich text' },
            ],
          },
          v: 'Native rich text',
        },
      ],
    ],
    id: 'sheet-rich-text',
    name: 'Rich text',
  };
  const changed = [
    {
      ...source,
      data: [[{ ct: { fa: 'General', t: 'g' }, v: 'NatXive rich text' }]],
    },
  ];

  const controlled = spreadsheetSheetsFromFortune(
    changed,
    [source],
    [
      {
        id: 'sheet-rich-text',
        op: 'replace',
        path: ['data', 0, 0],
        value: changed[0]?.data?.[0]?.[0],
      },
    ],
  );

  expect(controlled[0]?.data?.[0]?.[0]).toMatchObject({
    ct: {
      t: 'inlineStr',
      s: [
        {
          a3sXlsxColorOrigin: colorOrigin,
          bl: 1,
          fc: '#4472c4',
          v: 'NatXive ',
        },
        { it: 1, v: 'rich text' },
      ],
    },
    v: 'NatXive rich text',
  });
});

test('restores text-stable rich text when Fortune omits cell operations', () => {
  const colorOrigin = {
    baseColor: '#4472c4',
    index: 4,
    kind: 'theme',
    renderedColor: '#4472c4',
  } as const;
  const source = {
    data: [
      [
        {
          ct: {
            fa: 'General',
            t: 'inlineStr',
            s: [
              {
                a3sXlsxColorOrigin: colorOrigin,
                bl: 1,
                fc: '#4472c4',
                v: 'Native ',
              },
              { it: 1, v: 'rich text' },
            ],
          },
          v: 'Native rich text',
        },
      ],
    ],
    id: 'sheet-rich-text',
    name: 'Rich text',
  };
  const changed = [
    {
      ...source,
      data: [[{ ct: { fa: 'General', t: 'g' }, v: 'Native rich text' }]],
    },
  ];

  const controlled = spreadsheetSheetsFromFortune(changed, [source], []);

  expect(controlled[0]?.data?.[0]?.[0]).toEqual(source.data[0]?.[0]);
  expect(
    sameSpreadsheetHistoryContent(
      { sheets: [source], type: 'spreadsheet' },
      { sheets: controlled, type: 'spreadsheet' },
    ),
  ).toBe(true);
});

test('does not infer changed rich text when Fortune omits cell operations', () => {
  const source = {
    data: [
      [
        {
          ct: { t: 'inlineStr', s: [{ bl: 1, v: 'Source' }] },
          v: 'Source',
        },
      ],
    ],
    id: 'sheet-rich-text',
    name: 'Rich text',
  };
  const changed = [{ ...source, data: [[{ v: 'Changed' }]] }];

  expect(
    spreadsheetSheetsFromFortune(changed, [source], [])[0]?.data?.[0]?.[0],
  ).toEqual({ v: 'Changed' });
});

test('restores rich-text semantic metadata through authenticated cell operations', () => {
  const colorOrigin = {
    baseColor: '#4472c4',
    index: 4,
    kind: 'theme',
    renderedColor: '#4472c4',
  } as const;
  const sourceCell = freezeImportedSpreadsheetCell({
    ct: {
      t: 'inlineStr',
      s: [
        {
          a3sXlsxColorOrigin: colorOrigin,
          fc: '#4472c4',
          v: 'Blue',
        },
      ],
    },
    v: 'Blue',
  });
  const data = [[sourceCell]];
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 1,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 1,
    shownCommentCells: [],
  });
  const source = { data, id: 'sheet-rich-text', name: 'Rich text' };
  const changed = produce([source], (draft) => {
    const cell = draft[0]?.data?.[0]?.[0];
    if (!cell) throw new Error('Expected the rich-text cell.');
    delete cell.v;
    cell.ct = {
      t: 'inlineStr',
      s: [{ fc: '#4472c4', v: 'BluXe' }],
    };
  });

  const controlled = spreadsheetSheetsFromFortune(
    changed,
    [source],
    [
      {
        id: 'sheet-rich-text',
        op: 'replace',
        path: ['data', 0, 0],
        value: changed[0]?.data?.[0]?.[0],
      },
    ],
  );

  expect(controlled[0]?.data?.[0]?.[0]).toMatchObject({
    ct: {
      t: 'inlineStr',
      s: [
        {
          a3sXlsxColorOrigin: colorOrigin,
          fc: '#4472c4',
          v: 'BluXe',
        },
      ],
    },
    v: 'BluXe',
  });
});

test('projects authenticated formatted paste through registered and fallback matrices', () => {
  for (const register of [false, true]) {
    const sourceCell = freezeImportedSpreadsheetCell({
      ct: {
        t: 'inlineStr',
        s: [{ bl: 1, fc: '#4472c4', v: 'Blue' }],
      },
      v: 'Blue',
    });
    const data = [[sourceCell]];
    if (register) {
      registerImportedSpreadsheetMatrix(data, {
        columnCount: 1,
        formulaCells: [],
        fortuneReady: true,
        populatedCellCount: 1,
        protectionCellKey: '',
        rowCount: 1,
        shownCommentCells: [],
      });
    }
    const source = { data, id: `sheet-paste-${register}`, name: 'Paste' };
    stageSpreadsheetRichTextPaste(source, 0, 0, 'Blue', {
      end: 4,
      runs: [{ fc: '#c00000', it: 1, v: ' red' }],
      start: 4,
      text: ' red',
    });
    const changed = [{ ...source, data: [[{ v: 'Blue red' }]] }];

    const controlled = spreadsheetSheetsFromFortune(
      changed,
      [source],
      [
        {
          id: source.id,
          op: 'replace',
          path: ['data', 0, 0],
          value: changed[0]?.data?.[0]?.[0],
        },
      ],
    );

    expect(controlled[0]?.data?.[0]?.[0]).toMatchObject({
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, fc: '#4472c4', v: 'Blue' },
          { fc: '#c00000', it: 1, v: ' red' },
        ],
      },
      v: 'Blue red',
    });
  }
});

test('does not infer rich text across a structural operation batch', () => {
  const source = {
    data: [
      [
        {
          ct: { t: 'inlineStr', s: [{ bl: 1, v: 'Source' }] },
          v: 'Source',
        },
      ],
    ],
    id: 'sheet-rich-text',
    name: 'Rich text',
  };
  const changed = [{ ...source, data: [[{ v: 'Moved' }]] }];

  const controlled = spreadsheetSheetsFromFortune(
    changed,
    [source],
    [
      {
        id: 'sheet-rich-text',
        op: 'insertRowCol',
        path: [],
        value: {
          count: 1,
          direction: 'rightbottom',
          id: 'sheet-rich-text',
          index: 0,
          type: 'row',
        },
      },
    ],
  );

  expect(controlled[0]?.data?.[0]?.[0]).toEqual({ v: 'Moved' });
});

test('falls back to full reconciliation for structural operation batches', () => {
  const sourceCell = freezeImportedSpreadsheetCell({
    m: 'Before',
    v: 'Before',
  });
  const data = [[sourceCell]];
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 1,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 1,
    shownCommentCells: [],
  });
  const source = { data, id: 'sheet-imported', name: 'Imported' };
  const changed = produce([source], (draft) => {
    const cell = draft[0]?.data?.[0]?.[0];
    if (cell) {
      cell.m = 'After';
      cell.v = 'After';
    }
  });
  const controlled = spreadsheetSheetsFromFortune(
    changed,
    [source],
    [
      {
        id: 'sheet-imported',
        op: 'insertRowCol',
        path: [],
        value: {
          count: 1,
          direction: 'rightbottom',
          id: 'sheet-imported',
          index: 0,
          type: 'row',
        },
      },
    ],
  );

  expect(controlled[0]?.data).not.toBe(data);
  expect(controlled[0]?.data?.[0]).not.toBe(data[0]);
  expect(controlled[0]?.data?.[0]?.[0]?.v).toBe('After');
  expect(spreadsheetMatrixProfile(controlled[0]?.data)).toBeUndefined();
});

test('indexes initially visible comments without exposing enumerable metadata', () => {
  const projected = spreadsheetSheetsForFortune([
    {
      id: 'sheet-comments',
      name: 'Comments',
      data: [
        [
          {
            v: 'Commented',
            ps: {
              left: null,
              top: null,
              width: null,
              height: null,
              value: 'Visible',
              isShow: true,
            },
          },
        ],
      ],
    },
  ]);
  const data = projected[0]?.data as
    | (NonNullable<(typeof projected)[number]['data']> & {
        __a3sShownCommentCells?: readonly unknown[];
      })
    | undefined;

  expect(data?.[SPREADSHEET_SHOWN_COMMENT_CELLS_PROPERTY]).toEqual([
    { c: 0, r: 0 },
  ]);
  expect(Object.keys(data ?? [])).toEqual(['0']);
});

test('uses the imported protection summary without enumerating cells again', () => {
  let enumerations = 0;
  const row = [freezeImportedSpreadsheetCell({ v: 1 })];
  const data = new Proxy([row], {
    ownKeys(target) {
      enumerations += 1;
      return Reflect.ownKeys(target);
    },
  });
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 1,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 1,
    shownCommentCells: [],
  });
  enumerations = 0;

  expect(
    spreadsheetProtectionKey([
      { id: 'sheet-imported', name: 'Imported', data },
    ]),
  ).toContain('sheet-imported');
  expect(enumerations).toBe(0);
});

test('compares workbook state without JSON serialization', () => {
  const sentinel = {
    toJSON(): never {
      throw new Error('Workbook comparison serialized its input.');
    },
  };
  const sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    data: [[{ v: 1, m: '1' }]],
    sentinel,
  } as WorkSpreadsheetContent['sheets'][number] & { sentinel: typeof sentinel };

  expect(sameSpreadsheetWorkbookState([sheet], [{ ...sheet }])).toBe(true);
});

test('accepts shared Fortune matrices without enumerating their cells', () => {
  let enumerations = 0;
  const source = [[{ v: 'Shared', m: 'Shared' }]];
  const data = new Proxy(source, {
    ownKeys(target) {
      enumerations += 1;
      return Reflect.ownKeys(target);
    },
  });
  const sheet = { id: 'sheet-1', name: 'Shared', data };

  expect(sameSpreadsheetWorkbookState([sheet], [{ ...sheet }])).toBe(true);
  expect(enumerations).toBe(0);
});

test('compares only operation coordinates in a large changed matrix', () => {
  let enumerations = 0;
  const renderedData = [[{ m: 'First', v: 'First' }]];
  renderedData.length = 100_000;
  renderedData[99_999] = [];
  renderedData[99_999][9] = { m: 'Before', v: 'Before' };
  const changedStorage = renderedData.slice();
  changedStorage[99_999] = renderedData[99_999].slice();
  changedStorage[99_999][9] = { m: 'After', v: 'After' };
  const changedData = new Proxy(changedStorage, {
    ownKeys(target) {
      enumerations += 1;
      return Reflect.ownKeys(target);
    },
  });
  const renderedDataProxy = new Proxy(renderedData, {
    ownKeys(target) {
      enumerations += 1;
      return Reflect.ownKeys(target);
    },
  });
  const rendered = [{ id: 'sheet-1', name: 'Large', data: renderedDataProxy }];
  const changed = [{ id: 'sheet-1', name: 'Large', data: changedData }];

  expect(
    sameSpreadsheetWorkbookStateAfterOperations(changed, rendered, [
      {
        id: 'sheet-1',
        op: 'replace',
        path: ['data', 99_999, 9, 'v'],
        value: 'After',
      },
    ]),
  ).toBe(false);
  expect(enumerations).toBe(0);
});

test('accepts formula result-only operations without a full matrix scan', () => {
  const rendered = [
    {
      data: [[{ f: '=1+1', v: 2 }]],
      id: 'sheet-1',
      name: 'Formula',
    },
  ];
  const changed = [
    {
      data: [[{ ct: { fa: 'General', t: 'n' }, f: '=1+1', m: '2', v: 2 }]],
      id: 'sheet-1',
      name: 'Formula',
    },
  ];

  expect(
    sameSpreadsheetWorkbookStateAfterOperations(changed, rendered, [
      {
        id: 'sheet-1',
        op: 'replace',
        path: ['data', 0, 0],
        value: changed[0]?.data[0]?.[0],
      },
    ]),
  ).toBe(true);
});

test('rejects sheet metadata changes before enumerating large cell data', () => {
  let enumerations = 0;
  const data = new Proxy([[{ v: 'Value', m: 'Value' }]], {
    ownKeys(target) {
      enumerations += 1;
      return Reflect.ownKeys(target);
    },
  });

  expect(
    sameSpreadsheetWorkbookState(
      [{ id: 'sheet-1', name: 'Renamed', data }],
      [
        {
          id: 'sheet-1',
          name: 'Original',
          data: [[{ v: 'Value', m: 'Value' }]],
        },
      ],
    ),
  ).toBe(false);
  expect(enumerations).toBe(0);
});

test('ignores Fortune formula result normalization in workbook comparisons', () => {
  const rendered: WorkSpreadsheetContent['sheets'] = [
    {
      id: 'sheet-1',
      name: 'Formula',
      celldata: [{ r: 0, c: 0, v: { f: '=1+1', v: 2 } }],
    },
  ];
  const normalized: WorkSpreadsheetContent['sheets'] = [
    {
      ...rendered[0],
      celldata: [
        {
          r: 0,
          c: 0,
          v: { f: '=1+1', v: 2, m: '2', ct: { fa: 'General', t: 'n' } },
        },
      ],
    },
  ];

  expect(sameSpreadsheetWorkbookState(normalized, rendered)).toBe(true);
  const changedFormula: WorkSpreadsheetContent['sheets'] = [
    {
      ...normalized[0],
      celldata: [{ r: 0, c: 0, v: { f: '=2+2', v: 4, m: '4' } }],
    },
  ];
  expect(sameSpreadsheetWorkbookState(changedFormula, rendered)).toBe(false);
});

test('ignores Fortune default zoom normalization in workbook comparisons', () => {
  const rendered: WorkSpreadsheetContent['sheets'] = [
    {
      id: 'sheet-1',
      name: 'Default zoom',
      celldata: [{ r: 0, c: 0, v: { v: 'Value', m: 'Value' } }],
    },
  ];
  const normalized: WorkSpreadsheetContent['sheets'] = [
    { ...rendered[0], zoomRatio: 1 },
  ];

  expect(sameSpreadsheetWorkbookState(normalized, rendered)).toBe(true);
  expect(
    sameSpreadsheetWorkbookState(
      [{ ...rendered[0], zoomRatio: 1.25 }],
      rendered,
    ),
  ).toBe(false);
});

test('ignores Fortune calculation bookkeeping in workbook comparisons', () => {
  const rendered: WorkSpreadsheetContent['sheets'] = [
    {
      id: 'sheet-1',
      name: 'Calculated',
      celldata: [{ r: 0, c: 0, v: { f: '=1+1', v: 2 } }],
    },
  ];
  const normalized: WorkSpreadsheetContent['sheets'] = [
    {
      ...rendered[0],
      calcChain: [{ r: 0, c: 0, id: 'sheet-1' }],
      dynamicArray_compute: { '0_0': { r: 0, c: 0 } },
    },
  ];

  expect(sameSpreadsheetWorkbookState(normalized, rendered)).toBe(true);
});

test('treats Fortune visible-row caches as the same sparse workbook state', () => {
  const source = {
    id: 'sheet-1',
    name: 'Sparse',
    row: 1_000_000,
    column: 16_384,
    data: sparseMatrix(),
  };
  const projected = spreadsheetSheetsForFortune([source]);
  const fortuneData: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['data']
  > = [];
  fortuneData.length = 1_000_000;
  for (let row = 0; row < 100; row += 1) {
    fortuneData[row] = [];
    fortuneData[row].length = 16_384;
  }
  fortuneData[0][0] = { v: 'Anchor', m: 'Anchor' };
  fortuneData[999_999] = [];
  fortuneData[999_999].length = 16_384;
  fortuneData[999_999][16_383] = { v: 'Tail', m: 'Tail' };
  const projectedSheet = projected[0];
  if (!projectedSheet) throw new Error('Projected sheet is unavailable.');
  const changed = [
    { ...projectedSheet, celldata: undefined, data: fortuneData },
  ];

  expect(sameSpreadsheetWorkbookState(changed, projected)).toBe(true);
  const controlled = spreadsheetSheetsFromFortune(changed, [source]);
  expect(Object.keys(controlled[0]?.data ?? [])).toEqual(['0', '999999']);
});

test('materializes only a newly edited far row at the Fortune boundary', () => {
  const source = {
    id: 'sheet-1',
    name: 'Sparse',
    row: 1_000_000,
    column: 16_384,
    data: sparseMatrix(),
  };
  const fortuneData: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['data']
  > = [];
  fortuneData.length = 1_000_000;
  for (let row = 0; row < 100; row += 1) {
    fortuneData[row] = [];
    fortuneData[row].length = 16_384;
  }
  fortuneData[0][0] = source.data[0]?.[0] ?? null;
  fortuneData[800_000] = [];
  fortuneData[800_000].length = 16_384;
  fortuneData[800_000][12_000] = { v: 'Edited', m: 'Edited' };
  fortuneData[999_999] = [];
  fortuneData[999_999].length = 16_384;
  fortuneData[999_999][16_383] = source.data[999_999]?.[16_383] ?? null;

  const controlled = spreadsheetSheetsFromFortune(
    [{ ...source, data: fortuneData }],
    [source],
  );
  expect(Object.keys(controlled[0]?.data ?? [])).toEqual([
    '0',
    '800000',
    '999999',
  ]);
});

test('detects values added inside sparse history arrays', () => {
  const left = sparseHistoryWorkbook();
  const right = sparseHistoryWorkbook();
  const rightRow = right.sheets[0]?.data?.[0];
  if (!rightRow) throw new Error('Expected a sparse history row.');
  rightRow[1] = { v: 'new', m: 'new' };

  expect(sameSpreadsheetHistoryContent(left, right)).toBe(false);
});

function sparseMatrix(): NonNullable<
  WorkSpreadsheetContent['sheets'][number]['data']
> {
  const data: NonNullable<WorkSpreadsheetContent['sheets'][number]['data']> =
    [];
  data.length = 1_000_000;
  data[0] = [];
  data[0][0] = { v: 'Anchor', m: 'Anchor' };
  data[999_999] = [];
  data[999_999][16_383] = { v: 'Tail', m: 'Tail' };
  return data;
}

function sparseHistoryWorkbook(): WorkSpreadsheetContent {
  const row: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['data']
  >[number] = [];
  row.length = 2;
  row[0] = { v: 'existing', m: 'existing' };
  return {
    type: 'spreadsheet',
    sheets: [{ id: 'sheet-1', name: 'Sheet 1', data: [row] }],
  };
}
