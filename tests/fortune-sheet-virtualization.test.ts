import {
  api,
  checkCF,
  compute,
  type CellMatrix,
  type Context,
  defaultContext,
  getDataVerificationItem,
  locale,
  setCellValue,
  updateContextWithSheetData,
  updateSheet,
} from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';

test('loads cached formula help only when the active locale requests it', () => {
  const english = locale({ lang: 'en' } as Context);
  const chinese = locale({ lang: 'zh' } as Context);

  expect(Object.getOwnPropertyDescriptor(english, 'functionlist')?.get).toEqual(
    expect.any(Function),
  );
  expect(Object.getOwnPropertyDescriptor(chinese, 'functionlist')?.get).toEqual(
    expect.any(Function),
  );
  expect(english.functionlist).toBe(english.functionlist);
  expect(chinese.functionlist).toBe(chinese.functionlist);
  expect(english.functionlist.some((item) => item.n === 'SUM')).toBe(true);
  expect(chinese.functionlist.some((item) => item.n === 'SUM')).toBe(true);
  expect(locale({ lang: 'es' } as Context)).toBe(english);
  expect(locale({ lang: 'hi' } as Context)).toBe(english);
  expect(locale({ lang: 'ru' } as Context)).toBe(english);
  expect(locale({ lang: 'zh-TW' } as Context)).toBe(chinese);
});

test('converts Fortune celldata into a sparse logical matrix', () => {
  const data = api.celldataToData(
    [
      { r: 0, c: 0, v: { v: 'Anchor', m: 'Anchor' } },
      { r: 1_999, c: 499, v: { v: 'Tail', m: 'Tail' } },
    ],
    2_000,
    500,
  );

  expect(data).not.toBeNull();
  expect(data).toHaveLength(2_000);
  expect(Object.keys(data ?? [])).toEqual(['0', '1999']);
  expect(data?.[0]).toHaveLength(500);
  expect(Object.keys(data?.[0] ?? [])).toEqual(['0']);
  expect(Object.keys(data?.[1_999] ?? [])).toEqual(['499']);
});

test('converts sparse logical matrices by visiting populated cells only', () => {
  const data: CellMatrix = [];
  data.length = 1_000_000;
  data[0] = [];
  data[0].length = 16_384;
  data[0][0] = { v: 'Anchor', m: 'Anchor' };
  data[999_999] = [];
  data[999_999][16_383] = { v: 'Tail', m: 'Tail' };

  expect(api.dataToCelldata(data)).toEqual([
    { r: 0, c: 0, v: { v: 'Anchor', m: 'Anchor' } },
    { r: 999_999, c: 16_383, v: { v: 'Tail', m: 'Tail' } },
  ]);
});

test('resolves compact data-validation ranges without materializing cells', () => {
  const compactItem = {
    type: 'dropdown',
    type2: '',
    rangeTxt: 'A1:XFD1048576',
    value1: 'Ready,Blocked',
    value2: '',
    validity: '',
    remote: false,
    prohibitInput: true,
    hintShow: false,
    hintValue: '',
  };
  const directItem = { ...compactItem, value1: 'Direct' };
  const context = {
    currentSheetId: 'sheet-1',
    luckysheetfile: [
      {
        id: 'sheet-1',
        name: 'Sparse',
        dataVerification: { '999999_16383': directItem },
        dataValidationRanges: [
          {
            ranges: [{ row: [0, 1_048_575], column: [0, 16_383] }],
            item: compactItem,
          },
        ],
      },
    ],
  } as unknown as Context;

  expect(getDataVerificationItem(context, 800_000, 12_000)).toBe(compactItem);
  expect(getDataVerificationItem(context, 999_999, 16_383)).toBe(directItem);
  expect(getDataVerificationItem(context, 1_048_576, 0)).toBeUndefined();
});

test('computes conditional formatting from populated sparse cells only', () => {
  const data: CellMatrix = [];
  data.length = 1_048_576;
  data[0] = [];
  data[0].length = 16_384;
  data[0][0] = { v: 1, m: '1' };
  data[1_048_575] = [];
  data[1_048_575][16_383] = { v: 20, m: '20' };
  const computeMap = compute(
    {} as Context,
    [
      {
        type: 'default',
        cellrange: [{ row: [0, 1_048_575], column: [0, 16_383] }],
        format: { textColor: '#ffffff', cellColor: '#cc0000' },
        conditionName: 'greaterThan',
        conditionValue: [10],
      },
    ],
    data,
  );

  expect(Object.keys(computeMap)).toEqual(['1048575_16383']);
  expect(checkCF(1_048_575, 16_383, computeMap)).toMatchObject({
    textColor: '#ffffff',
    cellColor: '#cc0000',
  });
  expect(checkCF(500_000, 8_000, computeMap)).toBeNull();
});

test('updates a maximum worksheet without materializing every logical row', () => {
  const data: CellMatrix = [];
  data.length = 1_048_576;
  data[1_048_575] = [];
  data[1_048_575].length = 16_384;
  data[1_048_575][16_383] = { v: 'Tail', m: 'Tail' };
  const context = {
    defaultcolumnNum: 26,
    defaultrowNum: 60,
    formulaCache: { formulaCellInfoMap: {} },
    luckysheetfile: [],
  } as unknown as Context;

  updateSheet(context, [
    {
      id: 'sheet-1',
      name: 'Sparse',
      row: 1_048_576,
      column: 16_384,
      data,
    },
  ]);

  const updated = context.luckysheetfile[0]?.data;
  expect(updated).toHaveLength(1_048_576);
  expect(Object.keys(updated ?? [])).toHaveLength(101);
  expect(updated?.[1_048_575]?.[16_383]).toMatchObject({ v: 'Tail' });
});

test('adopts an authenticated frozen matrix without cloning populated cells', () => {
  const data: CellMatrix = [[Object.freeze({ v: 1, m: '1' })]];
  Object.freeze(data[0]);
  Object.defineProperty(data, '__a3sShownCommentCells', {
    configurable: false,
    enumerable: false,
    value: Object.freeze([]),
    writable: false,
  });
  Object.freeze(data);
  const context = {
    defaultcolumnNum: 26,
    defaultrowNum: 60,
    formulaCache: { formulaCellInfoMap: {} },
    luckysheetfile: [{ id: 'sheet-1', name: 'Before', data: [] }],
  } as unknown as Context;

  updateSheet(context, [
    {
      id: 'sheet-1',
      name: 'After',
      data,
    },
  ]);

  expect(context.luckysheetfile[0]?.data).toBe(data);
  expect(context.luckysheetfile[0]?.name).toBe('After');
});

test('calculates logical row and column offsets with compact local arrays', () => {
  const context = defaultContext({
    globalCache: {} as never,
    cellInput: { current: null },
    fxInput: { current: null },
    canvas: { current: null },
    cellArea: { current: null },
    workbookContainer: { current: null },
  });
  context.config = {
    rowlen: { 1: 30 },
    rowhidden: { 2: 0 },
    columnlen: { 1: 100 },
    colhidden: { 2: 0 },
  };
  const data: CellMatrix = [];
  data.length = 4;
  data[0] = [];
  data[0].length = 4;

  updateContextWithSheetData(context, data);

  expect(context.visibledatarow).toEqual([20, 51, 51, 71]);
  expect(context.rh_height).toBe(151);
  expect(context.visibledatacolumn).toEqual([74, 175, 175, 249]);
  expect(context.ch_width).toBe(369);
});

test('materializes only the edited far row for direct and API cell writes', () => {
  const data: CellMatrix = [];
  data.length = 1_048_576;
  data[0] = [];
  data[0].length = 16_384;
  const sheet = {
    id: 'sheet-1',
    name: 'Sparse',
    row: 1_048_576,
    column: 16_384,
    data,
  };
  const context = {
    currentSheetId: 'sheet-1',
    defaultcolumnNum: 16_384,
    formulaCache: {
      execFunctionExist: [],
      execFunctionGlobalData: null,
      formulaCellInfoMap: {},
    },
    luckysheetfile: [sheet],
  } as unknown as Context;

  expect(api.getCellValue(context, 900_000, 11_999, { id: 'sheet-1' })).toBe(
    null,
  );
  setCellValue(context, 900_000, 12_000, data, 'Direct');
  api.setCellValue(context, 900_001, 12_001, { v: 'API', m: 'API' }, null, {
    id: 'sheet-1',
  });

  expect(Object.keys(data)).toEqual(['0', '900000', '900001']);
  expect(data[900_000]?.[12_000]).toMatchObject({ v: 'Direct' });
  expect(data[900_001]?.[12_001]).toMatchObject({ v: 'API' });
});

test('writes cross-sheet formula results to the formula sheet', () => {
  const emptyRef = { current: null };
  const context = defaultContext({
    globalCache: {} as never,
    cellInput: emptyRef,
    fxInput: emptyRef,
    canvas: emptyRef,
    cellArea: emptyRef,
    workbookContainer: emptyRef,
  });
  context.currentSheetId = 'sheet-input';
  context.luckysheetfile = [
    {
      id: 'sheet-input',
      name: 'Inputs',
      status: 1,
      data: [
        [
          { v: 'Revenue', m: 'Revenue' },
          { v: 10, m: '10' },
        ],
      ],
    },
    {
      id: 'sheet-results',
      name: 'Results',
      status: 0,
      data: [
        [
          { v: 'Total', m: 'Total' },
          { f: '=Inputs!B1', v: 0 },
        ],
      ],
    },
  ];

  api.calculateFormula(context, 'sheet-results');

  expect(context.luckysheetfile[0]?.data?.[0]?.[1]).toMatchObject({ v: 10 });
  expect(context.luckysheetfile[1]?.data?.[0]?.[1]).toMatchObject({
    f: '=Inputs!B1',
    v: 10,
  });
});
