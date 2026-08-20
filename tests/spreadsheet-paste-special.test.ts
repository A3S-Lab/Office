import { describe, expect, test } from '@rstest/core';
import type { Cell } from '@fortune-sheet/core';
import {
  applySpreadsheetPasteSpecial,
  captureSpreadsheetClipboardSnapshot,
  createSpreadsheetTextClipboardSnapshot,
  spreadsheetPasteSpecialModeAvailable,
  spreadsheetPasteSpecialValidationError,
  type SpreadsheetPasteSpecialOptions,
} from '../src/internal/features/work/editors/spreadsheet-paste-special';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../src/internal/features/work/work-types';

const allOptions: SpreadsheetPasteSpecialOptions = {
  content: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
};

describe('Spreadsheet Paste Special model', () => {
  test('captures rich cells, borders, validation, protection, widths, and merges', () => {
    const content = workbook();
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [0, 1], column: [0, 1] },
      '10\t20\n30\tNorth',
    );

    expect(snapshot).not.toBeNull();
    expect(snapshot).toMatchObject({
      kind: 'rich',
      plainText: '10\t20\n30\tNorth',
      rowCount: 2,
      columnCount: 2,
      columnWidths: [120, 84],
      merges: [{ row: 1, column: 0, rowSpan: 1, columnSpan: 2 }],
    });
    expect(snapshot?.cells[0]?.[0]).toMatchObject({
      cell: { v: 10, m: '10', bg: '#fff2cc', bl: 1 },
      borders: {
        bottom: { color: '#4472c4', style: '1' },
      },
      validation: { type: 'number', value1: '0', value2: '100' },
      protection: { locked: false, hidden: true },
      hyperlink: { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
    });
    expect(snapshot?.cells[0]?.[1].cell?.f).toBe('=$A1+A$1');
  });

  test('pastes all rich metadata and translates relative formula references atomically', () => {
    const content = workbook();
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [0, 0], column: [0, 1] },
      '10\t20',
    );
    if (!snapshot) throw new Error('Expected a rich clipboard snapshot.');

    const result = applySpreadsheetPasteSpecial(content, {
      snapshot,
      targetSheetId: 'target',
      targetSelection: { row: [2, 2], column: [2, 2] },
      options: allOptions,
    });

    expect(result).not.toBeNull();
    expect(result?.targetRange).toEqual({ row: [2, 2], column: [2, 3] });
    const target = result?.content.sheets[1];
    expect(target?.data?.[2]?.[2]).toMatchObject({
      v: 10,
      bg: '#fff2cc',
      bl: 1,
      lo: 0,
    });
    expect((target?.data?.[2]?.[2] as Cell & { hi?: number })?.hi).toBe(1);
    expect(target?.data?.[2]?.[3]?.f).toBe('=$A3+C$1');
    expect(target?.hyperlink?.['2_2']).toEqual({
      linkType: 'webpage',
      linkAddress: 'https://a3s.dev',
    });
    expect(target?.dataVerification?.['2_2']).toMatchObject({
      type: 'number',
      value1: '0',
      value2: '100',
    });
    expect(target?.config?.borderInfo).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          rangeType: 'cell',
          value: expect.objectContaining({
            row_index: 2,
            col_index: 2,
            b: { color: '#4472c4', style: '1' },
          }),
        }),
      ]),
    );
    expect(target?.luckysheet_select_save?.at(-1)).toMatchObject({
      row: [2, 2],
      column: [2, 3],
      row_focus: 2,
      column_focus: 2,
    });
    expect(content.sheets[1]?.data?.[2]?.[2]?.v).toBe(100);
  });

  test('writes translated merge cells and canonical sheet merge metadata together', () => {
    const content = workbook();
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [1, 1], column: [0, 1] },
      '30\t',
    );
    if (!snapshot) throw new Error('Expected a rich clipboard snapshot.');

    const result = applySpreadsheetPasteSpecial(content, {
      snapshot,
      targetSheetId: 'target',
      targetSelection: { row: [4, 4], column: [3, 3] },
      options: allOptions,
    });
    const target = result?.content.sheets[1];

    expect(target?.config?.merge).toMatchObject({
      '4_3': { r: 4, c: 3, rs: 1, cs: 2 },
    });
    expect(target?.data?.[4]?.[3]?.mc).toEqual({
      r: 4,
      c: 3,
      rs: 1,
      cs: 2,
    });
    expect(target?.data?.[4]?.[4]?.mc).toEqual({ r: 4, c: 3 });
  });

  test('keeps destination content or formatting according to each paste mode', () => {
    const content = workbook();
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [0, 0], column: [0, 0] },
      '10',
    );
    if (!snapshot) throw new Error('Expected a rich clipboard snapshot.');

    const values = paste(content, snapshot, 'values');
    expect(values?.content.sheets[1]?.data?.[2]?.[2]).toMatchObject({
      v: 10,
      bg: '#ddebf7',
      fs: 18,
    });
    expect(values?.content.sheets[1]?.data?.[2]?.[2]?.ps?.value).toBe(
      'Destination comment',
    );

    const formats = paste(content, snapshot, 'formats');
    expect(formats?.content.sheets[1]?.data?.[2]?.[2]).toMatchObject({
      v: 100,
      bg: '#fff2cc',
      bl: 1,
      ps: { value: 'Destination comment' },
    });

    const comments = paste(content, snapshot, 'comments');
    expect(comments?.content.sheets[1]?.data?.[2]?.[2]).toMatchObject({
      v: 100,
      bg: '#ddebf7',
      ps: { value: 'Source comment' },
    });

    const validation = paste(content, snapshot, 'validation');
    expect(validation?.content.sheets[1]?.data?.[2]?.[2]?.v).toBe(100);
    expect(
      validation?.content.sheets[1]?.dataVerification?.['2_2'],
    ).toMatchObject({ type: 'number', value1: '0', value2: '100' });

    const withoutBorders = paste(content, snapshot, 'all-except-borders');
    expect(withoutBorders?.content.sheets[1]?.config?.borderInfo).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          rangeType: 'cell',
          value: expect.objectContaining({
            row_index: 2,
            col_index: 2,
            t: { color: '#c00000', style: '13' },
          }),
        }),
      ]),
    );
  });

  test('combines values or formulas with number formats without copying other styles', () => {
    const content = workbook();
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [0, 0], column: [0, 1] },
      '10\t20',
    );
    if (!snapshot) throw new Error('Expected a rich clipboard snapshot.');

    const values = paste(content, snapshot, 'values-and-number-formats');
    expect(values?.content.sheets[1]?.data?.[2]?.[2]).toMatchObject({
      v: 10,
      bg: '#ddebf7',
      fs: 18,
      ct: { fa: '0.00', t: 'n' },
    });

    const formulas = paste(content, snapshot, 'formulas-and-number-formats');
    expect(formulas?.content.sheets[1]?.data?.[2]?.[3]).toMatchObject({
      f: '=$A3+C$1',
      ct: { fa: '0%', t: 'n' },
    });
    expect(formulas?.content.sheets[1]?.data?.[2]?.[3]?.bg).toBe('#fce4d6');
  });

  test('tiles, transposes, skips blanks, and applies arithmetic as one result', () => {
    const content = workbook();
    const source = content.sheets[0];
    if (!source?.data) throw new Error('Expected source matrix.');
    source.data[0] = [{ v: 2, m: '2' }, null];
    source.data[1] = [
      { v: 3, m: '3' },
      { v: 4, m: '4' },
    ];
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [0, 1], column: [0, 1] },
      '2\t\n3\t4',
    );
    if (!snapshot) throw new Error('Expected a rich clipboard snapshot.');

    const result = applySpreadsheetPasteSpecial(content, {
      snapshot,
      targetSheetId: 'target',
      targetSelection: { row: [2, 5], column: [2, 5] },
      options: {
        content: 'values',
        operation: 'add',
        skipBlanks: true,
        transpose: true,
      },
    });

    expect(result?.targetRange).toEqual({ row: [2, 5], column: [2, 5] });
    expect(result?.content.sheets[1]?.data?.[2]?.[2]?.v).toBe(102);
    expect(result?.content.sheets[1]?.data?.[2]?.[3]?.v).toBe(103);
    expect(result?.content.sheets[1]?.data?.[3]?.[2]?.v).toBe(100);
    expect(result?.content.sheets[1]?.data?.[3]?.[3]?.v).toBe(104);
    expect(result?.content.sheets[1]?.data?.[4]?.[2]?.v).toBe(102);
  });

  test('pastes source column widths without changing cells', () => {
    const content = workbook();
    const snapshot = captureSpreadsheetClipboardSnapshot(
      content,
      'source',
      { row: [0, 0], column: [0, 1] },
      '10\t20',
    );
    if (!snapshot) throw new Error('Expected a rich clipboard snapshot.');

    const result = paste(content, snapshot, 'column-widths');
    expect(result?.content.sheets[1]?.config?.columnlen).toMatchObject({
      2: 120,
      3: 84,
    });
    expect(result?.content.sheets[1]?.data?.[2]?.[2]?.v).toBe(100);
  });

  test('keeps maximum-size sparse sheets sparse', () => {
    const content = workbook();
    const target = content.sheets[1];
    if (!target) throw new Error('Expected target sheet.');
    delete target.data;
    target.celldata = [{ r: 0, c: 0, v: { v: 'anchor', m: 'anchor' } }];
    target.row = 1_048_576;
    target.column = 16_384;
    const snapshot = createSpreadsheetTextClipboardSnapshot('42');
    if (!snapshot) throw new Error('Expected a text clipboard snapshot.');

    const result = applySpreadsheetPasteSpecial(content, {
      snapshot,
      targetSheetId: 'target',
      targetSelection: {
        row: [1_048_575, 1_048_575],
        column: [16_383, 16_383],
      },
      options: allOptions,
    });

    expect(result?.content.sheets[1]?.data).toBeUndefined();
    expect(result?.content.sheets[1]?.celldata).toHaveLength(2);
    expect(result?.content.sheets[1]?.celldata?.at(-1)).toEqual({
      r: 1_048_575,
      c: 16_383,
      v: { v: '42', m: '42' },
    });
  });

  test('fails closed for malformed targets, pivots, merges, protection, and arithmetic errors', () => {
    const content = workbook();
    const snapshot = createSpreadsheetTextClipboardSnapshot('0');
    if (!snapshot) throw new Error('Expected a text clipboard snapshot.');
    const base = {
      snapshot,
      targetSheetId: 'target',
      targetSelection: { row: [2, 2], column: [2, 2] },
      options: { ...allOptions, content: 'values' as const },
    };

    expect(
      spreadsheetPasteSpecialValidationError(content, {
        ...base,
        options: { ...base.options, operation: 'divide' },
      }),
    ).toBe('除数不能为 0。');

    const pivot = structuredClone(content);
    const pivotSheet = pivot.sheets[1];
    if (!pivotSheet) throw new Error('Expected the target worksheet.');
    pivotSheet.isPivotTable = true;
    expect(spreadsheetPasteSpecialValidationError(pivot, base)).toContain(
      '数据透视表',
    );

    const merged = structuredClone(content);
    const mergedSheet = merged.sheets[1];
    if (!mergedSheet?.config) throw new Error('Expected target sheet config.');
    mergedSheet.config.merge = {
      '2_2': { r: 2, c: 2, rs: 2, cs: 2 },
    };
    expect(spreadsheetPasteSpecialValidationError(merged, base)).toContain(
      '合并单元格',
    );

    const protectedContent = structuredClone(content);
    const protectedCell = protectedContent.sheets[1]?.data?.[2]?.[2];
    if (!protectedCell) throw new Error('Expected the target cell.');
    protectedCell.lo = 1;
    expect(
      spreadsheetPasteSpecialValidationError(protectedContent, base),
    ).toContain('保护');

    expect(
      spreadsheetPasteSpecialValidationError(content, {
        ...base,
        targetSelection: {
          row: [1_048_576, 1_048_576],
          column: [0, 0],
        },
      }),
    ).toContain('工作表边界');
  });

  test('limits rich-only modes for external text while retaining formulas and values', () => {
    const snapshot = createSpreadsheetTextClipboardSnapshot('=A1+1\t12');
    if (!snapshot) throw new Error('Expected a text clipboard snapshot.');
    expect(snapshot.cells[0]?.[0].cell).toMatchObject({
      f: '=A1+1',
      v: '=A1+1',
    });
    expect(spreadsheetPasteSpecialModeAvailable(snapshot, 'all')).toBe(true);
    expect(spreadsheetPasteSpecialModeAvailable(snapshot, 'values')).toBe(true);
    expect(spreadsheetPasteSpecialModeAvailable(snapshot, 'formats')).toBe(
      false,
    );
    expect(
      spreadsheetPasteSpecialModeAvailable(snapshot, 'column-widths'),
    ).toBe(false);

    const externalFormula = createSpreadsheetTextClipboardSnapshot(
      '=[Budget.xlsx]Sheet1!A1',
    );
    if (!externalFormula)
      throw new Error('Expected an external formula snapshot.');
    expect(externalFormula.containsUnsupportedFormulaState).toBe(true);
    expect(
      spreadsheetPasteSpecialValidationError(workbook(), {
        snapshot: externalFormula,
        targetSheetId: 'target',
        targetSelection: { row: [2, 2], column: [2, 2] },
        options: allOptions,
      }),
    ).toContain('外部公式');
  });
});

function paste(
  content: WorkSpreadsheetContent,
  snapshot: NonNullable<ReturnType<typeof captureSpreadsheetClipboardSnapshot>>,
  mode: SpreadsheetPasteSpecialOptions['content'],
) {
  return applySpreadsheetPasteSpecial(content, {
    snapshot,
    targetSheetId: 'target',
    targetSelection: { row: [2, 2], column: [2, 2] },
    options: { ...allOptions, content: mode },
  });
}

function workbook(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [sourceSheet(), targetSheet()],
  };
}

function sourceSheet(): WorkSpreadsheetSheet {
  return {
    id: 'source',
    name: 'Source',
    row: 20,
    column: 10,
    defaultColWidth: 96,
    config: {
      columnlen: { 0: 120, 1: 84 },
      merge: { '1_0': { r: 1, c: 0, rs: 1, cs: 2 } },
      borderInfo: [
        {
          rangeType: 'cell',
          value: {
            row_index: 0,
            col_index: 0,
            b: { color: '#4472c4', style: '1' },
          },
        },
      ],
      authority: {
        sheet: 0,
        cellProtectionRanges: [
          {
            range: { row: [0, 0], column: [0, 0] },
            locked: false,
            hidden: true,
          },
        ],
      },
    },
    data: [
      [
        {
          v: 10,
          m: '10',
          bg: '#fff2cc',
          bl: 1,
          ct: { fa: '0.00', t: 'n' },
          ps: comment('Source comment'),
        },
        {
          v: 20,
          m: '20',
          f: '=$A1+A$1',
          ct: { fa: '0%', t: 'n' },
        },
      ],
      [
        { v: 30, m: '30', mc: { r: 1, c: 0, rs: 1, cs: 2 } },
        { mc: { r: 1, c: 0 } },
      ],
    ],
    hyperlink: {
      '0_0': { linkType: 'webpage', linkAddress: 'https://a3s.dev' },
    },
    dataVerification: {
      '0_0': validation(),
    },
  };
}

function targetSheet(): WorkSpreadsheetSheet {
  return {
    id: 'target',
    name: 'Target',
    row: 20,
    column: 10,
    config: {
      columnlen: { 2: 60, 3: 60 },
      borderInfo: [
        {
          rangeType: 'cell',
          value: {
            row_index: 2,
            col_index: 2,
            t: { color: '#c00000', style: '13' },
          },
        },
      ],
    },
    data: Array.from({ length: 6 }, (_row, row) =>
      Array.from({ length: 6 }, (_cell, column) =>
        row >= 2 && column >= 2
          ? {
              v: 100,
              m: '100',
              bg: column === 2 ? '#ddebf7' : '#fce4d6',
              fs: 18,
              ps: comment('Destination comment'),
            }
          : null,
      ),
    ),
    dataVerification: {
      '2_2': { ...validation(), type: 'dropdown', value1: 'A,B' },
    },
  };
}

function comment(value: string) {
  return {
    left: null,
    top: null,
    width: null,
    height: null,
    value,
    isShow: false,
  };
}

function validation() {
  return {
    type: 'number',
    type2: 'between',
    rangeTxt: 'A1',
    value1: '0',
    value2: '100',
    validity: 'allow',
    remote: false,
    prohibitInput: true,
    hintShow: false,
    hintValue: '',
  };
}
