import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetDataValidation,
  canRemoveSpreadsheetDataValidation,
  createSpreadsheetDataValidationDialogSource,
  removeSpreadsheetDataValidation,
  validateSpreadsheetDataValidationRequest,
} from '../src/internal/features/work/editors/spreadsheet-data-validation';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet data validation', () => {
  test('applies one compact dropdown rule to multiple ranges without materializing cells', () => {
    const content = validationContent();
    const next = applySpreadsheetDataValidation(content, {
      sheetId: 'sheet-1',
      ranges: [
        { row: [1, 2], column: [1, 1] },
        { row: [4, 4], column: [3, 4] },
      ],
      activeCell: { row: 1, column: 1 },
      value: {
        type: 'dropdown',
        type2: '',
        value1: ' Ready, Blocked ',
        value2: '',
        prohibitInput: true,
        hintShow: true,
        hintValue: ' Choose a workflow state. ',
      },
    });

    expect(next?.sheets[0]?.data).toBe(content.sheets[0]?.data);
    expect(next?.sheets[0]?.celldata).toBeUndefined();
    expect(next?.sheets[0]?.dataValidationRanges).toEqual([
      {
        ranges: [
          { row: [0, 0], column: [0, 3] },
          { row: [3, 3], column: [0, 3] },
          { row: [1, 2], column: [0, 0] },
          { row: [1, 2], column: [2, 3] },
        ],
        item: expect.objectContaining({ value1: 'Keep' }),
      },
      {
        ranges: [
          { row: [1, 2], column: [1, 1] },
          { row: [4, 4], column: [3, 4] },
        ],
        item: {
          type: 'dropdown',
          type2: '',
          rangeTxt: 'B2:B3,D5:E5',
          value1: 'Ready,Blocked',
          value2: '',
          validity: '',
          remote: false,
          prohibitInput: true,
          hintShow: true,
          hintValue: 'Choose a workflow state.',
          checked: false,
        },
      },
    ]);
    expect(next?.sheets[0]?.dataVerification).toEqual({
      vendor: { opaque: true },
      '8_8': expect.objectContaining({ value1: 'Keep direct' }),
    });
    expect(content.sheets[0]?.dataValidationRanges?.[0]?.ranges).toEqual([
      { row: [0, 3], column: [0, 3] },
    ]);
  });

  test('edits and removes only the selected compact validation area', () => {
    const content = validationContent();
    const edited = applySpreadsheetDataValidation(content, {
      sheetId: 'sheet-1',
      ranges: [{ row: [1, 2], column: [1, 2] }],
      activeCell: { row: 1, column: 1 },
      value: {
        type: 'number_integer',
        type2: 'between',
        value1: ' 1 ',
        value2: '10',
        prohibitInput: true,
        hintShow: false,
        hintValue: 'ignored',
      },
    });

    expect(edited?.sheets[0]?.dataValidationRanges?.at(-1)).toEqual({
      ranges: [{ row: [1, 2], column: [1, 2] }],
      item: expect.objectContaining({
        type: 'number_integer',
        type2: 'between',
        value1: '1',
        value2: '10',
        hintShow: false,
        hintValue: '',
      }),
    });
    expect(
      edited &&
        canRemoveSpreadsheetDataValidation(edited, {
          sheetId: 'sheet-1',
          ranges: [{ row: [1, 1], column: [1, 1] }],
          activeCell: { row: 1, column: 1 },
        }),
    ).toBe(true);

    const removed = edited
      ? removeSpreadsheetDataValidation(edited, {
          sheetId: 'sheet-1',
          ranges: [{ row: [1, 1], column: [1, 1] }],
          activeCell: { row: 1, column: 1 },
        })
      : null;
    expect(
      removed?.sheets[0]?.dataValidationRanges?.some((entry) =>
        entry.ranges.some(
          (range) =>
            range.row[0] <= 1 &&
            range.row[1] >= 1 &&
            range.column[0] <= 1 &&
            range.column[1] >= 1,
        ),
      ),
    ).toBe(false);
    expect(removed?.sheets[0]?.data?.[1]?.[1]).toEqual({
      v: 'Ready',
      bg: '#e2f0d9',
    });
  });

  test('describes uniform and mixed selections from direct and compact rules', () => {
    const content = validationContent();
    const uniform = createSpreadsheetDataValidationDialogSource(content, {
      sheetId: 'sheet-1',
      ranges: [{ row: [0, 0], column: [0, 2] }],
      activeCell: { row: 0, column: 0 },
    });
    expect(uniform).toMatchObject({
      hasValidation: true,
      mixed: false,
      rangeReference: 'A1:C1',
      value: { type: 'dropdown', value1: 'Keep' },
    });

    const mixed = createSpreadsheetDataValidationDialogSource(content, {
      sheetId: 'sheet-1',
      ranges: [{ row: [0, 1], column: [0, 1] }],
      activeCell: { row: 0, column: 0 },
    });
    expect(mixed).toMatchObject({
      hasValidation: true,
      mixed: true,
      value: { type: 'dropdown', value1: 'Keep' },
    });
  });

  test('accepts a one-dimensional list source and rejects unsafe targets and values', () => {
    const content = validationContent();
    expect(
      validateSpreadsheetDataValidationRequest(content, {
        sheetId: 'sheet-1',
        ranges: [{ row: [1, 2], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: dropdownValue("='Lists'!$A$1:$A$4"),
      }),
    ).toMatchObject({
      ok: true,
      item: { value1: "'Lists'!A1:A4" },
    });
    expect(
      validateSpreadsheetDataValidationRequest(content, {
        sheetId: 'sheet-1',
        ranges: [{ row: [1, 2], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: dropdownValue("'Lists'!A1:B2"),
      }),
    ).toMatchObject({ ok: false, code: 'multiple-list-columns' });
    expect(
      validateSpreadsheetDataValidationRequest(content, {
        sheetId: 'sheet-1',
        ranges: [{ row: [1, 2], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: {
          ...dropdownValue('Ready'),
          type: 'number_integer',
          type2: 'between',
          value1: '1.5',
          value2: '2',
        },
      }),
    ).toMatchObject({ ok: false, code: 'invalid-number' });

    const protectedContent = validationContent();
    const sheet = protectedContent.sheets[0];
    if (!sheet) throw new Error('Expected validation fixture sheet');
    sheet.config = { authority: { sheet: 1 } };
    expect(
      validateSpreadsheetDataValidationRequest(protectedContent, {
        sheetId: 'sheet-1',
        ranges: [{ row: [1, 1], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: dropdownValue('Ready,Blocked'),
      }),
    ).toMatchObject({ ok: false, code: 'protected-range' });
  });

  test('bounds one gesture to ten thousand cells', () => {
    expect(
      validateSpreadsheetDataValidationRequest(validationContent(), {
        sheetId: 'sheet-1',
        ranges: [{ row: [0, 100], column: [0, 99] }],
        activeCell: { row: 0, column: 0 },
        value: dropdownValue('Ready,Blocked'),
      }),
    ).toMatchObject({ ok: false, code: 'range-too-large' });
  });

  test('normalizes date formulas and Excel serials to stable ISO boundaries', () => {
    const content = validationContent();
    expect(
      validateSpreadsheetDataValidationRequest(content, {
        sheetId: 'sheet-1',
        ranges: [{ row: [1, 2], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: {
          ...dropdownValue('unused'),
          type: 'date',
          type2: 'between',
          value1: 'DATE(2026, 8, 21)',
          value2: '46631',
        },
      }),
    ).toMatchObject({
      ok: true,
      item: {
        type: 'date',
        type2: 'between',
        value1: '2026-08-21',
        value2: '2027-09-01',
      },
    });
    expect(
      validateSpreadsheetDataValidationRequest(content, {
        sheetId: 'sheet-1',
        ranges: [{ row: [1, 1], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: {
          ...dropdownValue('unused'),
          type: 'date',
          type2: 'equal',
          value1: '60',
        },
      }),
    ).toMatchObject({ ok: false, code: 'invalid-date' });
  });
});

function dropdownValue(value1: string) {
  return {
    type: 'dropdown' as const,
    type2: '' as const,
    value1,
    value2: '',
    prohibitInput: true,
    hintShow: false,
    hintValue: '',
  };
}

function validationContent(): WorkSpreadsheetContent {
  const validation = {
    type: 'dropdown',
    type2: '',
    rangeTxt: 'A1:D4',
    value1: 'Keep',
    value2: '',
    validity: '',
    remote: false,
    prohibitInput: true,
    hintShow: false,
    hintValue: '',
  };
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Inputs',
        status: 1,
        row: 200,
        column: 200,
        data: [[{ v: 'State' }], [null, { v: 'Ready', bg: '#e2f0d9' }]],
        dataVerification: {
          '1_1': { ...validation, value1: 'Replace direct' },
          '8_8': { ...validation, value1: 'Keep direct' },
          vendor: { opaque: true },
        },
        dataValidationRanges: [
          {
            ranges: [{ row: [0, 3], column: [0, 3] }],
            item: validation,
          },
        ],
      },
      {
        id: 'lists',
        name: 'Lists',
        row: 20,
        column: 4,
        data: [[{ v: 'Ready' }], [{ v: 'Blocked' }]],
      },
    ],
  };
}
