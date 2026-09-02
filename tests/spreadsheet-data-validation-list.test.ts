import { describe, expect, test } from '@rstest/core';
import {
  materializeSpreadsheetDependentListsForFortune,
  normalizeSpreadsheetDependentListFormula,
  resolveSpreadsheetDependentListReference,
  restoreSpreadsheetDependentListProjections,
} from '../src/internal/features/work/editors/spreadsheet-data-validation-list';
import {
  applySpreadsheetDataValidation,
  validateSpreadsheetDataValidationRequest,
} from '../src/internal/features/work/editors/spreadsheet-data-validation';
import {
  spreadsheetSheetsForFortune,
  spreadsheetSheetsFromFortune,
} from '../src/internal/features/work/editors/spreadsheet-editor-support';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet dependent list validation', () => {
  test('normalizes a bounded INDIRECT concatenation grammar', () => {
    expect(
      normalizeSpreadsheetDependentListFormula(
        ' =indirect( "\'Lists\'!"&$A2 ) ',
      ),
    ).toEqual({
      ok: true,
      formula: `=INDIRECT("'Lists'!"&$A2)`,
    });
    expect(
      normalizeSpreadsheetDependentListFormula('=OFFSET(A1,1,0)'),
    ).toMatchObject({ ok: false });
    expect(
      normalizeSpreadsheetDependentListFormula('=INDIRECT(IF(A1="x",A2,A3))'),
    ).toMatchObject({ ok: false });
  });

  test('resolves relative drivers to different local named ranges', () => {
    const content = dependentListContent();
    expect(
      resolveSpreadsheetDependentListReference(
        content,
        'inputs',
        1,
        1,
        1,
        1,
        '=INDIRECT($A2)',
      ),
    ).toEqual({ ok: true, displayReference: "'Lists'!B1:B2" });
    expect(
      resolveSpreadsheetDependentListReference(
        content,
        'inputs',
        2,
        1,
        1,
        1,
        '=INDIRECT($A2)',
      ),
    ).toEqual({ ok: true, displayReference: "'Lists'!C1:C2" });
    expect(
      resolveSpreadsheetDependentListReference(
        content,
        'inputs',
        3,
        1,
        1,
        1,
        '=INDIRECT($A2)',
      ),
    ).toEqual({ ok: true, empty: true });
  });

  test('materializes runtime references and restores the controlled formula', () => {
    const content = dependentListContent();
    const projected = spreadsheetSheetsForFortune(
      content.sheets,
      content.namedRanges,
    );
    expect(projected[0]?.dataVerification?.['1_1']).toMatchObject({
      value1: "'Lists'!B1:B2",
    });
    expect(projected[0]?.dataVerification?.['2_1']).toMatchObject({
      value1: "'Lists'!C1:C2",
    });
    expect(projected[0]?.dataVerification?.['3_1']).toMatchObject({
      value1: '',
    });

    const edited = projected.map((sheet) => ({
      ...sheet,
      data: sheet.data?.map((row, rowIndex) =>
        row?.map((cell, columnIndex) =>
          rowIndex === 1 && columnIndex === 1 ? { v: 'B2' } : cell,
        ),
      ),
    }));
    const restored = spreadsheetSheetsFromFortune(edited, content.sheets, []);
    expect(restored[0]?.dataVerification).toBeUndefined();
    expect(restored[0]?.dataValidationRanges?.[0]?.item.value1).toBe(
      '=INDIRECT($A2)',
    );

    const direct = materializeSpreadsheetDependentListsForFortune(
      [
        {
          ...content.sheets[0],
          dataVerification: {
            '1_1': {
              ...content.sheets[0]?.dataValidationRanges?.[0]?.item,
              value1: '=INDIRECT($A2)',
            },
          },
        },
      ],
      content.namedRanges,
    );
    const directRestored = restoreSpreadsheetDependentListProjections(direct);
    expect(directRestored[0]?.dataVerification?.['1_1']).toMatchObject({
      value1: '=INDIRECT($A2)',
    });
  });

  test('fails closed when an imported compact rule exceeds the projection budget', () => {
    const content = dependentListContent();
    const oversized = {
      ...content.sheets[0],
      row: 20_000,
      dataValidationRanges: [
        {
          ranges: [{ row: [1, 10_001], column: [1, 1] }],
          item: {
            ...content.sheets[0]?.dataValidationRanges?.[0]?.item,
            value1: '=INDIRECT($A2)',
          },
        },
      ],
    };
    const projected = materializeSpreadsheetDependentListsForFortune(
      [oversized],
      content.namedRanges,
    );
    expect(projected[0]?.dataVerification).toBeUndefined();
    expect(projected[0]?.dataValidationRanges?.[0]?.item.value1).toBe('');
    expect(
      restoreSpreadsheetDependentListProjections(projected)[0]
        ?.dataValidationRanges?.[0]?.item.value1,
    ).toBe('=INDIRECT($A2)');
  });

  test('keeps the formula in the controlled rule and validates non-empty drivers', () => {
    const content = dependentListContent();
    const applied = applySpreadsheetDataValidation(content, {
      sheetId: 'inputs',
      ranges: [{ row: [1, 3], column: [1, 1] }],
      activeCell: { row: 1, column: 1 },
      value: {
        type: 'dropdown',
        type2: '',
        value1: '=INDIRECT($A2)',
        value2: '',
        allowBlank: false,
        showDropdownArrow: true,
        prohibitInput: true,
        errorStyle: 'stop',
        errorTitle: '',
        errorMessage: '',
        hintShow: false,
        hintTitle: '',
        hintValue: '',
      },
    });
    expect(applied?.sheets[0]?.dataValidationRanges?.[0]?.item.value1).toBe(
      '=INDIRECT($A2)',
    );
    expect(
      validateSpreadsheetDataValidationRequest(content, {
        sheetId: 'inputs',
        ranges: [{ row: [1, 2], column: [1, 1] }],
        activeCell: { row: 1, column: 1 },
        value: {
          type: 'dropdown',
          type2: '',
          value1: '=INDIRECT(A:A)',
          value2: '',
          allowBlank: true,
          showDropdownArrow: true,
          prohibitInput: true,
          errorStyle: 'stop',
          errorTitle: '',
          errorMessage: '',
          hintShow: false,
          hintTitle: '',
          hintValue: '',
        },
      }),
    ).toMatchObject({ ok: false, code: 'invalid-list-formula' });
  });
});

function dependentListContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    namedRanges: [
      {
        id: 'north',
        name: 'North',
        reference: "'Lists'!B1:B2",
      },
      {
        id: 'south',
        name: 'South',
        reference: "'Lists'!C1:C2",
      },
    ],
    sheets: [
      {
        id: 'inputs',
        name: 'Inputs',
        row: 10,
        column: 4,
        data: [
          [{ v: 'Region' }, { v: 'Owner' }],
          [{ v: 'North' }, null],
          [{ v: 'South' }, null],
          [null, null],
        ],
        dataValidationRanges: [
          {
            ranges: [{ row: [1, 3], column: [1, 1] }],
            item: {
              type: 'dropdown',
              type2: '',
              rangeTxt: 'B2:B4',
              value1: '=INDIRECT($A2)',
              value2: '',
              validity: '',
              remote: false,
              allowBlank: true,
              showDropdownArrow: true,
              prohibitInput: true,
              errorStyle: 'stop',
              errorTitle: '',
              errorMessage: '',
              hintShow: false,
              hintTitle: '',
              hintValue: '',
            },
          },
        ],
      },
      {
        id: 'lists',
        name: 'Lists',
        row: 4,
        column: 4,
        data: [
          [{ v: 'North-1' }, { v: 'Alice' }, { v: 'Cara' }],
          [{ v: 'North-2' }, { v: 'Bob' }, { v: 'Dana' }],
        ],
      },
    ],
  };
}
