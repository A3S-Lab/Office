import { expect, test } from '@rstest/core';
import { spreadsheetFormulaCount } from '../src/internal/features/work/work-spreadsheet-formula-analysis';
import { unsupportedSpreadsheetFormulaFunctions } from '../src/internal/features/work/work-spreadsheet-formula-support';
import {
  freezeImportedSpreadsheetCell,
  registerImportedSpreadsheetMatrix,
} from '../src/internal/features/work/work-spreadsheet-matrix-profile';

test('loads spreadsheet function compatibility metadata on demand', () => {
  expect(
    unsupportedSpreadsheetFormulaFunctions(
      '=SUM(A1:A2) + A3S_UNKNOWN_FUNCTION(1)',
    ),
  ).toEqual(['A3S_UNKNOWN_FUNCTION']);
});

test('counts formulas from an authenticated import summary', () => {
  let enumerations = 0;
  const data = new Proxy(
    [[freezeImportedSpreadsheetCell({ f: '=1+1', v: 2 })]],
    {
      ownKeys(target) {
        enumerations += 1;
        return Reflect.ownKeys(target);
      },
    },
  );
  registerImportedSpreadsheetMatrix(data, {
    columnCount: 1,
    formulaCells: [{ column: 0, row: 0 }],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: 1,
    shownCommentCells: [],
  });
  enumerations = 0;

  expect(
    spreadsheetFormulaCount({
      sheets: [{ data, id: 'sheet-1', name: 'Formula' }],
      type: 'spreadsheet',
    }),
  ).toBe(1);
  expect(enumerations).toBe(0);
});
