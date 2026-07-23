import { expect, test } from '@rstest/core';
import { unsupportedSpreadsheetFormulaFunctions } from '../src/internal/features/work/work-spreadsheet-formula-support';

test('loads spreadsheet function compatibility metadata on demand', () => {
  expect(
    unsupportedSpreadsheetFormulaFunctions(
      '=SUM(A1:A2) + A3S_UNKNOWN_FUNCTION(1)',
    ),
  ).toEqual(['A3S_UNKNOWN_FUNCTION']);
});
