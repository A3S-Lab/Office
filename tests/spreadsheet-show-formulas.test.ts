import { describe, expect, test } from '@rstest/core';
import {
  projectSpreadsheetCellShowingFormula,
  projectSpreadsheetSheetShowingFormulas,
  projectSpreadsheetSheetsShowingFormulas,
  spreadsheetFormulaDisplayText,
} from '../src/internal/features/work/editors/spreadsheet-show-formulas';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

describe('spreadsheet show formulas projection', () => {
  test('normalizes formula display text with a leading equals', () => {
    expect(spreadsheetFormulaDisplayText({ f: '=SUM(A1:A2)', v: 3 })).toBe(
      '=SUM(A1:A2)',
    );
    expect(spreadsheetFormulaDisplayText({ f: 'A1+1', v: 2 })).toBe('=A1+1');
    expect(spreadsheetFormulaDisplayText({ f: '   ', v: 1 })).toBeNull();
    expect(spreadsheetFormulaDisplayText({ v: 1 })).toBeNull();
    expect(spreadsheetFormulaDisplayText(null)).toBeNull();
  });

  test('projects formula text into m/v without mutating non-formula cells', () => {
    const formulaCell = { f: '=B1*2', m: '4', v: 4 };
    const projected = projectSpreadsheetCellShowingFormula(formulaCell);
    expect(projected).toEqual({ f: '=B1*2', m: '=B1*2', v: '=B1*2' });
    expect(projected).not.toBe(formulaCell);

    const valueCell = { m: '10', v: 10 };
    expect(projectSpreadsheetCellShowingFormula(valueCell)).toBe(valueCell);
  });

  test('projects dense data and sparse celldata rows', () => {
    const sheet = {
      id: 'sheet-1',
      name: 'Sheet 1',
      status: 1,
      data: [
        [
          { f: '=A2', m: '1', v: 1 },
          { m: 'plain', v: 'plain' },
        ],
      ],
      celldata: [{ r: 2, c: 0, v: { f: 'SUM(1,2)', m: '3', v: 3 } }],
    } as WorkSpreadsheetSheet;

    const projected = projectSpreadsheetSheetShowingFormulas(sheet);
    expect(projected).not.toBe(sheet);
    expect(projected.data?.[0]?.[0]).toEqual({
      f: '=A2',
      m: '=A2',
      v: '=A2',
    });
    expect(projected.data?.[0]?.[1]).toEqual({ m: 'plain', v: 'plain' });
    expect(projected.celldata?.[0]?.v).toEqual({
      f: 'SUM(1,2)',
      m: '=SUM(1,2)',
      v: '=SUM(1,2)',
    });

    expect(projectSpreadsheetSheetsShowingFormulas([sheet])[0]).toEqual(
      projected,
    );
  });

  test('returns the same sheet reference when nothing changes', () => {
    const sheet = {
      id: 'sheet-1',
      name: 'Sheet 1',
      status: 1,
      data: [[{ m: '1', v: 1 }]],
    } as WorkSpreadsheetSheet;
    expect(projectSpreadsheetSheetShowingFormulas(sheet)).toBe(sheet);
  });
});
