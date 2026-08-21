import { describe, expect, test } from '@rstest/core';
import {
  spreadsheetUnderlineCellValue,
  spreadsheetUnderlineCellValueFromSheetJs,
  spreadsheetUnderlineCellValueFromXlsx,
  spreadsheetUnderlineStyle,
  spreadsheetUnderlineStyles,
} from '../src/internal/features/work/work-spreadsheet-underline';

describe('spreadsheet underline styles', () => {
  test('maps every native Fortune underline value to one typed style', () => {
    expect(spreadsheetUnderlineStyles).toEqual([
      'none',
      'single',
      'double',
      'singleAccounting',
      'doubleAccounting',
    ]);
    expect(
      spreadsheetUnderlineStyles.map((style) => ({
        style,
        value: spreadsheetUnderlineCellValue(style),
      })),
    ).toEqual([
      { style: 'none', value: 0 },
      { style: 'single', value: 1 },
      { style: 'double', value: 2 },
      { style: 'singleAccounting', value: 3 },
      { style: 'doubleAccounting', value: 4 },
    ]);
    expect([0, 1, 2, 3, 4].map(spreadsheetUnderlineStyle)).toEqual(
      spreadsheetUnderlineStyles,
    );
  });

  test('normalizes OOXML and SheetJS underline encodings without collapsing variants', () => {
    expect(
      [
        null,
        '',
        'single',
        'double',
        'singleAccounting',
        'doubleAccounting',
        'none',
      ].map(spreadsheetUnderlineCellValueFromXlsx),
    ).toEqual([1, 1, 1, 2, 3, 4, 0]);
    expect(
      [
        false,
        true,
        1,
        2,
        3,
        4,
        0x21,
        0x22,
        'single',
        'double',
        'singleAccounting',
        'doubleAccounting',
        'none',
      ].map(spreadsheetUnderlineCellValueFromSheetJs),
    ).toEqual([0, 1, 1, 2, 3, 4, 3, 4, 1, 2, 3, 4, 0]);
  });
});
