import type { Sheet } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import { parseXml } from '../../../src/internal/features/work/work-ooxml-package';
import { spreadsheetConditionalFormatStyles } from '../../../src/internal/features/work/work-spreadsheet-conditional-format';
import {
  writeXlsxConditionalFormats,
  XlsxDifferentialFormatWriter,
} from '../../../src/internal/features/work/work-xlsx-conditional-format-write';

/**
 * R2 conditional-format precedence corpus.
 *
 * Pins Traditional Office / WPS daily priority ordering, Stop-if-true blocking,
 * non-conflicting property merge, and native XLSX priority / stopIfTrue export.
 */
describe('R2 conditional-format precedence corpus', () => {
  test('higher-priority rules win conflicting text and fill colors', () => {
    const sheet = {
      id: 'sheet-1',
      name: 'Sheet1',
      data: [[{ v: 10 }], [{ v: 20 }]],
      luckysheet_conditionformat_save: [
        {
          type: 'default',
          cellrange: [{ row: [0, 1], column: [0, 0] }],
          format: { textColor: '#006100', cellColor: '#c6efce' },
          conditionName: 'greaterThan',
          conditionValue: ['5'],
        },
        {
          type: 'default',
          cellrange: [{ row: [0, 1], column: [0, 0] }],
          format: { textColor: '#9c0006', cellColor: '#ffc7ce' },
          conditionName: 'greaterThan',
          conditionValue: ['5'],
        },
      ],
    } as unknown as Sheet;

    const styles = spreadsheetConditionalFormatStyles(sheet);
    expect(styles.get('0_0')).toEqual({
      textColor: '#006100',
      cellColor: '#c6efce',
    });
    expect(styles.get('1_0')).toEqual({
      textColor: '#006100',
      cellColor: '#c6efce',
    });
  });

  test('later rules still merge non-conflicting format properties', () => {
    const sheet = {
      id: 'sheet-1',
      name: 'Sheet1',
      data: [[{ v: 10 }]],
      luckysheet_conditionformat_save: [
        {
          type: 'default',
          cellrange: [{ row: [0, 0], column: [0, 0] }],
          format: { textColor: '#006100', cellColor: null },
          conditionName: 'greaterThan',
          conditionValue: ['5'],
        },
        {
          type: 'default',
          cellrange: [{ row: [0, 0], column: [0, 0] }],
          format: { textColor: null, cellColor: '#fff2cc' },
          conditionName: 'greaterThan',
          conditionValue: ['5'],
        },
      ],
    } as unknown as Sheet;

    expect(spreadsheetConditionalFormatStyles(sheet).get('0_0')).toEqual({
      textColor: '#006100',
      cellColor: '#fff2cc',
    });
  });

  test('stop-if-true blocks later rules including different properties', () => {
    const sheet = {
      id: 'sheet-1',
      name: 'Sheet1',
      data: [[{ v: 10 }], [{ v: 1 }]],
      luckysheet_conditionformat_save: [
        {
          type: 'default',
          cellrange: [{ row: [0, 1], column: [0, 0] }],
          format: { textColor: '#006100', cellColor: null },
          conditionName: 'greaterThan',
          conditionValue: ['5'],
          stopIfTrue: true,
        },
        {
          type: 'default',
          cellrange: [{ row: [0, 1], column: [0, 0] }],
          format: { textColor: null, cellColor: '#ffc7ce' },
          conditionName: 'greaterThan',
          conditionValue: ['0'],
        },
      ],
    } as unknown as Sheet;

    const styles = spreadsheetConditionalFormatStyles(sheet);
    expect(styles.get('0_0')).toEqual({ textColor: '#006100' });
    expect(styles.get('1_0')).toEqual({ cellColor: '#ffc7ce' });
  });

  test('exports ascending XLSX priority and stopIfTrue for ordered rules', () => {
    const worksheet = parseXml(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`,
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dxfs count="0"/>
</styleSheet>`,
      'xl/styles.xml',
    );
    const differentials = new XlsxDifferentialFormatWriter(styles);
    writeXlsxConditionalFormats(
      worksheet,
      [
        {
          type: 'default',
          cellrange: [{ row: [0, 2], column: [0, 0] }],
          format: { textColor: '#006100', cellColor: '#c6efce' },
          conditionName: 'greaterThan',
          conditionValue: ['5'],
          stopIfTrue: true,
        },
        {
          type: 'default',
          cellrange: [{ row: [0, 2], column: [0, 0] }],
          format: { textColor: '#9c0006', cellColor: '#ffc7ce' },
          conditionName: 'lessThan',
          conditionValue: ['0'],
        },
      ],
      differentials,
    );

    const rules = [
      ...worksheet.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'cfRule',
      ),
    ];
    expect(rules).toHaveLength(2);
    expect(rules[0]?.getAttribute('priority')).toBe('1');
    expect(rules[0]?.getAttribute('stopIfTrue')).toBe('1');
    expect(rules[1]?.getAttribute('priority')).toBe('2');
    expect(rules[1]?.hasAttribute('stopIfTrue')).toBe(false);
  });
});
