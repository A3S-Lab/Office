import { describe, expect, test } from '@rstest/core';
import {
  createSpreadsheetFormatCellsDialogSource,
  createSpreadsheetFormatCellsDraft,
  spreadsheetFormatCellsDraftErrors,
  spreadsheetFormatCellsPatch,
} from '../src/internal/features/work/editors/spreadsheet-format-cells-dialog-model';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet format cells dialog model', () => {
  test('describes mixed cells without emitting untouched fields', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          data: [
            [
              {
                v: 12,
                ct: { fa: '#,##0', t: 'n' },
                ht: '1',
                bl: 1,
                bg: '#fff2cc',
              },
              {
                v: 0.25,
                ct: { fa: '0.00%', t: 'n' },
                ht: '2',
                bl: 0,
              },
            ],
          ],
        },
      ],
    } satisfies WorkSpreadsheetContent;
    const cells = content.sheets[0]?.data ?? [];
    const source = createSpreadsheetFormatCellsDialogSource(
      content,
      'sheet-1',
      { row: [0, 0], column: [0, 1] },
      cells,
      { row: 0, column: 0 },
    );
    if (!source) throw new Error('Expected a format-cells dialog source.');
    const draft = createSpreadsheetFormatCellsDraft(source);

    expect(source.fields.numberFormat.mixed).toBe(true);
    expect(source.fields.horizontalAlignment.mixed).toBe(true);
    expect(source.fields.bold.mixed).toBe(true);
    expect(source.fields.fillColor.mixed).toBe(true);
    expect(draft.numberFormat).toBe('#,##0');
    expect(
      spreadsheetFormatCellsPatch(source, draft, {
        bold: true,
      }),
    ).toEqual({ bold: true });
  });

  test('emits an explicitly touched mixed value but drops a reverted uniform value', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          data: [
            [
              { v: 'A', ht: '1' },
              { v: 'B', ht: '2' },
            ],
          ],
        },
      ],
    } satisfies WorkSpreadsheetContent;
    const source = createSpreadsheetFormatCellsDialogSource(
      content,
      'sheet-1',
      { row: [0, 0], column: [0, 1] },
      content.sheets[0]?.data ?? [],
      { row: 0, column: 0 },
    );
    if (!source) throw new Error('Expected a format-cells dialog source.');
    const initial = createSpreadsheetFormatCellsDraft(source);

    expect(
      spreadsheetFormatCellsPatch(
        source,
        { ...initial, horizontalAlignment: 'left' },
        { horizontalAlignment: true },
      ),
    ).toEqual({ horizontalAlignment: 'left' });
    expect(
      spreadsheetFormatCellsPatch(
        source,
        { ...initial, fontFamily: 'Arial' },
        { fontFamily: true },
      ),
    ).toEqual({ fontFamily: 'Arial' });
    expect(
      spreadsheetFormatCellsPatch(source, initial, { fontFamily: true }),
    ).toEqual({});
  });

  test('keeps advanced underline variants distinct in mixed selections', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          data: [
            [
              { v: 'A', un: 4 },
              { v: 'B', un: 2 },
            ],
          ],
        },
      ],
    } satisfies WorkSpreadsheetContent;
    const source = createSpreadsheetFormatCellsDialogSource(
      content,
      'sheet-1',
      { row: [0, 0], column: [0, 1] },
      content.sheets[0]?.data ?? [],
      { row: 0, column: 0 },
    );
    if (!source) throw new Error('Expected a format-cells dialog source.');
    const draft = createSpreadsheetFormatCellsDraft(source);

    expect(source.fields.underline).toEqual({
      mixed: true,
      value: 'doubleAccounting',
    });
    expect(
      spreadsheetFormatCellsPatch(
        source,
        { ...draft, underline: 'singleAccounting' },
        { underline: true },
      ),
    ).toEqual({ underline: 'singleAccounting' });
  });

  test('validates custom number codes, font sizes, and rotation', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sheet 1', data: [[{ v: 1 }]] }],
    } satisfies WorkSpreadsheetContent;
    const source = createSpreadsheetFormatCellsDialogSource(
      content,
      'sheet-1',
      { row: [0, 0], column: [0, 0] },
      content.sheets[0]?.data ?? [],
      { row: 0, column: 0 },
    );
    if (!source) throw new Error('Expected a format-cells dialog source.');
    const draft = createSpreadsheetFormatCellsDraft(source);

    expect(
      spreadsheetFormatCellsDraftErrors({
        ...draft,
        numberFormat: ' ',
        fontSize: 0,
        rotation: 91,
      }),
    ).toEqual({
      fontSize: '字号需为 1–409 之间的数字。',
      numberFormat: '请输入数字格式代码。',
      rotation: '文字旋转角度需为 -90–90 之间的整数。',
    });
  });

  test('reads arbitrary negative Fortune rotations as their visible angle', () => {
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          data: [[{ v: 'Clockwise', rt: 120 }]],
        },
      ],
    } satisfies WorkSpreadsheetContent;
    const source = createSpreadsheetFormatCellsDialogSource(
      content,
      'sheet-1',
      { row: [0, 0], column: [0, 0] },
      content.sheets[0]?.data ?? [],
      { row: 0, column: 0 },
    );

    expect(source?.fields.rotation).toEqual({ mixed: false, value: -30 });
  });
});
