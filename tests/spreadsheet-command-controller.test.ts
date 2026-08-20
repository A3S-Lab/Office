import type { Cell, Sheet } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import {
  createSpreadsheetEditorExtensions,
  type SpreadsheetAutoFilterCommandPort,
  type SpreadsheetCalculationCommandPort,
  type SpreadsheetClipboardCommandPort,
  type SpreadsheetCommandContext,
  type SpreadsheetCommandRange,
  type SpreadsheetEditorCommands,
  type SpreadsheetFormatCellsCommandPort,
  type SpreadsheetFormatPainterCommandPort,
  type SpreadsheetNavigationCommandPort,
  type SpreadsheetWorkbookCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet command controller', () => {
  test('owns controlled workbook replacement in the document extension', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);
    const next: WorkSpreadsheetContent = {
      ...fixture.context.content,
      sheets: [
        ...fixture.context.content.sheets,
        { id: 'sheet-2', name: 'Sheet 2' },
      ],
    };

    expect(editor.extensionNames[0]).toBe('spreadsheetDocument');
    expect(editor.can().setSpreadsheetContent(next)).toBe(true);
    expect(editor.commands.setSpreadsheetContent(next)).toBe(true);
    expect(fixture.changes).toEqual([next]);

    editor.updateContext({ ...fixture.context, editable: false });
    expect(editor.can().setSpreadsheetContent(next)).toBe(false);
    expect(editor.commands.setSpreadsheetContent(next)).toBe(false);
    expect(fixture.changes).toEqual([next]);
  });

  test('routes cell formatting through the workbook command port', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [4, 2], column: [3, 1] }];
    const editor = spreadsheetEditor(fixture.context);

    const handled = editor.commands.setCellFormat('fs', 14);
    const rejected = editor.commands.setCellFormat('ct', '0.00%');
    const formatted = editor.commands.setCellFormat('ct', {
      fa: '0.00%',
      t: 'n',
    });

    expect(handled).toBe(true);
    expect(rejected).toBe(false);
    expect(formatted).toBe(true);
    expect(editor.extensionNames).toContain('spreadsheetCellFormatting');
    expect(fixture.workbook.formats).toEqual([
      {
        attribute: 'fs',
        range: { row: [2, 4], column: [1, 3] },
        sheetId: 'sheet-1',
        value: 14,
      },
      {
        attribute: 'ct',
        range: { row: [2, 4], column: [1, 3] },
        sheetId: 'sheet-1',
        value: { fa: '0.00%', t: 'n' },
      },
    ]);
  });

  test('owns the common WPS number-format shortcuts', () => {
    const fixture = commandFixture();
    fixture.context.toolbarCell = { v: 45_292.5 };
    const editor = spreadsheetEditor(fixture.context);
    const shortcuts = [
      { code: 'Backquote', key: '~' },
      { code: 'Digit1', key: '!' },
      { code: 'Digit4', key: '$' },
      { code: 'Digit5', key: '%' },
      { code: 'Digit3', key: '#' },
      { code: 'Digit2', key: '@' },
      { code: 'Digit6', key: '^' },
    ].map(
      ({ code, key }) =>
        new KeyboardEvent('keydown', {
          cancelable: true,
          code,
          key,
          metaKey: true,
          shiftKey: true,
        }),
    );

    expect(shortcuts.map((event) => editor.handleKeyDown(event))).toEqual([
      true,
      true,
      true,
      true,
      true,
      true,
      true,
    ]);
    expect(shortcuts.every((event) => event.defaultPrevented)).toBe(true);
    expect(fixture.workbook.formats).toEqual([
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: 'General', t: 'n' },
      },
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: '#,##0.00', t: 'n' },
      },
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: '[$¥-804]#,##0.00', t: 'n' },
      },
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: '0.00%', t: 'n' },
      },
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: 'yyyy-MM-dd', t: 'd' },
      },
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: 'hh:mm', t: 'd' },
      },
      {
        attribute: 'ct',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: { fa: '0.00E+00', t: 'n' },
      },
    ]);
  });

  test('adjusts mixed decimal formats per cell in one native batch', () => {
    const fixture = commandFixture();
    const cells: (Cell | null)[][] = [
      [
        { ct: { fa: '[$¥-804]#,##0.00', t: 'n' }, v: 12.5 },
        { ct: { fa: '0.0%', t: 'n' }, v: 0.25 },
        { ct: { fa: 'yyyy-MM-dd', t: 'd' }, v: 45_292 },
      ],
      [
        { ct: { fa: '[$¥-804]#,##0.00', t: 'n' }, v: 20.5 },
        { ct: { fa: '0.0%', t: 'n' }, v: 0.5 },
        { ct: { fa: '@', t: 's' }, v: '001' },
      ],
    ];
    fixture.context.content.sheets[0] = {
      ...fixture.context.content.sheets[0],
      column: 3,
      data: cells,
      row: 2,
    };
    fixture.workbook.cells = cells;
    fixture.workbook.selection = [{ row: [0, 1], column: [0, 2] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetNumberFormats');
    expect(editor.can().adjustDecimalPlaces('increase')).toBe(true);
    expect(editor.commands.adjustDecimalPlaces('increase')).toBe(true);
    expect(fixture.workbook.clearBatches).toEqual([
      [
        {
          name: 'setCellFormatByRange',
          args: [
            'ct',
            { fa: '[$¥-804]#,##0.000', t: 'n' },
            { row: [0, 1], column: [0, 0] },
            { id: 'sheet-1' },
          ],
        },
        {
          name: 'setCellFormatByRange',
          args: [
            'ct',
            { fa: '0.00%', t: 'n' },
            { row: [0, 1], column: [1, 1] },
            { id: 'sheet-1' },
          ],
        },
      ],
    ]);
  });

  test('opens and applies Format Cells through typed commands', () => {
    const fixture = commandFixture();
    fixture.workbook.cells = [[{ v: 'A3S' }, null]];
    fixture.workbook.selection = [
      {
        row: [2, 2],
        column: [4, 3],
        row_focus: 2,
        column_focus: 4,
      },
    ];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetCellFormat');
    expect(editor.can().openFormatCells()).toBe(true);
    expect(editor.commands.openFormatCells()).toBe(true);
    expect(fixture.formatCells.requests).toEqual([
      {
        sheetId: 'sheet-1',
        range: { row: [2, 2], column: [3, 4] },
        activeCell: { row: 2, column: 4 },
        cells: [[{ v: 'A3S' }, null]],
      },
    ]);
    expect(fixture.changes).toEqual([]);

    expect(
      editor.commands.applyCellFormat({
        sheetId: 'sheet-1',
        range: { row: [2, 2], column: [3, 4] },
        patch: { bold: true, fillColor: '#fff2cc' },
      }),
    ).toBe(true);
    expect(fixture.changes).toHaveLength(1);
    expect(fixture.changes[0]?.sheets[0]?.celldata).toEqual([
      { r: 2, c: 3, v: { bg: '#fff2cc', bl: 1 } },
      { r: 2, c: 4, v: { bg: '#fff2cc', bl: 1 } },
    ]);
  });

  test('owns Cmd/Ctrl+1 only on the spreadsheet editing surface', () => {
    const fixture = commandFixture();
    fixture.workbook.cells = [[{ v: 'A3S' }]];
    fixture.workbook.selection = [{ row: [0, 0], column: [0, 0] }];
    const editor = spreadsheetEditor(fixture.context);
    const gridShortcut = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: '1',
    });

    expect(editor.handleKeyDown(gridShortcut)).toBe(true);
    expect(gridShortcut.defaultPrevented).toBe(true);
    expect(fixture.formatCells.requests).toHaveLength(1);

    const input = document.createElement('input');
    document.body.append(input);
    let inputHandled = true;
    input.addEventListener('keydown', (event) => {
      inputHandled = editor.handleKeyDown(event);
    });
    input.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        metaKey: true,
        key: '1',
      }),
    );
    expect(inputHandled).toBe(false);
    expect(fixture.formatCells.requests).toHaveLength(1);
    input.remove();

    const dialog = document.createElement('div');
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    const dialogButton = document.createElement('button');
    dialog.append(dialogButton);
    document.body.append(dialog);
    const modalShortcut = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: '1',
    });
    Object.defineProperty(modalShortcut, 'target', { value: dialogButton });
    expect(editor.handleKeyDown(modalShortcut)).toBe(false);
    expect(fixture.formatCells.requests).toHaveLength(1);
    dialog.remove();

    editor.updateContext({ ...fixture.context, editable: false });
    const readOnlyShortcut = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: '1',
    });
    expect(editor.handleKeyDown(readOnlyShortcut)).toBe(false);
  });

  test('routes all four WPS fill directions through one native command port', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [3, 1], column: [4, 2] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetCellFill');
    for (const direction of ['down', 'right', 'up', 'left'] as const) {
      expect(editor.can().fillSelectedCells(direction)).toBe(true);
      expect(editor.commands.fillSelectedCells(direction)).toBe(true);
    }
    expect(fixture.workbook.fills).toEqual([
      {
        applyRange: { row: [2, 3], column: [2, 4] },
        copyRange: { row: [1, 1], column: [2, 4] },
        direction: 'down',
      },
      {
        applyRange: { row: [1, 3], column: [3, 4] },
        copyRange: { row: [1, 3], column: [2, 2] },
        direction: 'right',
      },
      {
        applyRange: { row: [1, 2], column: [2, 4] },
        copyRange: { row: [3, 3], column: [2, 4] },
        direction: 'up',
      },
      {
        applyRange: { row: [1, 3], column: [2, 3] },
        copyRange: { row: [1, 3], column: [4, 4] },
        direction: 'left',
      },
    ]);
    expect(fixture.workbook.selection).toEqual([
      { row: [3, 1], column: [4, 2] },
    ]);
    expect(fixture.workbook.fillRowsAtCall).toEqual([
      [1, 2, 3],
      [1, 2, 3],
      [1, 2, 3],
      [1, 2, 3],
    ]);
  });

  test('requires one editable unmerged range before native fill', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [
      { row: [0, 2], column: [0, 0] },
      { row: [0, 2], column: [2, 2] },
    ];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().fillSelectedCells('down')).toBe(false);
    expect(editor.commands.fillSelectedCells('down')).toBe(false);
    expect(fixture.workbook.fills).toEqual([]);

    fixture.workbook.selection = [{ row: [0, 2], column: [0, 0] }];
    fixture.workbook.failFill = true;
    expect(editor.can().fillSelectedCells('down')).toBe(true);
    expect(editor.commands.fillSelectedCells('down')).toBe(false);
    expect(fixture.workbook.fills).toEqual([]);
    expect(fixture.workbook.fillRowsAtCall).toEqual([[0, 1, 2]]);
    expect(fixture.workbook.sheet.data).toEqual([[null]]);

    editor.updateContext({ ...fixture.context, editable: false });
    fixture.workbook.failFill = false;
    expect(editor.can().fillSelectedCells('down')).toBe(false);
    expect(editor.commands.fillSelectedCells('down')).toBe(false);
    expect(fixture.workbook.fills).toEqual([]);
  });

  test('restores an absent native data matrix when fill fails', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 2], column: [0, 0] }];
    fixture.workbook.failFill = true;
    delete fixture.workbook.sheet.data;
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().fillSelectedCells('down')).toBe(true);
    expect(editor.commands.fillSelectedCells('down')).toBe(false);
    expect(fixture.workbook.fillRowsAtCall).toEqual([[0, 1, 2]]);
    expect(Object.hasOwn(fixture.workbook.sheet, 'data')).toBe(false);
  });

  test('owns Cmd/Ctrl+D and Cmd/Ctrl+R on the spreadsheet grid', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 2], column: [0, 2] }];
    const editor = spreadsheetEditor(fixture.context);
    const shortcuts = [
      new KeyboardEvent('keydown', {
        cancelable: true,
        ctrlKey: true,
        key: 'd',
      }),
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: 'r',
        metaKey: true,
      }),
    ];

    expect(shortcuts.map((event) => editor.handleKeyDown(event))).toEqual([
      true,
      true,
    ]);
    expect(shortcuts.every((event) => event.defaultPrevented)).toBe(true);
    expect(fixture.workbook.fills.map(({ direction }) => direction)).toEqual([
      'down',
      'right',
    ]);
  });

  test('keeps unavailable or failed WPS fill shortcuts out of the browser', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 0], column: [0, 0] }];
    const editor = spreadsheetEditor(fixture.context);
    const unavailable = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: 'r',
    });

    expect(editor.can().fillSelectedCells('right')).toBe(false);
    expect(editor.handleKeyDown(unavailable)).toBe(true);
    expect(unavailable.defaultPrevented).toBe(true);
    expect(fixture.workbook.fills).toEqual([]);

    fixture.workbook.selection = [{ row: [0, 2], column: [0, 0] }];
    fixture.workbook.failFill = true;
    const failed = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: 'd',
    });
    expect(editor.handleKeyDown(failed)).toBe(true);
    expect(failed.defaultPrevented).toBe(true);
    expect(fixture.workbook.fills).toEqual([]);
    expect(fixture.workbook.sheet.data).toEqual([[null]]);

    editor.updateContext({ ...fixture.context, editable: false });
    const readOnly = new KeyboardEvent('keydown', {
      cancelable: true,
      metaKey: true,
      key: 'r',
    });
    expect(editor.handleKeyDown(readOnly)).toBe(true);
    expect(readOnly.defaultPrevented).toBe(true);
  });

  test('commits cell borders through one immutable controlled update', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [4, 2], column: [3, 1] }];
    const editor = spreadsheetEditor(fixture.context);
    const format = {
      target: 'outside',
      color: '#2463eb',
      style: 'medium',
    } as const;

    expect(editor.extensionNames).toContain('spreadsheetCellBorders');
    expect(editor.can().setSelectedCellBorders(format)).toBe(true);
    expect(editor.commands.setSelectedCellBorders(format)).toBe(true);
    expect(fixture.changes).toHaveLength(1);
    expect(fixture.context.content.sheets[0]?.config).toBeUndefined();
    expect(fixture.changes[0]?.sheets[0]?.config?.borderInfo).toEqual([
      {
        rangeType: 'range',
        borderType: 'border-outside',
        color: '#2463eb',
        style: '8',
        range: [{ row: [2, 4], column: [1, 3] }],
      },
    ]);
    expect(fixture.workbook.clearBatches).toEqual([]);
    expect(fixture.workbook.formats).toEqual([]);

    editor.updateContext({ ...fixture.context, editable: false });
    expect(editor.can().setSelectedCellBorders(format)).toBe(false);
    expect(editor.commands.setSelectedCellBorders(format)).toBe(false);
    expect(fixture.changes).toHaveLength(1);
  });

  test('applies a WPS cell style through one controlled update', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [1, 0], column: [2, 1] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetCellStyles');
    expect(editor.can().applyCellStyle('good')).toBe(true);
    expect(editor.commands.applyCellStyle('good')).toBe(true);
    expect(fixture.changes).toHaveLength(1);
    expect(fixture.context.content.sheets[0]?.data).toBeUndefined();
    expect(fixture.changes[0]?.sheets[0]?.data?.[0]?.[1]).toMatchObject({
      bg: '#c6efce',
      fc: '#006100',
      ff: 'Aptos',
      fs: 10,
    });
    expect(fixture.changes[0]?.sheets[0]?.data?.[1]?.[2]).toMatchObject({
      bg: '#c6efce',
      fc: '#006100',
    });
    expect(fixture.workbook.formats).toEqual([]);

    editor.updateContext({ ...fixture.context, editable: false });
    expect(editor.can().applyCellStyle('bad')).toBe(false);
    expect(editor.commands.applyCellStyle('bad')).toBe(false);
    expect(fixture.changes).toHaveLength(1);
  });

  test('runs WPS merge commands as one workbook batch', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 1], column: [0, 2] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().mergeSelectedCells('merge-and-center')).toBe(true);
    expect(editor.commands.mergeSelectedCells('merge-and-center')).toBe(true);

    expect(fixture.workbook.clearBatches).toEqual([
      [
        {
          name: 'mergeCells',
          args: [
            [{ row: [0, 1], column: [0, 2] }],
            'merge-all',
            { id: 'sheet-1' },
          ],
        },
        {
          name: 'setCellFormatByRange',
          args: ['ht', '0', { row: [0, 1], column: [0, 2] }, { id: 'sheet-1' }],
        },
      ],
    ]);

    expect(editor.can().mergeSelectedCells('merge-across')).toBe(true);
    expect(editor.commands.mergeSelectedCells('merge-across')).toBe(true);
    expect(fixture.workbook.clearBatches.at(-1)).toEqual([
      {
        name: 'mergeCells',
        args: [
          [{ row: [0, 1], column: [0, 2] }],
          'merge-horizontal',
          { id: 'sheet-1' },
        ],
      },
    ]);
  });

  test('unmerges and fills through the native merge model', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets[0] = {
      ...fixture.context.content.sheets[0],
      config: {
        merge: {
          '0_0': { r: 0, c: 0, rs: 2, cs: 2 },
        },
      },
      data: [
        [
          {
            v: 'North',
            m: 'North',
            ht: '0',
            mc: { r: 0, c: 0, rs: 2, cs: 2 },
          },
          { mc: { r: 0, c: 0 } },
        ],
        [{ mc: { r: 0, c: 0 } }, { mc: { r: 0, c: 0 } }],
      ],
    };
    fixture.workbook.selection = [{ row: [0, 1], column: [0, 1] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().mergeSelectedCells('merge-cells')).toBe(false);
    expect(editor.can().mergeSelectedCells('unmerge-and-fill')).toBe(true);
    expect(editor.commands.mergeSelectedCells('unmerge-and-fill')).toBe(true);

    expect(fixture.workbook.clearBatches).toEqual([
      [
        {
          name: 'cancelMerge',
          args: [[{ row: [0, 1], column: [0, 1] }], { id: 'sheet-1' }],
        },
        {
          name: 'setCellValuesByRange',
          args: [
            [
              [
                { v: 'North', m: 'North', ht: '0' },
                { v: 'North', m: 'North', ht: '0' },
              ],
              [
                { v: 'North', m: 'North', ht: '0' },
                { v: 'North', m: 'North', ht: '0' },
              ],
            ],
            { row: [0, 1], column: [0, 1] },
            null,
            { id: 'sheet-1' },
          ],
        },
      ],
    ]);
  });

  test('uses Ctrl+M for the WPS merge-and-center command', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 1], column: [0, 2] }];
    const editor = spreadsheetEditor(fixture.context);
    const event = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: 'm',
    });

    expect(editor.handleKeyDown(event)).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(fixture.workbook.clearBatches).toHaveLength(1);
  });

  test('uses explicit recalculation commands', () => {
    const fixture = commandFixture();
    const { commands } = spreadsheetEditor(fixture.context);

    expect(commands.recalculateFormula('selection')).toBe(true);
    expect(commands.recalculateFormula('workbook')).toBe(true);

    expect(fixture.calculation.requests).toEqual([
      {
        scope: 'selection',
        range: { row: [2, 4], column: [1, 3] },
        sheetId: 'sheet-1',
      },
      { scope: 'workbook' },
    ]);
  });

  test('routes WPS clipboard actions through one typed command port', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetClipboard');
    expect(editor.can().pasteSelection()).toBe(true);
    expect(editor.can().cutSelection()).toBe(true);
    expect(editor.can().copySelection()).toBe(true);
    expect(editor.commands.pasteSelection()).toBe(true);
    expect(editor.commands.cutSelection()).toBe(true);
    expect(editor.commands.copySelection()).toBe(true);
    expect(fixture.clipboard.calls).toEqual(['paste', 'cut', 'copy']);

    editor.updateContext({
      ...fixture.context,
      clipboard: {
        ...fixture.clipboard,
        canCopySelection: false,
        canCutSelection: false,
        canPasteSelection: false,
      },
    });
    expect(editor.can().pasteSelection()).toBe(false);
    expect(editor.can().cutSelection()).toBe(false);
    expect(editor.can().copySelection()).toBe(false);
    expect(editor.commands.pasteSelection()).toBe(false);
    expect(editor.commands.cutSelection()).toBe(false);
    expect(editor.commands.copySelection()).toBe(false);
    expect(fixture.clipboard.calls).toEqual(['paste', 'cut', 'copy']);
  });

  test('owns WPS clipboard shortcuts through the same command port', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);
    for (const shortcut of [
      { key: 'c', metaKey: true },
      { ctrlKey: true, key: 'x' },
      { key: 'v', metaKey: true },
    ]) {
      const event = new KeyboardEvent('keydown', {
        ...shortcut,
        cancelable: true,
      });
      expect(editor.handleKeyDown(event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
    }
    expect(fixture.clipboard.calls).toEqual(['copy', 'cut', 'paste']);
  });

  test('routes the WPS format painter lifecycle through one typed command port', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);
    const target = {
      sheetId: 'sheet-2',
      selection: { row: [3, 4], column: [1, 2] },
    };

    expect(editor.extensionNames).toContain('spreadsheetFormatPainter');
    expect(editor.can().activateFormatPainter('once')).toBe(true);
    expect(editor.commands.activateFormatPainter('once')).toBe(true);
    expect(editor.can().applyFormatPainter(target)).toBe(true);
    expect(editor.commands.applyFormatPainter(target)).toBe(true);
    expect(editor.can().cancelFormatPainter()).toBe(true);
    expect(editor.commands.cancelFormatPainter()).toBe(true);
    expect(fixture.formatPainter.calls).toEqual([
      'activate:once',
      'apply:sheet-2:3-4:1-2',
      'cancel',
    ]);

    editor.updateContext({
      ...fixture.context,
      formatPainter: {
        ...fixture.formatPainter,
        active: false,
        canActivate: false,
      },
    });
    expect(editor.can().activateFormatPainter('locked')).toBe(false);
    expect(editor.commands.activateFormatPainter('locked')).toBe(false);
    expect(editor.can().applyFormatPainter(target)).toBe(false);
    expect(editor.commands.applyFormatPainter(target)).toBe(false);
    expect(editor.can().cancelFormatPainter()).toBe(false);
    expect(editor.commands.cancelFormatPainter()).toBe(false);
  });

  test('uses Escape to leave an active format painter without editing cells', () => {
    const fixture = commandFixture();
    fixture.formatPainter.active = true;
    const editor = spreadsheetEditor(fixture.context);
    const escapeEvent = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'Escape',
    });

    expect(editor.handleKeyDown(escapeEvent)).toBe(true);
    expect(escapeEvent.defaultPrevented).toBe(true);
    expect(fixture.formatPainter.calls).toEqual(['cancel']);
  });

  test('routes the AutoFilter lifecycle through one typed command port', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetAutoFilter');
    expect(editor.can().toggleAutoFilter()).toBe(true);
    expect(editor.commands.toggleAutoFilter()).toBe(true);
    fixture.autoFilter.active = true;
    fixture.autoFilter.canOpenMenu = true;
    expect(editor.can().openAutoFilterMenu()).toBe(true);
    expect(editor.commands.openAutoFilterMenu()).toBe(true);
    expect(fixture.autoFilter.calls).toEqual(['toggle', 'open']);

    editor.updateContext({
      ...fixture.context,
      autoFilter: {
        ...fixture.autoFilter,
        canOpenMenu: false,
        canToggle: false,
      },
    });
    expect(editor.can().toggleAutoFilter()).toBe(false);
    expect(editor.commands.toggleAutoFilter()).toBe(false);
    expect(editor.can().openAutoFilterMenu()).toBe(false);
    expect(editor.commands.openAutoFilterMenu()).toBe(false);
  });

  test('owns WPS AutoFilter toggle and header-menu shortcuts', () => {
    const fixture = commandFixture();
    fixture.autoFilter.active = true;
    fixture.autoFilter.canOpenMenu = true;
    const editor = spreadsheetEditor(fixture.context);

    for (const init of [
      { ctrlKey: true, key: 'l', shiftKey: true },
      { altKey: true, key: 'ArrowDown' },
    ]) {
      const event = new KeyboardEvent('keydown', {
        ...init,
        cancelable: true,
      });
      expect(editor.handleKeyDown(event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
    }
    expect(fixture.autoFilter.calls).toEqual(['toggle', 'open']);
  });

  test('leaves AutoFilter shortcuts with native text controls', () => {
    const fixture = commandFixture();
    fixture.autoFilter.active = true;
    fixture.autoFilter.canOpenMenu = true;
    const editor = spreadsheetEditor(fixture.context);
    const container = document.createElement('div');
    container.className = 'fortune-container';
    const input = document.createElement('input');
    container.append(input);
    document.body.append(container);
    const handled: boolean[] = [];
    input.addEventListener('keydown', (event) => {
      handled.push(editor.handleKeyDown(event));
    });

    input.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        ctrlKey: true,
        key: 'l',
        shiftKey: true,
      }),
    );
    input.dispatchEvent(
      new KeyboardEvent('keydown', {
        altKey: true,
        bubbles: true,
        cancelable: true,
        key: 'ArrowDown',
      }),
    );

    expect(handled).toEqual([false, false]);
    expect(fixture.autoFilter.calls).toEqual([]);
    container.remove();
  });

  test('owns the WPS F9 workbook recalculation shortcut', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);
    const recalculate = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'F9',
    });

    expect(editor.handleKeyDown(recalculate)).toBe(true);
    expect(recalculate.defaultPrevented).toBe(true);
    expect(fixture.calculation.requests).toEqual([{ scope: 'workbook' }]);
  });

  test('routes context-menu clear and paste through one workbook command port', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [3, 2], column: [4, 3] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().clearSelectedCells()).toBe(true);
    expect(editor.commands.clearSelectedCells()).toBe(true);
    expect(editor.commands.pasteCells([['A3S', 'Office']])).toBe(true);
    expect(editor.commands.pasteCells([])).toBe(false);

    expect(fixture.workbook.clearBatches).toHaveLength(1);
    expect(fixture.workbook.clearBatches[0]?.slice(0, 4)).toEqual([
      { name: 'clearCell', args: [2, 3, { id: 'sheet-1' }] },
      { name: 'clearCell', args: [2, 4, { id: 'sheet-1' }] },
      { name: 'clearCell', args: [3, 3, { id: 'sheet-1' }] },
      { name: 'clearCell', args: [3, 4, { id: 'sheet-1' }] },
    ]);
    expect(fixture.workbook.clearBatches[0]?.at(-1)).toMatchObject({
      name: 'updateSheet',
    });
    expect(fixture.workbook.pastes).toEqual([
      {
        range: { row: [2, 2], column: [3, 4] },
        sheetId: 'sheet-1',
        values: [['A3S', 'Office']],
      },
    ]);
    expect(fixture.workbook.selections).toEqual([
      {
        range: [{ row: [2, 2], column: [3, 4] }],
        sheetId: 'sheet-1',
      },
    ]);
    expect(fixture.formulaBarValues).toEqual(['', 'A3S']);
  });

  test('routes every WPS Clear variant through one typed workbook batch', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 0], column: [0, 0] }];
    const editor = spreadsheetEditor(fixture.context);

    for (const mode of [
      'contents',
      'formats',
      'comments',
      'hyperlinks',
      'all',
    ] as const) {
      expect(editor.can().clearSelectedCells(mode)).toBe(true);
      expect(editor.commands.clearSelectedCells(mode)).toBe(true);
    }

    expect(
      fixture.workbook.clearBatches.map((batch) =>
        batch.map((call) => call.name),
      ),
    ).toEqual([
      ['clearCell', 'updateSheet'],
      ['updateSheet'],
      ['updateSheet'],
      ['updateSheet'],
      ['clearCell', 'updateSheet'],
    ]);
    expect(fixture.formulaBarValues).toEqual(['', '']);
  });

  test('routes row and column structure actions through the workbook command port', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets[0] = {
      ...fixture.context.content.sheets[0],
      row: 12,
      column: 8,
    };
    fixture.workbook.selection = [{ row: [3, 4], column: [2, 3] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().insertSelectedStructure('row', 'before')).toBe(true);
    expect(editor.commands.insertSelectedStructure('row', 'before')).toBe(true);
    expect(editor.commands.insertSelectedStructure('column', 'after')).toBe(
      true,
    );
    expect(editor.commands.deleteSelectedStructure('row')).toBe(true);
    expect(editor.commands.setSelectedStructureHidden('column', true)).toBe(
      true,
    );
    expect(editor.commands.setSelectedStructureHidden('column', false)).toBe(
      true,
    );
    expect(editor.commands.setSelectedStructureSize('row', 36)).toBe(true);
    expect(editor.commands.setSelectedStructureSize('column', 128)).toBe(true);

    expect(fixture.workbook.structureChanges).toEqual([
      {
        action: 'insert',
        axis: 'row',
        count: 2,
        direction: 'lefttop',
        end: undefined,
        index: 3,
        sheetId: 'sheet-1',
        start: undefined,
      },
      {
        action: 'insert',
        axis: 'column',
        count: 2,
        direction: 'rightbottom',
        end: undefined,
        index: 3,
        sheetId: 'sheet-1',
        start: undefined,
      },
      {
        action: 'delete',
        axis: 'row',
        count: undefined,
        direction: undefined,
        end: 4,
        index: undefined,
        sheetId: 'sheet-1',
        start: 3,
      },
    ]);
    expect(fixture.workbook.selections).toEqual([
      {
        range: [{ row: [3, 3], column: [2, 3] }],
        sheetId: 'sheet-1',
      },
    ]);
    expect(fixture.workbook.visibilityChanges).toEqual([
      { axis: 'column', hidden: true, indices: ['2', '3'] },
      { axis: 'column', hidden: false, indices: ['2', '3'] },
    ]);
    expect(fixture.workbook.sizeChanges).toEqual([
      {
        axis: 'row',
        custom: true,
        sheetId: 'sheet-1',
        sizes: { 3: 36 },
      },
      {
        axis: 'column',
        custom: true,
        sheetId: 'sheet-1',
        sizes: { 2: 128, 3: 128 },
      },
    ]);
  });

  test('clamps the selection after deleting terminal rows and columns', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets[0] = {
      ...fixture.context.content.sheets[0],
      row: 12,
      column: 8,
    };
    fixture.workbook.selection = [{ row: [11, 11], column: [7, 7] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.commands.deleteSelectedStructure('row')).toBe(true);
    expect(fixture.workbook.selection).toEqual([
      { row: [10, 10], column: [7, 7] },
    ]);
    expect(editor.commands.deleteSelectedStructure('column')).toBe(true);
    expect(fixture.workbook.selection).toEqual([
      { row: [10, 10], column: [6, 6] },
    ]);
  });

  test('uses precomputed grid dimensions for structure capabilities', () => {
    const fixture = commandFixture();
    const data = new Proxy([[{ v: 'Anchor' }]], {
      ownKeys: () => {
        throw new Error('Structure capability rescanned the worksheet.');
      },
    });
    fixture.context.content = {
      ...fixture.context.content,
      sheets: [
        {
          ...fixture.context.content.sheets[0],
          data,
          row: 40,
          column: 12,
        },
      ],
    };
    fixture.context.targetSheetGridSize = {
      rowCount: 40,
      columnCount: 12,
    };
    const can = spreadsheetEditor(fixture.context).can();

    expect(can.insertSelectedStructure('column', 'before')).toBe(true);
    expect(can.deleteSelectedStructure('column')).toBe(true);
  });

  test('sorts the selected rows without dropping cell formatting', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [0, 2], column: [0, 1] }];
    fixture.workbook.cells = [
      [{ v: '研发', bl: 1 }, { v: 120 }],
      [{ v: '市场', fc: '#d84b4f' }, { v: 80 }],
      [{ v: '产品', bg: '#eef4ff' }, { v: 100 }],
    ];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().sortSelectedCells('ascending')).toBe(true);
    expect(editor.commands.sortSelectedCells('ascending')).toBe(true);
    expect(fixture.workbook.pastes.at(-1)).toEqual({
      range: { row: [0, 2], column: [0, 1] },
      sheetId: 'sheet-1',
      values: [
        [{ v: '产品', bg: '#eef4ff' }, { v: 100 }],
        [{ v: '市场', fc: '#d84b4f' }, { v: 80 }],
        [{ v: '研发', bl: 1 }, { v: 120 }],
      ],
    });
  });

  test('does not reuse a vendor-frozen paste range for the resulting selection', () => {
    const fixture = commandFixture();
    fixture.workbook.freezePastedRange = true;
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.commands.pasteCells([['A3S', 'Office']])).toBe(true);
    expect(fixture.workbook.selections).toEqual([
      {
        range: [{ row: [0, 0], column: [0, 1] }],
        sheetId: 'sheet-1',
      },
    ]);
  });

  test('updates controlled sheet view state without mutating the input', () => {
    const fixture = commandFixture();
    const previousSheet = fixture.context.content.sheets[0];
    const { commands } = spreadsheetEditor(fixture.context);

    expect(commands.setGridLines(false)).toBe(true);
    expect(commands.setZoom(175)).toBe(true);

    expect(previousSheet.showGridLines).toBe(true);
    expect(fixture.changes[0].sheets[0].showGridLines).toBe(false);
    expect(fixture.changes[1].sheets[0].zoomRatio).toBe(1.75);
  });

  test('routes WPS Freeze Panes presets through controlled workbook state', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [
      {
        row: [2, 2],
        column: [1, 1],
        row_focus: 2,
        column_focus: 1,
      },
    ];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.extensionNames).toContain('spreadsheetView');
    expect(editor.can().setFreezePanes('selection')).toBe(true);
    expect(editor.commands.setFreezePanes('selection')).toBe(true);
    expect(fixture.context.content.sheets[0]?.frozen).toBeUndefined();
    expect(fixture.changes[0]?.sheets[0]?.frozen).toEqual({
      type: 'rangeBoth',
      range: { row_focus: 1, column_focus: 0 },
    });

    editor.updateContext({
      ...fixture.context,
      content: fixture.changes[0],
    });
    expect(editor.can().setFreezePanes('selection')).toBe(false);
    expect(editor.can().setFreezePanes('none')).toBe(true);
    expect(editor.commands.setFreezePanes('none')).toBe(true);
    expect(fixture.changes[1]?.sheets[0]?.frozen).toBeUndefined();

    fixture.workbook.selection = [
      {
        row: [0, 0],
        column: [0, 0],
        row_focus: 0,
        column_focus: 0,
      },
    ];
    editor.updateContext(fixture.context);
    expect(editor.can().setFreezePanes('selection')).toBe(false);
  });

  test('routes history shortcuts through the history extension', () => {
    const fixture = commandFixture();
    const calls: string[] = [];
    fixture.context.history = {
      canRedo: false,
      canUndo: true,
      redo: () => {
        calls.push('redo');
        return true;
      },
      undo: () => {
        calls.push('undo');
        return true;
      },
    };
    const editor = spreadsheetEditor(fixture.context);
    const undo = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'z',
      metaKey: true,
    });

    expect(editor.handleKeyDown(undo)).toBe(true);
    expect(undo.defaultPrevented).toBe(true);
    expect(calls).toEqual(['undo']);

    const redo = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'z',
      metaKey: true,
      shiftKey: true,
    });
    expect(editor.handleKeyDown(redo)).toBe(false);
    expect(redo.defaultPrevented).toBe(false);
    expect(calls).toEqual(['undo']);
  });

  test('owns core cell-format and clear shortcuts', () => {
    const fixture = commandFixture();
    fixture.context.toolbarCell = { bl: 0, cl: 0, it: 1, un: 0 };
    const editor = spreadsheetEditor(fixture.context);

    const shortcuts = [
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: 'b',
        metaKey: true,
      }),
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: 'i',
        metaKey: true,
      }),
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: 'u',
        metaKey: true,
      }),
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: '5',
        metaKey: true,
      }),
      new KeyboardEvent('keydown', {
        cancelable: true,
        key: 'Delete',
      }),
    ];

    expect(shortcuts.map((event) => editor.handleKeyDown(event))).toEqual([
      true,
      true,
      true,
      true,
      true,
    ]);
    expect(shortcuts.every((event) => event.defaultPrevented)).toBe(true);
    expect(fixture.workbook.formats).toEqual([
      {
        attribute: 'bl',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: 1,
      },
      {
        attribute: 'it',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: 0,
      },
      {
        attribute: 'un',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: 1,
      },
      {
        attribute: 'cl',
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
        value: 1,
      },
    ]);
    expect(fixture.workbook.clearBatches).toHaveLength(1);
    expect(fixture.workbook.clearBatches[0]).toHaveLength(7);
  });

  test('owns worksheet creation and standard worksheet navigation shortcuts', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets.push({
      id: 'sheet-2',
      name: 'Sheet 2',
      order: 1,
      status: 0,
    });
    const editor = spreadsheetEditor(fixture.context);

    const createSheet = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'F11',
      shiftKey: true,
    });
    expect(editor.handleKeyDown(createSheet)).toBe(true);
    expect(createSheet.defaultPrevented).toBe(true);
    const created = fixture.changes.at(-1);
    if (!created) throw new Error('Expected the created workbook change.');
    expect(created.sheets.at(-1)).toMatchObject({
      id: 'sheet-3',
      name: '工作表 3',
      status: 1,
    });

    editor.updateContext({
      ...fixture.context,
      activeSheetId: 'sheet-3',
      content: created,
      targetSheetId: 'sheet-3',
    });
    const previousSheet = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: true,
      key: 'PageUp',
    });
    expect(editor.handleKeyDown(previousSheet)).toBe(true);
    expect(previousSheet.defaultPrevented).toBe(true);
    expect(fixture.changes.at(-1)?.sheets).toEqual([
      expect.objectContaining({ id: 'sheet-1', status: 0 }),
      expect.objectContaining({ id: 'sheet-2', status: 1 }),
      expect.objectContaining({ id: 'sheet-3', status: 0 }),
    ]);

    editor.updateContext({
      ...fixture.context,
      activeSheetId: 'sheet-2',
      content: fixture.changes.at(-1) ?? created,
      targetSheetId: 'sheet-2',
    });
    const nextSheetOnMac = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'PageDown',
      metaKey: true,
    });
    expect(editor.handleKeyDown(nextSheetOnMac)).toBe(true);
    expect(nextSheetOnMac.defaultPrevented).toBe(true);
    expect(fixture.changes.at(-1)?.sheets).toEqual([
      expect.objectContaining({ id: 'sheet-1', status: 0 }),
      expect.objectContaining({ id: 'sheet-2', status: 0 }),
      expect.objectContaining({ id: 'sheet-3', status: 1 }),
    ]);
  });

  test('keeps worksheet navigation available in read-only view without enabling edits', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets.push({
      id: 'sheet-2',
      name: 'Sheet 2',
      order: 1,
      status: 0,
    });
    fixture.context.content.sheets.push({
      id: 'sheet-3',
      name: 'Hidden',
      hide: 1,
      order: 2,
      status: 0,
    });
    fixture.context.editable = false;
    const activated: string[] = [];
    fixture.context.view = {
      activateSheet: (sheetId) => {
        activated.push(sheetId);
        return true;
      },
    };
    const editor = spreadsheetEditor(fixture.context);
    const nextSheet = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'PageDown',
      metaKey: true,
    });
    const createSheet = new KeyboardEvent('keydown', {
      cancelable: true,
      key: 'F11',
      shiftKey: true,
    });

    expect(editor.handleKeyDown(nextSheet)).toBe(true);
    expect(nextSheet.defaultPrevented).toBe(true);
    expect(editor.handleKeyDown(createSheet)).toBe(false);
    expect(createSheet.defaultPrevented).toBe(false);
    expect(editor.can().activateSheet('sheet-3')).toBe(false);
    expect(editor.commands.activateSheet('sheet-3')).toBe(false);
    expect(activated).toEqual(['sheet-2']);
    expect(fixture.changes).toEqual([]);
  });

  test('keeps worksheet activation local when an editable view owns it', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets.push({
      id: 'sheet-2',
      name: 'Sheet 2',
      order: 1,
      status: 0,
    });
    const activated: string[] = [];
    fixture.context.view = {
      activateSheet: (sheetId) => {
        activated.push(sheetId);
        return true;
      },
    };
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.commands.activateSheet('sheet-2')).toBe(true);
    expect(activated).toEqual(['sheet-2']);
    expect(fixture.changes).toEqual([]);
  });

  test('owns deterministic cell movement instead of relying on vendor key handlers', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets[0] = {
      ...fixture.context.content.sheets[0],
      row: 40,
      column: 12,
      data: Array.from({ length: 40 }, () => Array(12).fill(null)),
    };
    fixture.context.content.sheets[0].data?.[6]?.splice(6, 1, { v: 'last' });
    fixture.workbook.selection = [
      {
        row: [2, 2],
        column: [3, 3],
        row_focus: 2,
        column_focus: 3,
      },
    ];
    const editor = spreadsheetEditor(fixture.context);
    const press = (
      key: string,
      modifiers: Pick<KeyboardEventInit, 'metaKey' | 'shiftKey'> = {},
    ) => {
      const event = new KeyboardEvent('keydown', {
        cancelable: true,
        key,
        ...modifiers,
      });
      expect(editor.handleKeyDown(event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
    };

    press('ArrowRight');
    expect(fixture.workbook.selection?.[0]).toEqual({
      row: [2, 2],
      column: [4, 4],
      row_focus: 2,
      column_focus: 4,
    });
    press('Enter');
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [3, 3],
      column: [4, 4],
    });
    press('Enter', { shiftKey: true });
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [2, 2],
      column: [4, 4],
    });
    press('Tab');
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [2, 2],
      column: [5, 5],
    });
    press('Tab', { shiftKey: true });
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [2, 2],
      column: [4, 4],
    });
    press('Home');
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [2, 2],
      column: [0, 0],
    });
    press('End', { metaKey: true });
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [6, 6],
      column: [6, 6],
    });
    press('Home', { metaKey: true });
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [0, 0],
      column: [0, 0],
    });
    press('PageDown');
    expect(fixture.workbook.selection?.[0]).toMatchObject({
      row: [20, 20],
      column: [0, 0],
    });
  });

  test('owns row, column, all-cells, and extended selection shortcuts', () => {
    const fixture = commandFixture();
    fixture.context.content.sheets[0] = {
      ...fixture.context.content.sheets[0],
      row: 40,
      column: 12,
    };
    fixture.workbook.selection = [
      {
        row: [2, 2],
        column: [3, 3],
        row_focus: 2,
        column_focus: 3,
      },
    ];
    const editor = spreadsheetEditor(fixture.context);
    const press = (init: KeyboardEventInit) => {
      const event = new KeyboardEvent('keydown', {
        cancelable: true,
        ...init,
      });
      expect(editor.handleKeyDown(event)).toBe(true);
      expect(event.defaultPrevented).toBe(true);
    };

    press({ key: 'ArrowRight', shiftKey: true });
    press({ key: 'ArrowDown', shiftKey: true });
    expect(fixture.workbook.selection?.[0]).toEqual({
      row: [2, 3],
      column: [3, 4],
      row_focus: 3,
      column_focus: 4,
    });

    press({ ctrlKey: true, key: ' ' });
    expect(fixture.workbook.selection?.[0]).toEqual({
      row: [0, 39],
      column: [4, 4],
      row_focus: 3,
      column_focus: 4,
    });

    fixture.workbook.selection = [
      {
        row: [2, 2],
        column: [3, 3],
        row_focus: 2,
        column_focus: 3,
      },
    ];
    press({ key: ' ', shiftKey: true });
    expect(fixture.workbook.selection?.[0]).toEqual({
      row: [2, 2],
      column: [0, 11],
      row_focus: 2,
      column_focus: 3,
    });

    press({ key: 'a', metaKey: true });
    expect(fixture.workbook.selection?.[0]).toEqual({
      row: [0, 39],
      column: [0, 11],
      row_focus: 2,
      column_focus: 3,
    });
  });

  test('leaves grid navigation shortcuts with the focused toolbar control', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);
    const button = document.createElement('button');
    document.body.append(button);
    let handled = true;
    button.addEventListener('keydown', (event) => {
      handled = editor.handleKeyDown(event);
    });
    for (const key of ['ArrowRight', 'Delete']) {
      const event = new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key,
      });
      button.dispatchEvent(event);
      expect(handled).toBe(false);
      expect(event.defaultPrevented).toBe(false);
    }
    expect(fixture.workbook.selections).toEqual([]);
    expect(fixture.workbook.clearBatches).toEqual([]);
    button.remove();
  });

  test('owns formula-bar select-all in the keyboard extension', () => {
    const fixture = commandFixture();
    const editor = spreadsheetEditor(fixture.context);
    const formulaBar = document.createElement('div');
    formulaBar.className = 'fortune-fx-input';
    formulaBar.textContent = '=SUM(A1:A4)';
    const target = document.createElement('span');
    formulaBar.append(target);
    document.body.append(formulaBar);
    let handled = false;
    target.addEventListener('keydown', (event) => {
      handled = editor.handleKeyDown(event);
    });
    const selectAll = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'a',
      metaKey: true,
    });

    target.dispatchEvent(selectAll);

    expect(handled).toBe(true);
    expect(selectAll.defaultPrevented).toBe(true);
    expect(window.getSelection()?.toString()).toBe('=SUM(A1:A4)');
    window.getSelection()?.removeAllRanges();
    formulaBar.remove();
  });
});

function commandFixture(): {
  autoFilter: RecordingSpreadsheetAutoFilter;
  clipboard: RecordingSpreadsheetClipboard;
  changes: WorkSpreadsheetContent[];
  calculation: RecordingSpreadsheetCalculation;
  context: SpreadsheetCommandContext;
  formulaBarValues: unknown[];
  formatPainter: RecordingSpreadsheetFormatPainter;
  formatCells: RecordingSpreadsheetFormatCells;
  navigation: RecordingSpreadsheetNavigation;
  workbook: RecordingSpreadsheetWorkbook;
} {
  const content = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        showGridLines: true,
        zoomRatio: 1,
      },
    ],
  } satisfies WorkSpreadsheetContent;
  const changes: WorkSpreadsheetContent[] = [];
  const autoFilter = new RecordingSpreadsheetAutoFilter();
  const calculation = new RecordingSpreadsheetCalculation();
  const clipboard = new RecordingSpreadsheetClipboard();
  const formulaBarValues: unknown[] = [];
  const formatPainter = new RecordingSpreadsheetFormatPainter();
  const formatCells = new RecordingSpreadsheetFormatCells();
  const navigation = new RecordingSpreadsheetNavigation();
  const workbook = new RecordingSpreadsheetWorkbook();
  return {
    autoFilter,
    calculation,
    clipboard,
    changes,
    formulaBarValues,
    formatPainter,
    formatCells,
    navigation,
    workbook,
    context: {
      activeSheetId: 'sheet-1',
      autoFilter,
      calculation,
      clipboard,
      content,
      editable: true,
      fallbackRange: { row: [0, 1], column: [0, 2] },
      formulaBar: {
        setValue: (value) => formulaBarValues.push(value),
      },
      formatPainter,
      formatCells,
      history: null,
      navigation,
      onChange: (next) => changes.push(next),
      selection: {
        sheetId: 'sheet-1',
        selection: {
          row: [4, 2],
          column: [3, 1],
        },
      },
      targetSheetId: 'sheet-1',
      toolbarCell: null,
      view: null,
      workbook,
    },
  };
}

class RecordingSpreadsheetNavigation
  implements SpreadsheetNavigationCommandPort
{
  calls: string[] = [];
  canOpenFind = true;
  canOpenGoTo = true;

  openFind(): boolean {
    if (!this.canOpenFind) return false;
    this.calls.push('find');
    return true;
  }

  openGoTo(): boolean {
    if (!this.canOpenGoTo) return false;
    this.calls.push('go-to');
    return true;
  }
}

class RecordingSpreadsheetFormatCells
  implements SpreadsheetFormatCellsCommandPort
{
  canOpen = true;
  requests: Parameters<SpreadsheetFormatCellsCommandPort['open']>[0][] = [];

  open(
    request: Parameters<SpreadsheetFormatCellsCommandPort['open']>[0],
  ): boolean {
    if (!this.canOpen) return false;
    this.requests.push(request);
    return true;
  }
}

class RecordingSpreadsheetAutoFilter
  implements SpreadsheetAutoFilterCommandPort
{
  active = false;
  canOpenMenu = false;
  canToggle = true;
  calls: string[] = [];

  openMenu(): boolean {
    if (!this.canOpenMenu) return false;
    this.calls.push('open');
    return true;
  }

  toggle(): boolean {
    if (!this.canToggle) return false;
    this.calls.push('toggle');
    return true;
  }
}

class RecordingSpreadsheetFormatPainter
  implements SpreadsheetFormatPainterCommandPort
{
  active = false;
  canActivate = true;
  calls: string[] = [];
  mode: 'once' | 'locked' | null = null;

  activate(mode: 'once' | 'locked'): boolean {
    this.active = true;
    this.mode = mode;
    this.calls.push(`activate:${mode}`);
    return true;
  }

  applySelection(
    target: Parameters<
      SpreadsheetFormatPainterCommandPort['applySelection']
    >[0],
  ): boolean {
    if (!this.active) return false;
    const range = target.selection;
    this.calls.push(
      `apply:${target.sheetId}:${range.row[0]}-${range.row[1]}:${range.column[0]}-${range.column[1]}`,
    );
    return true;
  }

  cancel(): boolean {
    if (!this.active) return false;
    this.active = false;
    this.mode = null;
    this.calls.push('cancel');
    return true;
  }
}

class RecordingSpreadsheetClipboard implements SpreadsheetClipboardCommandPort {
  calls: string[] = [];
  canCopySelection = true;
  canCutSelection = true;
  canPasteSelection = true;

  copySelection(): boolean {
    this.calls.push('copy');
    return true;
  }

  cutSelection(): boolean {
    this.calls.push('cut');
    return true;
  }

  pasteSelection(): boolean {
    this.calls.push('paste');
    return true;
  }
}

function spreadsheetEditor(context: SpreadsheetCommandContext) {
  return createOfficeEditorRuntime<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >(context, createSpreadsheetEditorExtensions());
}

class RecordingSpreadsheetWorkbook implements SpreadsheetWorkbookCommandPort {
  cells: (Cell | null)[][] = [];
  clearBatches: Array<Array<{ name: string; args: unknown[] }>> = [];
  failFill = false;
  fillRowsAtCall: number[][] = [];
  fills: Array<{
    applyRange: SpreadsheetCommandRange;
    copyRange: SpreadsheetCommandRange;
    direction: 'down' | 'left' | 'right' | 'up';
  }> = [];
  sheet: Sheet = { id: 'sheet-1', name: 'Sheet 1', data: [[null]] };
  formats: Array<{
    attribute: string;
    range: SpreadsheetCommandRange;
    sheetId: string | undefined;
    value: unknown;
  }> = [];
  pastes: Array<{
    range: SpreadsheetCommandRange;
    sheetId: string | undefined;
    values: unknown[][];
  }> = [];
  selections: Array<{
    range: SpreadsheetCommandRange[];
    sheetId: string | undefined;
  }> = [];
  selection: SpreadsheetCommandRange[] | undefined;
  freezePastedRange = false;
  sizeChanges: Array<{
    axis: 'row' | 'column';
    custom: boolean;
    sheetId: string | undefined;
    sizes: Record<string, number>;
  }> = [];
  structureChanges: Array<{
    action: 'insert' | 'delete';
    axis: 'row' | 'column';
    count: number | undefined;
    direction: 'lefttop' | 'rightbottom' | undefined;
    end: number | undefined;
    index: number | undefined;
    sheetId: string | undefined;
    start: number | undefined;
  }> = [];
  visibilityChanges: Array<{
    axis: 'row' | 'column';
    hidden: boolean;
    indices: string[];
  }> = [];

  batchCallApis(apiCalls: Array<{ name: string; args: unknown[] }>): void {
    this.clearBatches.push(apiCalls);
    for (const apiCall of apiCalls) {
      if (apiCall.name === 'deleteRowOrColumn') {
        const axis = apiCall.args[0] as 'column' | 'row';
        const start = apiCall.args[1] as number;
        const end = apiCall.args[2] as number;
        const options = apiCall.args[3] as { id?: string } | undefined;
        this.structureChanges.push({
          action: 'delete',
          axis,
          count: undefined,
          direction: undefined,
          end,
          index: undefined,
          sheetId: options?.id,
          start,
        });
      }
      if (apiCall.name === 'setSelection') {
        this.setSelection(
          apiCall.args[0] as SpreadsheetCommandRange[],
          apiCall.args[1] as { id?: string } | undefined,
        );
      }
    }
  }

  autoFillCell(
    copyRange: SpreadsheetCommandRange,
    applyRange: SpreadsheetCommandRange,
    direction: 'down' | 'left' | 'right' | 'up',
  ): void {
    const firstRow = Math.min(copyRange.row[0] ?? 0, applyRange.row[0] ?? 0);
    const lastRow = Math.max(copyRange.row[1] ?? 0, applyRange.row[1] ?? 0);
    const rows = Array.from(
      { length: lastRow - firstRow + 1 },
      (_, index) => firstRow + index,
    );
    this.fillRowsAtCall.push(
      rows.filter((row) => Array.isArray(this.sheet.data?.[row])),
    );
    if (rows.some((row) => !Array.isArray(this.sheet.data?.[row]))) {
      throw new Error('Native fill requires materialized data rows.');
    }
    if (this.failFill) throw new Error('Native fill failed.');
    this.fills.push({ applyRange, copyRange, direction });
  }

  getSelection(): SpreadsheetCommandRange[] | undefined {
    return this.selection;
  }

  getCellsByRange(): (Cell | null)[][] {
    return this.cells;
  }

  getSheet(): Sheet {
    return this.sheet;
  }

  insertRowOrColumn(
    axis: 'row' | 'column',
    index: number,
    count: number,
    direction: 'lefttop' | 'rightbottom',
    options?: { id?: string },
  ): void {
    this.structureChanges.push({
      action: 'insert',
      axis,
      count,
      direction,
      end: undefined,
      index,
      sheetId: options?.id,
      start: undefined,
    });
  }

  hideRowOrColumn(indices: string[], axis: 'row' | 'column'): void {
    this.visibilityChanges.push({ axis, hidden: true, indices });
  }

  showRowOrColumn(indices: string[], axis: 'row' | 'column'): void {
    this.visibilityChanges.push({ axis, hidden: false, indices });
  }

  setRowHeight(
    sizes: Record<string, number>,
    options?: { id?: string },
    custom = false,
  ): void {
    this.sizeChanges.push({
      axis: 'row',
      custom,
      sheetId: options?.id,
      sizes,
    });
  }

  setColumnWidth(
    sizes: Record<string, number>,
    options?: { id?: string },
    custom = false,
  ): void {
    this.sizeChanges.push({
      axis: 'column',
      custom,
      sheetId: options?.id,
      sizes,
    });
  }

  setCellValuesByRange(
    values: unknown[][],
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ): void {
    this.pastes.push({ range, sheetId: options?.id, values });
    if (this.freezePastedRange) {
      Object.freeze(range.row);
      Object.freeze(range.column);
      Object.freeze(range);
    }
  }

  setSelection(
    range: SpreadsheetCommandRange[],
    options?: { id?: string },
  ): void {
    if (range.some((selection) => Object.isFrozen(selection))) {
      throw new TypeError('Workbook selection ranges must remain mutable.');
    }
    this.selections.push({ range, sheetId: options?.id });
    this.selection = range;
  }

  setCellFormatByRange(
    attribute: Parameters<
      SpreadsheetWorkbookCommandPort['setCellFormatByRange']
    >[0],
    value: unknown,
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ): void {
    this.formats.push({
      attribute,
      range,
      sheetId: options?.id,
      value,
    });
  }
}

class RecordingSpreadsheetCalculation
  implements SpreadsheetCalculationCommandPort
{
  requests: Parameters<SpreadsheetCalculationCommandPort['recalculate']>[0][] =
    [];

  recalculate(
    request: Parameters<SpreadsheetCalculationCommandPort['recalculate']>[0],
  ): void {
    this.requests.push(request);
  }
}
