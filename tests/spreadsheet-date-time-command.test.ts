import {
  type Sheet,
  update as formatFortuneCellValue,
} from '@fortune-sheet/core';
import { expect, rstest, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import {
  createSpreadsheetDateTimeExtension,
  spreadsheetDateTimeEntry,
} from '../src/internal/features/work/editors/spreadsheet-date-time-command';
import type {
  SpreadsheetCommandContext,
  SpreadsheetWorkbookCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';

test('creates stable Excel 1900 date and minute-precision time entries', () => {
  const now = new Date(2026, 7, 22, 9, 7, 58, 900);

  expect(spreadsheetDateTimeEntry('date', now)).toEqual({
    format: { fa: 'yyyy-MM-dd', t: 'd' },
    formulaBarValue: '2026-08-22',
    value: 46_256,
  });
  expect(spreadsheetDateTimeEntry('time', now)).toEqual({
    format: { fa: 'hh:mm', t: 'd' },
    formulaBarValue: '09:07',
    value: (9 * 60 + 7) / 1_440,
  });
  expect(spreadsheetDateTimeEntry('date', new Date(Number.NaN))).toBeNull();
});

test('renders Excel day-zero time fractions through the real Fortune formatter', () => {
  expect(formatFortuneCellValue('hh:mm', (9 * 60 + 7) / 1_440)).toBe('09:07');
  expect(formatFortuneCellValue('hh:mm', 0)).toBe('00:00');
});

test('inserts date and time through one native batch at the active cell', () => {
  const fixture = dateTimeFixture();
  const editor = createOfficeEditorRuntime(fixture.context, [
    createSpreadsheetDateTimeExtension(),
  ]);
  rstest.useFakeTimers();
  rstest.setSystemTime(new Date(2026, 7, 22, 9, 7, 58, 900));

  try {
    expect(editor.can().insertCurrentDateTime('date')).toBe(true);
    expect(editor.commands.insertCurrentDateTime('date')).toBe(true);
    expect(editor.commands.insertCurrentDateTime('time')).toBe(true);
  } finally {
    rstest.useRealTimers();
  }

  expect(fixture.batches).toEqual([
    [
      {
        name: 'setCellValuesByRange',
        args: [[[46_256]], { row: [3, 3], column: [2, 2] }, { id: 'sheet-1' }],
      },
      {
        name: 'setCellFormatByRange',
        args: [
          'ct',
          { fa: 'yyyy-MM-dd', t: 'd' },
          { row: [3, 3], column: [2, 2] },
          { id: 'sheet-1' },
        ],
      },
    ],
    [
      {
        name: 'setCellValuesByRange',
        args: [
          [[(9 * 60 + 7) / 1_440]],
          { row: [3, 3], column: [2, 2] },
          { id: 'sheet-1' },
        ],
      },
      {
        name: 'setCellFormatByRange',
        args: [
          'ct',
          { fa: 'hh:mm', t: 'd' },
          { row: [3, 3], column: [2, 2] },
          { id: 'sheet-1' },
        ],
      },
    ],
  ]);
  expect(fixture.formulaBarValues).toEqual(['2026-08-22', '09:07']);
});

test('owns the WPS date and time shortcuts only on the editable grid', () => {
  const fixture = dateTimeFixture();
  const editor = createOfficeEditorRuntime(fixture.context, [
    createSpreadsheetDateTimeExtension(),
  ]);
  const container = document.createElement('div');
  container.className = 'fortune-container';
  const grid = document.createElement('div');
  grid.className = 'fortune-sheet-overlay';
  container.append(grid);
  document.body.append(container);
  let handled = false;
  grid.addEventListener('keydown', (event) => {
    handled = editor.handleKeyDown(event);
  });

  const dateShortcut = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    code: 'Semicolon',
    ctrlKey: true,
    key: ';',
  });
  grid.dispatchEvent(dateShortcut);
  expect(handled).toBe(true);
  expect(dateShortcut.defaultPrevented).toBe(true);

  const timeShortcut = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    code: 'Semicolon',
    ctrlKey: true,
    key: ':',
    shiftKey: true,
  });
  grid.dispatchEvent(timeShortcut);
  expect(handled).toBe(true);
  expect(timeShortcut.defaultPrevented).toBe(true);
  expect(fixture.batches).toHaveLength(2);

  const hostInput = document.createElement('input');
  document.body.append(hostInput);
  handled = true;
  hostInput.addEventListener('keydown', (event) => {
    handled = editor.handleKeyDown(event);
  });
  hostInput.dispatchEvent(
    new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'Semicolon',
      ctrlKey: true,
      key: ';',
    }),
  );
  expect(handled).toBe(false);
  expect(fixture.batches).toHaveLength(2);

  editor.updateContext({ ...fixture.context, editable: false });
  handled = true;
  grid.dispatchEvent(
    new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      code: 'Semicolon',
      ctrlKey: true,
      key: ';',
    }),
  );
  expect(handled).toBe(false);
  expect(fixture.batches).toHaveLength(2);

  hostInput.remove();
  container.remove();
});

test('rejects unsafe date and time insertion targets before mutation', () => {
  for (const fixture of [
    dateTimeFixture({ config: { authority: { sheet: 1 } } }),
    dateTimeFixture({
      config: { merge: { '3_2': { r: 3, c: 2, rs: 1, cs: 2 } } },
    }),
    dateTimeFixture({ pivotTable: { enabled: true } }),
    dateTimeFixture({}, { row: [1_048_576, 1_048_576], column: [0, 0] }),
  ]) {
    const editor = createOfficeEditorRuntime(fixture.context, [
      createSpreadsheetDateTimeExtension(),
    ]);
    expect(editor.can().insertCurrentDateTime('date')).toBe(false);
    expect(editor.commands.insertCurrentDateTime('date')).toBe(false);
    expect(fixture.batches).toEqual([]);
  }

  const inactive = dateTimeFixture();
  inactive.context = { ...inactive.context, activeSheetId: 'sheet-2' };
  const inactiveEditor = createOfficeEditorRuntime(inactive.context, [
    createSpreadsheetDateTimeExtension(),
  ]);
  expect(inactiveEditor.can().insertCurrentDateTime('time')).toBe(false);
  expect(inactiveEditor.commands.insertCurrentDateTime('time')).toBe(false);
});

function dateTimeFixture(
  sheetOverride: Partial<Sheet> = {},
  selectionOverride: {
    row: number[];
    column: number[];
    row_focus?: number;
    column_focus?: number;
  } = {
    row: [2, 4],
    column: [1, 3],
    row_focus: 3,
    column_focus: 2,
  },
): {
  batches: Array<Array<{ name: string; args: unknown[] }>>;
  context: SpreadsheetCommandContext;
  formulaBarValues: unknown[];
} {
  const batches: Array<Array<{ name: string; args: unknown[] }>> = [];
  const formulaBarValues: unknown[] = [];
  const sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 10,
    column: 6,
    data: [[{ v: 'Existing' }]],
    ...sheetOverride,
  } satisfies Sheet;
  const workbook: SpreadsheetWorkbookCommandPort = {
    autoFillCell: () => undefined,
    batchCallApis: (calls) => batches.push(calls),
    getCellsByRange: () => [],
    getSelection: () => [selectionOverride],
    getSheet: () => sheet,
    hideRowOrColumn: () => undefined,
    insertRowOrColumn: () => undefined,
    setCellFormatByRange: () => undefined,
    setCellValuesByRange: () => undefined,
    setColumnWidth: () => undefined,
    setRowHeight: () => undefined,
    setSelection: () => undefined,
    showRowOrColumn: () => undefined,
  };
  const context: SpreadsheetCommandContext = {
    activeSheetId: 'sheet-1',
    autoFilter: {
      active: false,
      canOpenMenu: false,
      canToggle: false,
      openMenu: () => false,
      toggle: () => false,
    },
    calculation: null,
    clipboard: {
      canCopySelection: false,
      canCutSelection: false,
      canOpenPasteSpecial: false,
      canPasteSelection: false,
      canPasteSpecial: () => false,
      copySelection: () => false,
      cutSelection: () => false,
      openPasteSpecial: () => false,
      pasteSelection: () => false,
      pasteSpecial: () => false,
    },
    content: { type: 'spreadsheet', sheets: [sheet] },
    dataValidation: { canOpen: false, open: () => false },
    editable: true,
    fallbackRange: { row: [0, 0], column: [0, 0] },
    formulaBar: { setValue: (value) => formulaBarValues.push(value) },
    formatCells: { canOpen: false, open: () => false },
    formatPainter: {
      active: false,
      canActivate: false,
      mode: null,
      activate: () => false,
      applySelection: () => false,
      cancel: () => false,
    },
    history: null,
    hyperlink: { canOpen: false, open: () => false },
    navigation: {
      canOpenFind: false,
      canOpenGoTo: false,
      openFind: () => false,
      openGoTo: () => false,
    },
    onChange: () => undefined,
    selection: null,
    table: { canOpen: false, open: () => false },
    targetSheetId: 'sheet-1',
    toolbarCell: null,
    view: null,
    workbook,
  };
  return { batches, context, formulaBarValues };
}
