import type { Sheet } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import { createSpreadsheetCopyFromAboveExtension } from '../src/internal/features/work/editors/spreadsheet-copy-from-above-command';
import type {
  SpreadsheetCommandContext,
  SpreadsheetWorkbookCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';

test('copies a formula or cached value through one native batch', () => {
  const fixture = copyFromAboveFixture();
  const editor = createOfficeEditorRuntime(fixture.context, [
    createSpreadsheetCopyFromAboveExtension(),
  ]);

  expect(editor.can().copyCellFromAbove('formula')).toBe(true);
  expect(editor.commands.copyCellFromAbove('formula')).toBe(true);
  expect(editor.commands.copyCellFromAbove('value')).toBe(true);
  expect(fixture.batches).toEqual([
    [
      {
        name: 'setCellValuesByRange',
        args: [
          [['=$A$1+A1']],
          { row: [1, 1], column: [1, 1] },
          { id: 'sheet-1' },
        ],
      },
    ],
    [
      {
        name: 'setCellValuesByRange',
        args: [[[4]], { row: [1, 1], column: [1, 1] }, { id: 'sheet-1' }],
      },
    ],
  ]);
  expect(fixture.formulaBarValues).toEqual(['=$A$1+A1', 4]);
});

test('owns both physical apostrophe shortcuts only on the editable grid', () => {
  const fixture = copyFromAboveFixture();
  const editor = createOfficeEditorRuntime(fixture.context, [
    createSpreadsheetCopyFromAboveExtension(),
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

  const formula = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    code: 'Quote',
    ctrlKey: true,
    key: "'",
  });
  grid.dispatchEvent(formula);
  expect(handled).toBe(true);
  expect(formula.defaultPrevented).toBe(true);

  const value = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    code: 'Quote',
    ctrlKey: true,
    key: '"',
    shiftKey: true,
  });
  grid.dispatchEvent(value);
  expect(handled).toBe(true);
  expect(value.defaultPrevented).toBe(true);
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
      code: 'Quote',
      ctrlKey: true,
      key: "'",
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
      code: 'Quote',
      ctrlKey: true,
      key: "'",
    }),
  );
  expect(handled).toBe(false);
  expect(fixture.batches).toHaveLength(2);

  hostInput.remove();
  container.remove();
});

test('rejects unsafe targets and reports native batch failures', () => {
  for (const fixture of [
    copyFromAboveFixture({ config: { authority: { sheet: 1 } } }),
    copyFromAboveFixture(
      { config: { merge: { '1_1': { r: 1, c: 1, rs: 1, cs: 2 } } } },
      { row: [1, 1], column: [1, 1], row_focus: 1, column_focus: 1 },
    ),
    copyFromAboveFixture(
      {},
      {
        row: [0, 0],
        column: [1, 1],
        row_focus: 0,
        column_focus: 1,
      },
    ),
  ]) {
    const editor = createOfficeEditorRuntime(fixture.context, [
      createSpreadsheetCopyFromAboveExtension(),
    ]);
    expect(editor.can().copyCellFromAbove('formula')).toBe(false);
    expect(editor.commands.copyCellFromAbove('formula')).toBe(false);
    expect(fixture.batches).toEqual([]);
  }

  const inactive = copyFromAboveFixture();
  inactive.context = { ...inactive.context, activeSheetId: 'sheet-2' };
  const inactiveEditor = createOfficeEditorRuntime(inactive.context, [
    createSpreadsheetCopyFromAboveExtension(),
  ]);
  expect(inactiveEditor.can().copyCellFromAbove('value')).toBe(false);

  const failing = copyFromAboveFixture();
  const workbook = failing.context.workbook;
  if (!workbook) throw new Error('Fixture workbook is missing.');
  failing.context.workbook = {
    ...workbook,
    batchCallApis: () => {
      throw new Error('detached workbook');
    },
  };
  const failingEditor = createOfficeEditorRuntime(failing.context, [
    createSpreadsheetCopyFromAboveExtension(),
  ]);
  expect(failingEditor.can().copyCellFromAbove('formula')).toBe(true);
  expect(failingEditor.commands.copyCellFromAbove('formula')).toBe(false);
});

function copyFromAboveFixture(
  sheetOverride: Partial<Sheet> = {},
  selectionOverride: {
    row: number[];
    column: number[];
    row_focus?: number;
    column_focus?: number;
  } = {
    row: [1, 2],
    column: [0, 1],
    row_focus: 1,
    column_focus: 1,
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
    row: 4,
    column: 3,
    data: [
      [{ v: 2 }, { f: '=$A$1+A1', v: 4, m: '4.00', bl: 1 }],
      [null, { v: 'Old', it: 1 }],
    ],
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
    fallbackRange: selectionOverride,
    formatCells: { canOpen: false, open: () => false },
    formatPainter: {
      active: false,
      canActivate: false,
      mode: null,
      activate: () => false,
      applySelection: () => false,
      cancel: () => false,
    },
    formulaBar: { setValue: (value) => formulaBarValues.push(value) },
    gridSize: { columnCount: 3, rowCount: 4 },
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
