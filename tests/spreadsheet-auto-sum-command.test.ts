import type { Sheet } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import { createSpreadsheetAutoSumExtension } from '../src/internal/features/work/editors/spreadsheet-auto-sum-command';
import type {
  SpreadsheetCommandContext,
  SpreadsheetWorkbookCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';

test('applies AutoSum through one native batch and updates the formula bar', () => {
  const fixture = autoSumFixture();
  const editor = createOfficeEditorRuntime(fixture.context, [
    createSpreadsheetAutoSumExtension(),
  ]);

  expect(editor.can().applyAutoSum('sum')).toBe(true);
  expect(editor.commands.applyAutoSum('sum')).toBe(true);
  expect(fixture.batches).toEqual([
    [
      {
        name: 'setCellValuesByRange',
        args: [
          [['=SUM(A2:A3)']],
          { row: [3, 3], column: [0, 0] },
          { id: 'sheet-1' },
        ],
      },
      {
        name: 'setSelection',
        args: [
          [
            {
              row: [3, 3],
              column: [0, 0],
              row_focus: 3,
              column_focus: 0,
            },
          ],
          { id: 'sheet-1' },
        ],
      },
    ],
  ]);
  expect(fixture.formulaBarValues).toEqual(['=SUM(A2:A3)']);
});

test('owns Alt+= only on an editable spreadsheet grid', () => {
  const fixture = autoSumFixture();
  const editor = createOfficeEditorRuntime(fixture.context, [
    createSpreadsheetAutoSumExtension(),
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
  const gridShortcut = new KeyboardEvent('keydown', {
    altKey: true,
    bubbles: true,
    cancelable: true,
    code: 'Equal',
    key: '=',
  });

  grid.dispatchEvent(gridShortcut);
  expect(handled).toBe(true);
  expect(gridShortcut.defaultPrevented).toBe(true);
  expect(fixture.batches).toHaveLength(1);

  const hostInput = document.createElement('input');
  document.body.append(hostInput);
  handled = true;
  hostInput.addEventListener('keydown', (event) => {
    handled = editor.handleKeyDown(event);
  });
  hostInput.dispatchEvent(
    new KeyboardEvent('keydown', {
      altKey: true,
      bubbles: true,
      cancelable: true,
      code: 'Equal',
      key: '=',
    }),
  );
  expect(handled).toBe(false);
  expect(fixture.batches).toHaveLength(1);

  editor.updateContext({ ...fixture.context, editable: false });
  handled = true;
  grid.dispatchEvent(
    new KeyboardEvent('keydown', {
      altKey: true,
      bubbles: true,
      cancelable: true,
      code: 'Equal',
      key: '=',
    }),
  );
  expect(handled).toBe(false);
  expect(fixture.batches).toHaveLength(1);

  hostInput.remove();
  container.remove();
});

test('rejects protected or merged AutoSum targets', () => {
  const protectedFixture = autoSumFixture({
    config: { authority: { sheet: 1 } },
  });
  const mergedFixture = autoSumFixture({
    config: { merge: { '3_0': { r: 3, c: 0, rs: 1, cs: 2 } } },
  });
  const sparseLockedFixture = autoSumFixture({
    data: undefined,
    celldata: [
      { r: 1, c: 0, v: { v: 12 } },
      { r: 2, c: 0, v: { v: 18 } },
      { r: 3, c: 0, v: { lo: 1 } },
    ],
  });
  const protectedEditor = createOfficeEditorRuntime(protectedFixture.context, [
    createSpreadsheetAutoSumExtension(),
  ]);
  const mergedEditor = createOfficeEditorRuntime(mergedFixture.context, [
    createSpreadsheetAutoSumExtension(),
  ]);
  const sparseLockedEditor = createOfficeEditorRuntime(
    sparseLockedFixture.context,
    [createSpreadsheetAutoSumExtension()],
  );

  expect(protectedEditor.can().applyAutoSum('sum')).toBe(false);
  expect(protectedEditor.commands.applyAutoSum('sum')).toBe(false);
  expect(mergedEditor.can().applyAutoSum('sum')).toBe(false);
  expect(mergedEditor.commands.applyAutoSum('sum')).toBe(false);
  expect(sparseLockedEditor.can().applyAutoSum('sum')).toBe(false);
  expect(sparseLockedEditor.commands.applyAutoSum('sum')).toBe(false);
});

function autoSumFixture(sheetOverride: Partial<Sheet> = {}): {
  batches: Array<Array<{ name: string; args: unknown[] }>>;
  context: SpreadsheetCommandContext;
  formulaBarValues: unknown[];
} {
  const batches: Array<Array<{ name: string; args: unknown[] }>> = [];
  const formulaBarValues: unknown[] = [];
  const sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 8,
    column: 4,
    data: [[{ v: 'Revenue' }], [{ v: 12 }], [{ v: 18 }], [null]],
    ...sheetOverride,
  } satisfies Sheet;
  const workbook: SpreadsheetWorkbookCommandPort = {
    autoFillCell: () => undefined,
    batchCallApis: (calls) => batches.push(calls),
    getCellsByRange: () => [],
    getSelection: () => [{ row: [3, 3], column: [0, 0] }],
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
    fallbackRange: { row: [3, 3], column: [0, 0] },
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
    navigation: {
      canOpenFind: false,
      canOpenGoTo: false,
      openFind: () => false,
      openGoTo: () => false,
    },
    onChange: () => undefined,
    selection: null,
    targetSheetId: 'sheet-1',
    toolbarCell: null,
    view: null,
    workbook,
  };
  return { batches, context, formulaBarValues };
}
