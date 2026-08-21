import { expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
  SpreadsheetTableCommandPort,
  SpreadsheetWorkbookCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import { createSpreadsheetTableExtension } from '../src/internal/features/work/editors/spreadsheet-table-command';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('opens and commits a native table through typed commands', () => {
  const content = tableContent();
  const changes: WorkSpreadsheetContent[] = [];
  const requests: Parameters<SpreadsheetTableCommandPort['open']>[0][] = [];
  const table: SpreadsheetTableCommandPort = {
    canOpen: true,
    open: (target) => {
      requests.push(target);
      return true;
    },
  };
  const workbook = {
    getSelection: () => [
      {
        row: [1, 1],
        column: [1, 1],
        row_focus: 1,
        column_focus: 1,
      },
    ],
  } as SpreadsheetWorkbookCommandPort;
  const editor = createOfficeEditorRuntime<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >(
    tableContext(content, table, workbook, (next) => changes.push(next)),
    [createSpreadsheetTableExtension()],
  );

  expect(editor.extensionNames).toEqual(['spreadsheetTables']);
  expect(editor.can().openTable()).toBe(true);
  expect(editor.commands.openTable()).toBe(true);
  expect(requests).toEqual([
    {
      sheetId: 'sheet-1',
      selection: {
        row: [1, 1],
        column: [1, 1],
        row_focus: 1,
        column_focus: 1,
      },
    },
  ]);

  expect(
    editor.commands.applyTable({
      headerRow: true,
      name: 'Table1',
      range: { row: [0, 2], column: [0, 2] },
      sheetId: 'sheet-1',
    }),
  ).toBe(true);
  expect(changes).toHaveLength(1);
  expect(changes[0]?.sheets[0]?.tables?.[0]).toMatchObject({
    name: 'Table1',
    style: { family: 'medium', number: 2 },
  });
});

test('owns Cmd/Ctrl+T only while the spreadsheet grid has focus', () => {
  const requests: unknown[] = [];
  const table: SpreadsheetTableCommandPort = {
    canOpen: true,
    open: (target) => {
      requests.push(target);
      return true;
    },
  };
  const workbook = {
    getSelection: () => [{ row: [0, 0], column: [0, 0] }],
  } as SpreadsheetWorkbookCommandPort;
  const editor = createOfficeEditorRuntime<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >(
    tableContext(tableContent(), table, workbook, () => undefined),
    [createSpreadsheetTableExtension()],
  );
  const fortune = document.createElement('div');
  fortune.className = 'fortune-container';
  const grid = document.createElement('div');
  grid.className = 'fortune-sheet-overlay';
  fortune.append(grid);
  const input = document.createElement('input');
  document.body.append(fortune, input);
  const handled: boolean[] = [];
  grid.addEventListener('keydown', (event) => {
    handled.push(editor.handleKeyDown(event));
  });
  input.addEventListener('keydown', (event) => {
    handled.push(editor.handleKeyDown(event));
  });

  const gridShortcut = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    ctrlKey: true,
    key: 't',
  });
  grid.dispatchEvent(gridShortcut);
  const inputShortcut = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    ctrlKey: true,
    key: 't',
  });
  input.dispatchEvent(inputShortcut);

  expect(handled).toEqual([true, false]);
  expect(gridShortcut.defaultPrevented).toBe(true);
  expect(inputShortcut.defaultPrevented).toBe(false);
  expect(requests).toHaveLength(1);
  fortune.remove();
  input.remove();
});

function tableContext(
  content: WorkSpreadsheetContent,
  table: SpreadsheetTableCommandPort,
  workbook: SpreadsheetWorkbookCommandPort,
  onChange: (content: WorkSpreadsheetContent) => void,
): SpreadsheetCommandContext {
  return {
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
    content,
    dataValidation: { canOpen: false, open: () => false },
    editable: true,
    fallbackRange: { row: [0, 0], column: [0, 0] },
    formulaBar: null,
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
    onChange,
    selection: null,
    table,
    targetSheetId: 'sheet-1',
    toolbarCell: null,
    view: null,
    workbook,
  };
}

function tableContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sales',
        row: 8,
        column: 6,
        data: [
          [{ v: 'Region' }, { v: 'Revenue' }, { v: 'Status' }],
          [{ v: 'East' }, { v: 10 }, { v: 'Ready' }],
          [{ v: 'West' }, { v: 12 }, { v: 'Blocked' }],
        ],
      },
    ],
  };
}
