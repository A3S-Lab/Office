import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import {
  createSpreadsheetEditorExtensions,
  type SpreadsheetCalculationCommandPort,
  type SpreadsheetCommandContext,
  type SpreadsheetCommandRange,
  type SpreadsheetEditorCommands,
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

  test('uses explicit merge and recalculation commands', () => {
    const fixture = commandFixture();
    const { commands } = spreadsheetEditor(fixture.context);

    expect(commands.toggleCellMerge(false)).toBe(true);
    expect(commands.recalculateFormula('selection')).toBe(true);
    expect(commands.recalculateFormula('workbook')).toBe(true);

    expect(fixture.workbook.merges).toEqual([
      {
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
      },
    ]);
    expect(fixture.calculation.requests).toEqual([
      {
        scope: 'selection',
        range: { row: [2, 4], column: [1, 3] },
        sheetId: 'sheet-1',
      },
      { scope: 'workbook' },
    ]);
  });

  test('routes context-menu clear and paste through one workbook command port', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [3, 2], column: [4, 3] }];
    const editor = spreadsheetEditor(fixture.context);

    expect(editor.can().clearSelectedCells()).toBe(true);
    expect(editor.commands.clearSelectedCells()).toBe(true);
    expect(editor.commands.pasteCells([['A3S', 'Office']])).toBe(true);
    expect(editor.commands.pasteCells([])).toBe(false);

    expect(fixture.workbook.clearBatches).toEqual([
      [
        { name: 'clearCell', args: [2, 3, { id: 'sheet-1' }] },
        { name: 'clearCell', args: [2, 4, { id: 'sheet-1' }] },
        { name: 'clearCell', args: [3, 3, { id: 'sheet-1' }] },
        { name: 'clearCell', args: [3, 4, { id: 'sheet-1' }] },
      ],
    ]);
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
    expect(fixture.workbook.visibilityChanges).toEqual([
      { axis: 'column', hidden: true, indices: ['2', '3'] },
      { axis: 'column', hidden: false, indices: ['2', '3'] },
    ]);
    expect(fixture.workbook.sizeChanges).toEqual([
      {
        axis: 'row',
        custom: true,
        sheetId: 'sheet-1',
        sizes: { 3: 36, 4: 36 },
      },
      {
        axis: 'column',
        custom: true,
        sheetId: 'sheet-1',
        sizes: { 2: 128, 3: 128 },
      },
    ]);
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
    fixture.context.toolbarCell = { bl: 0, it: 1, un: 0 };
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
        key: 'Delete',
      }),
    ];

    expect(shortcuts.map((event) => editor.handleKeyDown(event))).toEqual([
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
    ]);
    expect(fixture.workbook.clearBatches).toHaveLength(1);
    expect(fixture.workbook.clearBatches[0]).toHaveLength(6);
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
  changes: WorkSpreadsheetContent[];
  calculation: RecordingSpreadsheetCalculation;
  context: SpreadsheetCommandContext;
  formulaBarValues: unknown[];
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
  const calculation = new RecordingSpreadsheetCalculation();
  const formulaBarValues: unknown[] = [];
  const workbook = new RecordingSpreadsheetWorkbook();
  return {
    calculation,
    changes,
    formulaBarValues,
    workbook,
    context: {
      activeSheetId: 'sheet-1',
      calculation,
      content,
      editable: true,
      fallbackRange: { row: [0, 1], column: [0, 2] },
      formulaBar: {
        setValue: (value) => formulaBarValues.push(value),
      },
      history: null,
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

function spreadsheetEditor(context: SpreadsheetCommandContext) {
  return createOfficeEditorRuntime<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >(context, createSpreadsheetEditorExtensions());
}

class RecordingSpreadsheetWorkbook implements SpreadsheetWorkbookCommandPort {
  cells: (Cell | null)[][] = [];
  clearBatches: Array<Array<{ name: string; args: unknown[] }>> = [];
  formats: Array<{
    attribute: string;
    range: SpreadsheetCommandRange;
    sheetId: string | undefined;
    value: unknown;
  }> = [];
  merges: Array<{
    range: SpreadsheetCommandRange;
    sheetId: string | undefined;
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

  cancelMerge(
    ranges: SpreadsheetCommandRange[],
    options?: { id?: string },
  ): void {
    this.merges.push(
      ...ranges.map((range) => ({ range, sheetId: options?.id })),
    );
  }

  batchCallApis(apiCalls: Array<{ name: string; args: unknown[] }>): void {
    this.clearBatches.push(apiCalls);
  }

  getSelection(): SpreadsheetCommandRange[] | undefined {
    return this.selection;
  }

  getCellsByRange(): (Cell | null)[][] {
    return this.cells;
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

  deleteRowOrColumn(
    axis: 'row' | 'column',
    start: number,
    end: number,
    options?: { id?: string },
  ): void {
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

  mergeCells(
    ranges: SpreadsheetCommandRange[],
    _type: string,
    options?: { id?: string },
  ): void {
    this.merges.push(
      ...ranges.map((range) => ({ range, sheetId: options?.id })),
    );
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
