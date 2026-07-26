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
  }

  setSelection(
    range: SpreadsheetCommandRange[],
    options?: { id?: string },
  ): void {
    this.selections.push({ range, sheetId: options?.id });
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
