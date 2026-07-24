import { describe, expect, test } from '@rstest/core';
import {
  executeSpreadsheetEditorCommand,
  type SpreadsheetCommandContext,
  type SpreadsheetCommandRange,
  type SpreadsheetWorkbookCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet command controller', () => {
  test('routes cell formatting through the workbook command port', () => {
    const fixture = commandFixture();
    fixture.workbook.selection = [{ row: [4, 2], column: [3, 1] }];

    const handled = executeSpreadsheetEditorCommand(fixture.context, {
      type: 'cell.format',
      attribute: 'fs',
      value: 14,
    });

    expect(handled).toBe(true);
    expect(fixture.workbook.formats).toEqual([
      {
        attribute: 'fs',
        range: { row: [2, 4], column: [1, 3] },
        sheetId: 'sheet-1',
        value: 14,
      },
    ]);
  });

  test('uses explicit merge and recalculation commands', () => {
    const fixture = commandFixture();

    expect(
      executeSpreadsheetEditorCommand(fixture.context, {
        type: 'cell.merge.toggle',
        merged: false,
      }),
    ).toBe(true);
    expect(
      executeSpreadsheetEditorCommand(fixture.context, {
        type: 'formula.recalculate',
        scope: 'selection',
      }),
    ).toBe(true);

    expect(fixture.workbook.merges).toEqual([
      {
        range: { row: [0, 1], column: [0, 2] },
        sheetId: 'sheet-1',
      },
    ]);
    expect(fixture.workbook.calculations).toEqual([
      {
        range: { row: [2, 4], column: [1, 3] },
        sheetId: 'sheet-1',
      },
    ]);
  });

  test('updates controlled sheet view state without mutating the input', () => {
    const fixture = commandFixture();
    const previousSheet = fixture.context.content.sheets[0];

    expect(
      executeSpreadsheetEditorCommand(fixture.context, {
        type: 'sheet.gridLines.set',
        visible: false,
      }),
    ).toBe(true);
    expect(
      executeSpreadsheetEditorCommand(fixture.context, {
        type: 'sheet.zoom.set',
        percent: 175,
      }),
    ).toBe(true);

    expect(previousSheet.showGridLines).toBe(true);
    expect(fixture.changes[0].sheets[0].showGridLines).toBe(false);
    expect(fixture.changes[1].sheets[0].zoomRatio).toBe(1.75);
  });
});

function commandFixture(): {
  changes: WorkSpreadsheetContent[];
  context: SpreadsheetCommandContext;
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
  const workbook = new RecordingSpreadsheetWorkbook();
  return {
    changes,
    workbook,
    context: {
      activeSheetId: 'sheet-1',
      content,
      fallbackRange: { row: [0, 1], column: [0, 2] },
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

class RecordingSpreadsheetWorkbook implements SpreadsheetWorkbookCommandPort {
  calculations: Array<{
    sheetId: string | undefined;
    range: SpreadsheetCommandRange | undefined;
  }> = [];
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
  selection: SpreadsheetCommandRange[] | undefined;

  calculateFormula(sheetId?: string, range?: SpreadsheetCommandRange): void {
    this.calculations.push({ sheetId, range });
  }

  cancelMerge(
    ranges: SpreadsheetCommandRange[],
    options?: { id?: string },
  ): void {
    this.merges.push(
      ...ranges.map((range) => ({ range, sheetId: options?.id })),
    );
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
