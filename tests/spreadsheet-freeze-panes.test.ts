import type { Selection } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  spreadsheetFreezePaneForSelection,
  spreadsheetFreezePanesSelectionLabel,
  updateSpreadsheetFreezePanes,
} from '../src/internal/features/work/editors/spreadsheet-freeze-panes';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import { createWorkArtifact } from '../src/internal/features/work/work-templates';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../src/internal/features/work/work-types';

describe('spreadsheet Freeze Panes model', () => {
  test('freezes rows above and columns left of the WPS current cell', () => {
    expect(spreadsheetFreezePaneForSelection(selection(2, 1))).toEqual({
      type: 'rangeBoth',
      range: { row_focus: 1, column_focus: 0 },
    });
    expect(spreadsheetFreezePaneForSelection(selection(3, 0))).toEqual({
      type: 'rangeRow',
      range: { row_focus: 2, column_focus: 0 },
    });
    expect(spreadsheetFreezePaneForSelection(selection(0, 2))).toEqual({
      type: 'rangeColumn',
      range: { row_focus: 0, column_focus: 1 },
    });
    expect(spreadsheetFreezePaneForSelection(selection(0, 0))).toBeNull();
  });

  test('describes the custom freeze boundary in worksheet language', () => {
    expect(spreadsheetFreezePanesSelectionLabel(selection(2, 1))).toBe(
      '冻结至第 2 行、A 列',
    );
    expect(spreadsheetFreezePanesSelectionLabel(selection(3, 0))).toBe(
      '冻结至第 3 行',
    );
    expect(spreadsheetFreezePanesSelectionLabel(selection(0, 2))).toBe(
      '冻结至 B 列',
    );
    expect(spreadsheetFreezePanesSelectionLabel(selection(0, 0))).toBe(
      '冻结至当前单元格',
    );
  });

  test('updates custom, preset, and unfreeze state without mutating input', () => {
    const content = workbook(sheet());
    const custom = updateSpreadsheetFreezePanes(
      content,
      'sheet-1',
      'selection',
      selection(2, 1),
    );

    expect(content.sheets[0]?.frozen).toBeUndefined();
    expect(custom).not.toBe(content);
    expect(custom?.sheets[0]?.frozen).toEqual({
      type: 'rangeBoth',
      range: { row_focus: 1, column_focus: 0 },
    });
    expect(custom?.sheets[0]?.luckysheet_select_save).toEqual([
      selection(2, 1),
    ]);
    expect(
      custom &&
        updateSpreadsheetFreezePanes(
          custom,
          'sheet-1',
          'selection',
          selection(2, 1),
        ),
    ).toBeNull();

    const topRow = updateSpreadsheetFreezePanes(
      content,
      'sheet-1',
      'topRow',
      selection(6, 4),
    );
    expect(topRow?.sheets[0]?.frozen).toEqual({
      type: 'rangeRow',
      range: { row_focus: 0, column_focus: 0 },
    });
    const firstColumn = updateSpreadsheetFreezePanes(
      topRow ?? content,
      'sheet-1',
      'firstColumn',
      selection(6, 4),
    );
    expect(firstColumn?.sheets[0]?.frozen).toEqual({
      type: 'rangeColumn',
      range: { row_focus: 0, column_focus: 0 },
    });
    const unfrozen = updateSpreadsheetFreezePanes(
      firstColumn ?? content,
      'sheet-1',
      'none',
      selection(6, 4),
    );
    expect(unfrozen?.sheets[0]?.frozen).toBeUndefined();
    expect(
      updateSpreadsheetFreezePanes(content, 'sheet-1', 'none', selection(6, 4)),
    ).toBeNull();
  });

  test('rejects a missing sheet and an A1 custom freeze', () => {
    const content = workbook(sheet());
    expect(
      updateSpreadsheetFreezePanes(
        content,
        'missing',
        'topRow',
        selection(2, 1),
      ),
    ).toBeNull();
    expect(
      updateSpreadsheetFreezePanes(
        content,
        'sheet-1',
        'selection',
        selection(0, 0),
      ),
    ).toBeNull();
  });

  test('preserves the selected WPS freeze boundary through XLSX export', async () => {
    const artifact = createWorkArtifact('quarterly-plan');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected the quarterly plan spreadsheet.');
    }
    const content = updateSpreadsheetFreezePanes(
      artifact.content,
      artifact.content.sheets[0]?.id ?? '',
      'selection',
      selection(2, 1),
    );
    if (!content) throw new Error('Expected a frozen workbook.');

    const blob = await createWorkArtifactBlob({ ...artifact, content });
    const imported = await importWorkFile(
      new File([blob], 'quarterly-plan-frozen.xlsx', { type: blob.type }),
    );
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected an imported spreadsheet.');
    }

    expect(imported.content.sheets[0]?.frozen).toEqual({
      type: 'rangeBoth',
      range: { row_focus: 1, column_focus: 0 },
    });
  });
});

function selection(row: number, column: number): Selection {
  return {
    row: [row, row],
    column: [column, column],
    row_focus: row,
    column_focus: column,
  };
}

function workbook(activeSheet: WorkSpreadsheetSheet): WorkSpreadsheetContent {
  return { type: 'spreadsheet', sheets: [activeSheet] };
}

function sheet(): WorkSpreadsheetSheet {
  return {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 40,
    column: 12,
    data: [[{ v: 'Header' }]],
  };
}
