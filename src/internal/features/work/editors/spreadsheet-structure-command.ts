import { spreadsheetGridSize } from '../spreadsheet-sparse';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import {
  canEditSpreadsheetSelection,
  spreadsheetLiveCommandRange,
} from './spreadsheet-command-selection';
import type {
  SpreadsheetCommandContext,
  SpreadsheetCommandRange,
  SpreadsheetEditorCommands,
  SpreadsheetStructureAxis,
  SpreadsheetStructureInsertPosition,
} from './spreadsheet-command-controller';
import { spreadsheetSelectionAfterStructureDeletion } from './spreadsheet-structure-selection';
import { canApplySpreadsheetTableStructureChange } from './spreadsheet-table-reconciliation';

export const MAX_SPREADSHEET_VISIBILITY_ROWS = 10_000;
export const MAX_SPREADSHEET_VISIBILITY_COLUMNS = 1_000;

export function createSpreadsheetStructureExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetStructure',
    addCommands: () => ({
      deleteSelectedStructure: {
        canExecute: canDeleteSelectedStructure,
        execute: deleteSelectedStructure,
      },
      insertSelectedStructure: {
        canExecute: canInsertSelectedStructure,
        execute: insertSelectedStructure,
      },
      setSelectedStructureHidden: {
        canExecute: canSetSelectedStructureHidden,
        execute: setSelectedStructureHidden,
      },
      setSelectedStructureSize: {
        canExecute: canSetSelectedStructureSize,
        execute: setSelectedStructureSize,
      },
    }),
  });
}

function canInsertSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  position: SpreadsheetStructureInsertPosition,
): boolean {
  if (!canEditSpreadsheetSelection(context)) return false;
  const range = spreadsheetLiveCommandRange(context);
  const [start, end] = spreadsheetStructureRange(range, axis);
  const count = end - start + 1;
  const maximum = axis === 'row' ? 10_000 : 1_000;
  return (
    spreadsheetStructureExtent(context, axis) + count < maximum &&
    canApplySpreadsheetTableStructureChange(
      context.content.sheets.find(
        (sheet) => sheet.id === context.targetSheetId,
      ),
      {
        axis,
        count,
        direction: position === 'before' ? 'lefttop' : 'rightbottom',
        index: position === 'before' ? start : end,
        kind: 'insert',
      },
    )
  );
}

function insertSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  position: SpreadsheetStructureInsertPosition,
): boolean {
  if (!canInsertSelectedStructure(context, axis, position)) return false;
  if (!context.workbook || !context.targetSheetId) return false;
  const [start, end] = spreadsheetStructureRange(
    spreadsheetLiveCommandRange(context),
    axis,
  );
  const before = position === 'before';
  try {
    context.workbook.insertRowOrColumn(
      axis,
      before ? start : end,
      end - start + 1,
      before ? 'lefttop' : 'rightbottom',
      { id: context.targetSheetId },
    );
    return true;
  } catch {
    return false;
  }
}

function canDeleteSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
): boolean {
  if (!canEditSpreadsheetSelection(context)) return false;
  const [start, end] = spreadsheetStructureRange(
    spreadsheetLiveCommandRange(context),
    axis,
  );
  return (
    end - start + 1 < spreadsheetStructureExtent(context, axis) &&
    canApplySpreadsheetTableStructureChange(
      context.content.sheets.find(
        (sheet) => sheet.id === context.targetSheetId,
      ),
      { axis, end, kind: 'delete', start },
    )
  );
}

function deleteSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
): boolean {
  if (!canDeleteSelectedStructure(context, axis)) return false;
  if (!context.workbook || !context.targetSheetId) return false;
  const range = spreadsheetLiveCommandRange(context);
  const [start, end] = spreadsheetStructureRange(range, axis);
  const selection = spreadsheetSelectionAfterStructureDeletion(
    range,
    axis,
    spreadsheetStructureExtent(context, axis),
    start,
    end,
  );
  try {
    context.workbook.batchCallApis([
      {
        name: 'deleteRowOrColumn',
        args: [axis, start, end, { id: context.targetSheetId }],
      },
      {
        name: 'setSelection',
        args: [[selection], { id: context.targetSheetId }],
      },
    ]);
    return true;
  } catch {
    return false;
  }
}

function canSetSelectedStructureHidden(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  hidden: boolean,
): boolean {
  if (!canEditSpreadsheetSelection(context)) return false;
  const [start, end] = spreadsheetStructureRange(
    spreadsheetLiveCommandRange(context),
    axis,
  );
  const count = end - start + 1;
  const maximum =
    axis === 'row'
      ? MAX_SPREADSHEET_VISIBILITY_ROWS
      : MAX_SPREADSHEET_VISIBILITY_COLUMNS;
  return (
    count <= maximum &&
    (!hidden || count < spreadsheetStructureExtent(context, axis))
  );
}

function setSelectedStructureHidden(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  hidden: boolean,
): boolean {
  if (
    !context.workbook ||
    !canSetSelectedStructureHidden(context, axis, hidden)
  ) {
    return false;
  }
  const [start, end] = spreadsheetStructureRange(
    spreadsheetLiveCommandRange(context),
    axis,
  );
  const indices = Array.from({ length: end - start + 1 }, (_, offset) =>
    String(start + offset),
  );
  try {
    if (hidden) context.workbook.hideRowOrColumn(indices, axis);
    else context.workbook.showRowOrColumn(indices, axis);
    return true;
  } catch {
    return false;
  }
}

function canSetSelectedStructureSize(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  size: number,
): boolean {
  const maximum = axis === 'row' ? 545 : 2_038;
  return Boolean(
    canEditSpreadsheetSelection(context) &&
      Number.isFinite(size) &&
      size >= 1 &&
      size <= maximum,
  );
}

function setSelectedStructureSize(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  size: number,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const [start, end] = spreadsheetStructureRange(
    spreadsheetLiveCommandRange(context),
    axis,
  );
  const sizes = Object.fromEntries(
    Array.from({ length: end - start + 1 }, (_, offset) => [
      String(start + offset),
      size,
    ]),
  );
  try {
    if (axis === 'row') {
      context.workbook.setRowHeight(sizes, { id: context.targetSheetId }, true);
    } else {
      context.workbook.setColumnWidth(
        sizes,
        { id: context.targetSheetId },
        true,
      );
    }
    return true;
  } catch {
    return false;
  }
}

function spreadsheetStructureRange(
  range: SpreadsheetCommandRange,
  axis: SpreadsheetStructureAxis,
): [number, number] {
  const values = axis === 'row' ? range.row : range.column;
  return [
    Math.min(values[0] ?? 0, values[1] ?? 0),
    Math.max(values[0] ?? 0, values[1] ?? 0),
  ];
}

function spreadsheetStructureExtent(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
): number {
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  const range = spreadsheetStructureRange(
    spreadsheetLiveCommandRange(context),
    axis,
  );
  if (!sheet) return range[1] + 1;
  const size = context.targetSheetGridSize ?? spreadsheetGridSize(sheet);
  const extent = axis === 'row' ? size?.rowCount : size?.columnCount;
  return Math.max(extent ?? 0, range[1] + 1);
}
