import type { Cell, Selection } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import { spreadsheetSingleRange } from './spreadsheet-editor-support';

export interface SpreadsheetWorkbookCommandPort {
  calculateFormula: (sheetId?: string, range?: SpreadsheetCommandRange) => void;
  cancelMerge: (
    ranges: SpreadsheetCommandRange[],
    options?: { id?: string },
  ) => void;
  getSelection: () => Array<Pick<Selection, 'row' | 'column'>> | undefined;
  mergeCells: (
    ranges: SpreadsheetCommandRange[],
    type: string,
    options?: { id?: string },
  ) => void;
  setCellFormatByRange: (
    attribute: keyof Cell,
    value: unknown,
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ) => void;
}

export interface SpreadsheetCommandRange {
  row: number[];
  column: number[];
}

export interface SpreadsheetCommandSelection {
  sheetId: string;
  selection: Selection;
}

export type SpreadsheetEditorCommand =
  | {
      type: 'cell.format';
      attribute: keyof Cell;
      value: unknown;
    }
  | {
      type: 'cell.merge.toggle';
      merged: boolean;
    }
  | {
      type: 'formula.recalculate';
      scope: 'selection' | 'workbook';
    }
  | {
      type: 'sheet.gridLines.set';
      visible: boolean;
    }
  | {
      type: 'sheet.zoom.set';
      percent: number;
    };

export interface SpreadsheetCommandContext {
  activeSheetId: string;
  content: WorkSpreadsheetContent;
  fallbackRange: SpreadsheetCommandRange;
  onChange: (content: WorkSpreadsheetContent) => void;
  selection: SpreadsheetCommandSelection | null;
  targetSheetId: string;
  workbook: SpreadsheetWorkbookCommandPort | null;
}

export function executeSpreadsheetEditorCommand(
  context: SpreadsheetCommandContext,
  command: SpreadsheetEditorCommand,
): boolean {
  switch (command.type) {
    case 'cell.format':
      return formatCells(context, command.attribute, command.value);
    case 'cell.merge.toggle':
      return toggleCellMerge(context, command.merged);
    case 'formula.recalculate':
      return recalculateSpreadsheet(context, command.scope);
    case 'sheet.gridLines.set':
      return updateSpreadsheetSheet(context, command.type, command.visible);
    case 'sheet.zoom.set':
      return updateSpreadsheetSheet(context, command.type, command.percent);
  }
}

function formatCells(
  context: SpreadsheetCommandContext,
  attribute: keyof Cell,
  value: unknown,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  context.workbook.setCellFormatByRange(attribute, value, liveRange(context), {
    id: context.targetSheetId,
  });
  return true;
}

function toggleCellMerge(
  context: SpreadsheetCommandContext,
  merged: boolean,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const ranges = [liveRange(context)];
  if (merged) {
    context.workbook.cancelMerge(ranges, { id: context.targetSheetId });
  } else {
    context.workbook.mergeCells(ranges, 'merge-all', {
      id: context.targetSheetId,
    });
  }
  return true;
}

function recalculateSpreadsheet(
  context: SpreadsheetCommandContext,
  scope: 'selection' | 'workbook',
): boolean {
  if (!context.workbook) return false;
  if (scope === 'workbook') {
    context.workbook.calculateFormula();
    return true;
  }
  if (!context.selection) return false;
  context.workbook.calculateFormula(
    context.selection.sheetId,
    spreadsheetSingleRange(context.selection.selection),
  );
  return true;
}

function updateSpreadsheetSheet(
  context: SpreadsheetCommandContext,
  type: 'sheet.gridLines.set' | 'sheet.zoom.set',
  value: boolean | number,
): boolean {
  if (
    !context.activeSheetId ||
    !context.content.sheets.some((sheet) => sheet.id === context.activeSheetId)
  ) {
    return false;
  }
  const next: WorkSpreadsheetContent = {
    ...context.content,
    sheets: context.content.sheets.map((sheet) => {
      if (sheet.id !== context.activeSheetId) return sheet;
      return type === 'sheet.gridLines.set'
        ? { ...sheet, showGridLines: Boolean(value) }
        : {
            ...sheet,
            zoomRatio: Math.min(400, Math.max(20, Number(value))) / 100,
          };
    }),
  };
  context.onChange(next);
  return true;
}

function liveRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange {
  const live = context.workbook?.getSelection()?.at(-1);
  return spreadsheetSingleRange(live ?? context.fallbackRange);
}
