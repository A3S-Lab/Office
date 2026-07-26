import type { Cell, Selection } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  createOfficeEditorExtension,
  type OfficeEditorCanCommands,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  isSpreadsheetNativeTextUndoTarget,
  selectSpreadsheetFormulaBarContents,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';

export interface SpreadsheetWorkbookCommandPort {
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

export type SpreadsheetCalculationCommand =
  | { scope: 'workbook' }
  | {
      scope: 'selection';
      sheetId: string;
      range: SpreadsheetCommandRange;
    };

export interface SpreadsheetCalculationCommandPort {
  recalculate: (request: SpreadsheetCalculationCommand) => void;
}

export interface SpreadsheetHistoryCommandPort {
  canRedo: boolean;
  canUndo: boolean;
  redo: () => boolean;
  undo: () => boolean;
}

export interface SpreadsheetEditorCommands {
  recalculateFormula: (scope: 'selection' | 'workbook') => boolean;
  redo: () => boolean;
  setCellFormat: (attribute: keyof Cell, value: unknown) => boolean;
  setGridLines: (visible: boolean) => boolean;
  setSpreadsheetContent: (content: WorkSpreadsheetContent) => boolean;
  setZoom: (percent: number) => boolean;
  toggleCellMerge: (merged: boolean) => boolean;
  undo: () => boolean;
}

export type SpreadsheetEditorCanCommands =
  OfficeEditorCanCommands<SpreadsheetEditorCommands>;

export interface SpreadsheetCommandContext {
  activeSheetId: string;
  calculation: SpreadsheetCalculationCommandPort | null;
  content: WorkSpreadsheetContent;
  editable: boolean;
  fallbackRange: SpreadsheetCommandRange;
  history: SpreadsheetHistoryCommandPort | null;
  onChange: (content: WorkSpreadsheetContent) => void;
  selection: SpreadsheetCommandSelection | null;
  targetSheetId: string;
  workbook: SpreadsheetWorkbookCommandPort | null;
}

export function createSpreadsheetEditorExtensions(): readonly OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
>[] {
  return [
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetDocument',
      addCommands: () => ({
        setSpreadsheetContent: {
          canExecute: (context) => context.editable,
          execute: (context, content) => {
            if (!context.editable) return false;
            context.onChange(content);
            return true;
          },
        },
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetHistory',
      addCommands: () => ({
        redo: {
          canExecute: (context) =>
            context.editable && (context.history?.canRedo ?? false),
          execute: (context) => context.history?.redo() ?? false,
        },
        undo: {
          canExecute: (context) =>
            context.editable && (context.history?.canUndo ?? false),
          execute: (context) => context.history?.undo() ?? false,
        },
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetKeyboardShortcuts',
      addKeyboardShortcuts: () => ({
        'Mod-a': (_props, event) => selectSpreadsheetFormulaBarContents(event),
        'Mod-z': ({ can, commands }, event) =>
          runSpreadsheetHistoryShortcut(event, can.undo, commands.undo),
        'Mod-Shift-z': ({ can, commands }, event) =>
          runSpreadsheetHistoryShortcut(event, can.redo, commands.redo),
        'Mod-y': ({ can, commands }, event) =>
          runSpreadsheetHistoryShortcut(event, can.redo, commands.redo),
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetCellFormatting',
      addCommands: () => ({
        setCellFormat: {
          canExecute: canEditSelectedCells,
          execute: formatCells,
        },
        toggleCellMerge: {
          canExecute: canEditSelectedCells,
          execute: toggleCellMerge,
        },
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetCalculation',
      addCommands: () => ({
        recalculateFormula: {
          canExecute: (context, scope) =>
            Boolean(
              context.editable &&
                context.calculation &&
                (scope === 'workbook' || context.selection),
            ),
          execute: recalculateSpreadsheet,
        },
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetView',
      addCommands: () => ({
        setGridLines: {
          canExecute: hasActiveSheet,
          execute: (context, visible) =>
            updateSpreadsheetSheet(context, 'gridLines', visible),
        },
        setZoom: {
          canExecute: hasActiveSheet,
          execute: (context, percent) =>
            updateSpreadsheetSheet(context, 'zoom', percent),
        },
      }),
    }),
  ];
}

function formatCells(
  context: SpreadsheetCommandContext,
  attribute: keyof Cell,
  value: unknown,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  if (attribute === 'ct' && !isSpreadsheetCellTypeFormat(value)) return false;
  try {
    context.workbook.setCellFormatByRange(
      attribute,
      value,
      liveRange(context),
      {
        id: context.targetSheetId,
      },
    );
    return true;
  } catch {
    return false;
  }
}

function isSpreadsheetCellTypeFormat(
  value: unknown,
): value is NonNullable<Cell['ct']> & { fa: string; t: string } {
  if (!value || typeof value !== 'object') return false;
  const format = value as Cell['ct'];
  return Boolean(
    typeof format?.fa === 'string' &&
      format.fa.trim() &&
      typeof format.t === 'string' &&
      format.t.trim(),
  );
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
  if (!context.calculation) return false;
  if (scope === 'workbook') {
    context.calculation.recalculate({ scope: 'workbook' });
    return true;
  }
  if (!context.selection) return false;
  context.calculation.recalculate({
    scope: 'selection',
    sheetId: context.selection.sheetId,
    range: spreadsheetSingleRange(context.selection.selection),
  });
  return true;
}

function updateSpreadsheetSheet(
  context: SpreadsheetCommandContext,
  property: 'gridLines' | 'zoom',
  value: boolean | number,
): boolean {
  if (!hasActiveSheet(context)) return false;
  const next: WorkSpreadsheetContent = {
    ...context.content,
    sheets: context.content.sheets.map((sheet) => {
      if (sheet.id !== context.activeSheetId) return sheet;
      return property === 'gridLines'
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

function canEditSelectedCells(context: SpreadsheetCommandContext): boolean {
  return Boolean(context.editable && context.workbook && context.targetSheetId);
}

function hasActiveSheet(context: SpreadsheetCommandContext): boolean {
  return Boolean(
    context.editable &&
      context.activeSheetId &&
      context.content.sheets.some(
        (sheet) => sheet.id === context.activeSheetId,
      ),
  );
}

function runSpreadsheetHistoryShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

function liveRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange {
  const live = context.workbook?.getSelection()?.at(-1);
  return spreadsheetSingleRange(live ?? context.fallbackRange);
}
