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
import {
  activateSpreadsheetSheet,
  addSpreadsheetSheet,
  adjacentSpreadsheetSheetId,
  deleteSpreadsheetSheet,
  duplicateSpreadsheetSheet,
  hideSpreadsheetSheet,
  moveSpreadsheetSheet,
  renameSpreadsheetSheet,
  setSpreadsheetSheetColor,
  type SpreadsheetSheetMoveDirection,
} from './spreadsheet-sheet-model';

export interface SpreadsheetWorkbookCommandPort {
  batchCallApis: (apiCalls: Array<{ name: string; args: unknown[] }>) => void;
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
  setCellValuesByRange: (
    values: unknown[][],
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ) => void;
  setCellFormatByRange: (
    attribute: keyof Cell,
    value: unknown,
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ) => void;
  setSelection: (
    range: SpreadsheetCommandRange[],
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

export interface SpreadsheetFormulaBarCommandPort {
  setValue: (value: unknown) => void;
}

export interface SpreadsheetEditorCommands {
  activateSheet: (sheetId: string) => boolean;
  addSheet: () => boolean;
  clearSelectedCells: () => boolean;
  deleteSheet: (sheetId: string) => boolean;
  duplicateSheet: (sheetId: string) => boolean;
  hideSheet: (sheetId: string) => boolean;
  moveSheet: (
    sheetId: string,
    direction: SpreadsheetSheetMoveDirection,
  ) => boolean;
  pasteCells: (values: readonly (readonly unknown[])[]) => boolean;
  recalculateFormula: (scope: 'selection' | 'workbook') => boolean;
  renameSheet: (sheetId: string, name: string) => boolean;
  redo: () => boolean;
  setCellFormat: (attribute: keyof Cell, value: unknown) => boolean;
  setGridLines: (visible: boolean) => boolean;
  setSheetColor: (sheetId: string, color: string | null) => boolean;
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
  formulaBar: SpreadsheetFormulaBarCommandPort | null;
  history: SpreadsheetHistoryCommandPort | null;
  onChange: (content: WorkSpreadsheetContent) => void;
  selection: SpreadsheetCommandSelection | null;
  targetSheetId: string;
  toolbarCell: Cell | null;
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
      name: 'spreadsheetSheets',
      addCommands: () => ({
        activateSheet: {
          canExecute: (context, sheetId) =>
            context.editable &&
            context.content.sheets.some((sheet) => sheet.id === sheetId),
          execute: (context, sheetId) =>
            applySpreadsheetSheetChange(
              context,
              activateSpreadsheetSheet(context.content, sheetId),
            ),
        },
        addSheet: {
          canExecute: (context) => context.editable,
          execute: (context) =>
            applySpreadsheetSheetChange(
              context,
              addSpreadsheetSheet(context.content),
            ),
        },
        deleteSheet: {
          canExecute: (context, sheetId) =>
            Boolean(
              context.editable &&
                context.content.sheets.length > 1 &&
                context.content.sheets.some((sheet) => sheet.id === sheetId),
            ),
          execute: (context, sheetId) =>
            applySpreadsheetSheetChange(
              context,
              deleteSpreadsheetSheet(context.content, sheetId),
            ),
        },
        duplicateSheet: {
          canExecute: (context, sheetId) =>
            context.editable &&
            context.content.sheets.some((sheet) => sheet.id === sheetId),
          execute: (context, sheetId) =>
            applySpreadsheetSheetChange(
              context,
              duplicateSpreadsheetSheet(context.content, sheetId),
            ),
        },
        hideSheet: {
          canExecute: (context, sheetId) =>
            context.editable &&
            context.content.sheets.filter((sheet) => sheet.hide !== 1).length >
              1 &&
            context.content.sheets.some(
              (sheet) => sheet.id === sheetId && sheet.hide !== 1,
            ),
          execute: (context, sheetId) =>
            applySpreadsheetSheetChange(
              context,
              hideSpreadsheetSheet(context.content, sheetId, true),
            ),
        },
        moveSheet: {
          canExecute: (context, sheetId, direction) =>
            Boolean(
              context.editable &&
                moveSpreadsheetSheet(context.content, sheetId, direction),
            ),
          execute: (context, sheetId, direction) =>
            applySpreadsheetSheetChange(
              context,
              moveSpreadsheetSheet(context.content, sheetId, direction),
            ),
        },
        renameSheet: {
          canExecute: (context, sheetId, name) =>
            Boolean(
              context.editable &&
                renameSpreadsheetSheet(context.content, sheetId, name),
            ),
          execute: (context, sheetId, name) =>
            applySpreadsheetSheetChange(
              context,
              renameSpreadsheetSheet(context.content, sheetId, name),
            ),
        },
        setSheetColor: {
          canExecute: (context, sheetId, color) =>
            Boolean(
              context.editable &&
                setSpreadsheetSheetColor(context.content, sheetId, color),
            ),
          execute: (context, sheetId, color) =>
            applySpreadsheetSheetChange(
              context,
              setSpreadsheetSheetColor(context.content, sheetId, color),
            ),
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
        'Mod-b': ({ can, commands, context }, event) =>
          runSpreadsheetCellFormatShortcut(
            event,
            context,
            can.setCellFormat,
            commands.setCellFormat,
            'bl',
          ),
        'Mod-i': ({ can, commands, context }, event) =>
          runSpreadsheetCellFormatShortcut(
            event,
            context,
            can.setCellFormat,
            commands.setCellFormat,
            'it',
          ),
        'Mod-u': ({ can, commands, context }, event) =>
          runSpreadsheetCellFormatShortcut(
            event,
            context,
            can.setCellFormat,
            commands.setCellFormat,
            'un',
          ),
        'Mod-z': ({ can, commands }, event) =>
          runSpreadsheetHistoryShortcut(event, can.undo, commands.undo),
        'Mod-Shift-z': ({ can, commands }, event) =>
          runSpreadsheetHistoryShortcut(event, can.redo, commands.redo),
        'Mod-y': ({ can, commands }, event) =>
          runSpreadsheetHistoryShortcut(event, can.redo, commands.redo),
        'Control-PageUp': ({ context, commands }, event) =>
          runSpreadsheetSheetNavigationShortcut(
            event,
            context,
            commands.activateSheet,
            -1,
          ),
        'Control-PageDown': ({ context, commands }, event) =>
          runSpreadsheetSheetNavigationShortcut(
            event,
            context,
            commands.activateSheet,
            1,
          ),
        Delete: ({ can, commands }, event) =>
          runSpreadsheetClearShortcut(
            event,
            can.clearSelectedCells,
            commands.clearSelectedCells,
          ),
        Backspace: ({ can, commands }, event) =>
          runSpreadsheetClearShortcut(
            event,
            can.clearSelectedCells,
            commands.clearSelectedCells,
          ),
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetCellFormatting',
      addCommands: () => ({
        clearSelectedCells: {
          canExecute: canEditSelectedCells,
          execute: clearSelectedCells,
        },
        pasteCells: {
          canExecute: (context, values) =>
            canEditSelectedCells(context) && isRectangularCellMatrix(values),
          execute: pasteCells,
        },
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

function clearSelectedCells(context: SpreadsheetCommandContext): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const range = liveRange(context);
  const calls: Array<{ name: string; args: unknown[] }> = [];
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      calls.push({
        name: 'clearCell',
        args: [row, column, { id: context.targetSheetId }],
      });
    }
  }
  try {
    context.workbook.batchCallApis(calls);
  } catch {
    return false;
  }
  syncSpreadsheetFormulaBar(context, '');
  return true;
}

function pasteCells(
  context: SpreadsheetCommandContext,
  values: readonly (readonly unknown[])[],
): boolean {
  if (
    !context.workbook ||
    !context.targetSheetId ||
    !isRectangularCellMatrix(values)
  ) {
    return false;
  }
  const source = liveRange(context);
  const data = values.map((row) => [...row]);
  const range = {
    row: [source.row[0], source.row[0] + data.length - 1],
    column: [source.column[0], source.column[0] + data[0].length - 1],
  };
  try {
    context.workbook.setCellValuesByRange(data, range, {
      id: context.targetSheetId,
    });
  } catch {
    return false;
  }
  try {
    context.workbook.setSelection([range], { id: context.targetSheetId });
  } catch {
    // The values were committed; selection highlighting is best effort.
  }
  syncSpreadsheetFormulaBar(context, data[0][0]);
  return true;
}

function isRectangularCellMatrix(
  values: readonly (readonly unknown[])[],
): boolean {
  const width = values[0]?.length ?? 0;
  return Boolean(width && values.every((row) => row.length === width));
}

function syncSpreadsheetFormulaBar(
  context: SpreadsheetCommandContext,
  value: unknown,
): void {
  try {
    context.formulaBar?.setValue(value);
  } catch {
    // Cell mutations remain valid if the vendor formula bar cannot refresh.
  }
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

function applySpreadsheetSheetChange(
  context: SpreadsheetCommandContext,
  next: WorkSpreadsheetContent | null,
): boolean {
  if (!context.editable || !next) return false;
  context.onChange(next);
  return true;
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

function runSpreadsheetCellFormatShortcut(
  event: KeyboardEvent,
  context: SpreadsheetCommandContext,
  canExecute: SpreadsheetEditorCanCommands['setCellFormat'],
  execute: SpreadsheetEditorCommands['setCellFormat'],
  attribute: 'bl' | 'it' | 'un',
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target)
  ) {
    return false;
  }
  const value = Number(context.toolbarCell?.[attribute]) === 1 ? 0 : 1;
  return canExecute(attribute, value) && execute(attribute, value);
}

function runSpreadsheetClearShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['clearSelectedCells'],
  execute: SpreadsheetEditorCommands['clearSelectedCells'],
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

function runSpreadsheetSheetNavigationShortcut(
  event: KeyboardEvent,
  context: SpreadsheetCommandContext,
  activateSheet: SpreadsheetEditorCommands['activateSheet'],
  direction: SpreadsheetSheetMoveDirection,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target)
  ) {
    return false;
  }
  const target = adjacentSpreadsheetSheetId(
    context.content,
    context.activeSheetId,
    direction,
  );
  return Boolean(
    target && target !== context.activeSheetId && activateSheet(target),
  );
}

function liveRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange {
  const live = context.workbook?.getSelection()?.at(-1);
  return spreadsheetSingleRange(live ?? context.fallbackRange);
}
