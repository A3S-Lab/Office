import type { Cell, Selection } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  createOfficeEditorExtension,
  type OfficeEditorCanCommands,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import { runSpreadsheetClipboardShortcut } from './spreadsheet-clipboard-shortcuts';
import {
  isSpreadsheetNativeTextUndoTarget,
  selectSpreadsheetFormulaBarContents,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  isSpreadsheetGridKeyboardTarget,
  moveSpreadsheetKeyboardSelection,
  runSpreadsheetSelectionMoveShortcut,
  runSpreadsheetSelectionScopeShortcut,
  type SpreadsheetKeyboardSelection,
  type SpreadsheetSelectionMove,
  type SpreadsheetSelectionScope,
  scopeSpreadsheetKeyboardSelection,
  spreadsheetSelectionContainsFocus,
} from './spreadsheet-keyboard-navigation';
import {
  activateSpreadsheetSheet,
  addSpreadsheetSheet,
  adjacentSpreadsheetSheetId,
  deleteSpreadsheetSheet,
  duplicateSpreadsheetSheet,
  hideSpreadsheetSheet,
  moveSpreadsheetSheet,
  renameSpreadsheetSheet,
  type SpreadsheetSheetMoveDirection,
  setSpreadsheetSheetColor,
} from './spreadsheet-sheet-model';

export interface SpreadsheetWorkbookCommandPort {
  batchCallApis: (apiCalls: Array<{ name: string; args: unknown[] }>) => void;
  cancelMerge: (
    ranges: SpreadsheetCommandRange[],
    options?: { id?: string },
  ) => void;
  getSelection: () => SpreadsheetCommandRange[] | undefined;
  getCellsByRange: (
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ) => (Cell | null)[][];
  insertRowOrColumn: (
    type: SpreadsheetStructureAxis,
    index: number,
    count: number,
    direction: 'lefttop' | 'rightbottom',
    options?: { id?: string },
  ) => void;
  deleteRowOrColumn: (
    type: SpreadsheetStructureAxis,
    start: number,
    end: number,
    options?: { id?: string },
  ) => void;
  hideRowOrColumn: (
    rowOrColumnInfo: string[],
    type: SpreadsheetStructureAxis,
  ) => void;
  showRowOrColumn: (
    rowOrColumnInfo: string[],
    type: SpreadsheetStructureAxis,
  ) => void;
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
  setRowHeight: (
    rowInfo: Record<string, number>,
    options?: { id?: string },
    custom?: boolean,
  ) => void;
  setColumnWidth: (
    columnInfo: Record<string, number>,
    options?: { id?: string },
    custom?: boolean,
  ) => void;
  setSelection: (
    range: SpreadsheetCommandRange[],
    options?: { id?: string },
  ) => void;
}

export interface SpreadsheetCommandRange extends SpreadsheetKeyboardSelection {
  row: number[];
  column: number[];
}

export interface SpreadsheetCommandSelection {
  sheetId: string;
  selection: Selection;
}

export type SpreadsheetStructureAxis = 'row' | 'column';
export type SpreadsheetStructureInsertPosition = 'before' | 'after';
export type SpreadsheetSortDirection = 'ascending' | 'descending';

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

export interface SpreadsheetClipboardCommandPort {
  canCopySelection: boolean;
  canCutSelection: boolean;
  canPasteSelection: boolean;
  copySelection: () => boolean;
  cutSelection: () => boolean;
  pasteSelection: () => boolean;
}

export interface SpreadsheetFormulaBarCommandPort {
  setValue: (value: unknown) => void;
}

export interface SpreadsheetViewCommandPort {
  activateSheet: (sheetId: string) => boolean;
}

export interface SpreadsheetEditorCommands {
  activateSheet: (sheetId: string) => boolean;
  addSheet: () => boolean;
  clearSelectedCells: () => boolean;
  copySelection: () => boolean;
  cutSelection: () => boolean;
  deleteSelectedStructure: (axis: SpreadsheetStructureAxis) => boolean;
  deleteSheet: (sheetId: string) => boolean;
  duplicateSheet: (sheetId: string) => boolean;
  hideSheet: (sheetId: string) => boolean;
  insertSelectedStructure: (
    axis: SpreadsheetStructureAxis,
    position: SpreadsheetStructureInsertPosition,
  ) => boolean;
  moveSheet: (
    sheetId: string,
    direction: SpreadsheetSheetMoveDirection,
  ) => boolean;
  moveSelection: (move: SpreadsheetSelectionMove, extend: boolean) => boolean;
  pasteCells: (values: readonly (readonly unknown[])[]) => boolean;
  pasteSelection: () => boolean;
  recalculateFormula: (scope: 'selection' | 'workbook') => boolean;
  renameSheet: (sheetId: string, name: string) => boolean;
  redo: () => boolean;
  setCellFormat: (attribute: keyof Cell, value: unknown) => boolean;
  setGridLines: (visible: boolean) => boolean;
  selectCellRange: (scope: SpreadsheetSelectionScope) => boolean;
  setSheetColor: (sheetId: string, color: string | null) => boolean;
  setSelectedStructureHidden: (
    axis: SpreadsheetStructureAxis,
    hidden: boolean,
  ) => boolean;
  setSelectedStructureSize: (
    axis: SpreadsheetStructureAxis,
    size: number,
  ) => boolean;
  setSpreadsheetContent: (content: WorkSpreadsheetContent) => boolean;
  setZoom: (percent: number) => boolean;
  sortSelectedCells: (direction: SpreadsheetSortDirection) => boolean;
  toggleCellMerge: (merged: boolean) => boolean;
  undo: () => boolean;
}

export type SpreadsheetEditorCanCommands =
  OfficeEditorCanCommands<SpreadsheetEditorCommands>;

export interface SpreadsheetCommandContext {
  activeSheetId: string;
  calculation: SpreadsheetCalculationCommandPort | null;
  clipboard: SpreadsheetClipboardCommandPort;
  content: WorkSpreadsheetContent;
  editable: boolean;
  fallbackRange: SpreadsheetCommandRange;
  formulaBar: SpreadsheetFormulaBarCommandPort | null;
  history: SpreadsheetHistoryCommandPort | null;
  onChange: (content: WorkSpreadsheetContent) => void;
  selection: SpreadsheetCommandSelection | null;
  targetSheetId: string;
  toolbarCell: Cell | null;
  view: SpreadsheetViewCommandPort | null;
  workbook: SpreadsheetWorkbookCommandPort | null;
}

interface SpreadsheetSelectionNavigationStorage {
  focus: { column: number; row: number } | null;
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
      name: 'spreadsheetClipboard',
      addCommands: () => ({
        copySelection: {
          canExecute: ({ clipboard }) => clipboard.canCopySelection,
          execute: ({ clipboard }) =>
            clipboard.canCopySelection && clipboard.copySelection(),
        },
        cutSelection: {
          canExecute: ({ clipboard }) => clipboard.canCutSelection,
          execute: ({ clipboard }) =>
            clipboard.canCutSelection && clipboard.cutSelection(),
        },
        pasteSelection: {
          canExecute: ({ clipboard }) => clipboard.canPasteSelection,
          execute: ({ clipboard }) =>
            clipboard.canPasteSelection && clipboard.pasteSelection(),
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
          canExecute: (context, sheetId) => {
            const sheet = context.content.sheets.find(
              (candidate) => candidate.id === sheetId,
            );
            return Boolean(
              sheet && (context.editable || (context.view && sheet.hide !== 1)),
            );
          },
          execute: (context, sheetId) => {
            if (!context.editable) {
              const sheet = context.content.sheets.find(
                (candidate) => candidate.id === sheetId,
              );
              if (!sheet || sheet.hide === 1) return false;
              return context.view?.activateSheet(sheetId) ?? false;
            }
            return applySpreadsheetSheetChange(
              context,
              activateSpreadsheetSheet(context.content, sheetId),
            );
          },
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
        'Mod-a': ({ can, commands }, event) =>
          selectSpreadsheetFormulaBarContents(event) ||
          runSpreadsheetSelectionScopeShortcut(
            event,
            can.selectCellRange,
            commands.selectCellRange,
            'all',
          ),
        'Mod-b': ({ can, commands, context }, event) =>
          runSpreadsheetCellFormatShortcut(
            event,
            context,
            can.setCellFormat,
            commands.setCellFormat,
            'bl',
          ),
        'Mod-c': ({ can, commands }, event) =>
          runSpreadsheetClipboardShortcut(
            event,
            can.copySelection,
            commands.copySelection,
          ),
        'Mod-i': ({ can, commands, context }, event) =>
          runSpreadsheetCellFormatShortcut(
            event,
            context,
            can.setCellFormat,
            commands.setCellFormat,
            'it',
          ),
        'Mod-v': ({ can, commands }, event) =>
          runSpreadsheetClipboardShortcut(
            event,
            can.pasteSelection,
            commands.pasteSelection,
          ),
        'Mod-x': ({ can, commands }, event) =>
          runSpreadsheetClipboardShortcut(
            event,
            can.cutSelection,
            commands.cutSelection,
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
        'Mod-PageUp': ({ context, commands }, event) =>
          runSpreadsheetSheetNavigationShortcut(
            event,
            context,
            commands.activateSheet,
            -1,
          ),
        'Mod-PageDown': ({ context, commands }, event) =>
          runSpreadsheetSheetNavigationShortcut(
            event,
            context,
            commands.activateSheet,
            1,
          ),
        'Shift-F11': ({ can, commands }, event) =>
          runSpreadsheetAddSheetShortcut(
            event,
            can.addSheet,
            commands.addSheet,
          ),
        'Alt-Shift-F1': ({ can, commands }, event) =>
          runSpreadsheetAddSheetShortcut(
            event,
            can.addSheet,
            commands.addSheet,
          ),
        F9: ({ can, commands }, event) =>
          runSpreadsheetRecalculationShortcut(
            event,
            can.recalculateFormula,
            commands.recalculateFormula,
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
        ArrowUp: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'up',
            false,
          ),
        ArrowDown: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'down',
            false,
          ),
        ArrowLeft: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'left',
            false,
          ),
        ArrowRight: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'right',
            false,
          ),
        'Shift-ArrowUp': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'up',
            true,
          ),
        'Shift-ArrowDown': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'down',
            true,
          ),
        'Shift-ArrowLeft': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'left',
            true,
          ),
        'Shift-ArrowRight': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'right',
            true,
          ),
        Enter: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'down',
            false,
          ),
        'Shift-Enter': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'up',
            false,
          ),
        Tab: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'next-cell',
            false,
          ),
        'Shift-Tab': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'previous-cell',
            false,
          ),
        Home: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'row-start',
            false,
          ),
        'Mod-Home': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'sheet-start',
            false,
          ),
        'Mod-End': ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'used-end',
            false,
          ),
        PageUp: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'page-up',
            false,
          ),
        PageDown: ({ can, commands }, event) =>
          runSpreadsheetSelectionMoveShortcut(
            event,
            can.moveSelection,
            commands.moveSelection,
            'page-down',
            false,
          ),
        'Control-Space': ({ can, commands }, event) =>
          runSpreadsheetSelectionScopeShortcut(
            event,
            can.selectCellRange,
            commands.selectCellRange,
            'column',
          ),
        'Shift-Space': ({ can, commands }, event) =>
          runSpreadsheetSelectionScopeShortcut(
            event,
            can.selectCellRange,
            commands.selectCellRange,
            'row',
          ),
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands,
      SpreadsheetSelectionNavigationStorage
    >({
      name: 'spreadsheetSelectionNavigation',
      addStorage: () => ({ focus: null }),
      addCommands: ({ storage }) => ({
        moveSelection: {
          canExecute: canNavigateSpreadsheetSelection,
          execute: (context, move, extend) =>
            updateSpreadsheetSelection(context, storage, (sheet, selection) =>
              moveSpreadsheetKeyboardSelection(sheet, selection, move, extend),
            ),
        },
        selectCellRange: {
          canExecute: canNavigateSpreadsheetSelection,
          execute: (context, scope) =>
            updateSpreadsheetSelection(context, storage, (sheet, selection) =>
              scopeSpreadsheetKeyboardSelection(sheet, selection, scope),
            ),
        },
      }),
    }),
    createOfficeEditorExtension<
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
        sortSelectedCells: {
          canExecute: canSortSelectedCells,
          execute: sortSelectedCells,
        },
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

function canInsertSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  _position: SpreadsheetStructureInsertPosition,
): boolean {
  if (!canEditSelectedCells(context)) return false;
  const range = liveRange(context);
  const [start, end] = spreadsheetStructureRange(range, axis);
  const count = end - start + 1;
  const maximum = axis === 'row' ? 10_000 : 1_000;
  return spreadsheetStructureExtent(context, axis) + count < maximum;
}

function insertSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  position: SpreadsheetStructureInsertPosition,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const [start, end] = spreadsheetStructureRange(liveRange(context), axis);
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
  if (!canEditSelectedCells(context)) return false;
  const [start, end] = spreadsheetStructureRange(liveRange(context), axis);
  return end - start + 1 < spreadsheetStructureExtent(context, axis);
}

function deleteSelectedStructure(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const [start, end] = spreadsheetStructureRange(liveRange(context), axis);
  try {
    context.workbook.deleteRowOrColumn(axis, start, end, {
      id: context.targetSheetId,
    });
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
  if (!canEditSelectedCells(context)) return false;
  if (!hidden) return true;
  const [start, end] = spreadsheetStructureRange(liveRange(context), axis);
  return end - start + 1 < spreadsheetStructureExtent(context, axis);
}

function setSelectedStructureHidden(
  context: SpreadsheetCommandContext,
  axis: SpreadsheetStructureAxis,
  hidden: boolean,
): boolean {
  if (!context.workbook) return false;
  const [start, end] = spreadsheetStructureRange(liveRange(context), axis);
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
    canEditSelectedCells(context) &&
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
  const [start, end] = spreadsheetStructureRange(liveRange(context), axis);
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

function canSortSelectedCells(
  context: SpreadsheetCommandContext,
  _direction: SpreadsheetSortDirection,
): boolean {
  if (!canEditSelectedCells(context)) return false;
  const [start, end] = spreadsheetStructureRange(liveRange(context), 'row');
  return end > start;
}

function sortSelectedCells(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetSortDirection,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const range = liveRange(context);
  try {
    const rows = context.workbook.getCellsByRange(range, {
      id: context.targetSheetId,
    });
    if (rows.length < 2) return false;
    const sorted = rows
      .map((cells, index) => ({ cells, index }))
      .sort((left, right) => {
        const result = compareSpreadsheetSortCells(
          left.cells[0] ?? null,
          right.cells[0] ?? null,
          direction,
        );
        return result || left.index - right.index;
      })
      .map(({ cells }) => cells);
    context.workbook.setCellValuesByRange(sorted, range, {
      id: context.targetSheetId,
    });
    return true;
  } catch {
    return false;
  }
}

function compareSpreadsheetSortCells(
  left: Cell | null,
  right: Cell | null,
  direction: SpreadsheetSortDirection,
): number {
  const leftValue = spreadsheetSortValue(left);
  const rightValue = spreadsheetSortValue(right);
  if (leftValue === null) return rightValue === null ? 0 : 1;
  if (rightValue === null) return -1;
  const order =
    typeof leftValue === 'number' && typeof rightValue === 'number'
      ? leftValue - rightValue
      : spreadsheetSortCollator.compare(String(leftValue), String(rightValue));
  return direction === 'ascending' ? order : -order;
}

function spreadsheetSortValue(cell: Cell | null): number | string | null {
  const value = cell?.v ?? cell?.m;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (typeof value === 'boolean') return value ? 1 : 0;
  return null;
}

const spreadsheetSortCollator = new Intl.Collator('zh-CN', {
  numeric: true,
  sensitivity: 'base',
});

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
  const range = spreadsheetStructureRange(liveRange(context), axis);
  if (!sheet) return range[1] + 1;
  if (axis === 'row') {
    return Math.max(sheet.row ?? 0, sheet.data?.length ?? 0, range[1] + 1);
  }
  const dataWidth = Math.max(0, ...(sheet.data ?? []).map((row) => row.length));
  return Math.max(sheet.column ?? 0, dataWidth, range[1] + 1);
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

function canNavigateSpreadsheetSelection(
  context: SpreadsheetCommandContext,
): boolean {
  return Boolean(
    context.workbook &&
      context.targetSheetId &&
      context.content.sheets.some(
        (sheet) => sheet.id === context.targetSheetId,
      ),
  );
}

function updateSpreadsheetSelection(
  context: SpreadsheetCommandContext,
  storage: SpreadsheetSelectionNavigationStorage,
  update: (
    sheet: WorkSpreadsheetContent['sheets'][number],
    selection: SpreadsheetCommandRange,
  ) => SpreadsheetCommandRange,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  if (!sheet) return false;
  const liveSelection =
    context.workbook.getSelection()?.at(-1) ?? context.fallbackRange;
  const rememberedFocus = storage.focus;
  const selection =
    rememberedFocus &&
    spreadsheetSelectionContainsFocus(liveSelection, rememberedFocus)
      ? {
          ...liveSelection,
          row_focus: rememberedFocus.row,
          column_focus: rememberedFocus.column,
        }
      : liveSelection;
  const next = update(sheet, selection);
  try {
    context.workbook.setSelection([next], { id: context.targetSheetId });
    storage.focus = {
      row: next.row_focus ?? next.row[1] ?? next.row[0] ?? 0,
      column: next.column_focus ?? next.column[1] ?? next.column[0] ?? 0,
    };
    return true;
  } catch {
    return false;
  }
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
    context.workbook.setSelection(
      [
        {
          row: [...range.row],
          column: [...range.column],
        },
      ],
      { id: context.targetSheetId },
    );
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
    !isSpreadsheetGridKeyboardTarget(event.target) ||
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

function runSpreadsheetAddSheetShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['addSheet'],
  execute: SpreadsheetEditorCommands['addSheet'],
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

function runSpreadsheetRecalculationShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['recalculateFormula'],
  execute: SpreadsheetEditorCommands['recalculateFormula'],
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute('workbook')
  ) {
    return false;
  }
  return execute('workbook');
}

function liveRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange {
  const live = context.workbook?.getSelection()?.at(-1);
  return spreadsheetSingleRange(live ?? context.fallbackRange);
}
