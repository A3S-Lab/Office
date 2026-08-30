import type { Cell, Selection, Sheet } from '@fortune-sheet/core';
import type { SpreadsheetGridSize } from '../spreadsheet-sparse';
import type { SpreadsheetTextOrientationId } from '../work-spreadsheet-text-orientation';
import type { WorkSpreadsheetContent } from '../work-types';
import { xlsxNativeFillCellKeys } from '../work-xlsx-native-fill';
import {
  createOfficeEditorExtension,
  type OfficeEditorCanCommands,
  type OfficeEditorExtension,
} from './office-editor-extension';
import type { SpreadsheetAutoSumFunction } from './spreadsheet-auto-sum';
import { createSpreadsheetAutoSumExtension } from './spreadsheet-auto-sum-command';
import type { SpreadsheetCellBorderFormat } from './spreadsheet-cell-border';
import { createSpreadsheetCellBorderExtension } from './spreadsheet-cell-border-command';
import {
  clearSpreadsheetSheetSelection,
  type SpreadsheetCellClearMode,
} from './spreadsheet-cell-clear';
import type { SpreadsheetCellFillDirection } from './spreadsheet-cell-fill';
import { createSpreadsheetCellFillExtension } from './spreadsheet-cell-fill-command';
import type { SpreadsheetCellFormatRequest } from './spreadsheet-cell-format';
import { createSpreadsheetCellFormatExtension } from './spreadsheet-cell-format-command';
import {
  canApplySpreadsheetCellMerge,
  type SpreadsheetCellMergeCommand,
  spreadsheetCellMergeApiCalls,
} from './spreadsheet-cell-merge';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import type { SpreadsheetCellStyleChoice } from './spreadsheet-cell-style';
import { createSpreadsheetCellStyleExtension } from './spreadsheet-cell-style-command';
import {
  canEditSpreadsheetSelection,
  rememberSpreadsheetCommandSelection,
  spreadsheetLiveCommandRange,
} from './spreadsheet-command-selection';
import type { SpreadsheetCopyFromAboveKind } from './spreadsheet-copy-from-above';
import { createSpreadsheetCopyFromAboveExtension } from './spreadsheet-copy-from-above-command';
import type {
  SpreadsheetDataValidationRequest,
  SpreadsheetDataValidationTarget,
} from './spreadsheet-data-validation';
import { createSpreadsheetDataValidationExtension } from './spreadsheet-data-validation-command';
import {
  createSpreadsheetDateTimeExtension,
  type SpreadsheetDateTimeKind,
} from './spreadsheet-date-time-command';
import {
  finiteSpreadsheetSelection,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  createSpreadsheetFontSizeExtension,
  type SpreadsheetFontSizeDirection,
} from './spreadsheet-font-size-command';
import type { SpreadsheetFormatCellsOpenIntent } from './spreadsheet-format-cells-intent';
import type { SpreadsheetFormatPainterMode } from './spreadsheet-format-painter';
import {
  type SpreadsheetFreezePanePreset,
  updateSpreadsheetFreezePanes,
} from './spreadsheet-freeze-panes';
import type {
  SpreadsheetHyperlinkCell,
  SpreadsheetHyperlinkRequest,
} from './spreadsheet-hyperlink';
import { createSpreadsheetHyperlinkExtension } from './spreadsheet-hyperlink-command';
import type {
  SpreadsheetKeyboardSelection,
  SpreadsheetSelectionMove,
  SpreadsheetSelectionScope,
} from './spreadsheet-keyboard-navigation';
import { createSpreadsheetKeyboardShortcutExtension } from './spreadsheet-keyboard-shortcuts';
import { createSpreadsheetNavigationExtension } from './spreadsheet-navigation-command';
import {
  createSpreadsheetNumberFormatExtension,
  type SpreadsheetDecimalPlacesDirection,
} from './spreadsheet-number-format-command';
import type { SpreadsheetPasteContent } from './spreadsheet-paste-special';
import type { SpreadsheetRichTextToggleAttribute } from './spreadsheet-rich-text-selection-format';
import { createSpreadsheetSelectionNavigationExtension } from './spreadsheet-selection-navigation-command';
import {
  activateSpreadsheetSheet,
  addSpreadsheetSheet,
  deleteSpreadsheetSheet,
  duplicateSpreadsheetSheet,
  hideSpreadsheetSheet,
  moveSpreadsheetSheet,
  renameSpreadsheetSheet,
  type SpreadsheetSheetMoveDirection,
  setSpreadsheetSheetColor,
} from './spreadsheet-sheet-model';
import type {
  SpreadsheetSortDirection,
  SpreadsheetSortOpenRequest,
  SpreadsheetSortRequest,
} from './spreadsheet-sort';
import { createSpreadsheetSortExtension } from './spreadsheet-sort-command';
import { createSpreadsheetStructureExtension } from './spreadsheet-structure-command';
import type {
  SpreadsheetTableDesignPatch,
  SpreadsheetTableRequest,
  SpreadsheetTableTarget,
} from './spreadsheet-table';
import { createSpreadsheetTableExtension } from './spreadsheet-table-command';
import { createSpreadsheetTextOrientationExtension } from './spreadsheet-text-orientation-command';
import { createSpreadsheetVisibilityShortcutExtension } from './spreadsheet-visibility-shortcuts';

export interface SpreadsheetWorkbookCommandPort {
  autoFillCell: (
    copyRange: SpreadsheetCommandRange,
    applyRange: SpreadsheetCommandRange,
    direction: SpreadsheetCellFillDirection,
  ) => void;
  batchCallApis: (apiCalls: Array<{ name: string; args: unknown[] }>) => void;
  getSelection: () => SpreadsheetCommandRange[] | undefined;
  getCellsByRange: (
    range: SpreadsheetCommandRange,
    options?: { id?: string },
  ) => (Cell | null)[][];
  getSheet: (options?: { id?: string }) => Sheet;
  insertRowOrColumn: (
    type: SpreadsheetStructureAxis,
    index: number,
    count: number,
    direction: 'lefttop' | 'rightbottom',
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

/**
 * Synchronous selection state shared by commands and the Fortune Sheet bridge.
 * Fortune applies API calls through React state, so reading its imperative API
 * twice in the same task can otherwise return the previous selection.
 */
export interface SpreadsheetSelectionRef {
  current: SpreadsheetCommandSelection | null;
  requested: SpreadsheetCommandSelection | null;
}

export type SpreadsheetStructureAxis = 'row' | 'column';
export type SpreadsheetStructureInsertPosition = 'before' | 'after';
export type { SpreadsheetSortDirection } from './spreadsheet-sort';

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
  canOpenPasteSpecial: boolean;
  canPasteSelection: boolean;
  canPasteSpecial: (content: SpreadsheetPasteContent) => boolean;
  copySelection: () => boolean;
  cutSelection: () => boolean;
  openPasteSpecial: () => boolean;
  pasteSelection: () => boolean;
  pasteSpecial: (content: SpreadsheetPasteContent) => boolean;
}

export interface SpreadsheetAutoFilterCommandPort {
  active: boolean;
  canOpenMenu: boolean;
  canToggle: boolean;
  openMenu: () => boolean;
  toggle: () => boolean;
}

export interface SpreadsheetFormatPainterCommandPort {
  active: boolean;
  canActivate: boolean;
  mode: SpreadsheetFormatPainterMode | null;
  activate: (mode: SpreadsheetFormatPainterMode) => boolean;
  applySelection: (target: SpreadsheetCommandSelection) => boolean;
  cancel: () => boolean;
}

export interface SpreadsheetFormatCellsOpenRequest {
  sheetId: string;
  range: SpreadsheetCellRange;
  activeCell: { row: number; column: number };
  cells: (Cell | null)[][];
  intent: SpreadsheetFormatCellsOpenIntent;
}

export interface SpreadsheetFormatCellsCommandPort {
  canOpen: boolean;
  open: (request: SpreadsheetFormatCellsOpenRequest) => boolean;
}

export interface SpreadsheetDataValidationCommandPort {
  canOpen: boolean;
  open: (request: SpreadsheetDataValidationTarget) => boolean;
}

export interface SpreadsheetHyperlinkCommandPort {
  canOpen: boolean;
  open: (request: SpreadsheetHyperlinkCell) => boolean;
}

export interface SpreadsheetTableCommandPort {
  canOpen: boolean;
  open: (target: SpreadsheetTableTarget) => boolean;
}

export interface SpreadsheetSortCommandPort {
  canApply: (
    request: SpreadsheetSortRequest,
    liveRange: SpreadsheetCommandRange,
  ) => boolean;
  canOpen: boolean;
  open: (request: SpreadsheetSortOpenRequest) => boolean;
}

export interface SpreadsheetNavigationCommandPort {
  canOpenFind: boolean;
  canOpenGoTo: boolean;
  openFind: () => boolean;
  openGoTo: () => boolean;
}

export interface SpreadsheetFormulaBarCommandPort {
  setValue: (value: unknown) => void;
}

export interface SpreadsheetRichTextFormatCommandPort {
  apply: (attribute: keyof Cell, value: unknown) => boolean;
  canApply: (attribute: keyof Cell, value: unknown) => boolean;
  canToggle: (attribute: SpreadsheetRichTextToggleAttribute) => boolean;
  toggle: (attribute: SpreadsheetRichTextToggleAttribute) => boolean;
}

export interface SpreadsheetViewCommandPort {
  activateSheet: (sheetId: string) => boolean;
}

export interface SpreadsheetEditorCommands {
  activateSheet: (sheetId: string) => boolean;
  activateFormatPainter: (mode: SpreadsheetFormatPainterMode) => boolean;
  addSheet: () => boolean;
  adjustDecimalPlaces: (
    direction: SpreadsheetDecimalPlacesDirection,
  ) => boolean;
  adjustFontSize: (direction: SpreadsheetFontSizeDirection) => boolean;
  applyCellStyle: (preset: SpreadsheetCellStyleChoice) => boolean;
  applyCellFormat: (request: SpreadsheetCellFormatRequest) => boolean;
  applyDataValidation: (request: SpreadsheetDataValidationRequest) => boolean;
  applyAutoSum: (functionName: SpreadsheetAutoSumFunction) => boolean;
  applyFormatPainter: (target: SpreadsheetCommandSelection) => boolean;
  applyHyperlink: (request: SpreadsheetHyperlinkRequest) => boolean;
  applyTable: (request: SpreadsheetTableRequest) => boolean;
  applyCustomSort: (request: SpreadsheetSortRequest) => boolean;
  cancelFormatPainter: () => boolean;
  clearSelectedCells: (mode?: SpreadsheetCellClearMode) => boolean;
  copyCellFromAbove: (kind: SpreadsheetCopyFromAboveKind) => boolean;
  copySelection: () => boolean;
  cutSelection: () => boolean;
  deleteSelectedStructure: (axis: SpreadsheetStructureAxis) => boolean;
  deleteSheet: (sheetId: string) => boolean;
  duplicateSheet: (sheetId: string) => boolean;
  fillSelectedCells: (direction: SpreadsheetCellFillDirection) => boolean;
  hideSheet: (sheetId: string) => boolean;
  insertSelectedStructure: (
    axis: SpreadsheetStructureAxis,
    position: SpreadsheetStructureInsertPosition,
  ) => boolean;
  insertCurrentDateTime: (kind: SpreadsheetDateTimeKind) => boolean;
  mergeSelectedCells: (command: SpreadsheetCellMergeCommand) => boolean;
  moveSheet: (
    sheetId: string,
    direction: SpreadsheetSheetMoveDirection,
  ) => boolean;
  moveSelection: (move: SpreadsheetSelectionMove, extend: boolean) => boolean;
  openAutoFilterMenu: () => boolean;
  openDataValidation: () => boolean;
  openFind: () => boolean;
  openFormatCells: (intent?: SpreadsheetFormatCellsOpenIntent) => boolean;
  openGoTo: () => boolean;
  openHyperlink: () => boolean;
  openPasteSpecial: () => boolean;
  openCustomSort: () => boolean;
  openTable: () => boolean;
  pasteCells: (values: readonly (readonly unknown[])[]) => boolean;
  pasteSelection: () => boolean;
  pasteSpecial: (content: SpreadsheetPasteContent) => boolean;
  recalculateFormula: (scope: 'selection' | 'workbook') => boolean;
  removeHyperlink: (target: SpreadsheetHyperlinkCell) => boolean;
  removeDataValidation: (target: SpreadsheetDataValidationTarget) => boolean;
  renameSheet: (sheetId: string, name: string) => boolean;
  redo: () => boolean;
  setCellFormat: (attribute: keyof Cell, value: unknown) => boolean;
  toggleCellFormat: (attribute: SpreadsheetRichTextToggleAttribute) => boolean;
  setTextOrientation: (orientation: SpreadsheetTextOrientationId) => boolean;
  setSelectedCellBorders: (format: SpreadsheetCellBorderFormat) => boolean;
  setFreezePanes: (preset: SpreadsheetFreezePanePreset) => boolean;
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
  toggleAutoFilter: () => boolean;
  updateTable: (
    sheetId: string,
    tableId: string,
    patch: SpreadsheetTableDesignPatch,
  ) => boolean;
  convertTableToRange: (sheetId: string, tableId: string) => boolean;
  undo: () => boolean;
}

export type SpreadsheetEditorCanCommands =
  OfficeEditorCanCommands<SpreadsheetEditorCommands>;

export interface SpreadsheetCommandContext {
  activeSheetId: string;
  autoFilter: SpreadsheetAutoFilterCommandPort;
  calculation: SpreadsheetCalculationCommandPort | null;
  clipboard: SpreadsheetClipboardCommandPort;
  content: WorkSpreadsheetContent;
  dataValidation: SpreadsheetDataValidationCommandPort;
  editable: boolean;
  fallbackRange: SpreadsheetCommandRange;
  formulaBar: SpreadsheetFormulaBarCommandPort | null;
  formatPainter: SpreadsheetFormatPainterCommandPort;
  formatCells: SpreadsheetFormatCellsCommandPort;
  hyperlink: SpreadsheetHyperlinkCommandPort;
  history: SpreadsheetHistoryCommandPort | null;
  navigation: SpreadsheetNavigationCommandPort;
  onChange: (content: WorkSpreadsheetContent) => void;
  richTextFormat?: SpreadsheetRichTextFormatCommandPort | null;
  selection: SpreadsheetCommandSelection | null;
  selectionRef?: SpreadsheetSelectionRef;
  sort: SpreadsheetSortCommandPort;
  table: SpreadsheetTableCommandPort;
  targetSheetGridSize?: SpreadsheetGridSize | null;
  targetSheetId: string;
  toolbarCell: Cell | null;
  view: SpreadsheetViewCommandPort | null;
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
    createSpreadsheetCellFormatExtension(),
    createSpreadsheetTextOrientationExtension(),
    createSpreadsheetFontSizeExtension(),
    createSpreadsheetCellBorderExtension(),
    createSpreadsheetCellStyleExtension(),
    createSpreadsheetNumberFormatExtension(),
    createSpreadsheetDateTimeExtension(),
    createSpreadsheetNavigationExtension(),
    createSpreadsheetDataValidationExtension(),
    createSpreadsheetHyperlinkExtension(),
    createSpreadsheetTableExtension(),
    createSpreadsheetAutoSumExtension(),
    createSpreadsheetCellFillExtension(),
    createSpreadsheetCopyFromAboveExtension(),
    createSpreadsheetVisibilityShortcutExtension(),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetFormatPainter',
      addCommands: () => ({
        activateFormatPainter: {
          canExecute: ({ formatPainter }) => formatPainter.canActivate,
          execute: ({ formatPainter }, mode) =>
            formatPainter.canActivate && formatPainter.activate(mode),
        },
        applyFormatPainter: {
          canExecute: ({ formatPainter }) => formatPainter.active,
          execute: ({ formatPainter }, target) =>
            formatPainter.active && formatPainter.applySelection(target),
        },
        cancelFormatPainter: {
          canExecute: ({ formatPainter }) => formatPainter.active,
          execute: ({ formatPainter }) =>
            formatPainter.active && formatPainter.cancel(),
        },
      }),
    }),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetAutoFilter',
      addCommands: () => ({
        openAutoFilterMenu: {
          canExecute: ({ autoFilter }) => autoFilter.canOpenMenu,
          execute: ({ autoFilter }) =>
            autoFilter.canOpenMenu && autoFilter.openMenu(),
        },
        toggleAutoFilter: {
          canExecute: ({ autoFilter }) => autoFilter.canToggle,
          execute: ({ autoFilter }) =>
            autoFilter.canToggle && autoFilter.toggle(),
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
        pasteSpecial: {
          canExecute: ({ clipboard }, content) =>
            clipboard.canPasteSpecial(content),
          execute: ({ clipboard }, content) =>
            clipboard.canPasteSpecial(content) &&
            clipboard.pasteSpecial(content),
        },
        openPasteSpecial: {
          canExecute: ({ clipboard }) => clipboard.canOpenPasteSpecial,
          execute: ({ clipboard }) =>
            clipboard.canOpenPasteSpecial && clipboard.openPasteSpecial(),
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
            const sheet = context.content.sheets.find(
              (candidate) => candidate.id === sheetId,
            );
            if (!sheet) return false;
            if (context.view && sheet.hide !== 1) {
              return context.view?.activateSheet(sheetId) ?? false;
            }
            if (!context.editable) return false;
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
    createSpreadsheetKeyboardShortcutExtension(),
    createSpreadsheetSelectionNavigationExtension(),
    createSpreadsheetSortExtension(),
    createSpreadsheetStructureExtension(),
    createOfficeEditorExtension<
      SpreadsheetCommandContext,
      SpreadsheetEditorCommands
    >({
      name: 'spreadsheetCellFormatting',
      addCommands: () => ({
        clearSelectedCells: {
          canExecute: (context) => canEditSpreadsheetSelection(context),
          execute: clearSelectedCells,
        },
        pasteCells: {
          canExecute: (context, values) =>
            canEditSpreadsheetSelection(context) &&
            isRectangularCellMatrix(values),
          execute: pasteCells,
        },
        setCellFormat: {
          canExecute: canEditSpreadsheetSelection,
          execute: formatCells,
        },
        toggleCellFormat: {
          canExecute: canToggleCellFormat,
          execute: toggleCellFormat,
        },
        mergeSelectedCells: {
          canExecute: canMergeSelectedCells,
          execute: mergeSelectedCells,
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
        setFreezePanes: {
          canExecute: canSetFreezePanes,
          execute: setFreezePanes,
        },
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

function clearSelectedCells(
  context: SpreadsheetCommandContext,
  mode: SpreadsheetCellClearMode = 'contents',
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const range = spreadsheetLiveCommandRange(context);
  const calls: Array<{ name: string; args: unknown[] }> = [];
  try {
    const sheet = context.workbook.getSheet({ id: context.targetSheetId });
    if (mode === 'contents' || mode === 'all') {
      for (let row = range.row[0]; row <= range.row[1]; row += 1) {
        for (
          let column = range.column[0];
          column <= range.column[1];
          column += 1
        ) {
          calls.push({
            name: 'clearCell',
            args: [row, column, { id: context.targetSheetId }],
          });
        }
      }
    }
    calls.push({
      name: 'updateSheet',
      args: [[clearSpreadsheetSheetSelection(sheet, range, mode)]],
    });
    context.workbook.batchCallApis(calls);
  } catch {
    return false;
  }
  if (mode === 'contents' || mode === 'all') {
    syncSpreadsheetFormulaBar(context, '');
  }
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
  const source = spreadsheetLiveCommandRange(context);
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
    rememberSpreadsheetCommandSelection(context, {
      row: [...range.row],
      column: [...range.column],
      row_focus: range.row[0],
      column_focus: range.column[0],
    });
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
  if (context.richTextFormat?.canApply(attribute, value)) {
    return context.richTextFormat.apply(attribute, value);
  }
  if (!context.workbook || !context.targetSheetId) return false;
  if (attribute === 'ct' && !isSpreadsheetCellTypeFormat(value)) return false;
  const range = spreadsheetLiveCommandRange(context);
  const options = { id: context.targetSheetId };
  try {
    if (attribute === 'bg') {
      context.workbook.batchCallApis([
        {
          name: 'setCellFormatByRange',
          args: [attribute, value, range, options],
        },
        ...xlsxNativeFillCellKeys.map((fillAttribute) => ({
          name: 'setCellFormatByRange',
          args: [fillAttribute, undefined, range, options],
        })),
      ]);
      return true;
    }
    context.workbook.setCellFormatByRange(attribute, value, range, options);
    return true;
  } catch {
    return false;
  }
}

function canToggleCellFormat(
  context: SpreadsheetCommandContext,
  attribute: SpreadsheetRichTextToggleAttribute,
): boolean {
  if (!['bl', 'cl', 'it', 'un'].includes(attribute)) return false;
  return Boolean(
    context.richTextFormat?.canToggle(attribute) ||
      canEditSpreadsheetSelection(context),
  );
}

function toggleCellFormat(
  context: SpreadsheetCommandContext,
  attribute: SpreadsheetRichTextToggleAttribute,
): boolean {
  if (context.richTextFormat?.canToggle(attribute)) {
    return context.richTextFormat.toggle(attribute);
  }
  const active =
    attribute === 'un'
      ? Number(context.toolbarCell?.un) >= 1
      : Number(context.toolbarCell?.[attribute]) === 1;
  return formatCells(context, attribute, active ? 0 : 1);
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

function canMergeSelectedCells(
  context: SpreadsheetCommandContext,
  command: SpreadsheetCellMergeCommand,
): boolean {
  if (!canEditSpreadsheetSelection(context)) return false;
  return canApplySpreadsheetCellMerge(
    context.content.sheets.find((sheet) => sheet.id === context.targetSheetId),
    spreadsheetLiveCommandRange(context),
    command,
  );
}

function mergeSelectedCells(
  context: SpreadsheetCommandContext,
  command: SpreadsheetCellMergeCommand,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const calls = spreadsheetCellMergeApiCalls(
    context.content.sheets.find((sheet) => sheet.id === context.targetSheetId),
    context.targetSheetId,
    spreadsheetLiveCommandRange(context),
    command,
  );
  if (!calls.length) return false;
  try {
    context.workbook.batchCallApis(calls);
    return true;
  } catch {
    return false;
  }
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

function canSetFreezePanes(
  context: SpreadsheetCommandContext,
  preset: SpreadsheetFreezePanePreset,
): boolean {
  return Boolean(
    context.editable &&
      updateSpreadsheetFreezePanes(
        context.content,
        context.activeSheetId,
        preset,
        liveFreezePanesSelection(context),
      ),
  );
}

function setFreezePanes(
  context: SpreadsheetCommandContext,
  preset: SpreadsheetFreezePanePreset,
): boolean {
  return applySpreadsheetSheetChange(
    context,
    updateSpreadsheetFreezePanes(
      context.content,
      context.activeSheetId,
      preset,
      liveFreezePanesSelection(context),
    ),
  );
}

function liveFreezePanesSelection(
  context: SpreadsheetCommandContext,
): Selection {
  return finiteSpreadsheetSelection(
    context.workbook?.getSelection()?.at(-1) ?? context.fallbackRange,
  );
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
