import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import {
  runSpreadsheetClipboardShortcut,
  runSpreadsheetPasteSpecialShortcut,
} from './spreadsheet-clipboard-shortcuts';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { selectSpreadsheetFormulaBarContents } from './spreadsheet-editor-support';
import {
  runSpreadsheetAddSheetShortcut,
  runSpreadsheetAutoFilterShortcut,
  runSpreadsheetCellFormatShortcut,
  runSpreadsheetClearShortcut,
  runSpreadsheetFillShortcut,
  runSpreadsheetFormatPainterEscape,
  runSpreadsheetHistoryShortcut,
  runSpreadsheetMergeShortcut,
  runSpreadsheetNumberFormatShortcut,
  runSpreadsheetRecalculationShortcut,
  runSpreadsheetSheetNavigationShortcut,
} from './spreadsheet-keyboard-shortcut-runners';
import {
  runSpreadsheetSelectionMoveShortcut,
  runSpreadsheetSelectionScopeShortcut,
} from './spreadsheet-keyboard-navigation';

export function createSpreadsheetKeyboardShortcutExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
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
      'Mod-d': ({ can, commands }, event) =>
        runSpreadsheetFillShortcut(
          event,
          can.fillSelectedCells,
          commands.fillSelectedCells,
          'down',
        ),
      'Mod-i': ({ can, commands, context }, event) =>
        runSpreadsheetCellFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'it',
        ),
      'Mod-r': ({ can, commands }, event) =>
        runSpreadsheetFillShortcut(
          event,
          can.fillSelectedCells,
          commands.fillSelectedCells,
          'right',
        ),
      'Mod-v': ({ can, commands }, event) =>
        runSpreadsheetClipboardShortcut(
          event,
          can.pasteSelection,
          commands.pasteSelection,
        ),
      [spreadsheetCommandCatalog.pasteSpecial.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetPasteSpecialShortcut(
          event,
          can.openPasteSpecial,
          commands.openPasteSpecial,
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
      'Mod-5': ({ can, commands, context }, event) =>
        runSpreadsheetCellFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'cl',
        ),
      [spreadsheetCommandCatalog.numberFormatGeneral.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'general',
        ),
      [spreadsheetCommandCatalog.numberFormatNumber.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'number',
        ),
      [spreadsheetCommandCatalog.numberFormatCurrency.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'currency',
        ),
      [spreadsheetCommandCatalog.numberFormatPercent.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'percent',
        ),
      [spreadsheetCommandCatalog.numberFormatDate.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'date',
        ),
      [spreadsheetCommandCatalog.numberFormatTime.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'time',
        ),
      [spreadsheetCommandCatalog.numberFormatScientific.shortcut.editor[0]]: (
        { can, commands, context },
        event,
      ) =>
        runSpreadsheetNumberFormatShortcut(
          event,
          context,
          can.setCellFormat,
          commands.setCellFormat,
          'scientific',
        ),
      'Mod-z': ({ can, commands }, event) =>
        runSpreadsheetHistoryShortcut(event, can.undo, commands.undo),
      'Mod-Shift-z': ({ can, commands }, event) =>
        runSpreadsheetHistoryShortcut(event, can.redo, commands.redo),
      'Mod-y': ({ can, commands }, event) =>
        runSpreadsheetHistoryShortcut(event, can.redo, commands.redo),
      'Control-m': ({ can, commands }, event) =>
        runSpreadsheetMergeShortcut(
          event,
          can.mergeSelectedCells,
          commands.mergeSelectedCells,
        ),
      'Mod-Shift-l': ({ can, commands }, event) =>
        runSpreadsheetAutoFilterShortcut(
          event,
          can.toggleAutoFilter,
          commands.toggleAutoFilter,
        ),
      'Alt-ArrowDown': ({ can, commands }, event) =>
        runSpreadsheetAutoFilterShortcut(
          event,
          can.openAutoFilterMenu,
          commands.openAutoFilterMenu,
        ),
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
        runSpreadsheetAddSheetShortcut(event, can.addSheet, commands.addSheet),
      'Alt-Shift-F1': ({ can, commands }, event) =>
        runSpreadsheetAddSheetShortcut(event, can.addSheet, commands.addSheet),
      Escape: ({ can, commands }, event) =>
        runSpreadsheetFormatPainterEscape(
          event,
          can.cancelFormatPainter,
          commands.cancelFormatPainter,
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
  });
}
