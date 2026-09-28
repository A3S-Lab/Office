import { resolveOfficeMessages } from '../../../i18n/office-locale';

const spreadsheetCommandMessages = resolveOfficeMessages();
export const spreadsheetRibbonTabs = [
  { id: 'home', label: spreadsheetCommandMessages['spreadsheet.ribbon.home'] },
  { id: 'insert', label: spreadsheetCommandMessages['spreadsheet.ribbon.insert'] },
  {
    id: 'pageLayout',
    label: spreadsheetCommandMessages['spreadsheet.ribbon.pageLayout'],
    compactLabel: spreadsheetCommandMessages['spreadsheet.ribbon.pageLayout.compact'],
  },
  { id: 'formulas', label: spreadsheetCommandMessages['spreadsheet.ribbon.formulas'] },
  { id: 'data', label: spreadsheetCommandMessages['spreadsheet.ribbon.data'] },
  { id: 'review', label: spreadsheetCommandMessages['spreadsheet.ribbon.review'] },
  { id: 'view', label: spreadsheetCommandMessages['spreadsheet.ribbon.view'] },
] as const;

export const spreadsheetTableDesignRibbonTab = {
  id: 'tableDesign',
  label: spreadsheetCommandMessages['spreadsheet.ribbon.tableDesign'],
  compactLabel: spreadsheetCommandMessages['spreadsheet.ribbon.tableDesign.compact'],
  contextual: true,
} as const;

export type SpreadsheetRibbonTabId =
  | (typeof spreadsheetRibbonTabs)[number]['id']
  | typeof spreadsheetTableDesignRibbonTab.id;

export interface SpreadsheetCommandShortcut {
  label: string;
  aria: string;
  editor?: readonly string[];
}

export type SpreadsheetCommandLocation =
  | { area: 'quickAccess' }
  | {
      area: 'ribbon';
      tab: SpreadsheetRibbonTabId;
      group: string;
    };

export interface SpreadsheetCommandDefinition {
  id: string;
  label: string;
  location: SpreadsheetCommandLocation;
  menuShortcut?: SpreadsheetCommandShortcut;
  shortcut?: SpreadsheetCommandShortcut;
}

export const spreadsheetCommandCatalog = {
  undo: {
    id: 'history.undo',
    label: spreadsheetCommandMessages['spreadsheet.command.undo'],
    location: { area: 'quickAccess' },
    shortcut: {
      label: 'Cmd/Ctrl+Z',
      aria: 'Control+Z Meta+Z',
      editor: ['Mod-z'],
    },
  },
  redo: {
    id: 'history.redo',
    label: spreadsheetCommandMessages['spreadsheet.command.redo'],
    location: { area: 'quickAccess' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+Z or Cmd/Ctrl+Y',
      aria: 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
      editor: ['Mod-Shift-z', 'Mod-y'],
    },
  },
  paste: {
    id: 'clipboard.paste',
    label: spreadsheetCommandMessages['spreadsheet.command.paste'],
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+V',
      aria: 'Control+V Meta+V',
      editor: ['Mod-v'],
    },
  },
  pasteSpecial: {
    id: 'clipboard.pasteSpecial',
    label: spreadsheetCommandMessages['spreadsheet.command.pasteSpecial'],
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+V',
      aria: 'Control+Alt+V Meta+Alt+V',
      editor: ['Mod-Alt-v'],
    },
  },
  cut: {
    id: 'clipboard.cut',
    label: spreadsheetCommandMessages['spreadsheet.command.cut'],
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+X',
      aria: 'Control+X Meta+X',
      editor: ['Mod-x'],
    },
  },
  copy: {
    id: 'clipboard.copy',
    label: spreadsheetCommandMessages['spreadsheet.command.copy'],
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+C',
      aria: 'Control+C Meta+C',
      editor: ['Mod-c'],
    },
  },
  formatPainter: {
    id: 'clipboard.formatPainter',
    label: spreadsheetCommandMessages['spreadsheet.command.formatPainter'],
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
  },
  bold: {
    id: 'font.bold',
    label: spreadsheetCommandMessages['spreadsheet.command.bold'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+B or Ctrl+2',
      aria: 'Control+B Meta+B Control+2',
      editor: ['Mod-b', 'Control-2'],
    },
  },
  italic: {
    id: 'font.italic',
    label: spreadsheetCommandMessages['spreadsheet.command.italic'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+I or Ctrl+3',
      aria: 'Control+I Meta+I Control+3',
      editor: ['Mod-i', 'Control-3'],
    },
  },
  underline: {
    id: 'font.underline',
    label: spreadsheetCommandMessages['spreadsheet.command.underline'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+U or Ctrl+4',
      aria: 'Control+U Meta+U Control+4',
      editor: ['Mod-u', 'Control-4'],
    },
  },
  strike: {
    id: 'font.strike',
    label: spreadsheetCommandMessages['spreadsheet.command.strike'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+5',
      aria: 'Control+5 Meta+5',
      editor: ['Mod-5'],
    },
  },
  growFont: {
    id: 'font.grow',
    label: spreadsheetCommandMessages['spreadsheet.command.growFont'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+. or Cmd/Ctrl+]',
      aria: 'Control+Shift+. Meta+Shift+. Control+] Meta+]',
      editor: ['Mod-Shift-.', 'Mod-]'],
    },
  },
  shrinkFont: {
    id: 'font.shrink',
    label: spreadsheetCommandMessages['spreadsheet.command.shrinkFont'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+, or Cmd/Ctrl+[',
      aria: 'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
      editor: ['Mod-Shift-,', 'Mod-['],
    },
  },
  numberFormatGeneral: {
    id: 'number.general',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatGeneral'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+~',
      aria: 'Control+Shift+~ Meta+Shift+~',
      editor: ['Mod-Shift-~'],
    },
  },
  numberFormatNumber: {
    id: 'number.number',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatNumber'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+!',
      aria: 'Control+Shift+! Meta+Shift+!',
      editor: ['Mod-Shift-!'],
    },
  },
  numberFormatCurrency: {
    id: 'number.currency',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatCurrency'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+$',
      aria: 'Control+Shift+$ Meta+Shift+$',
      editor: ['Mod-Shift-$'],
    },
  },
  numberFormatAccounting: {
    id: 'number.accounting',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatAccounting'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  numberFormatPercent: {
    id: 'number.percent',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatPercent'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+%',
      aria: 'Control+Shift+% Meta+Shift+%',
      editor: ['Mod-Shift-%'],
    },
  },
  numberFormatDate: {
    id: 'number.date',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatDate'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+#',
      aria: 'Control+Shift+# Meta+Shift+#',
      editor: ['Mod-Shift-#'],
    },
  },
  numberFormatTime: {
    id: 'number.time',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatTime'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+@',
      aria: 'Control+Shift+@ Meta+Shift+@',
      editor: ['Mod-Shift-@'],
    },
  },
  numberFormatScientific: {
    id: 'number.scientific',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatScientific'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+^',
      aria: 'Control+Shift+^ Meta+Shift+^',
      editor: ['Mod-Shift-^'],
    },
  },
  numberFormatFraction: {
    id: 'number.fraction',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatFraction'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  numberFormatText: {
    id: 'number.text',
    label: spreadsheetCommandMessages['spreadsheet.command.numberFormatText'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  decreaseDecimalPlaces: {
    id: 'number.decreaseDecimalPlaces',
    label: spreadsheetCommandMessages['spreadsheet.command.decreaseDecimalPlaces'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  increaseDecimalPlaces: {
    id: 'number.increaseDecimalPlaces',
    label: spreadsheetCommandMessages['spreadsheet.command.increaseDecimalPlaces'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  insertCurrentDate: {
    id: 'number.insertCurrentDate',
    label: spreadsheetCommandMessages['spreadsheet.command.insertCurrentDate'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Ctrl+;',
      aria: 'Control+;',
      editor: ['Control-;'],
    },
  },
  insertCurrentTime: {
    id: 'number.insertCurrentTime',
    label: spreadsheetCommandMessages['spreadsheet.command.insertCurrentTime'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Ctrl+Shift+;',
      aria: 'Control+Shift+;',
      editor: ['Control-Shift-;'],
    },
  },
  formatCells: {
    id: 'number.formatCells',
    label: spreadsheetCommandMessages['spreadsheet.command.formatCells'],
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+1',
      aria: 'Control+1 Meta+1',
      editor: ['Mod-1'],
    },
  },
  formatCellsFont: {
    id: 'font.formatCellsFont',
    label: spreadsheetCommandMessages['spreadsheet.command.formatCellsFont'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+F',
      aria: 'Control+Shift+F Meta+Shift+F',
      editor: ['Mod-Shift-f'],
    },
  },
  formatCellsFontSize: {
    id: 'font.formatCellsFontSize',
    label: spreadsheetCommandMessages['spreadsheet.command.formatCellsFontSize'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+P',
      aria: 'Control+Shift+P Meta+Shift+P',
      editor: ['Mod-Shift-p'],
    },
  },
  borderTop: {
    id: 'font.borderTop',
    label: spreadsheetCommandMessages['spreadsheet.command.borderTop'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderBottom: {
    id: 'font.borderBottom',
    label: spreadsheetCommandMessages['spreadsheet.command.borderBottom'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderLeft: {
    id: 'font.borderLeft',
    label: spreadsheetCommandMessages['spreadsheet.command.borderLeft'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderRight: {
    id: 'font.borderRight',
    label: spreadsheetCommandMessages['spreadsheet.command.borderRight'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderNone: {
    id: 'font.borderNone',
    label: spreadsheetCommandMessages['spreadsheet.command.borderNone'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+_',
      aria: 'Control+Shift+_ Meta+Shift+_',
      editor: ['Mod-Shift-_', 'Mod-Shift-Minus'],
    },
  },
  borderAll: {
    id: 'font.borderAll',
    label: spreadsheetCommandMessages['spreadsheet.command.borderAll'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderOutside: {
    id: 'font.borderOutside',
    label: spreadsheetCommandMessages['spreadsheet.command.borderOutside'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+&',
      aria: 'Control+Shift+& Meta+Shift+&',
      editor: ['Mod-Shift-&', 'Mod-Shift-7'],
    },
  },
  borderInside: {
    id: 'font.borderInside',
    label: spreadsheetCommandMessages['spreadsheet.command.borderInside'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderHorizontal: {
    id: 'font.borderHorizontal',
    label: spreadsheetCommandMessages['spreadsheet.command.borderHorizontal'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderVertical: {
    id: 'font.borderVertical',
    label: spreadsheetCommandMessages['spreadsheet.command.borderVertical'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderDiagonalDown: {
    id: 'font.borderDiagonalDown',
    label: spreadsheetCommandMessages['spreadsheet.command.borderDiagonalDown'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderDiagonalUp: {
    id: 'font.borderDiagonalUp',
    label: spreadsheetCommandMessages['spreadsheet.command.borderDiagonalUp'],
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  textOrientationHorizontal: {
    id: 'alignment.textOrientation.horizontal',
    label: spreadsheetCommandMessages['spreadsheet.command.textOrientationHorizontal'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationAngleCounterclockwise: {
    id: 'alignment.textOrientation.angleCounterclockwise',
    label: spreadsheetCommandMessages['spreadsheet.command.textOrientationAngleCounterclockwise'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationAngleClockwise: {
    id: 'alignment.textOrientation.angleClockwise',
    label: spreadsheetCommandMessages['spreadsheet.command.textOrientationAngleClockwise'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationVertical: {
    id: 'alignment.textOrientation.vertical',
    label: spreadsheetCommandMessages['spreadsheet.command.textOrientationVertical'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationRotateUp: {
    id: 'alignment.textOrientation.rotateUp',
    label: spreadsheetCommandMessages['spreadsheet.command.textOrientationRotateUp'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationRotateDown: {
    id: 'alignment.textOrientation.rotateDown',
    label: spreadsheetCommandMessages['spreadsheet.command.textOrientationRotateDown'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  mergeAndCenter: {
    id: 'alignment.mergeAndCenter',
    label: spreadsheetCommandMessages['spreadsheet.command.mergeAndCenter'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
    shortcut: {
      label: 'Ctrl+M',
      aria: 'Control+M',
      editor: ['Control-m'],
    },
  },
  mergeCells: {
    id: 'alignment.mergeCells',
    label: spreadsheetCommandMessages['spreadsheet.command.mergeCells'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  mergeAcross: {
    id: 'alignment.mergeAcross',
    label: spreadsheetCommandMessages['spreadsheet.command.mergeAcross'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  unmergeCells: {
    id: 'alignment.unmergeCells',
    label: spreadsheetCommandMessages['spreadsheet.command.unmergeCells'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  unmergeAndFill: {
    id: 'alignment.unmergeAndFill',
    label: spreadsheetCommandMessages['spreadsheet.command.unmergeAndFill'],
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  conditionalFormatting: {
    id: 'styles.conditionalFormatting',
    label: spreadsheetCommandMessages['spreadsheet.command.conditionalFormatting'],
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
  },
  cellStyles: {
    id: 'styles.cellStyles',
    label: spreadsheetCommandMessages['spreadsheet.command.cellStyles'],
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
  },
  insertRowsAbove: {
    id: 'cells.insertRowsAbove',
    label: spreadsheetCommandMessages['spreadsheet.command.insertRowsAbove'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  insertRowsBelow: {
    id: 'cells.insertRowsBelow',
    label: spreadsheetCommandMessages['spreadsheet.command.insertRowsBelow'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  insertColumnsLeft: {
    id: 'cells.insertColumnsLeft',
    label: spreadsheetCommandMessages['spreadsheet.command.insertColumnsLeft'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  insertColumnsRight: {
    id: 'cells.insertColumnsRight',
    label: spreadsheetCommandMessages['spreadsheet.command.insertColumnsRight'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  deleteRows: {
    id: 'cells.deleteRows',
    label: spreadsheetCommandMessages['spreadsheet.command.deleteRows'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  deleteColumns: {
    id: 'cells.deleteColumns',
    label: spreadsheetCommandMessages['spreadsheet.command.deleteColumns'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  hideRows: {
    id: 'cells.hideRows',
    label: spreadsheetCommandMessages['spreadsheet.command.hideRows'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+9',
      aria: 'Control+9 Meta+9',
      editor: ['Mod-9'],
    },
  },
  hideColumns: {
    id: 'cells.hideColumns',
    label: spreadsheetCommandMessages['spreadsheet.command.hideColumns'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+0',
      aria: 'Control+0 Meta+0',
      editor: ['Mod-0'],
    },
  },
  unhideRows: {
    id: 'cells.unhideRows',
    label: spreadsheetCommandMessages['spreadsheet.command.unhideRows'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+9',
      aria: 'Control+Shift+9 Meta+Shift+9',
      editor: ['Mod-Shift-9'],
    },
  },
  unhideColumns: {
    id: 'cells.unhideColumns',
    label: spreadsheetCommandMessages['spreadsheet.command.unhideColumns'],
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+0',
      aria: 'Control+Shift+0 Meta+Shift+0',
      editor: ['Mod-Shift-0'],
    },
  },
  autoSum: {
    id: 'editing.autoSum',
    label: spreadsheetCommandMessages['spreadsheet.command.autoSum'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Alt+=',
      aria: 'Alt+=',
      editor: ['Alt-equal'],
    },
  },
  autoAverage: {
    id: 'editing.autoAverage',
    label: spreadsheetCommandMessages['spreadsheet.command.autoAverage'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  autoCount: {
    id: 'editing.autoCount',
    label: spreadsheetCommandMessages['spreadsheet.command.autoCount'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  autoMaximum: {
    id: 'editing.autoMaximum',
    label: spreadsheetCommandMessages['spreadsheet.command.autoMaximum'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  autoMinimum: {
    id: 'editing.autoMinimum',
    label: spreadsheetCommandMessages['spreadsheet.command.autoMinimum'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  fillDown: {
    id: 'editing.fillDown',
    label: spreadsheetCommandMessages['spreadsheet.command.fillDown'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+D',
      aria: 'Control+D Meta+D',
      editor: ['Mod-d'],
    },
  },
  fillRight: {
    id: 'editing.fillRight',
    label: spreadsheetCommandMessages['spreadsheet.command.fillRight'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+R',
      aria: 'Control+R Meta+R',
      editor: ['Mod-r'],
    },
  },
  fillUp: {
    id: 'editing.fillUp',
    label: spreadsheetCommandMessages['spreadsheet.command.fillUp'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  fillLeft: {
    id: 'editing.fillLeft',
    label: spreadsheetCommandMessages['spreadsheet.command.fillLeft'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  copyFormulaFromAbove: {
    id: 'editing.copyFormulaFromAbove',
    label: spreadsheetCommandMessages['spreadsheet.command.copyFormulaFromAbove'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: "Ctrl+'",
      aria: "Control+'",
      editor: ["Control-'"],
    },
  },
  copyValueFromAbove: {
    id: 'editing.copyValueFromAbove',
    label: spreadsheetCommandMessages['spreadsheet.command.copyValueFromAbove'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: "Ctrl+Shift+'",
      aria: "Control+Shift+'",
      editor: ["Control-Shift-'"],
    },
  },
  clearAll: {
    id: 'editing.clearAll',
    label: spreadsheetCommandMessages['spreadsheet.command.clearAll'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  clearFormats: {
    id: 'editing.clearFormats',
    label: spreadsheetCommandMessages['spreadsheet.command.clearFormats'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  clearContents: {
    id: 'editing.clearContents',
    label: spreadsheetCommandMessages['spreadsheet.command.clearContents'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Delete/Backspace',
      aria: 'Delete Backspace',
      editor: ['Delete', 'Backspace'],
    },
  },
  clearComments: {
    id: 'editing.clearComments',
    label: spreadsheetCommandMessages['spreadsheet.command.clearComments'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  clearHyperlinks: {
    id: 'editing.clearHyperlinks',
    label: spreadsheetCommandMessages['spreadsheet.command.clearHyperlinks'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  find: {
    id: 'editing.find',
    label: spreadsheetCommandMessages['spreadsheet.command.find'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+F',
      aria: 'Control+F Meta+F',
      editor: ['Mod-f'],
    },
  },
  findAndSelect: {
    id: 'editing.findAndSelect',
    label: spreadsheetCommandMessages['spreadsheet.command.findAndSelect'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  goTo: {
    id: 'editing.goTo',
    label: spreadsheetCommandMessages['spreadsheet.command.goTo'],
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Ctrl+G or F5',
      aria: 'Control+G F5',
      editor: ['Control-g', 'F5'],
    },
  },
  insertChart: {
    id: 'insert.chart',
    label: spreadsheetCommandMessages['spreadsheet.command.insertChart'],
    location: { area: 'ribbon', tab: 'insert', group: 'charts' },
  },
  table: {
    id: 'insert.table',
    label: spreadsheetCommandMessages['spreadsheet.command.table'],
    location: { area: 'ribbon', tab: 'insert', group: 'tables' },
    shortcut: {
      label: 'Cmd/Ctrl+T',
      aria: 'Control+T Meta+T',
      editor: ['Mod-t'],
    },
  },
  hyperlink: {
    id: 'insert.hyperlink',
    label: spreadsheetCommandMessages['spreadsheet.command.hyperlink'],
    location: { area: 'ribbon', tab: 'insert', group: 'links' },
    shortcut: {
      label: 'Cmd/Ctrl+K',
      aria: 'Control+K Meta+K',
      editor: ['Mod-k'],
    },
  },
  printSettings: {
    id: 'pageLayout.printSettings',
    label: spreadsheetCommandMessages['spreadsheet.command.printSettings'],
    location: { area: 'ribbon', tab: 'pageLayout', group: 'pageSetup' },
  },
  nameManager: {
    id: 'formulas.nameManager',
    label: spreadsheetCommandMessages['spreadsheet.command.nameManager'],
    location: { area: 'ribbon', tab: 'formulas', group: 'definedNames' },
  },
  formulaManager: {
    id: 'formulas.manager',
    label: spreadsheetCommandMessages['spreadsheet.command.formulaManager'],
    location: { area: 'ribbon', tab: 'formulas', group: 'calculation' },
  },
  recalculateWorkbook: {
    id: 'formulas.recalculateWorkbook',
    label: spreadsheetCommandMessages['spreadsheet.command.recalculateWorkbook'],
    location: { area: 'ribbon', tab: 'formulas', group: 'calculation' },
    shortcut: { label: 'F9', aria: 'F9', editor: ['F9'] },
  },
  sortAscending: {
    id: 'data.sortAscending',
    label: spreadsheetCommandMessages['spreadsheet.command.sortAscending'],
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
  },
  sortDescending: {
    id: 'data.sortDescending',
    label: spreadsheetCommandMessages['spreadsheet.command.sortDescending'],
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
  },
  customSort: {
    id: 'data.customSort',
    label: spreadsheetCommandMessages['spreadsheet.command.customSort'],
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
  },
  autoFilter: {
    id: 'data.autoFilter',
    label: spreadsheetCommandMessages['spreadsheet.command.autoFilter'],
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+L',
      aria: 'Control+Shift+L Meta+Shift+L',
      editor: ['Mod-Shift-l'],
    },
    menuShortcut: {
      label: 'Alt+↓',
      aria: 'Alt+ArrowDown',
      editor: ['Alt-ArrowDown'],
    },
  },
  dataValidation: {
    id: 'data.validation',
    label: spreadsheetCommandMessages['spreadsheet.command.dataValidation'],
    location: { area: 'ribbon', tab: 'data', group: 'dataTools' },
  },
  pivotTable: {
    id: 'data.pivotTable',
    label: spreadsheetCommandMessages['spreadsheet.command.pivotTable'],
    location: { area: 'ribbon', tab: 'data', group: 'analysis' },
  },
  protectSheet: {
    id: 'review.protectSheet',
    label: spreadsheetCommandMessages['spreadsheet.command.protectSheet'],
    location: { area: 'ribbon', tab: 'review', group: 'protection' },
  },
  gridLines: {
    id: 'view.gridLines',
    label: spreadsheetCommandMessages['spreadsheet.command.gridLines'],
    location: { area: 'ribbon', tab: 'view', group: 'workbookViews' },
  },
  formulaBar: {
    id: 'view.formulaBar',
    label: spreadsheetCommandMessages['spreadsheet.command.formulaBar'],
    location: { area: 'ribbon', tab: 'view', group: 'workbookViews' },
  },
  showFormulas: {
    id: 'view.showFormulas',
    label: spreadsheetCommandMessages['spreadsheet.command.showFormulas'],
    location: { area: 'ribbon', tab: 'view', group: 'workbookViews' },
    shortcut: {
      label: 'Ctrl+`',
      aria: 'Control+` Meta+`',
      editor: ['Mod-`', 'Control-`'],
    },
  },
  headings: {
    id: 'view.headings',
    label: spreadsheetCommandMessages['spreadsheet.command.headings'],
    location: { area: 'ribbon', tab: 'view', group: 'workbookViews' },
  },
  freezePanes: {
    id: 'view.freezePanes',
    label: spreadsheetCommandMessages['spreadsheet.command.freezePanes'],
    location: { area: 'ribbon', tab: 'view', group: 'window' },
  },
} as const satisfies Record<string, SpreadsheetCommandDefinition>;

export type SpreadsheetCommandId = keyof typeof spreadsheetCommandCatalog;

export function getSpreadsheetCommandDefinition<T extends SpreadsheetCommandId>(
  id: T,
): SpreadsheetCommandDefinition {
  return spreadsheetCommandCatalog[id];
}
