import { spreadsheetNumberFormatPresetLabels } from './spreadsheet-number-format';

export const spreadsheetRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'pageLayout', label: '页面布局', compactLabel: '布局' },
  { id: 'formulas', label: '公式' },
  { id: 'data', label: '数据' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
] as const;

export const spreadsheetTableDesignRibbonTab = {
  id: 'tableDesign',
  label: '表格设计',
  compactLabel: '设计',
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
    label: '撤销',
    location: { area: 'quickAccess' },
    shortcut: {
      label: 'Cmd/Ctrl+Z',
      aria: 'Control+Z Meta+Z',
      editor: ['Mod-z'],
    },
  },
  redo: {
    id: 'history.redo',
    label: '重做',
    location: { area: 'quickAccess' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+Z 或 Cmd/Ctrl+Y',
      aria: 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
      editor: ['Mod-Shift-z', 'Mod-y'],
    },
  },
  paste: {
    id: 'clipboard.paste',
    label: '粘贴',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+V',
      aria: 'Control+V Meta+V',
      editor: ['Mod-v'],
    },
  },
  pasteSpecial: {
    id: 'clipboard.pasteSpecial',
    label: '选择性粘贴',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+V',
      aria: 'Control+Alt+V Meta+Alt+V',
      editor: ['Mod-Alt-v'],
    },
  },
  cut: {
    id: 'clipboard.cut',
    label: '剪切',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+X',
      aria: 'Control+X Meta+X',
      editor: ['Mod-x'],
    },
  },
  copy: {
    id: 'clipboard.copy',
    label: '复制',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+C',
      aria: 'Control+C Meta+C',
      editor: ['Mod-c'],
    },
  },
  formatPainter: {
    id: 'clipboard.formatPainter',
    label: '格式刷',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
  },
  bold: {
    id: 'font.bold',
    label: '加粗',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+B 或 Ctrl+2',
      aria: 'Control+B Meta+B Control+2',
      editor: ['Mod-b', 'Control-2'],
    },
  },
  italic: {
    id: 'font.italic',
    label: '斜体',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+I 或 Ctrl+3',
      aria: 'Control+I Meta+I Control+3',
      editor: ['Mod-i', 'Control-3'],
    },
  },
  underline: {
    id: 'font.underline',
    label: '下划线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+U 或 Ctrl+4',
      aria: 'Control+U Meta+U Control+4',
      editor: ['Mod-u', 'Control-4'],
    },
  },
  strike: {
    id: 'font.strike',
    label: '删除线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+5',
      aria: 'Control+5 Meta+5',
      editor: ['Mod-5'],
    },
  },
  growFont: {
    id: 'font.grow',
    label: '增大字号',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+. 或 Cmd/Ctrl+]',
      aria: 'Control+Shift+. Meta+Shift+. Control+] Meta+]',
      editor: ['Mod-Shift-.', 'Mod-]'],
    },
  },
  shrinkFont: {
    id: 'font.shrink',
    label: '减小字号',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+, 或 Cmd/Ctrl+[',
      aria: 'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
      editor: ['Mod-Shift-,', 'Mod-['],
    },
  },
  numberFormatGeneral: {
    id: 'number.general',
    label: spreadsheetNumberFormatPresetLabels.general,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+~',
      aria: 'Control+Shift+~ Meta+Shift+~',
      editor: ['Mod-Shift-~'],
    },
  },
  numberFormatNumber: {
    id: 'number.number',
    label: spreadsheetNumberFormatPresetLabels.number,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+!',
      aria: 'Control+Shift+! Meta+Shift+!',
      editor: ['Mod-Shift-!'],
    },
  },
  numberFormatCurrency: {
    id: 'number.currency',
    label: spreadsheetNumberFormatPresetLabels.currency,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+$',
      aria: 'Control+Shift+$ Meta+Shift+$',
      editor: ['Mod-Shift-$'],
    },
  },
  numberFormatAccounting: {
    id: 'number.accounting',
    label: spreadsheetNumberFormatPresetLabels.accounting,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  numberFormatPercent: {
    id: 'number.percent',
    label: spreadsheetNumberFormatPresetLabels.percent,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+%',
      aria: 'Control+Shift+% Meta+Shift+%',
      editor: ['Mod-Shift-%'],
    },
  },
  numberFormatDate: {
    id: 'number.date',
    label: spreadsheetNumberFormatPresetLabels.date,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+#',
      aria: 'Control+Shift+# Meta+Shift+#',
      editor: ['Mod-Shift-#'],
    },
  },
  numberFormatTime: {
    id: 'number.time',
    label: spreadsheetNumberFormatPresetLabels.time,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+@',
      aria: 'Control+Shift+@ Meta+Shift+@',
      editor: ['Mod-Shift-@'],
    },
  },
  numberFormatScientific: {
    id: 'number.scientific',
    label: spreadsheetNumberFormatPresetLabels.scientific,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+^',
      aria: 'Control+Shift+^ Meta+Shift+^',
      editor: ['Mod-Shift-^'],
    },
  },
  numberFormatFraction: {
    id: 'number.fraction',
    label: spreadsheetNumberFormatPresetLabels.fraction,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  numberFormatText: {
    id: 'number.text',
    label: spreadsheetNumberFormatPresetLabels.text,
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  decreaseDecimalPlaces: {
    id: 'number.decreaseDecimalPlaces',
    label: '减少小数位',
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  increaseDecimalPlaces: {
    id: 'number.increaseDecimalPlaces',
    label: '增加小数位',
    location: { area: 'ribbon', tab: 'home', group: 'number' },
  },
  insertCurrentDate: {
    id: 'number.insertCurrentDate',
    label: '插入当前日期',
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Ctrl+;',
      aria: 'Control+;',
      editor: ['Control-;'],
    },
  },
  insertCurrentTime: {
    id: 'number.insertCurrentTime',
    label: '插入当前时间',
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Ctrl+Shift+;',
      aria: 'Control+Shift+;',
      editor: ['Control-Shift-;'],
    },
  },
  formatCells: {
    id: 'number.formatCells',
    label: '设置单元格格式',
    location: { area: 'ribbon', tab: 'home', group: 'number' },
    shortcut: {
      label: 'Cmd/Ctrl+1',
      aria: 'Control+1 Meta+1',
      editor: ['Mod-1'],
    },
  },
  formatCellsFont: {
    id: 'font.formatCellsFont',
    label: '设置字体格式',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+F',
      aria: 'Control+Shift+F Meta+Shift+F',
      editor: ['Mod-Shift-f'],
    },
  },
  formatCellsFontSize: {
    id: 'font.formatCellsFontSize',
    label: '设置字号格式',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+P',
      aria: 'Control+Shift+P Meta+Shift+P',
      editor: ['Mod-Shift-p'],
    },
  },
  borderTop: {
    id: 'font.borderTop',
    label: '上框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderBottom: {
    id: 'font.borderBottom',
    label: '下框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderLeft: {
    id: 'font.borderLeft',
    label: '左框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderRight: {
    id: 'font.borderRight',
    label: '右框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderNone: {
    id: 'font.borderNone',
    label: '无框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+_',
      aria: 'Control+Shift+_ Meta+Shift+_',
      editor: ['Mod-Shift-_', 'Mod-Shift-Minus'],
    },
  },
  borderAll: {
    id: 'font.borderAll',
    label: '所有框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderOutside: {
    id: 'font.borderOutside',
    label: '外侧框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+&',
      aria: 'Control+Shift+& Meta+Shift+&',
      editor: ['Mod-Shift-&', 'Mod-Shift-7'],
    },
  },
  borderInside: {
    id: 'font.borderInside',
    label: '内部框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderHorizontal: {
    id: 'font.borderHorizontal',
    label: '内部横框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderVertical: {
    id: 'font.borderVertical',
    label: '内部竖框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderDiagonalDown: {
    id: 'font.borderDiagonalDown',
    label: '斜下框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  borderDiagonalUp: {
    id: 'font.borderDiagonalUp',
    label: '斜上框线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  textOrientationHorizontal: {
    id: 'alignment.textOrientation.horizontal',
    label: '横排文字',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationAngleCounterclockwise: {
    id: 'alignment.textOrientation.angleCounterclockwise',
    label: '逆时针倾斜',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationAngleClockwise: {
    id: 'alignment.textOrientation.angleClockwise',
    label: '顺时针倾斜',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationVertical: {
    id: 'alignment.textOrientation.vertical',
    label: '竖排文字',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationRotateUp: {
    id: 'alignment.textOrientation.rotateUp',
    label: '向上旋转文字',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  textOrientationRotateDown: {
    id: 'alignment.textOrientation.rotateDown',
    label: '向下旋转文字',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  mergeAndCenter: {
    id: 'alignment.mergeAndCenter',
    label: '合并居中',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
    shortcut: {
      label: 'Ctrl+M',
      aria: 'Control+M',
      editor: ['Control-m'],
    },
  },
  mergeCells: {
    id: 'alignment.mergeCells',
    label: '合并单元格',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  mergeAcross: {
    id: 'alignment.mergeAcross',
    label: '跨行合并',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  unmergeCells: {
    id: 'alignment.unmergeCells',
    label: '取消合并单元格',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  unmergeAndFill: {
    id: 'alignment.unmergeAndFill',
    label: '取消合并并填充',
    location: { area: 'ribbon', tab: 'home', group: 'alignment' },
  },
  conditionalFormatting: {
    id: 'styles.conditionalFormatting',
    label: '条件格式',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
  },
  cellStyles: {
    id: 'styles.cellStyles',
    label: '单元格样式',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
  },
  insertRowsAbove: {
    id: 'cells.insertRowsAbove',
    label: '在上方插入行',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  insertRowsBelow: {
    id: 'cells.insertRowsBelow',
    label: '在下方插入行',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  insertColumnsLeft: {
    id: 'cells.insertColumnsLeft',
    label: '在左侧插入列',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  insertColumnsRight: {
    id: 'cells.insertColumnsRight',
    label: '在右侧插入列',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  deleteRows: {
    id: 'cells.deleteRows',
    label: '删除所选行',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  deleteColumns: {
    id: 'cells.deleteColumns',
    label: '删除所选列',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
  },
  hideRows: {
    id: 'cells.hideRows',
    label: '隐藏所选行',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+9',
      aria: 'Control+9 Meta+9',
      editor: ['Mod-9'],
    },
  },
  hideColumns: {
    id: 'cells.hideColumns',
    label: '隐藏所选列',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+0',
      aria: 'Control+0 Meta+0',
      editor: ['Mod-0'],
    },
  },
  unhideRows: {
    id: 'cells.unhideRows',
    label: '取消隐藏所选行',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+9',
      aria: 'Control+Shift+9 Meta+Shift+9',
      editor: ['Mod-Shift-9'],
    },
  },
  unhideColumns: {
    id: 'cells.unhideColumns',
    label: '取消隐藏所选列',
    location: { area: 'ribbon', tab: 'home', group: 'cells' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+0',
      aria: 'Control+Shift+0 Meta+Shift+0',
      editor: ['Mod-Shift-0'],
    },
  },
  autoSum: {
    id: 'editing.autoSum',
    label: '自动求和',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Alt+=',
      aria: 'Alt+=',
      editor: ['Alt-equal'],
    },
  },
  autoAverage: {
    id: 'editing.autoAverage',
    label: '平均值',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  autoCount: {
    id: 'editing.autoCount',
    label: '计数',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  autoMaximum: {
    id: 'editing.autoMaximum',
    label: '最大值',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  autoMinimum: {
    id: 'editing.autoMinimum',
    label: '最小值',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  fillDown: {
    id: 'editing.fillDown',
    label: '向下填充',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+D',
      aria: 'Control+D Meta+D',
      editor: ['Mod-d'],
    },
  },
  fillRight: {
    id: 'editing.fillRight',
    label: '向右填充',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+R',
      aria: 'Control+R Meta+R',
      editor: ['Mod-r'],
    },
  },
  fillUp: {
    id: 'editing.fillUp',
    label: '向上填充',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  fillLeft: {
    id: 'editing.fillLeft',
    label: '向左填充',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  copyFormulaFromAbove: {
    id: 'editing.copyFormulaFromAbove',
    label: '复制上方公式',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: "Ctrl+'",
      aria: "Control+'",
      editor: ["Control-'"],
    },
  },
  copyValueFromAbove: {
    id: 'editing.copyValueFromAbove',
    label: '复制上方值',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: "Ctrl+Shift+'",
      aria: "Control+Shift+'",
      editor: ["Control-Shift-'"],
    },
  },
  clearAll: {
    id: 'editing.clearAll',
    label: '清除全部',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  clearFormats: {
    id: 'editing.clearFormats',
    label: '清除格式',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  clearContents: {
    id: 'editing.clearContents',
    label: '清除内容',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Delete/Backspace',
      aria: 'Delete Backspace',
      editor: ['Delete', 'Backspace'],
    },
  },
  clearComments: {
    id: 'editing.clearComments',
    label: '清除批注',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  clearHyperlinks: {
    id: 'editing.clearHyperlinks',
    label: '清除超链接',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  find: {
    id: 'editing.find',
    label: '查找',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+F',
      aria: 'Control+F Meta+F',
      editor: ['Mod-f'],
    },
  },
  findAndSelect: {
    id: 'editing.findAndSelect',
    label: '查找和选择',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
  },
  goTo: {
    id: 'editing.goTo',
    label: '定位',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Ctrl+G 或 F5',
      aria: 'Control+G F5',
      editor: ['Control-g', 'F5'],
    },
  },
  insertChart: {
    id: 'insert.chart',
    label: '插入图表',
    location: { area: 'ribbon', tab: 'insert', group: 'charts' },
  },
  table: {
    id: 'insert.table',
    label: '表格',
    location: { area: 'ribbon', tab: 'insert', group: 'tables' },
    shortcut: {
      label: 'Cmd/Ctrl+T',
      aria: 'Control+T Meta+T',
      editor: ['Mod-t'],
    },
  },
  hyperlink: {
    id: 'insert.hyperlink',
    label: '超链接',
    location: { area: 'ribbon', tab: 'insert', group: 'links' },
    shortcut: {
      label: 'Cmd/Ctrl+K',
      aria: 'Control+K Meta+K',
      editor: ['Mod-k'],
    },
  },
  printSettings: {
    id: 'pageLayout.printSettings',
    label: '打印设置',
    location: { area: 'ribbon', tab: 'pageLayout', group: 'pageSetup' },
  },
  nameManager: {
    id: 'formulas.nameManager',
    label: '名称管理器',
    location: { area: 'ribbon', tab: 'formulas', group: 'definedNames' },
  },
  formulaManager: {
    id: 'formulas.manager',
    label: '公式与计算',
    location: { area: 'ribbon', tab: 'formulas', group: 'calculation' },
  },
  recalculateWorkbook: {
    id: 'formulas.recalculateWorkbook',
    label: '重新计算工作簿',
    location: { area: 'ribbon', tab: 'formulas', group: 'calculation' },
    shortcut: { label: 'F9', aria: 'F9', editor: ['F9'] },
  },
  sortAscending: {
    id: 'data.sortAscending',
    label: '升序',
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
  },
  sortDescending: {
    id: 'data.sortDescending',
    label: '降序',
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
  },
  customSort: {
    id: 'data.customSort',
    label: '自定义排序',
    location: { area: 'ribbon', tab: 'data', group: 'sortAndFilter' },
  },
  autoFilter: {
    id: 'data.autoFilter',
    label: '自动筛选',
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
    label: '数据验证',
    location: { area: 'ribbon', tab: 'data', group: 'dataTools' },
  },
  pivotTable: {
    id: 'data.pivotTable',
    label: '数据透视表',
    location: { area: 'ribbon', tab: 'data', group: 'analysis' },
  },
  protectSheet: {
    id: 'review.protectSheet',
    label: '工作表保护',
    location: { area: 'ribbon', tab: 'review', group: 'protection' },
  },
  gridLines: {
    id: 'view.gridLines',
    label: '网格线',
    location: { area: 'ribbon', tab: 'view', group: 'workbookViews' },
  },
  freezePanes: {
    id: 'view.freezePanes',
    label: '冻结窗格',
    location: { area: 'ribbon', tab: 'view', group: 'window' },
  },
} as const satisfies Record<string, SpreadsheetCommandDefinition>;

export type SpreadsheetCommandId = keyof typeof spreadsheetCommandCatalog;

export function getSpreadsheetCommandDefinition<T extends SpreadsheetCommandId>(
  id: T,
): SpreadsheetCommandDefinition {
  return spreadsheetCommandCatalog[id];
}
