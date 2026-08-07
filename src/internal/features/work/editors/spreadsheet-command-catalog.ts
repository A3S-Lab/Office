export const spreadsheetRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'pageLayout', label: '页面布局', compactLabel: '布局' },
  { id: 'formulas', label: '公式' },
  { id: 'data', label: '数据' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
] as const;

export type SpreadsheetRibbonTabId =
  (typeof spreadsheetRibbonTabs)[number]['id'];

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
  bold: {
    id: 'font.bold',
    label: '加粗',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+B',
      aria: 'Control+B Meta+B',
      editor: ['Mod-b'],
    },
  },
  italic: {
    id: 'font.italic',
    label: '斜体',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+I',
      aria: 'Control+I Meta+I',
      editor: ['Mod-i'],
    },
  },
  underline: {
    id: 'font.underline',
    label: '下划线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+U',
      aria: 'Control+U Meta+U',
      editor: ['Mod-u'],
    },
  },
  conditionalFormatting: {
    id: 'styles.conditionalFormatting',
    label: '条件格式',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
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
  insertChart: {
    id: 'insert.chart',
    label: '插入图表',
    location: { area: 'ribbon', tab: 'insert', group: 'charts' },
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
} as const satisfies Record<string, SpreadsheetCommandDefinition>;

export type SpreadsheetCommandId = keyof typeof spreadsheetCommandCatalog;

export function getSpreadsheetCommandDefinition<T extends SpreadsheetCommandId>(
  id: T,
): SpreadsheetCommandDefinition {
  return spreadsheetCommandCatalog[id];
}
