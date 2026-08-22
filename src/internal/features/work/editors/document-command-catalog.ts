import { documentTextCaseKeyboardShortcuts } from '../work-document-text-case';
import { documentUnderlineKeyboardShortcuts } from '../work-document-underline';

export const documentRibbonTabs = [
  { id: 'home', label: '开始' },
  { id: 'insert', label: '插入' },
  { id: 'page', label: '页面布局', compactLabel: '布局' },
  { id: 'references', label: '引用' },
  { id: 'review', label: '审阅' },
  { id: 'view', label: '视图' },
] as const;

export const documentPictureRibbonTab = {
  id: 'picture',
  label: '图片',
  contextual: true,
} as const;

export const documentTableRibbonTabs = [
  {
    id: 'tableDesign',
    label: '表格设计',
    compactLabel: '设计',
    contextual: true,
  },
  {
    id: 'tableLayout',
    label: '表格布局',
    compactLabel: '布局',
    contextual: true,
  },
] as const;

export const documentPageChromeRibbonTab = {
  id: 'pageChrome',
  label: '页眉和页脚',
  compactLabel: '页眉页脚',
  contextual: true,
} as const;

export type DocumentStandardRibbonTabId =
  (typeof documentRibbonTabs)[number]['id'];

export type DocumentRibbonTabId =
  | DocumentStandardRibbonTabId
  | typeof documentPictureRibbonTab.id
  | (typeof documentTableRibbonTabs)[number]['id']
  | typeof documentPageChromeRibbonTab.id;

export interface DocumentCommandShortcut {
  label: string;
  aria: string;
  editor?: readonly string[];
}

export type DocumentCommandLocation =
  | { area: 'quickAccess' }
  | { area: 'status' }
  | {
      area: 'ribbon';
      tab: DocumentStandardRibbonTabId;
      group: string;
    };

export interface DocumentCommandDefinition {
  id: string;
  label: string;
  location: DocumentCommandLocation;
  shortcut?: DocumentCommandShortcut;
}

export const documentCommandCatalog = {
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
  copyFormat: {
    id: 'clipboard.copyFormat',
    label: '复制格式',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+C',
      aria: 'Control+Shift+C Meta+Shift+C',
      editor: ['Mod-Shift-c'],
    },
  },
  pasteFormat: {
    id: 'clipboard.pasteFormat',
    label: '粘贴格式',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+V',
      aria: 'Control+Shift+V Meta+Shift+V',
      editor: ['Mod-Shift-v'],
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
  doubleUnderline: {
    id: 'font.doubleUnderline',
    label: '双下划线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+D',
      aria: 'Control+Shift+D Meta+Shift+D',
      editor: [documentUnderlineKeyboardShortcuts.double],
    },
  },
  wordsUnderline: {
    id: 'font.wordsUnderline',
    label: '仅字下划线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+W',
      aria: 'Control+Shift+W Meta+Shift+W',
      editor: [documentUnderlineKeyboardShortcuts.words],
    },
  },
  strike: {
    id: 'font.strike',
    label: '删除线',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  allCaps: {
    id: 'font.allCaps',
    label: '全部大写',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+A',
      aria: 'Control+Shift+A Meta+Shift+A',
      editor: [documentTextCaseKeyboardShortcuts.allCaps],
    },
  },
  smallCaps: {
    id: 'font.smallCaps',
    label: '小型大写',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+K',
      aria: 'Control+Shift+K Meta+Shift+K',
      editor: [documentTextCaseKeyboardShortcuts.smallCaps],
    },
  },
  subscript: {
    id: 'font.subscript',
    label: '下标',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+=',
      aria: 'Control+= Meta+=',
      editor: ['Mod-='],
    },
  },
  superscript: {
    id: 'font.superscript',
    label: '上标',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+=',
      aria: 'Control+Shift+= Meta+Shift+=',
      editor: ['Mod-Shift-='],
    },
  },
  highlight: {
    id: 'font.highlight',
    label: '突出显示',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  clearFormatting: {
    id: 'font.clearFormatting',
    label: '清除格式',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  alignLeft: {
    id: 'paragraph.alignLeft',
    label: '左对齐',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+L',
      aria: 'Control+L Meta+L',
      editor: ['Mod-l'],
    },
  },
  alignCenter: {
    id: 'paragraph.alignCenter',
    label: '居中',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+E',
      aria: 'Control+E Meta+E',
      editor: ['Mod-e'],
    },
  },
  alignRight: {
    id: 'paragraph.alignRight',
    label: '右对齐',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+R',
      aria: 'Control+R Meta+R',
      editor: ['Mod-r'],
    },
  },
  alignJustify: {
    id: 'paragraph.alignJustify',
    label: '两端对齐',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+J',
      aria: 'Control+J Meta+J',
      editor: ['Mod-j'],
    },
  },
  lineSpacingSingle: {
    id: 'paragraph.lineSpacingSingle',
    label: '单倍行距',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+1',
      aria: 'Control+1 Meta+1',
      editor: ['Mod-1'],
    },
  },
  lineSpacingOneAndHalf: {
    id: 'paragraph.lineSpacingOneAndHalf',
    label: '1.5 倍行距',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+5',
      aria: 'Control+5 Meta+5',
      editor: ['Mod-5'],
    },
  },
  lineSpacingDouble: {
    id: 'paragraph.lineSpacingDouble',
    label: '双倍行距',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+2',
      aria: 'Control+2 Meta+2',
      editor: ['Mod-2'],
    },
  },
  heading1: {
    id: 'styles.heading1',
    label: '标题 1',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+1',
      aria: 'Control+Alt+1 Meta+Alt+1',
      editor: ['Mod-Alt-1'],
    },
  },
  heading2: {
    id: 'styles.heading2',
    label: '标题 2',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+2',
      aria: 'Control+Alt+2 Meta+Alt+2',
      editor: ['Mod-Alt-2'],
    },
  },
  heading3: {
    id: 'styles.heading3',
    label: '标题 3',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+3',
      aria: 'Control+Alt+3 Meta+Alt+3',
      editor: ['Mod-Alt-3'],
    },
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
  replace: {
    id: 'editing.replace',
    label: '替换',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+H',
      aria: 'Control+H Meta+H',
      editor: ['Mod-h'],
    },
  },
  insertPageBreak: {
    id: 'insert.pageBreak',
    label: '插入分页符',
    location: { area: 'ribbon', tab: 'insert', group: 'pages' },
    shortcut: {
      label: 'Cmd/Ctrl+Enter',
      aria: 'Control+Enter Meta+Enter',
      editor: ['Mod-Enter'],
    },
  },
  hyperlink: {
    id: 'insert.hyperlink',
    label: '添加链接',
    location: { area: 'ribbon', tab: 'insert', group: 'links' },
    shortcut: {
      label: 'Cmd/Ctrl+K',
      aria: 'Control+K Meta+K',
      editor: ['Mod-k'],
    },
  },
  bookmark: {
    id: 'insert.bookmark',
    label: '添加书签',
    location: { area: 'ribbon', tab: 'insert', group: 'links' },
  },
  spelling: {
    id: 'review.spelling',
    label: '拼写检查',
    location: { area: 'ribbon', tab: 'review', group: 'proofing' },
    shortcut: { label: 'F7', aria: 'F7', editor: ['F7'] },
  },
  insertComment: {
    id: 'review.insertComment',
    label: '添加批注',
    location: { area: 'ribbon', tab: 'review', group: 'comments' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+M',
      aria: 'Control+Alt+M Meta+Alt+M',
      editor: ['Mod-Alt-m'],
    },
  },
  trackChanges: {
    id: 'review.trackChanges',
    label: '修订模式',
    location: { area: 'ribbon', tab: 'review', group: 'tracking' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+E',
      aria: 'Control+Shift+E Meta+Shift+E',
      editor: ['Mod-Shift-e'],
    },
  },
  refreshFields: {
    id: 'references.refreshFields',
    label: '更新域',
    location: { area: 'ribbon', tab: 'references', group: 'update' },
    shortcut: { label: 'F9', aria: 'F9', editor: ['F9'] },
  },
  navigationPane: {
    id: 'view.navigationPane',
    label: '导航窗格',
    location: { area: 'ribbon', tab: 'view', group: 'show' },
  },
  wordCount: {
    id: 'status.wordCount',
    label: '字数统计',
    location: { area: 'status' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+G',
      aria: 'Control+Shift+G Meta+Shift+G',
      editor: ['Mod-Shift-g'],
    },
  },
} as const satisfies Record<string, DocumentCommandDefinition>;

export type DocumentCommandId = keyof typeof documentCommandCatalog;

export function getDocumentCommandDefinition<T extends DocumentCommandId>(
  id: T,
): DocumentCommandDefinition {
  return documentCommandCatalog[id];
}
