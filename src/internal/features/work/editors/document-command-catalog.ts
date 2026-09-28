import { documentHiddenTextKeyboardShortcut } from '../work-document-hidden-text';
import { documentTextCaseKeyboardShortcuts } from '../work-document-text-case';
import { documentUnderlineKeyboardShortcuts } from '../work-document-underline';

export const documentRibbonTabs = [
  { id: 'home' },
  { id: 'insert' },
  { id: 'page', hasCompactLabel: true },
  { id: 'references' },
  { id: 'review' },
  { id: 'view' },
] as const;

export const documentPictureRibbonTab = {
  id: 'picture',
  contextual: true,
} as const;

export const documentTextBoxRibbonTab = {
  id: 'textBox',
  contextual: true,
} as const;

export const documentConnectorRibbonTab = {
  id: 'connector',
  contextual: true,
} as const;

export const documentTableRibbonTabs = [
  {
    id: 'tableDesign',
    hasCompactLabel: true,
    contextual: true,
  },
  {
    id: 'tableLayout',
    hasCompactLabel: true,
    contextual: true,
  },
] as const;

export const documentPageChromeRibbonTab = {
  id: 'pageChrome',
  hasCompactLabel: true,
  contextual: true,
} as const;

export type DocumentStandardRibbonTabId =
  (typeof documentRibbonTabs)[number]['id'];

export type DocumentRibbonTabId =
  | DocumentStandardRibbonTabId
  | typeof documentPictureRibbonTab.id
  | typeof documentTextBoxRibbonTab.id
  | typeof documentConnectorRibbonTab.id
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
  location: DocumentCommandLocation;
  shortcut?: DocumentCommandShortcut;
}

export const documentCommandCatalog = {
  undo: {
    id: 'history.undo',
    location: { area: 'quickAccess' },
    shortcut: {
      label: 'Cmd/Ctrl+Z',
      aria: 'Control+Z Meta+Z',
      editor: ['Mod-z'],
    },
  },
  redo: {
    id: 'history.redo',
    location: { area: 'quickAccess' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+Z or Cmd/Ctrl+Y',
      aria: 'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
      editor: ['Mod-Shift-z', 'Mod-y'],
    },
  },
  growFont: {
    id: 'font.grow',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+. or Cmd/Ctrl+]',
      aria: 'Control+Shift+. Meta+Shift+. Control+] Meta+]',
      editor: ['Mod-Shift-.', 'Mod-]'],
    },
  },
  shrinkFont: {
    id: 'font.shrink',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+, or Cmd/Ctrl+[',
      aria: 'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
      editor: ['Mod-Shift-,', 'Mod-['],
    },
  },
  copyFormat: {
    id: 'clipboard.copyFormat',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+C',
      aria: 'Control+Shift+C Meta+Shift+C',
      editor: ['Mod-Shift-c'],
    },
  },
  pasteFormat: {
    id: 'clipboard.pasteFormat',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+V',
      aria: 'Control+Shift+V Meta+Shift+V',
      editor: ['Mod-Shift-v'],
    },
  },
  fontDialog: {
    id: 'font.dialog',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+D',
      aria: 'Control+D Meta+D',
      editor: ['Mod-d'],
    },
  },
  formatPainter: {
    id: 'clipboard.formatPainter',
    location: { area: 'ribbon', tab: 'home', group: 'clipboard' },
  },
  bold: {
    id: 'font.bold',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+B',
      aria: 'Control+B Meta+B',
      editor: ['Mod-b'],
    },
  },
  italic: {
    id: 'font.italic',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+I',
      aria: 'Control+I Meta+I',
      editor: ['Mod-i'],
    },
  },
  underline: {
    id: 'font.underline',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+U',
      aria: 'Control+U Meta+U',
      editor: ['Mod-u'],
    },
  },
  doubleUnderline: {
    id: 'font.doubleUnderline',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+D',
      aria: 'Control+Shift+D Meta+Shift+D',
      editor: [documentUnderlineKeyboardShortcuts.double],
    },
  },
  wordsUnderline: {
    id: 'font.wordsUnderline',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+W',
      aria: 'Control+Shift+W Meta+Shift+W',
      editor: [documentUnderlineKeyboardShortcuts.words],
    },
  },
  strike: {
    id: 'font.strike',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  doubleStrike: {
    id: 'font.doubleStrike',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+X',
      aria: 'Control+Shift+X Meta+Shift+X',
      editor: ['Mod-Shift-x'],
    },
  },
  allCaps: {
    id: 'font.allCaps',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+A',
      aria: 'Control+Shift+A Meta+Shift+A',
      editor: [documentTextCaseKeyboardShortcuts.allCaps],
    },
  },
  smallCaps: {
    id: 'font.smallCaps',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+K',
      aria: 'Control+Shift+K Meta+Shift+K',
      editor: [documentTextCaseKeyboardShortcuts.smallCaps],
    },
  },
  changeCase: {
    id: 'font.changeCase',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Shift+F3',
      aria: 'Shift+F3',
      editor: [documentTextCaseKeyboardShortcuts.changeCase],
    },
  },
  subscript: {
    id: 'font.subscript',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+=',
      aria: 'Control+= Meta+=',
      editor: ['Mod-='],
    },
  },
  superscript: {
    id: 'font.superscript',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+=',
      aria: 'Control+Shift+= Meta+Shift+=',
      editor: ['Mod-Shift-='],
    },
  },
  highlight: {
    id: 'font.highlight',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  hiddenText: {
    id: 'font.hiddenText',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+H',
      aria: 'Control+Shift+H Meta+Shift+H',
      editor: [documentHiddenTextKeyboardShortcut],
    },
  },
  runBorder: {
    id: 'font.runBorder',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
  },
  clearFormatting: {
    id: 'font.clearFormatting',
    location: { area: 'ribbon', tab: 'home', group: 'font' },
    shortcut: {
      label: 'Cmd/Ctrl+Space',
      aria: 'Control+Space Meta+Space',
      editor: ['Mod-Space'],
    },
  },
  increaseIndent: {
    id: 'paragraph.increaseIndent',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+M or Shift+Alt+. or Alt+Shift+→',
      aria: 'Control+M Meta+M Shift+Alt+. Alt+Shift+ArrowRight',
      editor: ['Mod-m', 'Shift-Alt-.', 'Alt-Shift-ArrowRight'],
    },
  },
  decreaseIndent: {
    id: 'paragraph.decreaseIndent',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+M or Shift+Alt+, or Alt+Shift+←',
      aria: 'Control+Shift+M Meta+Shift+M Shift+Alt+, Alt+Shift+ArrowLeft',
      editor: ['Mod-Shift-m', 'Shift-Alt-,', 'Alt-Shift-ArrowLeft'],
    },
  },
  bulletList: {
    id: 'paragraph.bulletList',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+L',
      aria: 'Control+Shift+L Meta+Shift+L',
      editor: ['Mod-Shift-l'],
    },
  },
  moveBlockUp: {
    id: 'paragraph.moveBlockUp',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Alt+Shift+↑',
      aria: 'Alt+Shift+ArrowUp',
      editor: ['Alt-Shift-ArrowUp'],
    },
  },
  moveBlockDown: {
    id: 'paragraph.moveBlockDown',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Alt+Shift+↓',
      aria: 'Alt+Shift+ArrowDown',
      editor: ['Alt-Shift-ArrowDown'],
    },
  },
  alignLeft: {
    id: 'paragraph.alignLeft',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+L',
      aria: 'Control+L Meta+L',
      editor: ['Mod-l'],
    },
  },
  alignCenter: {
    id: 'paragraph.alignCenter',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+E',
      aria: 'Control+E Meta+E',
      editor: ['Mod-e'],
    },
  },
  alignRight: {
    id: 'paragraph.alignRight',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+R',
      aria: 'Control+R Meta+R',
      editor: ['Mod-r'],
    },
  },
  alignJustify: {
    id: 'paragraph.alignJustify',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+J',
      aria: 'Control+J Meta+J',
      editor: ['Mod-j'],
    },
  },
  alignDistribute: {
    id: 'paragraph.alignDistribute',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+J',
      aria: 'Control+Shift+J Meta+Shift+J',
      editor: ['Mod-Shift-j'],
    },
  },
  lineSpacingSingle: {
    id: 'paragraph.lineSpacingSingle',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+1',
      aria: 'Control+1 Meta+1',
      editor: ['Mod-1'],
    },
  },
  lineSpacingOneAndHalf: {
    id: 'paragraph.lineSpacingOneAndHalf',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+5',
      aria: 'Control+5 Meta+5',
      editor: ['Mod-5'],
    },
  },
  lineSpacingDouble: {
    id: 'paragraph.lineSpacingDouble',
    location: { area: 'ribbon', tab: 'home', group: 'paragraph' },
    shortcut: {
      label: 'Cmd/Ctrl+2',
      aria: 'Control+2 Meta+2',
      editor: ['Mod-2'],
    },
  },
  normalStyle: {
    id: 'styles.normal',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+N',
      aria: 'Control+Shift+N Meta+Shift+N',
      editor: ['Mod-Shift-n'],
    },
  },
  heading1: {
    id: 'styles.heading1',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+1',
      aria: 'Control+Alt+1 Meta+Alt+1',
      editor: ['Mod-Alt-1'],
    },
  },
  heading2: {
    id: 'styles.heading2',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+2',
      aria: 'Control+Alt+2 Meta+Alt+2',
      editor: ['Mod-Alt-2'],
    },
  },
  heading3: {
    id: 'styles.heading3',
    location: { area: 'ribbon', tab: 'home', group: 'styles' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+3',
      aria: 'Control+Alt+3 Meta+Alt+3',
      editor: ['Mod-Alt-3'],
    },
  },
  find: {
    id: 'editing.find',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+F',
      aria: 'Control+F Meta+F',
      editor: ['Mod-f'],
    },
  },
  replace: {
    id: 'editing.replace',
    location: { area: 'ribbon', tab: 'home', group: 'editing' },
    shortcut: {
      label: 'Cmd/Ctrl+H',
      aria: 'Control+H Meta+H',
      editor: ['Mod-h'],
    },
  },
  insertPageBreak: {
    id: 'insert.pageBreak',
    location: { area: 'ribbon', tab: 'insert', group: 'pages' },
    shortcut: {
      label: 'Cmd/Ctrl+Enter',
      aria: 'Control+Enter Meta+Enter',
      editor: ['Mod-Enter'],
    },
  },
  insertTextBox: {
    id: 'insert.textBox',
    location: { area: 'ribbon', tab: 'insert', group: 'text' },
  },
  insertConnector: {
    id: 'insert.connector',
    location: { area: 'ribbon', tab: 'insert', group: 'text' },
  },
  insertContentControl: {
    id: 'insert.contentControl',
    location: { area: 'ribbon', tab: 'insert', group: 'text' },
  },
  hyperlink: {
    id: 'insert.hyperlink',
    location: { area: 'ribbon', tab: 'insert', group: 'links' },
    shortcut: {
      label: 'Cmd/Ctrl+K',
      aria: 'Control+K Meta+K',
      editor: ['Mod-k'],
    },
  },
  bookmark: {
    id: 'insert.bookmark',
    location: { area: 'ribbon', tab: 'insert', group: 'links' },
  },
  spelling: {
    id: 'review.spelling',
    location: { area: 'ribbon', tab: 'review', group: 'proofing' },
    shortcut: { label: 'F7', aria: 'F7', editor: ['F7'] },
  },
  insertComment: {
    id: 'review.insertComment',
    location: { area: 'ribbon', tab: 'review', group: 'comments' },
    shortcut: {
      label: 'Cmd/Ctrl+Alt+M',
      aria: 'Control+Alt+M Meta+Alt+M',
      editor: ['Mod-Alt-m'],
    },
  },
  trackChanges: {
    id: 'review.trackChanges',
    location: { area: 'ribbon', tab: 'review', group: 'tracking' },
    shortcut: {
      label: 'Cmd/Ctrl+Shift+E',
      aria: 'Control+Shift+E Meta+Shift+E',
      editor: ['Mod-Shift-e'],
    },
  },
  tableOfContents: {
    id: 'references.tableOfContents',
    location: { area: 'ribbon', tab: 'references', group: 'tableOfContents' },
  },
  refreshTableOfContents: {
    id: 'references.refreshTableOfContents',
    location: { area: 'ribbon', tab: 'references', group: 'tableOfContents' },
  },
  markIndexEntry: {
    id: 'references.markIndexEntry',
    location: { area: 'ribbon', tab: 'references', group: 'index' },
  },
  index: {
    id: 'references.index',
    location: { area: 'ribbon', tab: 'references', group: 'index' },
  },
  refreshIndex: {
    id: 'references.refreshIndex',
    location: { area: 'ribbon', tab: 'references', group: 'index' },
  },
  refreshFields: {
    id: 'references.refreshFields',
    location: { area: 'ribbon', tab: 'references', group: 'update' },
    shortcut: { label: 'F9', aria: 'F9', editor: ['F9'] },
  },
  unlinkFields: {
    id: 'references.unlinkFields',
    location: { area: 'ribbon', tab: 'references', group: 'update' },
    shortcut: {
      label: 'Ctrl+Shift+F9',
      aria: 'Control+Shift+F9 Meta+Shift+F9',
      editor: ['Mod-Shift-F9'],
    },
  },
  lockFields: {
    id: 'references.lockFields',
    location: { area: 'ribbon', tab: 'references', group: 'update' },
    shortcut: {
      label: 'Ctrl+F11',
      aria: 'Control+F11 Meta+F11',
      editor: ['Mod-F11'],
    },
  },
  unlockFields: {
    id: 'references.unlockFields',
    location: { area: 'ribbon', tab: 'references', group: 'update' },
    shortcut: {
      label: 'Ctrl+Shift+F11',
      aria: 'Control+Shift+F11 Meta+Shift+F11',
      editor: ['Mod-Shift-F11'],
    },
  },
  toggleFieldCodes: {
    id: 'view.toggleFieldCodes',
    location: { area: 'ribbon', tab: 'view', group: 'show' },
    shortcut: {
      label: 'Alt+F9 / Shift+F9',
      aria: 'Alt+F9 Shift+F9',
      editor: ['Alt-F9', 'Shift-F9'],
    },
  },
  navigationPane: {
    id: 'view.navigationPane',
    location: { area: 'ribbon', tab: 'view', group: 'show' },
  },
  showHiddenText: {
    id: 'view.showHiddenText',
    location: { area: 'ribbon', tab: 'view', group: 'show' },
  },
  wordCount: {
    id: 'status.wordCount',
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
