import type { Editor } from '@tiptap/core';
import { TableRow } from '@tiptap/extension-table';

export interface DocumentTableRowOptions {
  cantSplit: boolean;
  repeatHeader: boolean;
}

export type DocumentTableRowHeightRule = 'atLeast' | 'exact';

export interface SetDocumentTableRowCommandOptions {
  restoreFocus?: boolean;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTableRow: {
      setDocumentTableRowOptions: (
        options: DocumentTableRowOptions,
        commandOptions?: SetDocumentTableRowCommandOptions,
      ) => ReturnType;
    };
  }
}

export const DocumentTableRow = TableRow.extend({
  addAttributes() {
    return {
      ...(this.parent?.() ?? {}),
      cantSplit: booleanRowAttribute(
        'cantSplit',
        'officeCantSplit',
        'data-office-cant-split',
      ),
      repeatHeader: booleanRowAttribute(
        'repeatHeader',
        'officeRepeatHeader',
        'data-office-repeat-header',
      ),
      rowHeight: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentTableRowHeight(
            element.dataset.officeRowHeight || element.style.height,
          ),
        renderHTML: (attributes: Record<string, unknown>) => {
          const height = normalizeDocumentTableRowHeight(attributes.rowHeight);
          return height === null
            ? {}
            : {
                'data-office-row-height': String(height),
                style: `height: ${height}px`,
              };
        },
      },
      rowHeightRule: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentTableRowHeightRule(
            element.dataset.officeRowHeightRule,
          ),
        renderHTML: (attributes: Record<string, unknown>) => {
          const rule = normalizeDocumentTableRowHeightRule(
            attributes.rowHeightRule,
          );
          return rule === null ? {} : { 'data-office-row-height-rule': rule };
        },
      },
    };
  },

  addCommands() {
    return {
      ...(this.parent?.() ?? {}),
      setDocumentTableRowOptions:
        (options, commandOptions = {}) =>
        ({ chain, editor }) => {
          if (!editor.isActive('table')) return false;
          let commandChain = chain();
          if (commandOptions.restoreFocus !== false) {
            commandChain = commandChain.focus();
          }
          return commandChain
            .updateAttributes('tableRow', {
              cantSplit: Boolean(options.cantSplit),
              repeatHeader: Boolean(options.repeatHeader),
            })
            .run();
        },
    };
  },
});

export function documentTableRowOptions(
  editor: Editor,
): DocumentTableRowOptions {
  const attributes = editor.getAttributes('tableRow');
  return {
    cantSplit: directBoolean(attributes.cantSplit) ?? false,
    repeatHeader:
      directBoolean(attributes.repeatHeader) ?? editor.isActive('tableHeader'),
  };
}

export function canSetDocumentTableRowRepeatHeader(editor: Editor): boolean {
  const resolved = editor.state.selection.$from;
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    if (resolved.node(depth).type.name !== 'tableRow') continue;
    const tableDepth = depth - 1;
    const table = resolved.node(tableDepth);
    if (table.type.name !== 'table') return false;
    let leadingHeaderCount = 0;
    for (let index = 0; index < table.childCount; index += 1) {
      if (!documentTableRowRepeats(table.child(index))) break;
      leadingHeaderCount += 1;
    }
    return resolved.index(tableDepth) <= leadingHeaderCount;
  }
  return false;
}

export function documentTableRowRepeats(node: {
  attrs: Record<string, unknown>;
  childCount: number;
  child(index: number): { type: { name: string } };
}): boolean {
  const explicit = directBoolean(node.attrs.repeatHeader);
  if (explicit !== null) return explicit;
  if (node.childCount === 0) return false;
  for (let index = 0; index < node.childCount; index += 1) {
    if (node.child(index).type.name !== 'tableHeader') return false;
  }
  return true;
}

export function documentTableRowCantSplit(node: {
  attrs: Record<string, unknown>;
}): boolean {
  return directBoolean(node.attrs.cantSplit) ?? false;
}

export function documentTableRowHeight(attributes: Record<string, unknown>): {
  height: number | null;
  rule: DocumentTableRowHeightRule | null;
} {
  const height = normalizeDocumentTableRowHeight(attributes.rowHeight);
  return {
    height,
    rule:
      height === null
        ? null
        : (normalizeDocumentTableRowHeightRule(attributes.rowHeightRule) ??
          'atLeast'),
  };
}

export function normalizeDocumentTableRowHeight(value: unknown): number | null {
  const height = Number.parseFloat(String(value ?? '').replace(/px$/i, ''));
  return Number.isFinite(height) && height > 0
    ? Math.round(height * 100) / 100
    : null;
}

export function normalizeDocumentTableRowHeightRule(
  value: unknown,
): DocumentTableRowHeightRule | null {
  return value === 'exact' || value === 'atLeast' ? value : null;
}

function booleanRowAttribute(
  modelKey: 'cantSplit' | 'repeatHeader',
  datasetKey: 'officeCantSplit' | 'officeRepeatHeader',
  htmlName: 'data-office-cant-split' | 'data-office-repeat-header',
) {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      directBoolean(element.dataset[datasetKey]),
    renderHTML: (attributes: Record<string, unknown>) => {
      const value = directBoolean(attributes[modelKey]);
      return value === null ? {} : { [htmlName]: String(value) };
    },
  };
}

function directBoolean(value: unknown): boolean | null {
  if (value === true || value === 'true' || value === '1') return true;
  if (value === false || value === 'false' || value === '0') return false;
  return null;
}
