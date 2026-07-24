import type { Editor } from '@tiptap/core';
import { TableRow } from '@tiptap/extension-table';

export interface DocumentTableRowOptions {
  cantSplit: boolean;
  repeatHeader: boolean;
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

export function setDocumentTableRowOptions(
  editor: Editor,
  options: DocumentTableRowOptions,
  commandOptions: { restoreFocus?: boolean } = {},
): boolean {
  if (!editor.isActive('table')) return false;
  const chain = editor.chain();
  if (commandOptions.restoreFocus !== false) chain.focus();
  return chain
    .updateAttributes('tableRow', {
      cantSplit: Boolean(options.cantSplit),
      repeatHeader: Boolean(options.repeatHeader),
    })
    .run();
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
