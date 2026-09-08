import type { Editor } from '@tiptap/core';
import { TableRow } from '@tiptap/extension-table';
import {
  normalizeDocumentTableAlignment,
  type DocumentTableAlignment,
} from './work-document-table-geometry';
import {
  DOCUMENT_ROW_PROPERTY_REVISION_OMML_ATTRIBUTE,
  documentRowPropertyRevisionOmmlFromElement,
  encodeDocumentTablePropertyRevisionOmml,
} from './work-document-table-property-revision';

export interface DocumentTableRowOptions {
  cantSplit: boolean;
  repeatHeader: boolean;
  hidden?: boolean;
  alignment?: DocumentTableAlignment;
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
      hidden: booleanRowAttribute(
        'hidden',
        'officeRowHidden',
        'data-office-row-hidden',
      ),
      alignment: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          normalizeDocumentTableAlignment(element.dataset.officeRowAlignment),
        renderHTML: (attributes: Record<string, unknown>) => {
          const alignment = normalizeDocumentTableAlignment(
            attributes.alignment,
          );
          return alignment
            ? { 'data-office-row-alignment': alignment }
            : {};
        },
      },
      gridBefore: {
        default: null,
        parseHTML: (element: HTMLElement) => {
          const raw = element.dataset.officeRowGridBefore;
          if (raw === undefined || raw === '') return null;
          const value = Number(raw);
          return Number.isInteger(value) && value >= 0 ? value : null;
        },
        renderHTML: (attributes: Record<string, unknown>) => {
          const value =
            typeof attributes.gridBefore === 'number' &&
            Number.isInteger(attributes.gridBefore) &&
            attributes.gridBefore >= 0
              ? attributes.gridBefore
              : null;
          return value === null
            ? {}
            : { 'data-office-row-grid-before': String(value) };
        },
      },
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
      propertyRevisionOmml: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentRowPropertyRevisionOmmlFromElement(element),
        renderHTML: (attributes: Record<string, unknown>) => {
          const omml =
            typeof attributes.propertyRevisionOmml === 'string'
              ? attributes.propertyRevisionOmml
              : '';
          const encoded = encodeDocumentTablePropertyRevisionOmml(omml);
          return encoded
            ? { [DOCUMENT_ROW_PROPERTY_REVISION_OMML_ATTRIBUTE]: encoded }
            : {};
        },
      },
      rowChangeKind: rowChangeAttribute('kind', 'data-change-kind'),
      rowChangeId: rowChangeAttribute('id', 'data-change-id'),
      rowChangeAuthor: rowChangeAttribute('author', 'data-change-author'),
      rowChangeDate: rowChangeAttribute('date', 'data-change-date'),
      rowChangeBefore: rowChangeAttribute('before', 'data-change-before'),
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
              ...(options.hidden !== undefined
                ? { hidden: Boolean(options.hidden) }
                : {}),
              ...(options.alignment !== undefined
                ? {
                    alignment:
                      normalizeDocumentTableAlignment(options.alignment) ??
                      null,
                  }
                : {}),
              propertyRevisionOmml: null,
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
    hidden: directBoolean(attributes.hidden) ?? false,
    alignment: normalizeDocumentTableAlignment(attributes.alignment) ?? undefined,
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
  modelKey: 'cantSplit' | 'repeatHeader' | 'hidden',
  datasetKey: 'officeCantSplit' | 'officeRepeatHeader' | 'officeRowHidden',
  htmlName:
    | 'data-office-cant-split'
    | 'data-office-repeat-header'
    | 'data-office-row-hidden',
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

function rowChangeAttribute(
  field: 'kind' | 'id' | 'author' | 'date' | 'before',
  htmlName:
    | 'data-change-kind'
    | 'data-change-id'
    | 'data-change-author'
    | 'data-change-date'
    | 'data-change-before',
) {
  const modelName = `rowChange${field[0]?.toUpperCase() ?? ''}${field.slice(1)}`;
  return {
    default: field === 'kind' ? null : '',
    parseHTML: (element: HTMLElement) => {
      if (
        element.getAttribute('data-document-change') !== 'true' ||
        element.getAttribute('data-change-kind') !== 'row-formatting'
      ) {
        return field === 'kind' ? null : '';
      }
      return field === 'kind'
        ? 'row-formatting'
        : (element.getAttribute(htmlName) ?? '');
    },
    renderHTML: (attributes: Record<string, unknown>) => {
      if (attributes.rowChangeKind !== 'row-formatting') return {};
      if (field === 'kind') {
        return {
          'data-document-change': 'true',
          'data-change-kind': 'row-formatting',
        };
      }
      const value =
        typeof attributes[modelName] === 'string'
          ? attributes[modelName].trim()
          : '';
      return value ? { [htmlName]: value } : {};
    },
  };
}

function directBoolean(value: unknown): boolean | null {
  if (value === true || value === 'true' || value === '1') return true;
  if (value === false || value === 'false' || value === '0') return false;
  return null;
}
