import { Extension } from '@tiptap/core';
import { isHistoryTransaction } from '@tiptap/pm/history';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { type EditorState, Plugin, type Transaction } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import {
  createDocumentParagraphIdentity,
  createDocumentParagraphIdentityRegistry,
  normalizeDocumentParagraphId,
  uniqueDocumentParagraphIdentity,
} from './work-document-paragraph-identity';

export interface WorkDocumentTableRowIdentity {
  rowId: string;
  rowTextId: string;
}

export interface WorkDocumentTableRowIdentitySource {
  rowId?: unknown;
  rowTextId?: unknown;
}

interface DocumentTableRowAtPosition {
  node: ProseMirrorNode;
  position: number;
}

export const DOCUMENT_TABLE_ROW_ID_ATTRIBUTE = 'data-office-row-id';
export const DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE = 'data-office-row-text-id';

export const DocumentTableRowIdentity = Extension.create({
  name: 'documentTableRowIdentity',

  addGlobalAttributes() {
    return [
      {
        types: ['tableRow'],
        attributes: {
          rowId: identityAttribute(DOCUMENT_TABLE_ROW_ID_ATTRIBUTE),
          rowTextId: identityAttribute(DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE),
        },
      },
    ];
  },

  addProseMirrorPlugins() {
    return [createDocumentTableRowIdentityPlugin()];
  },
});

export function normalizeDocumentTableRowIdentity(
  source: WorkDocumentTableRowIdentitySource,
): WorkDocumentTableRowIdentity | null {
  const rowId = normalizeDocumentParagraphId(source.rowId);
  const rowTextId = normalizeDocumentParagraphId(source.rowTextId);
  return rowId && rowTextId ? { rowId, rowTextId } : null;
}

export function documentTableRowIdentityFromElement(
  element: Element,
): WorkDocumentTableRowIdentity | null {
  return normalizeDocumentTableRowIdentity({
    rowId: element.getAttribute(DOCUMENT_TABLE_ROW_ID_ATTRIBUTE),
    rowTextId: element.getAttribute(DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE),
  });
}

export function applyDocumentTableRowIdentityToElement(
  element: Element,
  source: WorkDocumentTableRowIdentitySource,
): WorkDocumentTableRowIdentity | null {
  const identity = normalizeDocumentTableRowIdentity(source);
  if (!identity) {
    element.removeAttribute(DOCUMENT_TABLE_ROW_ID_ATTRIBUTE);
    element.removeAttribute(DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE);
    return null;
  }
  element.setAttribute(DOCUMENT_TABLE_ROW_ID_ATTRIBUTE, identity.rowId);
  element.setAttribute(
    DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE,
    identity.rowTextId,
  );
  return identity;
}

function identityAttribute(htmlName: string) {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      normalizeDocumentParagraphId(element.getAttribute(htmlName)),
    renderHTML: (attributes: Record<string, unknown>) => {
      const modelName =
        htmlName === DOCUMENT_TABLE_ROW_ID_ATTRIBUTE ? 'rowId' : 'rowTextId';
      const value = normalizeDocumentParagraphId(attributes[modelName]);
      return value ? { [htmlName]: value } : {};
    },
  };
}

function createDocumentTableRowIdentityPlugin(): Plugin {
  return new Plugin({
    view(view) {
      const transaction = normalizeDocumentTableRowIdentities(view.state);
      if (transaction) {
        transaction.setMeta('addToHistory', false);
        view.dispatch(transaction);
      }
      return {};
    },
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      return normalizeDocumentTableRowIdentities(
        newState,
        oldState,
        transactions,
      );
    },
  });
}

function normalizeDocumentTableRowIdentities(
  state: EditorState,
  oldState?: EditorState,
  transactions: readonly Transaction[] = [],
): Transaction | null {
  const rows = documentTableRows(state.doc);
  const identifiedRows = rows.filter(({ node }) =>
    hasDocumentTableRowIdentityComponent(node),
  );
  if (!identifiedRows.length) return null;
  const retainedPositions = oldState
    ? retainedDocumentTableRowPositions(oldState, state, transactions)
    : new Set<number>();
  const editedPositions =
    oldState && !transactions.some(isHistoryTransaction)
      ? editedDocumentTableRowPositions(oldState, state, transactions)
      : new Set<number>();
  const ordered = [
    ...identifiedRows.filter((item) => retainedPositions.has(item.position)),
    ...identifiedRows.filter((item) => !retainedPositions.has(item.position)),
  ];
  const registry = createDocumentParagraphIdentityRegistry();
  const updates = new Map<number, WorkDocumentTableRowIdentity>();
  for (const row of ordered) {
    const current = tableRowIdentitySourceFromNode(row.node);
    const source = editedPositions.has(row.position)
      ? {
          rowId: current.rowId,
          rowTextId: createDocumentParagraphIdentity().textId,
        }
      : current;
    const wordIdentity = uniqueDocumentParagraphIdentity(
      { paragraphId: source.rowId, textId: source.rowTextId },
      registry,
    );
    const identity = {
      rowId: wordIdentity.paragraphId,
      rowTextId: wordIdentity.textId,
    };
    if (!sameDocumentTableRowIdentity(current, identity)) {
      updates.set(row.position, identity);
    }
  }
  if (!updates.size) return null;
  const transaction = state.tr;
  for (const row of rows) {
    const identity = updates.get(row.position);
    if (!identity) continue;
    transaction.setNodeMarkup(row.position, undefined, {
      ...row.node.attrs,
      ...identity,
    });
  }
  if (!transaction.docChanged) return null;
  transaction.setMeta('addToHistory', false);
  return transaction;
}

function hasDocumentTableRowIdentityComponent(node: ProseMirrorNode): boolean {
  return Boolean(
    normalizeDocumentParagraphId(node.attrs.rowId) ||
      normalizeDocumentParagraphId(node.attrs.rowTextId),
  );
}

function retainedDocumentTableRowPositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
): Set<number> {
  const mapping = transactionMapping(transactions);
  const retained = new Set<number>();
  for (const previous of documentTableRows(oldState.doc)) {
    const rowId = normalizeDocumentParagraphId(previous.node.attrs.rowId);
    if (!rowId) continue;
    const mapped = mapping.mapResult(previous.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current?.type.name === 'tableRow' &&
      normalizeDocumentParagraphId(current.attrs.rowId) === rowId
    ) {
      retained.add(mapped.pos);
    }
  }
  return retained;
}

function editedDocumentTableRowPositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
): Set<number> {
  const mapping = transactionMapping(transactions);
  const previousRows = documentTableRows(oldState.doc);
  const currentRows = documentTableRows(newState.doc);
  const previousById = tableRowsById(previousRows);
  const currentById = tableRowsById(currentRows);
  const edited = new Set<number>();
  for (const previous of previousRows) {
    const identity = normalizeDocumentTableRowIdentity(
      tableRowIdentitySourceFromNode(previous.node),
    );
    if (!identity) continue;
    const mapped = mapping.mapResult(previous.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current?.type.name === 'tableRow' &&
      normalizeDocumentParagraphId(current.attrs.rowId) === identity.rowId
    ) {
      if (tableRowContentChanged(previous.node, current, identity.rowTextId)) {
        edited.add(mapped.pos);
      }
      continue;
    }
    const previousMatches = previousById.get(identity.rowId) ?? [];
    const currentMatches = currentById.get(identity.rowId) ?? [];
    if (
      previousMatches.length === 1 &&
      currentMatches.length === 1 &&
      tableRowContentChanged(
        previous.node,
        currentMatches[0].node,
        identity.rowTextId,
      )
    ) {
      edited.add(currentMatches[0].position);
    }
  }
  return edited;
}

function tableRowContentChanged(
  previous: ProseMirrorNode,
  current: ProseMirrorNode,
  previousTextId: string,
): boolean {
  return (
    tableRowContentSignature(previous) !== tableRowContentSignature(current) &&
    normalizeDocumentParagraphId(current.attrs.rowTextId) === previousTextId
  );
}

function tableRowContentSignature(row: ProseMirrorNode): string {
  const parts = ['tableRow'];
  row.descendants((node) => {
    if (node.isText) {
      parts.push(`text:${node.text ?? ''}`);
      return;
    }
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      parts.push(
        `${node.type.name}:${String(node.attrs.colspan ?? 1)}:${String(
          node.attrs.rowspan ?? 1,
        )}`,
      );
      return;
    }
    parts.push(node.type.name);
  });
  return parts.join('\u0000');
}

function tableRowsById(
  rows: readonly DocumentTableRowAtPosition[],
): Map<string, DocumentTableRowAtPosition[]> {
  const result = new Map<string, DocumentTableRowAtPosition[]>();
  for (const row of rows) {
    const id = normalizeDocumentParagraphId(row.node.attrs.rowId);
    if (!id) continue;
    const matches = result.get(id) ?? [];
    matches.push(row);
    result.set(id, matches);
  }
  return result;
}

function documentTableRows(
  document: ProseMirrorNode,
): DocumentTableRowAtPosition[] {
  const rows: DocumentTableRowAtPosition[] = [];
  document.descendants((node, position) => {
    if (node.type.name === 'tableRow') rows.push({ node, position });
  });
  return rows;
}

function tableRowIdentitySourceFromNode(
  node: ProseMirrorNode,
): WorkDocumentTableRowIdentitySource {
  return { rowId: node.attrs.rowId, rowTextId: node.attrs.rowTextId };
}

function sameDocumentTableRowIdentity(
  source: WorkDocumentTableRowIdentitySource,
  identity: WorkDocumentTableRowIdentity,
): boolean {
  return (
    normalizeDocumentParagraphId(source.rowId) === identity.rowId &&
    normalizeDocumentParagraphId(source.rowTextId) === identity.rowTextId
  );
}

function transactionMapping(transactions: readonly Transaction[]): Mapping {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  return mapping;
}
