import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { type EditorState, Plugin, type Transaction } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import {
  DOCUMENT_INTEGRITY_FIELD,
  documentHasIntegrityFeature,
} from './work-document-integrity-index';
import { createWorkId } from './work-templates';

interface DocumentFieldAtPosition {
  node: ProseMirrorNode;
  position: number;
  id: string;
}

export function createDocumentFieldIdentityPlugin(
  fieldNodeName = 'documentField',
): Plugin {
  return new Plugin({
    view(view) {
      const transaction = normalizeDocumentFieldIdentities(
        view.state,
        fieldNodeName,
      );
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
      return normalizeDocumentFieldIdentities(
        newState,
        fieldNodeName,
        oldState,
        transactions,
      );
    },
  });
}

function normalizeDocumentFieldIdentities(
  state: EditorState,
  fieldNodeName: string,
  oldState?: EditorState,
  transactions: readonly Transaction[] = [],
): Transaction | null {
  if (
    fieldNodeName === 'documentField' &&
    !documentHasIntegrityFeature(state.doc, DOCUMENT_INTEGRITY_FIELD)
  ) {
    return null;
  }
  const fields = documentFields(state.doc, fieldNodeName);
  if (!fields.length) return null;
  const retained = oldState
    ? retainedDocumentFieldPositions(
        oldState,
        state,
        transactions,
        fieldNodeName,
      )
    : new Set<number>();
  const ordered = [
    ...fields.filter((field) => retained.has(field.position)),
    ...fields.filter((field) => !retained.has(field.position)),
  ];
  const usedIds = new Set<string>();
  const updates = new Map<number, string>();
  for (const field of ordered) {
    let id = field.id;
    if (!id || usedIds.has(id)) {
      do id = createWorkId('field');
      while (usedIds.has(id));
    }
    usedIds.add(id);
    if (id !== field.id) updates.set(field.position, id);
  }
  if (!updates.size) return null;
  const transaction = state.tr;
  for (const field of fields) {
    const id = updates.get(field.position);
    if (!id) continue;
    transaction.setNodeMarkup(field.position, undefined, {
      ...field.node.attrs,
      id,
    });
  }
  return transaction.docChanged ? transaction : null;
}

function retainedDocumentFieldPositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  fieldNodeName: string,
): Set<number> {
  const mapping = transactionMapping(transactions);
  const retained = new Set<number>();
  const currentById = new Map<string, DocumentFieldAtPosition[]>();
  for (const field of documentFields(newState.doc, fieldNodeName)) {
    if (!field.id) continue;
    const matches = currentById.get(field.id) ?? [];
    matches.push(field);
    currentById.set(field.id, matches);
  }
  for (const previous of documentFields(oldState.doc, fieldNodeName)) {
    if (!previous.id) continue;
    const mapped = mapping.mapResult(previous.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current?.type.name === fieldNodeName &&
      stringAttribute(current.attrs.id) === previous.id
    ) {
      retained.add(mapped.pos);
      continue;
    }
    const matches = currentById.get(previous.id);
    if (matches?.length === 1) retained.add(matches[0]?.position ?? -1);
  }
  retained.delete(-1);
  return retained;
}

function transactionMapping(transactions: readonly Transaction[]): Mapping {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  return mapping;
}

function documentFields(
  document: ProseMirrorNode,
  fieldNodeName: string,
): DocumentFieldAtPosition[] {
  const fields: DocumentFieldAtPosition[] = [];
  document.descendants((node, position) => {
    if (node.type.name !== fieldNodeName) return;
    fields.push({
      node,
      position,
      id: stringAttribute(node.attrs.id),
    });
  });
  return fields;
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}
