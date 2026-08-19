import { Extension } from '@tiptap/core';
import { isHistoryTransaction } from '@tiptap/pm/history';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { type EditorState, Plugin, type Transaction } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import {
  DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY,
  documentHasIntegrityFeature,
} from './work-document-integrity-index';

export interface WorkDocumentParagraphIdentity {
  paragraphId: string;
  textId: string;
}

export interface WorkDocumentParagraphIdentitySource {
  paragraphId?: unknown;
  textId?: unknown;
}

export interface WorkDocumentParagraphIdentityRegistry {
  paragraphIds: Set<string>;
}

interface DocumentParagraphIdentityOptions {
  rotateTextId: () => boolean;
  types: string[];
}

interface DocumentParagraphAtPosition {
  node: ProseMirrorNode;
  position: number;
}

export const DOCUMENT_PARAGRAPH_ID_ATTRIBUTE = 'data-office-paragraph-id';
export const DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE =
  'data-office-paragraph-text-id';

const PARAGRAPH_ID_PATTERN = /^[0-9A-F]{8}$/;
const MAX_PARAGRAPH_ID = 0x7fff_ffff;
let fallbackIdentitySequence = 0;

export const DocumentParagraphIdentity =
  Extension.create<DocumentParagraphIdentityOptions>({
    name: 'documentParagraphIdentity',

    addOptions() {
      return {
        rotateTextId: () => true,
        types: ['paragraph', 'heading', 'documentCaption'],
      };
    },

    addGlobalAttributes() {
      return [
        {
          types: this.options.types,
          attributes: {
            paragraphId: {
              default: null,
              parseHTML: (element: HTMLElement) =>
                normalizeDocumentParagraphId(
                  element.getAttribute(DOCUMENT_PARAGRAPH_ID_ATTRIBUTE),
                ),
              renderHTML: (attributes: Record<string, unknown>) => {
                const value = normalizeDocumentParagraphId(
                  attributes.paragraphId,
                );
                return value
                  ? { [DOCUMENT_PARAGRAPH_ID_ATTRIBUTE]: value }
                  : {};
              },
            },
            textId: {
              default: null,
              parseHTML: (element: HTMLElement) =>
                normalizeDocumentParagraphId(
                  element.getAttribute(DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE),
                ),
              renderHTML: (attributes: Record<string, unknown>) => {
                const value = normalizeDocumentParagraphId(attributes.textId);
                return value
                  ? { [DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE]: value }
                  : {};
              },
            },
          },
        },
      ];
    },

    addProseMirrorPlugins() {
      return [
        createDocumentParagraphIdentityPlugin(
          this.options.types,
          this.options.rotateTextId,
        ),
      ];
    },
  });

export function normalizeDocumentParagraphId(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim().toUpperCase();
  if (!PARAGRAPH_ID_PATTERN.test(normalized)) return null;
  const number = Number.parseInt(normalized, 16);
  return number > 0 && number <= MAX_PARAGRAPH_ID ? normalized : null;
}

export function createDocumentParagraphIdentityRegistry(): WorkDocumentParagraphIdentityRegistry {
  return { paragraphIds: new Set() };
}

export function normalizeDocumentParagraphIdentity(
  source: WorkDocumentParagraphIdentitySource,
): WorkDocumentParagraphIdentity | null {
  const paragraphId = normalizeDocumentParagraphId(source.paragraphId);
  const textId = normalizeDocumentParagraphId(source.textId);
  return paragraphId && textId ? { paragraphId, textId } : null;
}

export function documentParagraphIdentityFromElement(
  element: Element,
): WorkDocumentParagraphIdentity | null {
  return normalizeDocumentParagraphIdentity({
    paragraphId: element.getAttribute(DOCUMENT_PARAGRAPH_ID_ATTRIBUTE),
    textId: element.getAttribute(DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE),
  });
}

export function applyDocumentParagraphIdentityToElement(
  element: Element,
  source: WorkDocumentParagraphIdentitySource,
): WorkDocumentParagraphIdentity | null {
  const identity = normalizeDocumentParagraphIdentity(source);
  if (!identity) {
    element.removeAttribute(DOCUMENT_PARAGRAPH_ID_ATTRIBUTE);
    element.removeAttribute(DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE);
    return null;
  }
  element.setAttribute(DOCUMENT_PARAGRAPH_ID_ATTRIBUTE, identity.paragraphId);
  element.setAttribute(DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE, identity.textId);
  return identity;
}

function createDocumentParagraphIdentityPlugin(
  types: readonly string[],
  rotateTextId: () => boolean,
): Plugin {
  const trackedTypes = new Set(types);
  return new Plugin({
    view(view) {
      const transaction = normalizeDocumentParagraphIdentities(
        view.state,
        trackedTypes,
        rotateTextId,
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
      return normalizeDocumentParagraphIdentities(
        newState,
        trackedTypes,
        rotateTextId,
        oldState,
        transactions,
      );
    },
  });
}

function normalizeDocumentParagraphIdentities(
  state: EditorState,
  trackedTypes: ReadonlySet<string>,
  rotateTextId: () => boolean,
  oldState?: EditorState,
  transactions: readonly Transaction[] = [],
): Transaction | null {
  if (
    !documentHasIntegrityFeature(
      state.doc,
      DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY,
    )
  ) {
    return null;
  }
  const paragraphs = documentParagraphs(state.doc, trackedTypes);
  if (!paragraphs.length) return null;
  const retainedPositions = oldState
    ? retainedDocumentParagraphPositions(
        oldState,
        state,
        transactions,
        trackedTypes,
      )
    : new Set<number>();
  const editedPositions =
    oldState && rotateTextId() && !transactions.some(isHistoryTransaction)
      ? editedDocumentParagraphPositions(
          oldState,
          state,
          transactions,
          trackedTypes,
        )
      : new Set<number>();
  const ordered = [
    ...paragraphs.filter((item) => retainedPositions.has(item.position)),
    ...paragraphs.filter((item) => !retainedPositions.has(item.position)),
  ];
  const registry = createDocumentParagraphIdentityRegistry();
  const updates = new Map<number, WorkDocumentParagraphIdentity>();
  for (const paragraph of ordered) {
    const current = paragraphIdentitySourceFromNode(paragraph.node);
    const source = editedPositions.has(paragraph.position)
      ? {
          ...current,
          textId: nextDocumentParagraphTextId(current.textId),
        }
      : current;
    const identity = uniqueDocumentParagraphIdentity(source, registry);
    if (!sameDocumentParagraphIdentity(current, identity)) {
      updates.set(paragraph.position, identity);
    }
  }
  if (!updates.size) return null;
  const transaction = state.tr;
  for (const paragraph of paragraphs) {
    const identity = updates.get(paragraph.position);
    if (!identity) continue;
    transaction.setNodeMarkup(paragraph.position, undefined, {
      ...paragraph.node.attrs,
      ...identity,
    });
  }
  if (!transaction.docChanged) return null;
  transaction.setMeta('addToHistory', false);
  return transaction;
}

function hasDocumentParagraphIdentityComponent(node: ProseMirrorNode): boolean {
  return Boolean(
    normalizeDocumentParagraphId(node.attrs.paragraphId) ||
      normalizeDocumentParagraphId(node.attrs.textId),
  );
}

function retainedDocumentParagraphPositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  trackedTypes: ReadonlySet<string>,
): Set<number> {
  const mapping = transactionMapping(transactions);
  const retained = new Set<number>();
  for (const previous of documentParagraphs(oldState.doc, trackedTypes)) {
    const paragraphId = normalizeDocumentParagraphId(
      previous.node.attrs.paragraphId,
    );
    if (!paragraphId) continue;
    const mapped = mapping.mapResult(previous.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current &&
      trackedTypes.has(current.type.name) &&
      normalizeDocumentParagraphId(current.attrs.paragraphId) === paragraphId
    ) {
      retained.add(mapped.pos);
    }
  }
  return retained;
}

function editedDocumentParagraphPositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  trackedTypes: ReadonlySet<string>,
): Set<number> {
  const mapping = transactionMapping(transactions);
  const previousParagraphs = documentParagraphs(oldState.doc, trackedTypes);
  const currentParagraphs = documentParagraphs(newState.doc, trackedTypes);
  const previousById = paragraphsById(previousParagraphs);
  const currentById = paragraphsById(currentParagraphs);
  const edited = new Set<number>();
  for (const previous of previousParagraphs) {
    const identity = normalizeDocumentParagraphIdentity(
      paragraphIdentitySourceFromNode(previous.node),
    );
    if (!identity) continue;
    const mapped = mapping.mapResult(previous.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current &&
      trackedTypes.has(current.type.name) &&
      normalizeDocumentParagraphId(current.attrs.paragraphId) ===
        identity.paragraphId
    ) {
      if (paragraphTextChanged(previous.node, current, identity.textId)) {
        edited.add(mapped.pos);
      }
      continue;
    }
    const previousMatches = previousById.get(identity.paragraphId) ?? [];
    const currentMatches = currentById.get(identity.paragraphId) ?? [];
    if (
      previousMatches.length === 1 &&
      currentMatches.length === 1 &&
      paragraphTextChanged(
        previous.node,
        currentMatches[0].node,
        identity.textId,
      )
    ) {
      edited.add(currentMatches[0].position);
    }
  }
  return edited;
}

function paragraphTextChanged(
  previous: ProseMirrorNode,
  current: ProseMirrorNode,
  previousTextId: string,
): boolean {
  return (
    previous.textContent !== current.textContent &&
    normalizeDocumentParagraphId(current.attrs.textId) === previousTextId
  );
}

function paragraphsById(
  paragraphs: readonly DocumentParagraphAtPosition[],
): Map<string, DocumentParagraphAtPosition[]> {
  const result = new Map<string, DocumentParagraphAtPosition[]>();
  for (const paragraph of paragraphs) {
    const id = normalizeDocumentParagraphId(paragraph.node.attrs.paragraphId);
    if (!id) continue;
    const matches = result.get(id) ?? [];
    matches.push(paragraph);
    result.set(id, matches);
  }
  return result;
}

function documentParagraphs(
  document: ProseMirrorNode,
  trackedTypes: ReadonlySet<string>,
): DocumentParagraphAtPosition[] {
  const paragraphs: DocumentParagraphAtPosition[] = [];
  document.descendants((node, position) => {
    if (
      trackedTypes.has(node.type.name) &&
      hasDocumentParagraphIdentityComponent(node)
    ) {
      paragraphs.push({ node, position });
    }
  });
  return paragraphs;
}

function paragraphIdentitySourceFromNode(
  node: ProseMirrorNode,
): WorkDocumentParagraphIdentitySource {
  return {
    paragraphId: node.attrs.paragraphId,
    textId: node.attrs.textId,
  };
}

export function uniqueDocumentParagraphIdentity(
  source: WorkDocumentParagraphIdentitySource,
  registry: WorkDocumentParagraphIdentityRegistry,
): WorkDocumentParagraphIdentity {
  const preferred = normalizeDocumentParagraphIdentity(source);
  if (preferred && !registry.paragraphIds.has(preferred.paragraphId)) {
    registry.paragraphIds.add(preferred.paragraphId);
    return preferred;
  }
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const identity = createDocumentParagraphIdentity();
    if (registry.paragraphIds.has(identity.paragraphId)) continue;
    registry.paragraphIds.add(identity.paragraphId);
    return identity;
  }
  return sequentialDocumentParagraphIdentity(registry);
}

export function createDocumentParagraphIdentity(): WorkDocumentParagraphIdentity {
  return {
    paragraphId: randomDocumentParagraphId(),
    textId: randomDocumentParagraphId(),
  };
}

function sequentialDocumentParagraphIdentity(
  registry: WorkDocumentParagraphIdentityRegistry,
): WorkDocumentParagraphIdentity {
  for (let value = 1; value <= MAX_PARAGRAPH_ID; value += 1) {
    const paragraphId = formatDocumentParagraphId(value);
    if (registry.paragraphIds.has(paragraphId)) continue;
    registry.paragraphIds.add(paragraphId);
    return { paragraphId, textId: randomDocumentParagraphId() };
  }
  throw new Error('No unique Word paragraph identity is available.');
}

function nextDocumentParagraphTextId(value: unknown): string {
  const current = normalizeDocumentParagraphId(value);
  if (!current) return randomDocumentParagraphId();
  const number = Number.parseInt(current, 16);
  return formatDocumentParagraphId(number >= MAX_PARAGRAPH_ID ? 1 : number + 1);
}

function randomDocumentParagraphId(): string {
  let value = 0;
  const cryptoApi = globalThis.crypto;
  if (typeof cryptoApi?.getRandomValues === 'function') {
    const values = new Uint32Array(1);
    cryptoApi.getRandomValues(values);
    value = values[0] ?? 0;
  } else if (typeof cryptoApi?.randomUUID === 'function') {
    value = Number.parseInt(
      cryptoApi.randomUUID().replaceAll('-', '').slice(0, 8),
      16,
    );
  }
  if (!value) {
    fallbackIdentitySequence += 1;
    value = stableIdentityHash(
      `${Date.now()}:${Math.random()}:${fallbackIdentitySequence}`,
    );
  }
  return formatDocumentParagraphId(value || fallbackIdentitySequence || 1);
}

function stableIdentityHash(value: string): number {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0 || 1;
}

function formatDocumentParagraphId(value: number): string {
  return ((value >>> 0) & MAX_PARAGRAPH_ID || 1)
    .toString(16)
    .toUpperCase()
    .padStart(8, '0');
}

function sameDocumentParagraphIdentity(
  source: WorkDocumentParagraphIdentitySource,
  identity: WorkDocumentParagraphIdentity,
): boolean {
  return (
    normalizeDocumentParagraphId(source.paragraphId) === identity.paragraphId &&
    normalizeDocumentParagraphId(source.textId) === identity.textId
  );
}

function transactionMapping(transactions: readonly Transaction[]): Mapping {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  return mapping;
}
