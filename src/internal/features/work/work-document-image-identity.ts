import { isHistoryTransaction } from '@tiptap/pm/history';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { type EditorState, Plugin, type Transaction } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import {
  DOCUMENT_INTEGRITY_IMAGE,
  documentHasIntegrityFeature,
} from './work-document-integrity-index';

export interface WorkDocumentImageIdentity {
  objectId: string;
  docPropertiesId: number;
  anchorId: string;
  editId: string;
}

export interface WorkDocumentImageIdentitySource {
  objectId?: unknown;
  docPropertiesId?: unknown;
  anchorId?: unknown;
  editId?: unknown;
}

export interface WorkDocumentImageIdentityRegistry {
  objectIds: Set<string>;
  docPropertiesIds: Set<number>;
  anchorIds: Set<string>;
}

const IMAGE_OBJECT_ID_PATTERN = /^[0-9A-F]{8}$/;
const MAX_DOC_PROPERTIES_ID = 0xffff_ffff;
const MAX_WORDPROCESSING_ID = 0x7fff_ffff;
const IMAGE_OBJECT_ID_ATTRIBUTE = 'data-office-image-object-id';
const IMAGE_DOC_PROPERTIES_ID_ATTRIBUTE = 'data-office-image-doc-properties-id';
const IMAGE_ANCHOR_ID_ATTRIBUTE = 'data-office-image-anchor-id';
const IMAGE_EDIT_ID_ATTRIBUTE = 'data-office-image-edit-id';
const IMAGE_IDENTITY_ATTRIBUTE_NAMES = new Set([
  'objectId',
  'docPropertiesId',
  'anchorId',
  'editId',
]);
let fallbackIdentitySequence = 0;

export function createDocumentImageIdentityRegistry(): WorkDocumentImageIdentityRegistry {
  return {
    objectIds: new Set(),
    docPropertiesIds: new Set(),
    anchorIds: new Set(),
  };
}

export function normalizeDocumentImageObjectId(value: unknown): string | null {
  const normalized = normalizeHexIdentifier(value);
  if (!normalized || !IMAGE_OBJECT_ID_PATTERN.test(normalized)) return null;
  const number = Number.parseInt(normalized, 16);
  return number > 0 && number <= MAX_WORDPROCESSING_ID ? normalized : null;
}

export function normalizeDocumentImageDocPropertiesId(
  value: unknown,
): number | null {
  if (
    value === null ||
    value === undefined ||
    value === '' ||
    (typeof value === 'string' && !value.trim())
  ) {
    return null;
  }
  const number = typeof value === 'number' ? value : Number(value);
  return Number.isSafeInteger(number) &&
    number >= 0 &&
    number <= MAX_DOC_PROPERTIES_ID
    ? number
    : null;
}

export function normalizeDocumentImageAnchorId(value: unknown): string | null {
  return normalizeDocumentImageObjectId(value);
}

export function normalizeDocumentImageEditId(value: unknown): string | null {
  return normalizeDocumentImageObjectId(value);
}

export function normalizeDocumentImageIdentity(
  source: WorkDocumentImageIdentitySource,
): WorkDocumentImageIdentity | null {
  const anchorId = normalizeDocumentImageAnchorId(source.anchorId);
  const docPropertiesId = normalizeDocumentImageDocPropertiesId(
    source.docPropertiesId,
  );
  const objectId =
    normalizeDocumentImageObjectId(source.objectId) ??
    anchorId ??
    (docPropertiesId === null
      ? null
      : documentImageObjectIdFromDocPropertiesId(docPropertiesId));
  if (!objectId) return null;
  return {
    objectId,
    docPropertiesId:
      docPropertiesId ?? documentImageDocPropertiesIdFromObjectId(objectId),
    anchorId: anchorId ?? objectId,
    editId:
      normalizeDocumentImageEditId(source.editId) ??
      documentImageEditIdFromObjectId(objectId),
  };
}

export function createDocumentImageIdentity(): WorkDocumentImageIdentity {
  const objectId = randomDocumentImageObjectId();
  return normalizeDocumentImageIdentity({
    objectId,
  }) as WorkDocumentImageIdentity;
}

export function uniqueDocumentImageIdentity(
  source: WorkDocumentImageIdentitySource,
  registry: WorkDocumentImageIdentityRegistry,
): WorkDocumentImageIdentity {
  const preferred = normalizeDocumentImageIdentity(source);
  if (preferred && !documentImageIdentityConflicts(preferred, registry)) {
    reserveDocumentImageIdentity(preferred, registry);
    return preferred;
  }
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const identity = createDocumentImageIdentity();
    if (documentImageIdentityConflicts(identity, registry)) continue;
    reserveDocumentImageIdentity(identity, registry);
    return identity;
  }
  return sequentialDocumentImageIdentity(registry);
}

export function documentImageIdentityFromAttributes(
  attributes: Record<string, unknown>,
): WorkDocumentImageIdentity | null {
  return normalizeDocumentImageIdentity({
    objectId: attributes.objectId,
    docPropertiesId: attributes.docPropertiesId,
    anchorId: attributes.anchorId,
    editId: attributes.editId,
  });
}

export function documentImageIdentityFromElement(
  element: Element,
): WorkDocumentImageIdentity {
  const identity =
    normalizeDocumentImageIdentity({
      objectId: element.getAttribute(IMAGE_OBJECT_ID_ATTRIBUTE),
      docPropertiesId: element.getAttribute(IMAGE_DOC_PROPERTIES_ID_ATTRIBUTE),
      anchorId: element.getAttribute(IMAGE_ANCHOR_ID_ATTRIBUTE),
      editId: element.getAttribute(IMAGE_EDIT_ID_ATTRIBUTE),
    }) ?? createDocumentImageIdentity();
  applyDocumentImageIdentityToElement(element, identity);
  return identity;
}

export function applyDocumentImageIdentityToElement(
  element: Element,
  source: WorkDocumentImageIdentitySource,
): WorkDocumentImageIdentity {
  const identity =
    normalizeDocumentImageIdentity(source) ?? createDocumentImageIdentity();
  element.setAttribute(IMAGE_OBJECT_ID_ATTRIBUTE, identity.objectId);
  element.setAttribute(
    IMAGE_DOC_PROPERTIES_ID_ATTRIBUTE,
    String(identity.docPropertiesId),
  );
  element.setAttribute(IMAGE_ANCHOR_ID_ATTRIBUTE, identity.anchorId);
  element.setAttribute(IMAGE_EDIT_ID_ATTRIBUTE, identity.editId);
  return identity;
}

export function createDocumentImageIdentityPlugin(
  imageNodeName = 'image',
): Plugin {
  return new Plugin({
    view(view) {
      const transaction = normalizeDocumentImageIdentities(
        view.state,
        imageNodeName,
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
      return normalizeDocumentImageIdentities(
        newState,
        imageNodeName,
        oldState,
        transactions,
      );
    },
  });
}

interface DocumentImageAtPosition {
  node: ProseMirrorNode;
  position: number;
}

function normalizeDocumentImageIdentities(
  state: EditorState,
  imageNodeName: string,
  oldState?: EditorState,
  transactions: readonly Transaction[] = [],
): Transaction | null {
  if (
    imageNodeName === 'image' &&
    !documentHasIntegrityFeature(state.doc, DOCUMENT_INTEGRITY_IMAGE)
  ) {
    return null;
  }
  const images = documentImages(state.doc, imageNodeName);
  if (!images.length) return null;
  const retainedPositions = oldState
    ? retainedDocumentImagePositions(
        oldState,
        state,
        transactions,
        imageNodeName,
      )
    : new Set<number>();
  const editedPositions =
    oldState &&
    !transactions.some(isHistoryTransaction) &&
    transactions.some(
      (transaction) => transaction.getMeta('addToHistory') !== false,
    )
      ? editedDocumentImagePositions(
          oldState,
          state,
          transactions,
          imageNodeName,
        )
      : new Set<number>();
  const ordered = [
    ...images.filter((image) => retainedPositions.has(image.position)),
    ...images.filter((image) => !retainedPositions.has(image.position)),
  ];
  const registry = createDocumentImageIdentityRegistry();
  const updates = new Map<number, WorkDocumentImageIdentity>();
  for (const image of ordered) {
    const currentSource = documentImageIdentitySourceFromNode(image.node);
    let source = currentSource;
    if (editedPositions.has(image.position)) {
      source = {
        ...source,
        editId: nextDocumentImageEditId(source.editId),
      };
    }
    const identity = uniqueDocumentImageIdentity(source, registry);
    if (!sameDocumentImageIdentity(currentSource, identity)) {
      updates.set(image.position, identity);
    }
  }
  if (!updates.size) return null;
  const transaction = state.tr;
  for (const image of images) {
    const identity = updates.get(image.position);
    if (!identity) continue;
    transaction.setNodeMarkup(image.position, undefined, {
      ...image.node.attrs,
      ...identity,
    });
  }
  return transaction.docChanged ? transaction : null;
}

function retainedDocumentImagePositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  imageNodeName: string,
): Set<number> {
  const mapping = transactionMapping(transactions);
  const retained = new Set<number>();
  for (const image of documentImages(oldState.doc, imageNodeName)) {
    const previousObjectId = normalizeDocumentImageObjectId(
      image.node.attrs.objectId,
    );
    if (!previousObjectId) continue;
    const mapped = mapping.mapResult(image.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current?.type.name === imageNodeName &&
      normalizeDocumentImageObjectId(current.attrs.objectId) ===
        previousObjectId
    ) {
      retained.add(mapped.pos);
    }
  }
  return retained;
}

function editedDocumentImagePositions(
  oldState: EditorState,
  newState: EditorState,
  transactions: readonly Transaction[],
  imageNodeName: string,
): Set<number> {
  const mapping = transactionMapping(transactions);
  const currentImages = documentImages(newState.doc, imageNodeName);
  const previousImages = documentImages(oldState.doc, imageNodeName);
  const previousObjectIds = new Set(
    previousImages.flatMap((image) => {
      const objectId = normalizeDocumentImageObjectId(
        image.node.attrs.objectId,
      );
      return objectId ? [objectId] : [];
    }),
  );
  const currentByObjectId = new Map<string, DocumentImageAtPosition[]>();
  for (const image of currentImages) {
    const objectId = normalizeDocumentImageObjectId(image.node.attrs.objectId);
    if (!objectId) continue;
    const matching = currentByObjectId.get(objectId) ?? [];
    matching.push(image);
    currentByObjectId.set(objectId, matching);
  }
  const edited = new Set<number>();
  for (const previous of previousImages) {
    const objectId = normalizeDocumentImageObjectId(
      previous.node.attrs.objectId,
    );
    if (!objectId) continue;
    const mapped = mapping.mapResult(previous.position, 1);
    const current = newState.doc.nodeAt(mapped.pos);
    if (
      current?.type.name === imageNodeName &&
      normalizeDocumentImageObjectId(current.attrs.objectId) === objectId
    ) {
      if (
        sameImageEditId(previous.node, current) &&
        !sameDocumentImageContent(previous.node, current)
      ) {
        edited.add(mapped.pos);
      }
      continue;
    }
    const moved = currentByObjectId.get(objectId);
    if (moved?.length === 1 && sameImageEditId(previous.node, moved[0].node)) {
      edited.add(moved[0].position);
    }
  }
  for (const current of currentImages) {
    const objectId = normalizeDocumentImageObjectId(
      current.node.attrs.objectId,
    );
    if (objectId && !previousObjectIds.has(objectId)) {
      edited.add(current.position);
    }
  }
  return edited;
}

function transactionMapping(transactions: readonly Transaction[]): Mapping {
  const mapping = new Mapping();
  for (const transaction of transactions) {
    mapping.appendMapping(transaction.mapping);
  }
  return mapping;
}

function documentImages(
  document: ProseMirrorNode,
  imageNodeName: string,
): DocumentImageAtPosition[] {
  const images: DocumentImageAtPosition[] = [];
  document.descendants((node, position) => {
    if (node.type.name === imageNodeName) images.push({ node, position });
  });
  return images;
}

function documentImageIdentitySourceFromNode(
  node: ProseMirrorNode,
): WorkDocumentImageIdentitySource {
  return {
    objectId: node.attrs.objectId,
    docPropertiesId: node.attrs.docPropertiesId,
    anchorId: node.attrs.anchorId,
    editId: node.attrs.editId,
  };
}

function sameDocumentImageContent(
  previous: ProseMirrorNode,
  current: ProseMirrorNode,
): boolean {
  const names = new Set([
    ...Object.keys(previous.attrs),
    ...Object.keys(current.attrs),
  ]);
  for (const name of names) {
    if (IMAGE_IDENTITY_ATTRIBUTE_NAMES.has(name)) continue;
    if (!Object.is(previous.attrs[name], current.attrs[name])) return false;
  }
  return true;
}

function sameImageEditId(
  previous: ProseMirrorNode,
  current: ProseMirrorNode,
): boolean {
  return (
    normalizeDocumentImageEditId(previous.attrs.editId) ===
    normalizeDocumentImageEditId(current.attrs.editId)
  );
}

function sameDocumentImageIdentity(
  source: WorkDocumentImageIdentitySource,
  identity: WorkDocumentImageIdentity,
): boolean {
  return (
    normalizeDocumentImageObjectId(source.objectId) === identity.objectId &&
    normalizeDocumentImageDocPropertiesId(source.docPropertiesId) ===
      identity.docPropertiesId &&
    normalizeDocumentImageAnchorId(source.anchorId) === identity.anchorId &&
    normalizeDocumentImageEditId(source.editId) === identity.editId
  );
}

function documentImageIdentityConflicts(
  identity: WorkDocumentImageIdentity,
  registry: WorkDocumentImageIdentityRegistry,
): boolean {
  return (
    registry.objectIds.has(identity.objectId) ||
    registry.docPropertiesIds.has(identity.docPropertiesId) ||
    registry.anchorIds.has(identity.anchorId)
  );
}

function reserveDocumentImageIdentity(
  identity: WorkDocumentImageIdentity,
  registry: WorkDocumentImageIdentityRegistry,
): void {
  registry.objectIds.add(identity.objectId);
  registry.docPropertiesIds.add(identity.docPropertiesId);
  registry.anchorIds.add(identity.anchorId);
}

function sequentialDocumentImageIdentity(
  registry: WorkDocumentImageIdentityRegistry,
): WorkDocumentImageIdentity {
  for (let value = 1; value <= MAX_WORDPROCESSING_ID; value += 1) {
    const identity = normalizeDocumentImageIdentity({
      objectId: formatHexIdentifier(value),
    }) as WorkDocumentImageIdentity;
    if (documentImageIdentityConflicts(identity, registry)) continue;
    reserveDocumentImageIdentity(identity, registry);
    return identity;
  }
  throw new Error('No unique Word image identity is available.');
}

function documentImageDocPropertiesIdFromObjectId(objectId: string): number {
  return Number.parseInt(objectId, 16);
}

function documentImageEditIdFromObjectId(objectId: string): string {
  return formatWordprocessingIdentifier(
    Math.imul(Number.parseInt(objectId, 16), 0x045d_9f3b),
  );
}

function nextDocumentImageEditId(value: unknown): string {
  const current = normalizeDocumentImageEditId(value);
  if (!current) return formatWordprocessingIdentifier(Date.now());
  const number = Number.parseInt(current, 16);
  return formatHexIdentifier(number >= MAX_WORDPROCESSING_ID ? 1 : number + 1);
}

function documentImageObjectIdFromDocPropertiesId(id: number): string {
  return formatWordprocessingIdentifier(id);
}

function randomDocumentImageObjectId(): string {
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
  return formatWordprocessingIdentifier(value || fallbackIdentitySequence || 1);
}

function stableIdentityHash(value: string): number {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0 || 1;
}

function normalizeHexIdentifier(value: unknown): string | null {
  return typeof value === 'string' && value.trim()
    ? value.trim().toUpperCase()
    : null;
}

function formatHexIdentifier(value: number): string {
  return (value >>> 0).toString(16).toUpperCase().padStart(8, '0');
}

function formatWordprocessingIdentifier(value: number): string {
  return formatHexIdentifier(value & MAX_WORDPROCESSING_ID || 1);
}
