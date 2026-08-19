import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

export const DOCUMENT_INTEGRITY_BOOKMARK = 1 << 0;
export const DOCUMENT_INTEGRITY_FIELD = 1 << 1;
export const DOCUMENT_INTEGRITY_IMAGE = 1 << 2;
export const DOCUMENT_INTEGRITY_NOTE = 1 << 3;
export const DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY = 1 << 4;
export const DOCUMENT_INTEGRITY_TABLE_ROW_IDENTITY = 1 << 5;

export type DocumentIntegrityFeature =
  | typeof DOCUMENT_INTEGRITY_BOOKMARK
  | typeof DOCUMENT_INTEGRITY_FIELD
  | typeof DOCUMENT_INTEGRITY_IMAGE
  | typeof DOCUMENT_INTEGRITY_NOTE
  | typeof DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY
  | typeof DOCUMENT_INTEGRITY_TABLE_ROW_IDENTITY;

// ProseMirror nodes are persistent. A transaction recreates only the path to
// the edited node, so unchanged document chunks can safely reuse their feature
// summaries without retaining indexes for every paragraph, cell, and text node.
const featureCache = new WeakMap<ProseMirrorNode, number>();

export function documentHasIntegrityFeature(
  document: ProseMirrorNode,
  feature: DocumentIntegrityFeature,
): boolean {
  return (documentIntegrityFeatures(document, true) & feature) !== 0;
}

/**
 * Seeds only parser-authenticated chunk hints. Edited chunks are new
 * ProseMirror nodes and therefore fall back to a real subtree scan.
 */
export function primeDocumentIntegrityFeatures(
  document: ProseMirrorNode,
): void {
  document.forEach(primeDocumentIntegrityNode);
}

function primeDocumentIntegrityNode(node: ProseMirrorNode): void {
  if (node.type.name === 'documentChunk') {
    const features = nonNegativeInteger(node.attrs.integrityFeatures);
    if (features !== null) featureCache.set(node, features);
    if (node.attrs.windowContainer !== true) return;
  }
  node.forEach(primeDocumentIntegrityNode);
}

function documentIntegrityFeatures(
  node: ProseMirrorNode,
  cacheNode: boolean,
): number {
  if (cacheNode) {
    const cached = featureCache.get(node);
    if (cached !== undefined) return cached;
  }

  let features = nodeIntegrityFeatures(node);
  node.forEach((child) => {
    features |= documentIntegrityFeatures(
      child,
      child.type.name === 'documentChunk',
    );
  });

  if (cacheNode) featureCache.set(node, features);
  return features;
}

function nodeIntegrityFeatures(node: ProseMirrorNode): number {
  let features = 0;
  switch (node.type.name) {
    case 'documentBookmarkBoundary':
      features |= DOCUMENT_INTEGRITY_BOOKMARK;
      break;
    case 'documentCrossReference':
      if (node.attrs.targetType === 'bookmark') {
        features |= DOCUMENT_INTEGRITY_BOOKMARK;
      }
      break;
    case 'documentField':
      features |= DOCUMENT_INTEGRITY_FIELD;
      break;
    case 'image':
      features |= DOCUMENT_INTEGRITY_IMAGE;
      break;
    case 'documentNote':
    case 'documentNoteReference':
      features |= DOCUMENT_INTEGRITY_NOTE;
      break;
  }
  if (
    node.isText &&
    node.marks.some(
      (mark) =>
        mark.type.name === 'link' &&
        typeof mark.attrs.href === 'string' &&
        mark.attrs.href.startsWith('#'),
    )
  ) {
    features |= DOCUMENT_INTEGRITY_BOOKMARK;
  }
  if (
    identityComponent(node.attrs.paragraphId) ||
    identityComponent(node.attrs.textId)
  ) {
    features |= DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY;
  }
  if (
    identityComponent(node.attrs.rowId) ||
    identityComponent(node.attrs.rowTextId)
  ) {
    features |= DOCUMENT_INTEGRITY_TABLE_ROW_IDENTITY;
  }
  return features;
}

function identityComponent(value: unknown): boolean {
  return typeof value === 'string' && value.trim().length > 0;
}

function nonNegativeInteger(value: unknown): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : null;
}
