import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  documentLazyHtmlChunkFragment,
  patchDocumentLazyHtmlProjection,
} from './work-document-lazy-html';
import {
  documentLazyHtmlProjection,
  invalidateDocumentLazyHtmlProjection,
  materializeLazyDocumentEditorRoot,
} from './work-document-lazy-model';
import { serializeWorkDocumentNode } from './work-document-model-codec';
import { transferChangedDocumentTextStatistics } from './work-document-text-statistics';
import type { WorkDocumentModel, WorkDocumentNode } from './work-types';

export interface MaterializedLazyDocumentUpdate {
  html: string | null;
  root: WorkDocumentNode;
}

interface ChangedLazyDocumentChunk {
  after: ProseMirrorNode;
  before: ProseMirrorNode;
  id: string;
}

/**
 * Materializes the complete controlled model and patches canonical HTML from
 * small changed chunks when the transaction only edits plain text.
 */
export function materializeLazyDocumentUpdate(
  previousDocument: ProseMirrorNode,
  nextDocument: ProseMirrorNode,
  compactRoot: WorkDocumentNode,
  model: WorkDocumentModel,
): MaterializedLazyDocumentUpdate {
  const root = materializeLazyDocumentEditorRoot(compactRoot, model);
  const projection = documentLazyHtmlProjection(model);
  const changed = changedLazyDocumentChunks(previousDocument, nextDocument);
  if (
    !projection ||
    !changed?.length ||
    changed.some(
      ({ after, before }) => !documentChunkOnlyChangesText(before, after),
    )
  ) {
    invalidateDocumentLazyHtmlProjection(model);
    return { html: null, root };
  }

  try {
    const replacements = new Map<string, string>();
    for (const { id } of changed) {
      const range = projection.ranges.get(id);
      const isolated = isolatedDocumentChunk(root, id);
      if (!range || !isolated) throw new Error('Lazy chunk is not indexed.');
      const serialized = serializeWorkDocumentNode(isolated);
      const sectionOpen = serialized.indexOf('>') + 1;
      const sectionClose = serialized.lastIndexOf('</section>');
      if (sectionOpen <= 0 || sectionClose < sectionOpen) {
        throw new Error('Lazy chunk serialization has no section boundary.');
      }
      const fragment = documentLazyHtmlChunkFragment(
        serialized.slice(sectionOpen, sectionClose),
        range.tablePart,
      );
      if (fragment === null) {
        throw new Error('Lazy table chunk serialization is malformed.');
      }
      replacements.set(id, fragment);
    }
    const html = patchDocumentLazyHtmlProjection(projection, replacements);
    if (html === null) throw new Error('Lazy HTML patch could not be applied.');
    return { html, root };
  } catch {
    invalidateDocumentLazyHtmlProjection(model);
    return { html: null, root };
  }
}

/** Carries cached document statistics across a parser-authenticated text edit. */
export function transferLazyDocumentTextStatistics(
  previousDocument: ProseMirrorNode,
  nextDocument: ProseMirrorNode,
): boolean {
  const changed = changedLazyDocumentChunks(previousDocument, nextDocument);
  if (
    !changed?.length ||
    changed.some(
      ({ after, before }) => !documentChunkOnlyChangesText(before, after),
    )
  ) {
    return false;
  }
  return transferChangedDocumentTextStatistics(
    previousDocument,
    nextDocument,
    changed,
  );
}

function changedLazyDocumentChunks(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): ChangedLazyDocumentChunk[] | null {
  const changed: ChangedLazyDocumentChunk[] = [];
  const visit = (previous: ProseMirrorNode, next: ProseMirrorNode): boolean => {
    if (previous === next || previous.eq(next)) return true;
    if (previous.type !== next.type) return false;
    if (
      previous.type.name === 'documentChunk' &&
      previous.attrs.windowContainer !== true &&
      next.attrs.windowContainer !== true
    ) {
      const previousId = documentChunkId(previous);
      const nextId = documentChunkId(next);
      if (!previousId || previousId !== nextId) return false;
      changed.push({ after: next, before: previous, id: nextId });
      return true;
    }
    if (!previous.sameMarkup(next) || previous.childCount !== next.childCount) {
      return false;
    }
    for (let index = 0; index < previous.childCount; index += 1) {
      const previousChild = previous.child(index);
      const nextChild = next.child(index);
      if (!visit(previousChild, nextChild)) return false;
    }
    return true;
  };
  return visit(before, after) ? changed : null;
}

function documentChunkOnlyChangesText(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  let textChanged = false;
  const visit = (previous: ProseMirrorNode, next: ProseMirrorNode): boolean => {
    if (
      previous.type !== next.type ||
      !previous.sameMarkup(next) ||
      previous.childCount !== next.childCount
    ) {
      return false;
    }
    if (previous.isText || next.isText) {
      if (!previous.isText || !next.isText) return false;
      if (previous.text !== next.text) textChanged = true;
      return true;
    }
    for (let index = 0; index < previous.childCount; index += 1) {
      if (!visit(previous.child(index), next.child(index))) return false;
    }
    return true;
  };
  return visit(before, after) && textChanged;
}

function isolatedDocumentChunk(
  root: WorkDocumentNode,
  id: string,
): WorkDocumentNode | null {
  let section: WorkDocumentNode | null = null;
  let chunk: WorkDocumentNode | null = null;
  const visit = (
    node: WorkDocumentNode,
    activeSection: WorkDocumentNode | null,
  ): void => {
    if (chunk) return;
    const nextSection = node.type === 'documentSection' ? node : activeSection;
    if (
      node.type === 'documentChunk' &&
      node.attrs?.windowContainer !== true &&
      node.attrs?.id === id
    ) {
      section = nextSection;
      chunk = node;
      return;
    }
    for (const child of node.content ?? []) visit(child, nextSection);
  };
  visit(root, null);
  const foundSection = section as WorkDocumentNode | null;
  const foundChunk = chunk as WorkDocumentNode | null;
  if (!foundSection || !foundChunk) return null;
  return {
    type: 'doc',
    content: [{ ...foundSection, content: [foundChunk] }],
  };
}

function documentChunkId(node: ProseMirrorNode): string | null {
  const id = node.attrs.id;
  return typeof id === 'string' && id ? id : null;
}
