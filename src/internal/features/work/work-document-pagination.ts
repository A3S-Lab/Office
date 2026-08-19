import { Extension } from '@tiptap/core';
import {
  Plugin,
  PluginKey,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state';
import { DecorationSet } from '@tiptap/pm/view';
import { pageBreakDecorations } from './work-document-pagination-decorations';
import {
  type DocumentChunkHydrationMeta,
  documentChunkMountedIds,
  DOCUMENT_CHUNK_HYDRATION_META,
  DOCUMENT_CHUNK_PAGINATION_META,
  DOCUMENT_CHUNK_VISIBLE_IDS_META,
  selectedDocumentChunkId,
  type DocumentChunkVisibleIdsMeta,
} from './work-document-chunk-node';
import type { DocumentPaginationVisualBreak } from './work-document-pagination-types';

export {
  createDocumentLineFragments,
  createShapedDocumentLineFragments,
  findDocumentLineStartOffset,
} from './work-document-line-measurement';
export {
  reusableDocumentChunkLayoutBlocks,
  reusableDocumentLayoutBlocks,
} from './work-document-pagination-dom';
export {
  documentPageBodyHeight,
  documentPageMetrics,
  documentPaginationSurfaceHeight,
  measureDocumentLayoutBlocks,
  measureDocumentLayoutBlocksIncrementally,
} from './work-document-pagination-measurement';
export type { IncrementalDocumentLayoutMeasurementOptions } from './work-document-pagination-measurement';
export type {
  DocumentPaginationSection,
  DocumentPaginationSnapshot,
  DocumentPaginationVisualBreak,
  DocumentPaginationVisualPageChrome,
  DocumentTableCellBoundary,
  DocumentTableCellFragmentMeasurement,
  DocumentTableCellPageBreak,
  DocumentTablePaginationBreak,
  DocumentTableRowFragmentPlan,
  DocumentTextLayoutCollection,
  MeasuredDocumentLayoutBlock,
} from './work-document-pagination-types';
export { createDocumentTableRowFragmentPlan } from './work-document-table-pagination';
export {
  collectDocumentTextLayoutParagraphs,
  collectDocumentTextLayoutRuns,
  documentTextLayoutBatches,
} from './work-document-text-layout';

interface DocumentPaginationPluginState {
  breakIndex: DocumentPaginationBreakIndex;
  breaks: readonly DocumentPaginationVisualBreak[];
  decorations: DecorationSet;
  revision: number;
  visibleIds: readonly string[];
}

interface DocumentPaginationBreakIndex {
  byChunkId: ReadonlyMap<string, readonly DocumentPaginationVisualBreak[]>;
  outsideChunks: readonly DocumentPaginationVisualBreak[];
}

interface DocumentPaginationMeta {
  kind: 'apply' | 'clear';
  revision: number;
  breaks?: DocumentPaginationVisualBreak[];
  visibleIds?: readonly string[];
}

const documentPaginationPluginKey =
  new PluginKey<DocumentPaginationPluginState>('documentPagination');

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentPagination: {
      applyDocumentPagination: (
        revision: number,
        breaks: DocumentPaginationVisualBreak[],
      ) => ReturnType;
      clearDocumentPagination: (revision: number) => ReturnType;
    };
  }
}

export const DocumentPagination = Extension.create({
  name: 'documentPagination',

  addCommands() {
    return {
      applyDocumentPagination:
        (revision, breaks) =>
        ({ state, tr }) => {
          tr.setMeta(documentPaginationPluginKey, {
            kind: 'apply',
            revision,
            breaks,
            visibleIds: documentChunkMountedIds(state) ?? [],
          } satisfies DocumentPaginationMeta);
          tr.setMeta(DOCUMENT_CHUNK_PAGINATION_META, { breaks });
          tr.setMeta('addToHistory', false);
          return true;
        },
      clearDocumentPagination:
        (revision) =>
        ({ state, tr }) => {
          const current = documentPaginationState(state);
          if (!current.breaks.length && current.revision === revision) {
            return false;
          }
          tr.setMeta(documentPaginationPluginKey, {
            kind: 'clear',
            revision,
          } satisfies DocumentPaginationMeta);
          tr.setMeta(DOCUMENT_CHUNK_PAGINATION_META, { breaks: [] });
          tr.setMeta('addToHistory', false);
          return true;
        },
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin<DocumentPaginationPluginState>({
        key: documentPaginationPluginKey,
        state: {
          init: () => ({
            breakIndex: emptyDocumentPaginationBreakIndex(),
            breaks: [],
            decorations: DecorationSet.empty,
            revision: 0,
            visibleIds: [],
          }),
          apply: (transaction, current, _previous, next) => {
            const meta = transaction.getMeta(documentPaginationPluginKey) as
              | DocumentPaginationMeta
              | undefined;
            const visibleMeta = transaction.getMeta(
              DOCUMENT_CHUNK_VISIBLE_IDS_META,
            ) as DocumentChunkVisibleIdsMeta | undefined;
            const hydrationMeta = transaction.getMeta(
              DOCUMENT_CHUNK_HYDRATION_META,
            ) as DocumentChunkHydrationMeta | undefined;
            if (meta?.kind === 'clear') {
              return {
                breakIndex: emptyDocumentPaginationBreakIndex(),
                breaks: [],
                decorations: DecorationSet.empty,
                revision: meta.revision,
                visibleIds: documentChunkMountedIds(next) ?? [],
              };
            }
            if (meta?.kind === 'apply') {
              return createDocumentPaginationPluginState(
                next,
                meta.breaks ?? [],
                meta.revision,
                meta.visibleIds,
              );
            }
            if (transaction.docChanged && hydrationMeta) {
              const visibleIds =
                visibleMeta?.visibleIds ??
                documentChunkMountedIds(next) ??
                current.visibleIds;
              return {
                ...current,
                decorations: documentPaginationDecorationSet(
                  next.doc,
                  current.breakIndex,
                  visibleIds,
                ),
                visibleIds: [...visibleIds],
              };
            }
            if (transaction.docChanged) {
              const appendedTo = transaction.getMeta('appendedTransaction') as
                | Transaction
                | undefined;
              const sourceMeta = appendedTo?.getMeta(
                documentPaginationPluginKey,
              ) as DocumentPaginationMeta | undefined;
              if (sourceMeta?.kind === 'apply') {
                return createDocumentPaginationPluginState(
                  next,
                  mapDocumentPaginationBreaks(
                    current.breaks,
                    transaction.mapping,
                  ),
                  current.revision,
                );
              }
              return {
                breakIndex: emptyDocumentPaginationBreakIndex(),
                breaks: [],
                decorations: DecorationSet.empty,
                revision: current.revision,
                visibleIds: documentChunkMountedIds(next) ?? [],
              };
            }
            const visibleIds =
              visibleMeta?.visibleIds ??
              (transaction.selectionSet
                ? withSelectedDocumentChunk(current.visibleIds, next)
                : current.visibleIds);
            if (arraysEqual(visibleIds, current.visibleIds)) return current;
            return {
              ...current,
              decorations: documentPaginationDecorationSet(
                next.doc,
                current.breakIndex,
                visibleIds,
              ),
              visibleIds: [...visibleIds],
            };
          },
        },
        props: {
          decorations(state) {
            return documentPaginationState(state).decorations;
          },
        },
        view: (initialView) => {
          const publishDiagnostics = (state: EditorState) => {
            const pagination = documentPaginationState(state);
            initialView.dom.dataset.paginationVisualBreaks = String(
              pagination.breaks.length,
            );
            initialView.dom.dataset.paginationVisibleBreaks = String(
              visibleDocumentPaginationBreakCount(
                pagination.breakIndex,
                pagination.visibleIds,
              ),
            );
            initialView.dom.dataset.paginationVisibleDecorations = String(
              pagination.decorations.find().length,
            );
            initialView.dom.dataset.paginationIndexedChunks = String(
              pagination.breakIndex.byChunkId.size,
            );
          };
          publishDiagnostics(initialView.state);
          return {
            update(nextView) {
              publishDiagnostics(nextView.state);
            },
          };
        },
      }),
    ];
  },
});

function documentPaginationState(
  state: EditorState,
): DocumentPaginationPluginState {
  return (
    documentPaginationPluginKey.getState(state) ?? {
      breakIndex: emptyDocumentPaginationBreakIndex(),
      breaks: [],
      decorations: DecorationSet.empty,
      revision: 0,
      visibleIds: [],
    }
  );
}

function createDocumentPaginationPluginState(
  state: EditorState,
  breaks: readonly DocumentPaginationVisualBreak[],
  revision: number,
  requestedVisibleIds?: readonly string[],
): DocumentPaginationPluginState {
  const visibleIds =
    requestedVisibleIds ?? documentChunkMountedIds(state) ?? [];
  const breakIndex = indexDocumentPaginationBreaks(state.doc, breaks);
  return {
    breakIndex,
    breaks: [...breaks],
    decorations: documentPaginationDecorationSet(
      state.doc,
      breakIndex,
      visibleIds,
    ),
    revision,
    visibleIds,
  };
}

function indexDocumentPaginationBreaks(
  document: EditorState['doc'],
  breaks: readonly DocumentPaginationVisualBreak[],
): DocumentPaginationBreakIndex {
  if (!breaks.length) return emptyDocumentPaginationBreakIndex();
  const chunks: Array<{ from: number; id: string; to: number }> = [];
  document.descendants((node, position) => {
    if (node.type.name !== 'documentChunk') return true;
    if (node.attrs.windowContainer === true) return true;
    chunks.push({
      from: position,
      id: typeof node.attrs.id === 'string' ? node.attrs.id : '',
      to: position + node.nodeSize,
    });
    return false;
  });
  if (!chunks.length) {
    return { byChunkId: new Map(), outsideChunks: [...breaks] };
  }

  const byChunkId = new Map<string, DocumentPaginationVisualBreak[]>();
  const outsideChunks: DocumentPaginationVisualBreak[] = [];
  const orderedBreaks = [...breaks].sort(compareDocumentPaginationBreaks);
  let chunkIndex = 0;
  for (const pageBreak of orderedBreaks) {
    while (
      chunkIndex < chunks.length &&
      pageBreak.position >= (chunks[chunkIndex]?.to ?? Number.POSITIVE_INFINITY)
    ) {
      chunkIndex += 1;
    }
    const chunk = chunks[chunkIndex];
    if (
      chunk &&
      pageBreak.position > chunk.from &&
      pageBreak.position < chunk.to
    ) {
      const chunkBreaks = byChunkId.get(chunk.id) ?? [];
      chunkBreaks.push(pageBreak);
      byChunkId.set(chunk.id, chunkBreaks);
    } else {
      outsideChunks.push(pageBreak);
    }
  }
  return { byChunkId, outsideChunks };
}

function documentPaginationDecorationSet(
  document: EditorState['doc'],
  breakIndex: DocumentPaginationBreakIndex,
  visibleIds: readonly string[],
): DecorationSet {
  const visibleBreaks = [...breakIndex.outsideChunks];
  for (const id of visibleIds) {
    const chunkBreaks = breakIndex.byChunkId.get(id);
    if (chunkBreaks) visibleBreaks.push(...chunkBreaks);
  }
  visibleBreaks.sort(compareDocumentPaginationBreaks);
  return DecorationSet.create(
    document,
    visibleBreaks.flatMap(pageBreakDecorations),
  );
}

function visibleDocumentPaginationBreakCount(
  breakIndex: DocumentPaginationBreakIndex,
  visibleIds: readonly string[],
): number {
  let count = breakIndex.outsideChunks.length;
  for (const id of visibleIds) {
    count += breakIndex.byChunkId.get(id)?.length ?? 0;
  }
  return count;
}

function emptyDocumentPaginationBreakIndex(): DocumentPaginationBreakIndex {
  return { byChunkId: new Map(), outsideChunks: [] };
}

function compareDocumentPaginationBreaks(
  left: DocumentPaginationVisualBreak,
  right: DocumentPaginationVisualBreak,
): number {
  return left.position - right.position || left.pageIndex - right.pageIndex;
}

function mapDocumentPaginationBreaks(
  breaks: readonly DocumentPaginationVisualBreak[],
  mapping: Transaction['mapping'],
): DocumentPaginationVisualBreak[] {
  return breaks.map((pageBreak) => ({
    ...pageBreak,
    position: mapping.map(pageBreak.position, -1),
    ...(pageBreak.tableBreak?.cellBreaks
      ? {
          tableBreak: {
            ...pageBreak.tableBreak,
            cellBreaks: pageBreak.tableBreak.cellBreaks.map((cellBreak) => ({
              ...cellBreak,
              position: mapping.map(cellBreak.position, -1),
            })),
          },
        }
      : {}),
  }));
}

function arraysEqual(
  left: readonly string[],
  right: readonly string[],
): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

function withSelectedDocumentChunk(
  visibleIds: readonly string[],
  state: EditorState,
): readonly string[] {
  const selected = selectedDocumentChunkId(state);
  return selected && !visibleIds.includes(selected)
    ? [...visibleIds, selected]
    : visibleIds;
}
