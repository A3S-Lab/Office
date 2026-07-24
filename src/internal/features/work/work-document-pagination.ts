import { Extension, type Editor } from '@tiptap/core';
import {
  Plugin,
  PluginKey,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state';
import { DecorationSet } from '@tiptap/pm/view';
import { pageBreakDecorations } from './work-document-pagination-decorations';
import type { DocumentPaginationVisualBreak } from './work-document-pagination-types';

export {
  createDocumentLineFragments,
  createShapedDocumentLineFragments,
  findDocumentLineStartOffset,
} from './work-document-line-measurement';
export { reusableDocumentLayoutBlocks } from './work-document-pagination-dom';
export {
  documentPageChromeHeights,
  documentPageMetrics,
  measureDocumentLayoutBlocks,
} from './work-document-pagination-measurement';
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
} from './work-document-text-layout';

interface DocumentPaginationPluginState {
  decorations: DecorationSet;
  revision: number;
}

interface DocumentPaginationMeta {
  kind: 'apply' | 'clear';
  revision: number;
  breaks?: DocumentPaginationVisualBreak[];
}

const documentPaginationPluginKey =
  new PluginKey<DocumentPaginationPluginState>('documentPagination');

export const DocumentPagination = Extension.create({
  name: 'documentPagination',

  addProseMirrorPlugins() {
    return [
      new Plugin<DocumentPaginationPluginState>({
        key: documentPaginationPluginKey,
        state: {
          init: () => ({
            decorations: DecorationSet.empty,
            revision: 0,
          }),
          apply: (transaction, current, _previous, next) => {
            const meta = transaction.getMeta(documentPaginationPluginKey) as
              | DocumentPaginationMeta
              | undefined;
            if (meta?.kind === 'clear') {
              return {
                decorations: DecorationSet.empty,
                revision: meta.revision,
              };
            }
            if (meta?.kind === 'apply') {
              return {
                decorations: DecorationSet.create(
                  next.doc,
                  (meta.breaks ?? []).flatMap(pageBreakDecorations),
                ),
                revision: meta.revision,
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
                return {
                  ...current,
                  decorations: current.decorations.map(
                    transaction.mapping,
                    next.doc,
                  ),
                };
              }
              return {
                decorations: DecorationSet.empty,
                revision: current.revision,
              };
            }
            return {
              ...current,
              decorations: current.decorations.map(
                transaction.mapping,
                next.doc,
              ),
            };
          },
        },
        props: {
          decorations(state) {
            return documentPaginationState(state).decorations;
          },
        },
      }),
    ];
  },
});

export function clearDocumentPagination(
  editor: Editor,
  revision: number,
): void {
  if (editor.isDestroyed) return;
  const current = documentPaginationState(editor.state);
  if (!current.decorations.find().length && current.revision === revision)
    return;
  editor.view.dispatch(
    editor.state.tr
      .setMeta(documentPaginationPluginKey, {
        kind: 'clear',
        revision,
      } satisfies DocumentPaginationMeta)
      .setMeta('addToHistory', false),
  );
}

export function applyDocumentPagination(
  editor: Editor,
  revision: number,
  breaks: DocumentPaginationVisualBreak[],
): void {
  if (editor.isDestroyed) return;
  editor.view.dispatch(
    editor.state.tr
      .setMeta(documentPaginationPluginKey, {
        kind: 'apply',
        revision,
        breaks,
      } satisfies DocumentPaginationMeta)
      .setMeta('addToHistory', false),
  );
}

function documentPaginationState(
  state: EditorState,
): DocumentPaginationPluginState {
  return (
    documentPaginationPluginKey.getState(state) ?? {
      decorations: DecorationSet.empty,
      revision: 0,
    }
  );
}
