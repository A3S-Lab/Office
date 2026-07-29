import type { Editor } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export interface DocumentFindHighlightRange {
  from: number;
  to: number;
}

interface DocumentFindHighlightMeta {
  activeIndex: number;
  ranges: readonly DocumentFindHighlightRange[];
}

const documentFindHighlightPluginKey = new PluginKey<DecorationSet>(
  'documentFindHighlight',
);

export function registerDocumentFindHighlight(editor: Editor): void {
  if (editor.isDestroyed || documentFindHighlightPluginKey.get(editor.state)) {
    return;
  }
  editor.registerPlugin(
    new Plugin<DecorationSet>({
      key: documentFindHighlightPluginKey,
      state: {
        init: () => DecorationSet.empty,
        apply: (transaction, current) => {
          const meta = transaction.getMeta(documentFindHighlightPluginKey) as
            | DocumentFindHighlightMeta
            | null
            | undefined;
          if (meta) {
            return DecorationSet.create(
              transaction.doc,
              meta.ranges.flatMap((range, index) => {
                if (
                  range.from < 0 ||
                  range.to <= range.from ||
                  range.to > transaction.doc.content.size
                ) {
                  return [];
                }
                const active = index === meta.activeIndex;
                return [
                  Decoration.inline(range.from, range.to, {
                    class: `work-document-find-match${active ? ' active' : ''}`,
                    'data-document-find-active': active ? 'true' : 'false',
                    'data-document-find-index': String(index),
                  }),
                ];
              }),
            );
          }
          return transaction.docChanged
            ? current.map(transaction.mapping, transaction.doc)
            : current;
        },
      },
      props: {
        decorations: (state) =>
          documentFindHighlightPluginKey.getState(state) ?? DecorationSet.empty,
      },
    }),
  );
}

export function updateDocumentFindHighlights(
  editor: Editor,
  ranges: readonly DocumentFindHighlightRange[],
  activeIndex: number,
): void {
  if (editor.isDestroyed || !documentFindHighlightPluginKey.get(editor.state)) {
    return;
  }
  editor.view.dispatch(
    editor.state.tr
      .setMeta(documentFindHighlightPluginKey, { activeIndex, ranges })
      .setMeta('addToHistory', false),
  );
}

export function unregisterDocumentFindHighlight(editor: Editor): void {
  if (!editor.isDestroyed)
    editor.unregisterPlugin(documentFindHighlightPluginKey);
}
