import { Extension } from '@tiptap/core';
import type { Mark, Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import {
  documentScriptFontFamilyForRendering,
  documentScriptFontSegments,
  parseDocumentScriptFonts,
  serializeDocumentScriptFonts,
} from './work-document-script-fonts';
import { documentWordLineHeightFactor } from './work-document-word-line-metrics';

const SCRIPT_FONT_NORMALIZATION_META = 'documentScriptFontNormalization';
const documentScriptFontPluginKey = new PluginKey(
  'documentScriptFontNormalization',
);

interface DocumentRange {
  from: number;
  to: number;
}

interface ScriptFontMarkUpdate {
  from: number;
  to: number;
  current: Mark;
  next: Mark;
}

export const DocumentScriptFontFormatting = Extension.create({
  name: 'documentScriptFontFormatting',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: documentScriptFontPluginKey,
        view(view) {
          const transaction = normalizeDocumentScriptFontMarks(view.state);
          if (transaction) view.dispatch(transaction);
          return {};
        },
        appendTransaction(transactions, _oldState, newState) {
          if (
            !transactions.some((transaction) => transaction.docChanged) ||
            transactions.some((transaction) =>
              transaction.getMeta(SCRIPT_FONT_NORMALIZATION_META),
            )
          ) {
            return null;
          }
          return normalizeDocumentScriptFontMarks(
            newState,
            changedDocumentRanges(transactions, newState),
            true,
          );
        },
      }),
    ];
  },
});

function normalizeDocumentScriptFontMarks(
  state: EditorState,
  ranges?: readonly DocumentRange[],
  includeInHistory = false,
): Transaction | null {
  const textStyle = state.schema.marks.textStyle;
  if (!textStyle) return null;
  const updates: ScriptFontMarkUpdate[] = [];
  const visited = new Set<number>();
  const visit = (node: ProseMirrorNode, position: number) => {
    if (!node.isText || !node.text || visited.has(position)) return;
    visited.add(position);
    const current = node.marks.find((mark) => mark.type === textStyle);
    const fonts = parseDocumentScriptFonts(current?.attrs.scriptFonts);
    if (!current || !fonts) return;
    const serialized = serializeDocumentScriptFonts(fonts);
    if (!serialized) return;
    for (const segment of documentScriptFontSegments(node.text, fonts.hint)) {
      const family = documentScriptFontFamilyForRendering(
        fonts,
        segment.slot,
        current.attrs.fontFamily,
      );
      const attributes = {
        ...current.attrs,
        scriptFonts: serialized,
        scriptFontSlot: segment.slot,
        ...(family
          ? {
              fontFamily: family,
              wordLineHeightFactor: documentWordLineHeightFactor(family),
            }
          : {}),
      };
      const next = textStyle.create(attributes);
      if (current.eq(next)) continue;
      updates.push({
        from: position + segment.from,
        to: position + segment.to,
        current,
        next,
      });
    }
  };
  if (!ranges?.length) {
    state.doc.descendants(visit);
  } else {
    for (const range of ranges) {
      state.doc.nodesBetween(range.from, range.to, visit);
    }
  }
  if (!updates.length) return null;
  const transaction = state.tr;
  for (const update of updates) {
    transaction.removeMark(update.from, update.to, update.current);
    transaction.addMark(update.from, update.to, update.next);
  }
  transaction.setMeta(SCRIPT_FONT_NORMALIZATION_META, true);
  if (!includeInHistory) transaction.setMeta('addToHistory', false);
  return transaction;
}

function changedDocumentRanges(
  transactions: readonly Transaction[],
  state: EditorState,
): DocumentRange[] {
  const maximum = state.doc.content.size;
  const ranges: DocumentRange[] = [];
  for (const transaction of transactions) {
    for (const map of transaction.mapping.maps) {
      map.forEach((_oldFrom, _oldTo, newFrom, newTo) => {
        ranges.push(expandedRange(newFrom, newTo, maximum));
      });
    }
  }
  ranges.push(expandedRange(state.selection.from, state.selection.to, maximum));
  return mergeRanges(ranges);
}

function expandedRange(
  from: number,
  to: number,
  maximum: number,
): DocumentRange {
  return {
    from: Math.max(0, Math.min(maximum, Math.min(from, to) - 2)),
    to: Math.max(0, Math.min(maximum, Math.max(from, to) + 2)),
  };
}

function mergeRanges(ranges: readonly DocumentRange[]): DocumentRange[] {
  const ordered = [...ranges]
    .filter((range) => range.to >= range.from)
    .sort((left, right) => left.from - right.from || left.to - right.to);
  const merged: DocumentRange[] = [];
  for (const range of ordered) {
    const previous = merged[merged.length - 1];
    if (previous && range.from <= previous.to) {
      previous.to = Math.max(previous.to, range.to);
    } else {
      merged.push({ ...range });
    }
  }
  return merged;
}
