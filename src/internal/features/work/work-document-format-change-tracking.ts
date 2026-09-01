import { isHistoryTransaction } from '@tiptap/pm/history';
import type {
  Mark as ProseMirrorMark,
  Node as ProseMirrorNode,
} from '@tiptap/pm/model';
import type { EditorState, PluginKey, Transaction } from '@tiptap/pm/state';
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from '@tiptap/pm/transform';
import { ySyncPluginKey } from '@tiptap/y-tiptap';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import {
  isDocumentCharacterFormatMark,
  serializeDocumentCharacterFormatting,
} from './work-document-format-changes';
import { trackDocumentNumberingChangeTransaction } from './work-document-numbering-change-tracking';
import { trackDocumentParagraphFormattingTransaction } from './work-document-paragraph-format-change-tracking';

interface DocumentFormattingChangeTrackingOptions {
  isTracking: () => boolean;
  createChange: (
    kind: 'formatting' | 'paragraph-formatting' | 'numbering',
  ) => WorkDocumentChangeIdentity;
}

export function trackDocumentFormattingTransaction(
  transaction: Transaction,
  state: EditorState,
  type: ProseMirrorMark['type'],
  options: DocumentFormattingChangeTrackingOptions,
  pluginKey: PluginKey,
): void {
  if (!options.isTracking() || state.doc.eq(transaction.doc)) return;
  const sync = transaction.getMeta(ySyncPluginKey) as
    | { isChangeOrigin?: boolean }
    | undefined;
  if (
    transaction.getMeta(pluginKey) ||
    sync?.isChangeOrigin ||
    isHistoryTransaction(transaction)
  ) {
    return;
  }
  const ranges = formattingStepRanges(transaction);
  const formattedDocument = transaction.doc;
  let identity: WorkDocumentChangeIdentity | null = null;
  for (const range of ranges) {
    for (const segment of formattingSegments(
      state.doc,
      formattedDocument,
      range,
    )) {
      const beforeMarks = textMarksAt(state.doc, segment.from, segment.to);
      const afterMarks = textMarksAt(
        formattedDocument,
        segment.from,
        segment.to,
      );
      if (!beforeMarks || !afterMarks) continue;
      if (
        state.doc.textBetween(segment.from, segment.to) !==
        formattedDocument.textBetween(segment.from, segment.to)
      ) {
        continue;
      }
      const before = serializeDocumentCharacterFormatting(beforeMarks);
      const after = serializeDocumentCharacterFormatting(afterMarks);
      if (before === after || documentChangeMark(afterMarks)) continue;
      identity ??= options.createChange('formatting');
      transaction.addMark(
        segment.from,
        segment.to,
        type.create({
          kind: 'formatting',
          id: identity.id,
          actorId: identity.actorId ?? '',
          author: identity.author || 'A3S Work',
          date: identity.date || new Date().toISOString(),
          before,
        }),
      );
    }
  }
  const paragraphFormatting = trackDocumentParagraphFormattingTransaction(
    transaction,
    state,
    {
      createChange: () => options.createChange('paragraph-formatting'),
    },
  );
  const numbering = trackDocumentNumberingChangeTransaction(
    transaction,
    state,
    {
      createChange: () => options.createChange('numbering'),
    },
  );
  if (identity || paragraphFormatting || numbering) {
    transaction.setMeta(pluginKey, {
      ...(identity ? { formatting: true } : {}),
      ...(paragraphFormatting ? { paragraphFormatting: true } : {}),
      ...(numbering ? { numbering: true } : {}),
    });
  }
}

function formattingStepRanges(
  transaction: Transaction,
): Array<{ from: number; to: number }> {
  if (transaction.steps.some((step) => step instanceof ReplaceStep)) {
    return [];
  }
  const ranges: Array<{ from: number; to: number }> = [];
  for (const step of transaction.steps) {
    if (
      (step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
      isDocumentCharacterFormatMark(step.mark.type.name) &&
      step.from < step.to
    ) {
      ranges.push({ from: step.from, to: step.to });
    }
  }
  return mergeFormattingRanges(ranges);
}

function mergeFormattingRanges(
  ranges: Array<{ from: number; to: number }>,
): Array<{ from: number; to: number }> {
  const merged: Array<{ from: number; to: number }> = [];
  for (const range of ranges.sort((left, right) => left.from - right.from)) {
    const previous = merged.at(-1);
    if (!previous || range.from > previous.to) {
      merged.push({ ...range });
      continue;
    }
    previous.to = Math.max(previous.to, range.to);
  }
  return merged;
}

function formattingSegments(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
  range: { from: number; to: number },
): Array<{ from: number; to: number }> {
  const boundaries = new Set([range.from, range.to]);
  collectFormattingBoundaries(before, range, boundaries);
  collectFormattingBoundaries(after, range, boundaries);
  const ordered = Array.from(boundaries).sort((left, right) => left - right);
  return ordered.slice(0, -1).flatMap((from, index) => {
    const to = ordered[index + 1];
    return to !== undefined && from < to ? [{ from, to }] : [];
  });
}

function collectFormattingBoundaries(
  document: ProseMirrorNode,
  range: { from: number; to: number },
  boundaries: Set<number>,
): void {
  document.nodesBetween(range.from, range.to, (node, position) => {
    if (!node.isText) return;
    boundaries.add(Math.max(range.from, position));
    boundaries.add(Math.min(range.to, position + node.nodeSize));
  });
}

function textMarksAt(
  document: ProseMirrorNode,
  from: number,
  to: number,
): readonly ProseMirrorMark[] | null {
  let marks: readonly ProseMirrorMark[] | null = null;
  document.nodesBetween(from, to, (node, position) => {
    if (
      marks ||
      !node.isText ||
      position > from ||
      position + node.nodeSize < to
    ) {
      return;
    }
    marks = node.marks;
  });
  return marks;
}

function documentChangeMark(
  marks: readonly ProseMirrorMark[],
): ProseMirrorMark | undefined {
  return marks.find((mark) => mark.type.name === 'documentChange');
}
