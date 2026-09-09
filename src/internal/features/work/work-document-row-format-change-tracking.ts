import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import {
  normalizeDocumentRowFormattingSnapshot,
  serializeDocumentRowFormatting,
} from './work-document-row-format-changes';
import {
  normalizeDocumentTableAlignment,
  normalizeDocumentTablePreferredWidth,
  type DocumentTableAlignment,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';
import {
  normalizeDocumentTableRowHeight,
  normalizeDocumentTableRowHeightRule,
} from './work-document-table-row';

interface DocumentRowFormattingTrackingOptions {
  createChange: () => WorkDocumentChangeIdentity;
}

/**
 * When track-changes is on, row cantSplit / repeatHeader / height / hidden /
 * alignment / gridBefore / gridAfter / widthBefore / widthAfter edits become reviewable `row-formatting` revisions.
 */
export function trackDocumentRowFormattingTransaction(
  transaction: Transaction,
  state: EditorState,
  options: DocumentRowFormattingTrackingOptions,
): boolean {
  let identity: WorkDocumentChangeIdentity | null = null;
  let tracked = false;
  state.doc.descendants((before, position) => {
    if (before.type.name !== 'tableRow') return;
    const after = transaction.doc.nodeAt(position);
    if (!after || after.type.name !== 'tableRow') return;
    if (after.attrs.rowChangeKind === 'row-formatting') return;
    if (after.attrs.propertyRevisionOmml) return;
    const previousSnapshot = rowFormatting(before);
    const currentSnapshot = rowFormatting(after);
    if (!previousSnapshot || !currentSnapshot) return;
    const previous = serializeDocumentRowFormatting(previousSnapshot);
    const current = serializeDocumentRowFormatting(currentSnapshot);
    if (previous === current) return;
    if (!sameRowContentIgnoringFormatting(before, after)) return;
    if (!onlyRowFormattingChanged(before, after)) return;
    identity ??= options.createChange();
    transaction.setNodeMarkup(position, undefined, {
      ...after.attrs,
      propertyRevisionOmml: null,
      rowChangeKind: 'row-formatting',
      rowChangeId: identity.id,
      rowChangeAuthor: identity.author || 'A3S Work',
      rowChangeDate: identity.date || new Date().toISOString(),
      rowChangeBefore: previous,
    });
    tracked = true;
  });
  return tracked;
}

function rowFormatting(node: ProseMirrorNode): {
  cantSplit?: boolean;
  repeatHeader?: boolean;
  height?: { value: number; rule: 'exact' | 'atLeast' };
  hidden?: boolean;
  alignment?: DocumentTableAlignment;
  gridBefore?: number;
  gridAfter?: number;
  widthBefore?: DocumentTablePreferredWidth;
  widthAfter?: DocumentTablePreferredWidth;
} | null {
  const height = normalizeDocumentTableRowHeight(node.attrs.rowHeight);
  const rule =
    height === null
      ? null
      : (normalizeDocumentTableRowHeightRule(node.attrs.rowHeightRule) ??
        'atLeast');
  const alignment =
    normalizeDocumentTableAlignment(node.attrs.alignment) ?? 'left';
  const gridBefore =
    typeof node.attrs.gridBefore === 'number' &&
    Number.isInteger(node.attrs.gridBefore) &&
    node.attrs.gridBefore >= 0
      ? node.attrs.gridBefore
      : 0;
  const gridAfter =
    typeof node.attrs.gridAfter === 'number' &&
    Number.isInteger(node.attrs.gridAfter) &&
    node.attrs.gridAfter >= 0
      ? node.attrs.gridAfter
      : 0;
  const widthBefore = normalizeDocumentTablePreferredWidth(
    node.attrs.widthBefore,
  ) ?? {
    type: 'auto' as const,
    value: null,
  };
  const widthAfter = normalizeDocumentTablePreferredWidth(
    node.attrs.widthAfter,
  ) ?? {
    type: 'auto' as const,
    value: null,
  };
  return normalizeDocumentRowFormattingSnapshot({
    cantSplit:
      typeof node.attrs.cantSplit === 'boolean' ? node.attrs.cantSplit : false,
    repeatHeader:
      typeof node.attrs.repeatHeader === 'boolean'
        ? node.attrs.repeatHeader
        : false,
    hidden: typeof node.attrs.hidden === 'boolean' ? node.attrs.hidden : false,
    alignment,
    gridBefore,
    gridAfter,
    widthBefore,
    widthAfter,
    ...(height !== null && rule ? { height: { value: height, rule } } : {}),
  });
}

function onlyRowFormattingChanged(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  const beforeAttrs = { ...before.attrs };
  const afterAttrs = { ...after.attrs };
  delete beforeAttrs.cantSplit;
  delete beforeAttrs.repeatHeader;
  delete beforeAttrs.rowHeight;
  delete beforeAttrs.rowHeightRule;
  delete beforeAttrs.hidden;
  delete beforeAttrs.alignment;
  delete beforeAttrs.gridBefore;
  delete beforeAttrs.gridAfter;
  delete beforeAttrs.widthBefore;
  delete beforeAttrs.widthAfter;
  delete beforeAttrs.rowChangeKind;
  delete beforeAttrs.rowChangeId;
  delete beforeAttrs.rowChangeAuthor;
  delete beforeAttrs.rowChangeDate;
  delete beforeAttrs.rowChangeBefore;
  delete beforeAttrs.propertyRevisionOmml;
  delete afterAttrs.cantSplit;
  delete afterAttrs.repeatHeader;
  delete afterAttrs.rowHeight;
  delete afterAttrs.rowHeightRule;
  delete afterAttrs.hidden;
  delete afterAttrs.alignment;
  delete afterAttrs.gridBefore;
  delete afterAttrs.gridAfter;
  delete afterAttrs.widthBefore;
  delete afterAttrs.widthAfter;
  delete afterAttrs.rowChangeKind;
  delete afterAttrs.rowChangeId;
  delete afterAttrs.rowChangeAuthor;
  delete afterAttrs.rowChangeDate;
  delete afterAttrs.rowChangeBefore;
  delete afterAttrs.propertyRevisionOmml;
  return JSON.stringify(beforeAttrs) === JSON.stringify(afterAttrs);
}

function sameRowContentIgnoringFormatting(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  if (before.childCount !== after.childCount) return false;
  for (let index = 0; index < before.childCount; index += 1) {
    if (!before.child(index).eq(after.child(index))) return false;
  }
  return true;
}
