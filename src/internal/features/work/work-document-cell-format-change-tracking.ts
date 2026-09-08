import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import {
  normalizeDocumentCellFormattingSnapshot,
  normalizeDocumentTableCellTextDirection,
  preferredWidthFromCellAttributes,
  serializeDocumentCellFormatting,
  type DocumentTableCellTextDirection,
} from './work-document-cell-format-changes';
import { normalizeTableColor } from './work-document-table-borders';
import {
  normalizeDocumentTableVerticalAlign,
  type DocumentTableVerticalAlign,
} from './work-document-table-cell-formatting';
import {
  normalizeDocumentTableCellMarginOverrides,
  type DocumentTableCellMarginOverrides,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';

interface DocumentCellFormattingTrackingOptions {
  createChange: () => WorkDocumentChangeIdentity;
}

/**
 * When track-changes is on, cell verticalAlign / solid fill / margin /
 * preferred-width / noWrap / textDirection / fitText / hideMark edits become reviewable
 * `cell-formatting` revisions.
 */
export function trackDocumentCellFormattingTransaction(
  transaction: Transaction,
  state: EditorState,
  options: DocumentCellFormattingTrackingOptions,
): boolean {
  let identity: WorkDocumentChangeIdentity | null = null;
  let tracked = false;
  state.doc.descendants((before, position) => {
    if (!isTableCell(before)) return;
    const after = transaction.doc.nodeAt(position);
    if (!after || !isTableCell(after)) return;
    if (after.attrs.cellChangeKind === 'cell-formatting') return;
    if (after.attrs.propertyRevisionOmml) return;
    if (after.attrs.themeFill) return;
    const previousSnapshot = cellFormatting(before);
    const currentSnapshot = cellFormatting(after);
    if (!previousSnapshot || !currentSnapshot) return;
    const previous = serializeDocumentCellFormatting(previousSnapshot);
    const current = serializeDocumentCellFormatting(currentSnapshot);
    if (previous === current) return;
    if (!sameCellContentIgnoringFormatting(before, after)) return;
    if (!onlyReviewableCellFormattingChanged(before, after)) return;
    identity ??= options.createChange();
    transaction.setNodeMarkup(position, undefined, {
      ...after.attrs,
      propertyRevisionOmml: null,
      cellChangeKind: 'cell-formatting',
      cellChangeId: identity.id,
      cellChangeAuthor: identity.author || 'A3S Work',
      cellChangeDate: identity.date || new Date().toISOString(),
      cellChangeBefore: previous,
    });
    tracked = true;
  });
  return tracked;
}

function cellFormatting(node: ProseMirrorNode): {
  verticalAlign?: DocumentTableVerticalAlign;
  fill?: string;
  margins?: DocumentTableCellMarginOverrides;
  width?: DocumentTablePreferredWidth;
  noWrap?: boolean;
  textDirection?: DocumentTableCellTextDirection;
  fitText?: boolean;
  hideMark?: boolean;
} | null {
  const verticalAlign =
    normalizeDocumentTableVerticalAlign(node.attrs.verticalAlign) ?? 'top';
  const fill =
    normalizeTableColor(
      typeof node.attrs.backgroundColor === 'string'
        ? node.attrs.backgroundColor
        : null,
    ) ?? undefined;
  const margins = normalizeDocumentTableCellMarginOverrides(node.attrs.margins);
  const width = preferredWidthFromCellAttributes(node.attrs) ?? undefined;
  const noWrap =
    typeof node.attrs.noWrap === 'boolean' ? node.attrs.noWrap : false;
  const textDirection =
    normalizeDocumentTableCellTextDirection(node.attrs.textDirection) ?? 'lrTb';
  const fitText =
    typeof node.attrs.fitText === 'boolean' ? node.attrs.fitText : false;
  const hideMark =
    typeof node.attrs.hideMark === 'boolean' ? node.attrs.hideMark : false;
  return normalizeDocumentCellFormattingSnapshot({
    verticalAlign,
    ...(fill ? { fill } : {}),
    ...(margins ? { margins } : {}),
    ...(width ? { width } : {}),
    noWrap,
    textDirection,
    fitText,
    hideMark,
  });
}

function onlyReviewableCellFormattingChanged(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  const beforeAttrs = { ...before.attrs };
  const afterAttrs = { ...after.attrs };
  for (const key of [
    'verticalAlign',
    'backgroundColor',
    'margins',
    'colwidth',
    'columnWidthPercentages',
    'noWrap',
    'textDirection',
    'fitText',
    'hideMark',
    'cellChangeKind',
    'cellChangeId',
    'cellChangeAuthor',
    'cellChangeDate',
    'cellChangeBefore',
    'propertyRevisionOmml',
  ] as const) {
    delete beforeAttrs[key];
    delete afterAttrs[key];
  }
  return JSON.stringify(beforeAttrs) === JSON.stringify(afterAttrs);
}

function sameCellContentIgnoringFormatting(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  if (before.childCount !== after.childCount) return false;
  for (let index = 0; index < before.childCount; index += 1) {
    const left = before.child(index);
    const right = after.child(index);
    if (!left || !right || !left.eq(right)) return false;
  }
  return true;
}

function isTableCell(node: ProseMirrorNode): boolean {
  return node.type.name === 'tableCell' || node.type.name === 'tableHeader';
}
