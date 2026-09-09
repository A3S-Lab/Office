import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import {
  normalizeDocumentTableFormattingSnapshot,
  normalizeDocumentTableOverlap,
  normalizeDocumentTableStyleId,
  normalizeDocumentTableCellSpacing,
  serializeDocumentTableFormatting,
  type DocumentTableFormattingSnapshot,
} from './work-document-table-format-changes';
import { normalizeTableColor } from './work-document-table-borders';
import {
  normalizeDocumentTableFormattingBorders,
  orderedDocumentTableFormattingBorders,
} from './work-document-table-formatting-borders';
import {
  normalizeDocumentTableLook,
  orderedDocumentTableLook,
} from './work-document-table-look';

interface DocumentTableFormattingTrackingOptions {
  createChange: () => WorkDocumentChangeIdentity;
}

/**
 * When track-changes is on, table layout / alignment / preferred-width /
 * indent / default cell-margin / bidiVisual / solid-fill / tblLook / tblOverlap / tblStyle / tblCellSpacing / tblBorders edits become
 * reviewable `table-formatting` revisions.
 *
 * Layout-mode switches (`setDocumentTableLayoutMode`) also rewrite cell
 * `colwidth` / percentage attrs; those companions are ignored when deciding
 * whether the table body stayed the same so the geometry edit stays
 * reviewable.
 */
export function trackDocumentTableFormattingTransaction(
  transaction: Transaction,
  state: EditorState,
  options: DocumentTableFormattingTrackingOptions,
): boolean {
  let identity: WorkDocumentChangeIdentity | null = null;
  let tracked = false;
  state.doc.descendants((before, position) => {
    if (before.type.name !== 'table') return;
    const after = transaction.doc.nodeAt(position);
    if (!after || after.type.name !== 'table') return;
    if (after.attrs.tableChangeKind === 'table-formatting') return;
    if (after.attrs.propertyRevisionOmml) return;
    const previousSnapshot = geometryFormatting(before);
    const currentSnapshot = geometryFormatting(after);
    if (!previousSnapshot || !currentSnapshot) return;
    const previous = serializeDocumentTableFormatting(previousSnapshot);
    const current = serializeDocumentTableFormatting(currentSnapshot);
    if (previous === current) return;
    if (!sameTableContentIgnoringGeometry(before, after)) return;
    if (!onlyReviewableGeometryChanged(before, after)) return;
    identity ??= options.createChange();
    transaction.setNodeMarkup(position, undefined, {
      ...after.attrs,
      propertyRevisionOmml: null,
      tableChangeKind: 'table-formatting',
      tableChangeId: identity.id,
      tableChangeAuthor: identity.author || 'A3S Work',
      tableChangeDate: identity.date || new Date().toISOString(),
      tableChangeBefore: previous,
    });
    tracked = true;
  });
  return tracked;
}

function geometryFormatting(
  node: ProseMirrorNode,
): DocumentTableFormattingSnapshot | null {
  const geometry =
    node.attrs.geometry &&
    typeof node.attrs.geometry === 'object' &&
    !Array.isArray(node.attrs.geometry)
      ? (node.attrs.geometry as Record<string, unknown>)
      : null;
  if (!geometry) return null;
  return normalizeDocumentTableFormattingSnapshot({
    layout: geometry.layout,
    alignment: geometry.alignment,
    width: geometry.width,
    indent: geometry.indent,
    cellMargins: geometry.cellMargins,
    bidiVisual:
      typeof node.attrs.bidiVisual === 'boolean' ? node.attrs.bidiVisual : false,
    ...(normalizeTableColor(
      typeof node.attrs.fill === 'string' ? node.attrs.fill : null,
    )
      ? {
          fill: normalizeTableColor(
            typeof node.attrs.fill === 'string' ? node.attrs.fill : null,
          )!,
        }
      : {}),
    ...(normalizeDocumentTableLook(node.attrs.look)
      ? { look: orderedDocumentTableLook(normalizeDocumentTableLook(node.attrs.look)!) }
      : {}),
    ...(normalizeDocumentTableOverlap(node.attrs.overlap)
      ? { overlap: normalizeDocumentTableOverlap(node.attrs.overlap)! }
      : {}),
    ...(normalizeDocumentTableStyleId(node.attrs.styleId)
      ? { styleId: normalizeDocumentTableStyleId(node.attrs.styleId)! }
      : {}),
    ...(normalizeDocumentTableCellSpacing(node.attrs.cellSpacing) !== null
      ? {
          cellSpacing: normalizeDocumentTableCellSpacing(
            node.attrs.cellSpacing,
          )!,
        }
      : {}),
    ...(normalizeDocumentTableFormattingBorders(node.attrs.borders)
      ? {
          borders: orderedDocumentTableFormattingBorders(
            normalizeDocumentTableFormattingBorders(node.attrs.borders)!,
          ),
        }
      : {}),
  });
}

function onlyReviewableGeometryChanged(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  const beforeGeometry = geometryRecord(before);
  const afterGeometry = geometryRecord(after);
  if (!beforeGeometry || !afterGeometry) return false;
  const beforeRest = { ...beforeGeometry };
  const afterRest = { ...afterGeometry };
  delete beforeRest.layout;
  delete beforeRest.alignment;
  delete beforeRest.width;
  delete beforeRest.indent;
  delete beforeRest.cellMargins;
  delete afterRest.layout;
  delete afterRest.alignment;
  delete afterRest.width;
  delete afterRest.indent;
  delete afterRest.cellMargins;
  return JSON.stringify(beforeRest) === JSON.stringify(afterRest);
}

function geometryRecord(
  node: ProseMirrorNode,
): Record<string, unknown> | null {
  return node.attrs.geometry &&
    typeof node.attrs.geometry === 'object' &&
    !Array.isArray(node.attrs.geometry)
    ? (node.attrs.geometry as Record<string, unknown>)
    : null;
}

function sameTableContentIgnoringGeometry(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  return (
    JSON.stringify(tableSnapshotIgnoringGeometry(before)) ===
    JSON.stringify(tableSnapshotIgnoringGeometry(after))
  );
}

function tableSnapshotIgnoringGeometry(node: ProseMirrorNode): unknown {
  const attrs = { ...node.attrs };
  delete attrs.geometry;
  delete attrs.layoutMode;
  delete attrs.bidiVisual;
  delete attrs.fill;
  delete attrs.look;
  delete attrs.overlap;
  delete attrs.styleId;
  delete attrs.cellSpacing;
  delete attrs.borders;
  delete attrs.tableChangeKind;
  delete attrs.tableChangeId;
  delete attrs.tableChangeAuthor;
  delete attrs.tableChangeDate;
  delete attrs.tableChangeBefore;
  const content: unknown[] = [];
  node.forEach((child) => {
    content.push(nodeSnapshot(child));
  });
  return { type: node.type.name, attrs, content };
}

function nodeSnapshot(node: ProseMirrorNode): unknown {
  if (node.isText) return { type: node.type.name, text: node.text ?? '' };
  const content: unknown[] = [];
  node.forEach((child) => {
    content.push(nodeSnapshot(child));
  });
  return {
    type: node.type.name,
    attrs: snapshotAttrsIgnoringLayoutCompanions(node),
    ...(content.length ? { content } : {}),
  };
}

function snapshotAttrsIgnoringLayoutCompanions(
  node: ProseMirrorNode,
): Record<string, unknown> {
  const attrs = { ...node.attrs };
  if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
    delete attrs.colwidth;
    delete attrs.columnWidthPercentages;
  }
  return attrs;
}
