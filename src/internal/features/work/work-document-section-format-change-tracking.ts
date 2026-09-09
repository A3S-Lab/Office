import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';
import type { WorkDocumentChangeIdentity } from './work-document-changes';
import {
  documentSectionLayoutFromNodeAttributes,
  type DocumentSectionNodeAttributes,
} from './work-document-section';
import {
  sectionFormattingSnapshotFromLayout,
  serializeDocumentSectionFormatting,
} from './work-document-section-format-changes';

interface DocumentSectionFormattingTrackingOptions {
  createChange: () => WorkDocumentChangeIdentity;
}

/**
 * When track-changes is on, section orientation, page-geometry, page-margin,
 * paper-source, equal-width column, and/or different-first-page edits become
 * reviewable `section-formatting` revisions. Companion `pageGeometry` / `pageSize`
 * swaps that follow orientation or geometry edits are allowed; other section
 * layout fields must stay unchanged.
 */
export function trackDocumentSectionFormattingTransaction(
  transaction: Transaction,
  state: EditorState,
  options: DocumentSectionFormattingTrackingOptions,
): boolean {
  let identity: WorkDocumentChangeIdentity | null = null;
  let tracked = false;
  state.doc.descendants((before, position) => {
    if (before.type.name !== 'documentSection') return;
    const after = transaction.doc.nodeAt(position);
    if (!after || after.type.name !== 'documentSection') return;
    if (after.attrs.sectionChangeKind === 'section-formatting') return;
    if (after.attrs.propertyRevisionOmml) return;
    const previousSnapshot = sectionFormatting(before);
    const currentSnapshot = sectionFormatting(after);
    if (!previousSnapshot || !currentSnapshot) return;
    const previous = serializeDocumentSectionFormatting(previousSnapshot);
    const current = serializeDocumentSectionFormatting(currentSnapshot);
    if (previous === current) return;
    if (!sameSectionContentIgnoringFormatting(before, after)) return;
    if (!onlyReviewableSectionFormattingChanged(before, after)) return;
    identity ??= options.createChange();
    transaction.setNodeMarkup(position, undefined, {
      ...after.attrs,
      propertyRevisionOmml: null,
      sectionChangeKind: 'section-formatting',
      sectionChangeId: identity.id,
      sectionChangeAuthor: identity.author || 'A3S Work',
      sectionChangeDate: identity.date || new Date().toISOString(),
      sectionChangeBefore: previous,
    });
    tracked = true;
  });
  return tracked;
}

function sectionFormatting(node: ProseMirrorNode) {
  const layout = documentSectionLayoutFromNodeAttributes(
    node.attrs as Partial<DocumentSectionNodeAttributes>,
  );
  return sectionFormattingSnapshotFromLayout(layout);
}

function onlyReviewableSectionFormattingChanged(
  before: ProseMirrorNode,
  after: ProseMirrorNode,
): boolean {
  return (
    JSON.stringify(layoutSnapshotIgnoringFormatting(before)) ===
    JSON.stringify(layoutSnapshotIgnoringFormatting(after))
  );
}

function layoutSnapshotIgnoringFormatting(node: ProseMirrorNode): unknown {
  const layout = documentSectionLayoutFromNodeAttributes(
    node.attrs as Partial<DocumentSectionNodeAttributes>,
  );
  const {
    orientation: _orientation,
    pageGeometry: _pageGeometry,
    pageSize: _pageSize,
    pageMargins: _pageMargins,
    margins: _margins,
    paperSource: _paperSource,
    columns: _columns,
    pageChrome,
    propertyRevisionOmml: _propertyRevisionOmml,
    formattingChange: _formattingChange,
    ...rest
  } = layout;
  // differentFirstPage is reviewable via section-formatting; ignore its flips.
  const chrome =
    pageChrome && typeof pageChrome === 'object' && !Array.isArray(pageChrome)
      ? { ...pageChrome, differentFirstPage: false as const }
      : pageChrome;
  return { ...rest, pageChrome: chrome };
}

function sameSectionContentIgnoringFormatting(
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
