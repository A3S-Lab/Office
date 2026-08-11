import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { EditorState, Transaction } from '@tiptap/pm/state';

export interface WorkDocumentBookmarkReferenceTarget {
  type: 'bookmark';
  id: string;
  name: string;
  title: string;
  display: string;
  instruction: string;
}

export interface WorkDocumentBookmarkReferenceRange {
  id: string;
  name: string;
  from: number;
  to: number;
}

export interface WorkDocumentBookmarkReferenceRename {
  from: number;
  to: number;
  previousId: string;
  nextId: string;
  previousName: string;
  nextName: string;
}

const BOOKMARK_REFERENCE_SELECTOR =
  'span[data-document-cross-reference][data-reference-target-type="bookmark"]';
const BLOCK_TEXT_CONTAINER_NAMES = new Set([
  'ADDRESS',
  'BLOCKQUOTE',
  'DIV',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'LI',
  'P',
  'PRE',
  'TD',
  'TH',
]);

export function docxBookmarkReferenceTarget(
  instruction: string,
): string | null {
  return /^\s*REF\s+([^\s\\]+)/i.exec(instruction)?.[1] ?? null;
}

export function supportedDocxBookmarkReferenceInstruction(
  instruction: string,
): boolean {
  if (!docxBookmarkReferenceTarget(instruction)) return false;
  const switches = Array.from(
    instruction.matchAll(/\\([^\s\\])/g),
    ([, value]) => value.toLowerCase(),
  );
  return switches.every((value) => value === 'h' || value === '*');
}

export function documentBookmarkReferenceInstruction(
  name: string,
  source = '',
): string {
  const target = name.trim();
  if (!target) return source.trim();
  const hyperlink = !source.trim() || /(?:^|\s)\\h(?:\s|$)/i.test(source);
  return `REF ${target}${hyperlink ? ' \\h' : ''}`;
}

export function documentBookmarkReferenceDisplay(
  document: ProseMirrorNode,
  bookmark: Pick<WorkDocumentBookmarkReferenceRange, 'from' | 'to' | 'name'>,
): string {
  return (
    normalizedReferenceText(
      document.textBetween(bookmark.from + 1, bookmark.to, ' ', ' '),
    ) || bookmark.name
  );
}

export function synchronizeDocumentBookmarkReferenceNodes(
  state: EditorState,
  transaction: Transaction,
  bookmarks: readonly WorkDocumentBookmarkReferenceRange[],
  renames: readonly WorkDocumentBookmarkReferenceRename[],
): void {
  const referenceType = state.schema.nodes.documentCrossReference;
  if (!referenceType) return;
  const byId = new Map(bookmarks.map((bookmark) => [bookmark.id, bookmark]));
  const byName = new Map(
    bookmarks.map((bookmark) => [bookmark.name.toLowerCase(), bookmark]),
  );
  state.doc.descendants((node, position) => {
    if (node.type !== referenceType || node.attrs.targetType !== 'bookmark') {
      return;
    }
    let targetId = stringAttribute(node.attrs.targetId);
    let targetName = stringAttribute(node.attrs.targetName);
    const rename = renames
      .filter(
        (candidate) =>
          position > candidate.from &&
          position < candidate.to &&
          ((targetId && targetId === candidate.previousId) ||
            (targetName &&
              targetName.toLowerCase() ===
                candidate.previousName.toLowerCase())),
      )
      .sort((left, right) => left.to - left.from - (right.to - right.from))[0];
    if (rename) {
      targetId = rename.nextId;
      targetName = rename.nextName;
    }
    const target =
      byId.get(targetId) ?? byName.get(targetName.toLowerCase()) ?? null;
    const nextTargetId = target?.id ?? targetId;
    const nextTargetName = target?.name ?? targetName;
    const display = target
      ? documentBookmarkReferenceDisplay(state.doc, target)
      : stringAttribute(node.attrs.display) || targetName || '引用缺失';
    const instruction = documentBookmarkReferenceInstruction(
      nextTargetName,
      stringAttribute(node.attrs.instruction),
    );
    const orphaned = !target;
    if (
      node.attrs.targetId === nextTargetId &&
      node.attrs.targetName === nextTargetName &&
      node.attrs.display === display &&
      node.attrs.instruction === instruction &&
      node.attrs.orphaned === orphaned
    ) {
      return;
    }
    transaction.setNodeMarkup(position, undefined, {
      ...node.attrs,
      targetId: nextTargetId,
      targetName: nextTargetName,
      display,
      instruction,
      orphaned,
    });
  });
}

export function normalizeDocumentBookmarkReferencesHtml(
  source: string,
): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const bookmarks = domBookmarkReferenceTargets(document);
  const byId = new Map(bookmarks.map((bookmark) => [bookmark.id, bookmark]));
  const byName = new Map(
    bookmarks.map((bookmark) => [bookmark.name.toLowerCase(), bookmark]),
  );
  for (const reference of document.body.querySelectorAll<HTMLElement>(
    BOOKMARK_REFERENCE_SELECTOR,
  )) {
    const targetId = reference.dataset.referenceTargetId?.trim() ?? '';
    const targetName = reference.dataset.referenceTargetName?.trim() ?? '';
    const target =
      byId.get(targetId) ?? byName.get(targetName.toLowerCase()) ?? null;
    applyDomBookmarkReference(
      reference,
      target?.id ?? targetId,
      target?.name ?? targetName,
      target?.display ??
        reference.dataset.referenceDisplay?.trim() ??
        targetName,
      !target,
    );
  }
  return document.body.innerHTML;
}

export function retargetDomBookmarkReferences(
  root: HTMLElement,
  start: HTMLElement,
  end: HTMLElement,
  previous: { id: string; name: string },
  next: { id: string; name: string },
): void {
  const range = root.ownerDocument.createRange();
  range.setStartAfter(start);
  range.setEndBefore(end);
  for (const reference of root.querySelectorAll<HTMLElement>(
    BOOKMARK_REFERENCE_SELECTOR,
  )) {
    if (!range.intersectsNode(reference)) continue;
    const id = reference.dataset.referenceTargetId?.trim() ?? '';
    const name = reference.dataset.referenceTargetName?.trim() ?? '';
    if (
      id !== previous.id &&
      name.toLowerCase() !== previous.name.toLowerCase()
    ) {
      continue;
    }
    reference.dataset.referenceTargetId = next.id;
    reference.dataset.referenceTargetName = next.name;
    reference.dataset.referenceInstruction =
      documentBookmarkReferenceInstruction(
        next.name,
        reference.dataset.referenceInstruction,
      );
  }
}

function domBookmarkReferenceTargets(document: Document): Array<{
  id: string;
  name: string;
  display: string;
}> {
  const open = new Map<string, HTMLElement[]>();
  const targets: Array<{ id: string; name: string; display: string }> = [];
  for (const boundary of document.body.querySelectorAll<HTMLElement>(
    'span[data-document-bookmark-boundary]',
  )) {
    const id = boundary.dataset.bookmarkId?.trim() ?? '';
    if (!id) continue;
    if (boundary.dataset.bookmarkKind !== 'end') {
      const stack = open.get(id) ?? [];
      stack.push(boundary);
      open.set(id, stack);
      continue;
    }
    const start = open.get(id)?.pop();
    const name = start?.dataset.bookmarkName?.trim() ?? '';
    if (!start || !name) continue;
    targets.push({
      id,
      name,
      display: normalizedReferenceText(domTextBetween(start, boundary)) || name,
    });
  }
  return targets;
}

function domTextBetween(start: HTMLElement, end: HTMLElement): string {
  const root = start.ownerDocument.body;
  const walker = start.ownerDocument.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
  );
  const values: string[] = [];
  let previousBlock: Element | null = null;
  while (walker.nextNode()) {
    const node = walker.currentNode;
    const followsStart = Boolean(
      start.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING,
    );
    const precedesEnd = Boolean(
      node.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING,
    );
    if (!followsStart || !precedesEnd) continue;
    const block = closestTextBlock(node.parentElement, root);
    if (values.length && previousBlock && block !== previousBlock) {
      values.push(' ');
    }
    values.push(node.textContent ?? '');
    previousBlock = block;
  }
  return values.join('');
}

function closestTextBlock(
  element: Element | null,
  root: HTMLElement,
): Element | null {
  let current = element;
  while (current && current !== root) {
    if (BLOCK_TEXT_CONTAINER_NAMES.has(current.tagName)) return current;
    current = current.parentElement;
  }
  return root;
}

function applyDomBookmarkReference(
  reference: HTMLElement,
  targetId: string,
  targetName: string,
  display: string,
  orphaned: boolean,
): void {
  reference.dataset.documentCrossReference = 'true';
  reference.dataset.referenceTargetType = 'bookmark';
  reference.dataset.referenceTargetId = targetId;
  reference.dataset.referenceTargetName = targetName;
  reference.dataset.referenceInstruction = documentBookmarkReferenceInstruction(
    targetName,
    reference.dataset.referenceInstruction,
  );
  reference.dataset.referenceDisplay = display;
  reference.classList.add('work-document-cross-reference');
  reference.textContent = orphaned ? '引用缺失' : display;
  if (orphaned) reference.dataset.referenceOrphaned = 'true';
  else delete reference.dataset.referenceOrphaned;
}

function normalizedReferenceText(value: string | null | undefined): string {
  return value?.replace(/\s+/g, ' ').trim() ?? '';
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}
