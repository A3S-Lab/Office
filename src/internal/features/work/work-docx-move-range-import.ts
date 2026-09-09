import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants } from './work-ooxml-package';

const MOVE_FROM = 'moveFrom';
const MOVE_TO = 'moveTo';
const FROM_START = 'moveFromRangeStart';
const FROM_END = 'moveFromRangeEnd';
const TO_START = 'moveToRangeStart';
const TO_END = 'moveToRangeEnd';
const MAX_MOVE_DATE_LENGTH = 64;
const START_ATTRIBUTES = new Set(['id', 'author', 'date', 'name']);
const END_ATTRIBUTES = new Set(['id']);
const SECTION_BREAK = 'sectPr';
const BLOCKING_CONTAINERS = new Set(['tbl', 'sdt', SECTION_BREAK]);

export interface DocxMoveRangeCompanion {
  rangeId: string;
  rangeName: string;
  from: Element;
  to: Element;
  markers: readonly Element[];
}

/**
 * Word often wraps bounded `w:moveFrom`/`w:moveTo` pairs with matching
 * `w:move*Range*` bookmarks. Companion bookmarks are relationship-free and
 * uniquely sandwich each supported move wrapper in document order within one
 * section (immediate siblings or same-section cross-paragraph placement).
 * They may be stripped after capture so the text move stays reviewable.
 */
export function companionDocxMoveRangeBookmarks(
  document: Document,
  movePairs: ReadonlyArray<{ from: Element; to: Element }>,
): DocxMoveRangeCompanion[] {
  const usedMarkers = new Set<Element>();
  const companions: DocxMoveRangeCompanion[] = [];
  for (const pair of movePairs) {
    const companion = companionForMovePair(pair.from, pair.to, usedMarkers);
    if (!companion) continue;
    for (const marker of companion.markers) usedMarkers.add(marker);
    companions.push(companion);
  }
  return companions;
}

export function stripDocxMoveRangeMarkers(markers: readonly Element[]): void {
  for (const marker of markers) marker.remove();
}

export function inspectUnpairedDocxMoveRangeMarkers(
  document: Document,
  companions: readonly DocxMoveRangeCompanion[],
): number {
  const companionMarkers = new Set(
    companions.flatMap((companion) => companion.markers),
  );
  return [
    ...descendants(document, FROM_START),
    ...descendants(document, FROM_END),
    ...descendants(document, TO_START),
    ...descendants(document, TO_END),
  ].filter(
    (marker) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(marker.namespaceURI ?? '') &&
      !companionMarkers.has(marker),
  ).length;
}

function companionForMovePair(
  from: Element,
  to: Element,
  usedMarkers: ReadonlySet<Element>,
): DocxMoveRangeCompanion | null {
  const fromBound = findSandwichMarkers(from, FROM_START, FROM_END);
  const toBound = findSandwichMarkers(to, TO_START, TO_END);
  if (!fromBound || !toBound) return null;
  const markers = [fromBound.start, fromBound.end, toBound.start, toBound.end];
  if (markers.some((marker) => usedMarkers.has(marker))) return null;
  if (new Set(markers).size !== 4) return null;
  const rangeId = wordAttribute(fromBound.start, 'id')?.trim() ?? '';
  const rangeName = wordAttribute(fromBound.start, 'name')?.trim() ?? '';
  const author = wordAttribute(fromBound.start, 'author')?.trim() ?? '';
  const date = normalizeDate(wordAttribute(fromBound.start, 'date'));
  if (
    !/^\d{1,10}$/.test(rangeId) ||
    !rangeName ||
    rangeName.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(rangeName) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author)
  ) {
    return null;
  }
  if (
    !sameStartMarker(fromBound.start, rangeId, rangeName, author, date) ||
    !sameStartMarker(toBound.start, rangeId, rangeName, author, date) ||
    !sameEndMarker(fromBound.end, rangeId) ||
    !sameEndMarker(toBound.end, rangeId)
  ) {
    return null;
  }
  const moveAuthor = wordAttribute(from, 'author')?.trim() ?? '';
  const moveDate = normalizeDate(wordAttribute(from, 'date'));
  if (moveAuthor !== author || moveDate !== date) return null;
  if (
    wordAttribute(to, 'author')?.trim() !== author ||
    normalizeDate(wordAttribute(to, 'date')) !== date
  ) {
    return null;
  }
  if (crossesSectionBoundary(markers)) return null;
  return { rangeId, rangeName, from, to, markers };
}

function findSandwichMarkers(
  move: Element,
  startLocalName: string,
  endLocalName: string,
): { start: Element; end: Element } | null {
  const immediateStart = previousSiblingNamed(move, startLocalName);
  const immediateEnd = nextSiblingNamed(move, endLocalName);
  if (immediateStart && immediateEnd) {
    return { start: immediateStart, end: immediateEnd };
  }

  const namespace = move.namespaceURI ?? '';
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return null;
  const story = storyRoot(move);
  if (!story) return null;

  const starts = descendants(story, startLocalName).filter(
    (element) => element.namespaceURI === namespace,
  );
  const ends = descendants(story, endLocalName).filter(
    (element) => element.namespaceURI === namespace,
  );
  const candidates: Array<{ start: Element; end: Element }> = [];
  for (const start of starts) {
    if (!precedes(start, move)) continue;
    const rangeId = wordAttribute(start, 'id')?.trim() ?? '';
    if (!/^\d{1,10}$/.test(rangeId)) continue;
    const matchingEnds = ends.filter(
      (end) =>
        precedes(move, end) &&
        wordAttribute(end, 'id')?.trim() === rangeId &&
        sandwichContainsOnlyMove(start, end, move),
    );
    if (matchingEnds.length !== 1) continue;
    const end = matchingEnds[0];
    if (!end) continue;
    candidates.push({ start, end });
  }
  return candidates.length === 1 ? (candidates[0] ?? null) : null;
}

/**
 * Every element strictly between the range bookmarks must lie on the move
 * wrapper's ancestor chain, be the wrapper, or be inside the wrapper. That
 * admits bookmarks placed around the containing paragraph while rejecting
 * extra sibling content or nested tracked ranges.
 */
function sandwichContainsOnlyMove(
  start: Element,
  end: Element,
  move: Element,
): boolean {
  let sawMove = false;
  for (const element of elementsBetween(start, end)) {
    if (element === move) {
      sawMove = true;
      continue;
    }
    if (move.contains(element)) continue;
    if (element.contains(move)) {
      if (BLOCKING_CONTAINERS.has(element.localName)) return false;
      continue;
    }
    return false;
  }
  return sawMove;
}

function elementsBetween(start: Element, end: Element): Element[] {
  const root = start.ownerDocument?.documentElement;
  if (!root || !start.ownerDocument) return [];
  const walker = start.ownerDocument.createTreeWalker(
    root,
    NodeFilter.SHOW_ELEMENT,
  );
  while (walker.currentNode !== start && walker.nextNode()) {
    // Seek the start bookmark.
  }
  if (walker.currentNode !== start) return [];
  const elements: Element[] = [];
  let node: Node | null = walker.nextNode();
  while (node && node !== end) {
    if (node instanceof Element) elements.push(node);
    node = walker.nextNode();
  }
  return node === end ? elements : [];
}

function crossesSectionBoundary(markers: readonly Element[]): boolean {
  const ordered = [...markers].sort(compareDocumentOrder);
  const first = ordered[0];
  const last = ordered[ordered.length - 1];
  if (!first || !last) return true;
  for (const element of elementsBetween(first, last)) {
    if (element.localName === SECTION_BREAK) return true;
  }
  return Boolean(ordered.some((marker) => marker.localName === SECTION_BREAK));
}

function storyRoot(element: Element): Element | null {
  let current: Element | null = element;
  while (current) {
    if (
      current.localName === 'body' ||
      current.localName === 'hdr' ||
      current.localName === 'ftr' ||
      current.localName === 'footnote' ||
      current.localName === 'endnote' ||
      current.localName === 'comments'
    ) {
      return current;
    }
    current = current.parentElement;
  }
  return element.ownerDocument?.documentElement ?? null;
}

function precedes(left: Element, right: Element): boolean {
  const position = left.compareDocumentPosition(right);
  return (position & Node.DOCUMENT_POSITION_FOLLOWING) !== 0;
}

function compareDocumentOrder(left: Element, right: Element): number {
  if (left === right) return 0;
  const position = left.compareDocumentPosition(right);
  if (position & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
  if (position & Node.DOCUMENT_POSITION_PRECEDING) return 1;
  return 0;
}

function sameStartMarker(
  element: Element,
  rangeId: string,
  rangeName: string,
  author: string,
  date: string,
): boolean {
  return (
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') &&
    wordAttribute(element, 'id')?.trim() === rangeId &&
    wordAttribute(element, 'name')?.trim() === rangeName &&
    wordAttribute(element, 'author')?.trim() === author &&
    normalizeDate(wordAttribute(element, 'date')) === date &&
    !hasUnsupportedAttributes(element, START_ATTRIBUTES)
  );
}

function sameEndMarker(element: Element, rangeId: string): boolean {
  return (
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') &&
    wordAttribute(element, 'id')?.trim() === rangeId &&
    !hasUnsupportedAttributes(element, END_ATTRIBUTES)
  );
}

function previousSiblingNamed(
  element: Element,
  localName: string,
): Element | null {
  const sibling = element.previousElementSibling;
  return sibling?.localName === localName &&
    sibling.namespaceURI === element.namespaceURI
    ? sibling
    : null;
}

function nextSiblingNamed(element: Element, localName: string): Element | null {
  const sibling = element.nextElementSibling;
  return sibling?.localName === localName &&
    sibling.namespaceURI === element.namespaceURI
    ? sibling
    : null;
}

function wordAttribute(element: Element, localName: string): string | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === localName &&
      xmlAttributeNamespace(element, candidate) === namespace,
  );
  return matches.length === 1 ? (matches[0]?.value ?? null) : null;
}

function hasUnsupportedAttributes(
  element: Element,
  supported: ReadonlySet<string>,
): boolean {
  const namespace = element.namespaceURI;
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !supported.has(xmlAttributeLocalName(candidate)),
  );
}

function normalizeDate(value: string | null): string {
  if (!value) return '';
  if (value.length > MAX_MOVE_DATE_LENGTH) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

export const DOCX_MOVE_RANGE_MARKER_NAMES = [
  FROM_START,
  FROM_END,
  TO_START,
  TO_END,
] as const;

export { MOVE_FROM, MOVE_TO };
