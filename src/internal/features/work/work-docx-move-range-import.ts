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
const ALWAYS_BLOCKED_CONTAINERS = new Set(['sdt', SECTION_BREAK]);
const TABLE_CHROME = new Set([
  'tblPr',
  'tblGrid',
  'tblGridCol',
  'trPr',
  'tcPr',
]);
const OFF_PATH_CONTENT = new Set(['p', MOVE_FROM, MOVE_TO, 'ins', 'del']);
const TRACKED_REVISION_NAMES = new Set([MOVE_FROM, MOVE_TO, 'ins', 'del']);
const BOOKMARK_START_ATTRIBUTES = new Set(['id', 'name']);
const BOOKMARK_END_ATTRIBUTES = new Set(['id']);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);

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
 * uniquely sandwich each supported move wrapper in document order (immediate
 * siblings, cross-paragraph placement including across section breaks, or one
 * table with the move in exactly one cell and allowlisted chrome / sibling
 * cell text-only content). They may be stripped after capture so the text
 * move stays reviewable. Nested-table, SDT, section-sandwich, tracked or
 * rich sibling-cell content, and unpaired markers stay fail-closed. */
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
  // Each side's sandwich must stay move-only aside from allowlisted table
  // chrome and sibling-cell text-only content (nested tables, SDT, section
  // breaks, and tracked/rich sibling cells remain fail-closed via
  // sandwichContainsOnlyMove). The destination may live in a later section
  // than the source; that does not block companion admission.
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
 * wrapper's ancestor chain, be the wrapper, be inside the wrapper, be
 * allowlisted table chrome on the move table, or be untracked text-only
 * content in a sibling cell of that same table. At most one `w:tbl` may
 * appear, it must contain the move in exactly one `w:tc`, and `w:sdt` /
 * `w:sectPr` stay blocked. Nested tables, tracked revisions, and rich
 * sibling-cell content reject.
 */
function sandwichContainsOnlyMove(
  start: Element,
  end: Element,
  move: Element,
): boolean {
  const moveTable = nearestAncestorNamed(move, 'tbl');
  const moveCell = nearestAncestorNamed(move, 'tc');
  if (moveTable && !moveCell) return false;

  let sawMove = false;
  let tableCount = 0;
  for (const element of elementsBetween(start, end)) {
    if (element === move) {
      sawMove = true;
      continue;
    }
    if (move.contains(element)) continue;
    if (ALWAYS_BLOCKED_CONTAINERS.has(element.localName)) return false;
    if (element.localName === 'tbl') {
      tableCount += 1;
      if (tableCount > 1 || !element.contains(move)) return false;
      continue;
    }
    if (element.contains(move)) continue;
    if (isTableChromeOnMovePath(element, moveTable)) {
      continue;
    }
    if (isAdmittedSiblingTableContent(element, moveTable, moveCell)) {
      continue;
    }
    if (OFF_PATH_CONTENT.has(element.localName)) return false;
    return false;
  }
  return sawMove;
}

function nearestAncestorNamed(
  element: Element,
  localName: string,
): Element | null {
  let current: Element | null = element.parentElement;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function isTableChromeOnMovePath(
  element: Element,
  moveTable: Element | null,
): boolean {
  const chrome = nearestTableChrome(element);
  if (!chrome || chromeContainsTrackedRevision(chrome)) return false;
  const parent = chrome.parentElement;
  switch (chrome.localName) {
    case 'tcPr':
      return (
        moveTable !== null &&
        parent?.localName === 'tc' &&
        moveTable.contains(parent)
      );
    case 'trPr':
      return (
        moveTable !== null &&
        parent?.localName === 'tr' &&
        moveTable.contains(parent)
      );
    case 'tblPr':
    case 'tblGrid':
      return moveTable !== null && parent === moveTable;
    case 'tblGridCol':
      return (
        moveTable !== null &&
        parent?.localName === 'tblGrid' &&
        parent.parentElement === moveTable
      );
    default:
      return false;
  }
}

/**
 * Sibling cells (and rows that do not contain the move) in the one admitted
 * table may carry untracked text-only paragraphs, empty/`rPr`-only runs, and
 * relationship-free bookmarks. Drawings, hyperlinks, nested containers, and
 * tracked revisions stay fail-closed.
 */
function isAdmittedSiblingTableContent(
  element: Element,
  moveTable: Element | null,
  moveCell: Element | null,
): boolean {
  if (!moveTable || !moveCell) return false;
  if (!moveTable.contains(element) || moveCell.contains(element)) return false;
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
    return false;
  }
  if (TRACKED_REVISION_NAMES.has(element.localName)) return false;

  switch (element.localName) {
    case 'tr':
    case 'tc':
    case 'p':
      return true;
    case 'pPr':
    case 'rPr':
      return isWordprocessingOnlySubtree(element);
    case 'r':
      return isEmptyOrTextOnlyUntrackedRun(element);
    case 't':
      return element.children.length === 0;
    case 'bookmarkStart':
    case 'bookmarkEnd':
      return isRelationshipFreeBookmarkMarker(element);
    default:
      return false;
  }
}

function isEmptyOrTextOnlyUntrackedRun(run: Element): boolean {
  const properties = Array.from(run.children).filter(
    (child) => child.localName === 'rPr',
  );
  if (properties.length > 1) return false;
  for (const child of Array.from(run.children)) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return false;
    }
    if (TRACKED_REVISION_NAMES.has(child.localName)) return false;
    if (child.localName === 'rPr') {
      if (!isWordprocessingOnlySubtree(child)) return false;
      continue;
    }
    if (child.localName === 't') {
      if (child.children.length > 0) return false;
      continue;
    }
    return false;
  }
  return true;
}

function isWordprocessingOnlySubtree(element: Element): boolean {
  return !Array.from(element.querySelectorAll('*')).some(
    (descendant) =>
      !DOCX_WORDPROCESSING_NAMESPACES.has(descendant.namespaceURI ?? ''),
  );
}

function isRelationshipFreeBookmarkMarker(element: Element): boolean {
  if (
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') ||
    element.children.length > 0
  ) {
    return false;
  }
  const allowed =
    element.localName === 'bookmarkStart'
      ? BOOKMARK_START_ATTRIBUTES
      : element.localName === 'bookmarkEnd'
        ? BOOKMARK_END_ATTRIBUTES
        : null;
  if (!allowed) return false;
  let hasId = false;
  let hasName = element.localName !== 'bookmarkStart';
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return false;
    if (namespace && namespace !== element.namespaceURI) return false;
    const localName = xmlAttributeLocalName(attribute);
    if (!(namespace === element.namespaceURI || !namespace)) return false;
    if (!allowed.has(localName)) return false;
    const value = attribute.value.trim();
    if (!value || value.length > 255 || /[\u0000-\u001f\u007f]/.test(value)) {
      return false;
    }
    if (localName === 'id') {
      if (!/^\+?\d{1,10}$/.test(value)) return false;
      hasId = true;
    }
    if (localName === 'name') hasName = true;
  }
  return hasId && hasName;
}

function nearestTableChrome(element: Element): Element | null {
  let current: Element | null = element;
  while (current) {
    if (TABLE_CHROME.has(current.localName)) return current;
    if (
      current.localName === 'tbl' ||
      current.localName === 'tr' ||
      current.localName === 'tc' ||
      current.localName === 'p' ||
      current.localName === 'body' ||
      current.localName === 'hdr' ||
      current.localName === 'ftr'
    ) {
      return null;
    }
    current = current.parentElement;
  }
  return null;
}

function chromeContainsTrackedRevision(chrome: Element): boolean {
  return (
    descendants(chrome, MOVE_FROM).length > 0 ||
    descendants(chrome, MOVE_TO).length > 0 ||
    descendants(chrome, 'ins').length > 0 ||
    descendants(chrome, 'del').length > 0
  );
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
