import type {
  WorkDocumentChangeIdentity,
  WorkDocumentChangeKind,
  WorkDocumentMoveRole,
} from './work-document-changes';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  companionDocxMoveRangeBookmarks,
  inspectUnpairedDocxMoveRangeMarkers,
  stripDocxMoveRangeMarkers,
} from './work-docx-move-range-import';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute, descendants } from './work-ooxml-package';

export interface ImportedDocxChangeMarker extends WorkDocumentChangeIdentity {
  kind: WorkDocumentChangeKind;
  moveRole?: WorkDocumentMoveRole;
  moveRangeId?: string;
  moveRangeName?: string;
  start: string;
  end: string;
}

export interface ImportedDocxChangeMarkers {
  changes: ImportedDocxChangeMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MOVE_FROM = 'moveFrom';
const MOVE_TO = 'moveTo';
const MAX_MOVE_REVISIONS = 65_536;
const MAX_MOVE_TEXT_LENGTH = 1_000_000;
const MAX_MOVE_DATE_LENGTH = 64;

interface SupportedMoveRevision {
  element: Element;
  role: WorkDocumentMoveRole;
  id: string;
  author: string;
  date: string;
  text: string;
}

interface SupportedMovePair {
  from: SupportedMoveRevision;
  to: SupportedMoveRevision;
}

export function markDocxTextChanges(
  document: Document,
): ImportedDocxChangeMarkers {
  const changes: ImportedDocxChangeMarker[] = [];
  const movePairs = supportedMovePairs(document);
  const companions = companionDocxMoveRangeBookmarks(
    document,
    movePairs.map((pair) => ({
      from: pair.from.element,
      to: pair.to.element,
    })),
  );
  const companionByMove = new Map(
    companions.flatMap((companion) => [
      [companion.from, companion] as const,
      [companion.to, companion] as const,
    ]),
  );
  stripDocxMoveRangeMarkers(
    companions.flatMap((companion) => companion.markers),
  );
  const moveByElement = new Map<Element, SupportedMovePair>();
  for (const pair of movePairs) {
    moveByElement.set(pair.from.element, pair);
    moveByElement.set(pair.to.element, pair);
  }
  const processedMoves = new Set<SupportedMovePair>();
  let markerIndex = 0;
  const revisions = [
    ...descendants(document, 'ins'),
    ...descendants(document, 'del'),
    ...movePairs.flatMap((pair) => [pair.from.element, pair.to.element]),
  ].sort(compareDocumentOrder);
  for (const revision of revisions) {
    if (!revision.parentNode || closestRevision(revision.parentElement))
      continue;
    const move = moveByElement.get(revision);
    if (move) {
      if (processedMoves.has(move)) continue;
      processedMoves.add(move);
      const id = uniqueChangeId(`docx-move-${move.from.id}`, changes);
      const ordered = [move.from, move.to].sort((left, right) =>
        compareDocumentOrder(left.element, right.element),
      );
      const companion = companionByMove.get(move.from.element);
      for (const side of ordered) {
        markerIndex += 1;
        if (side.role === 'from') convertDeletedText(document, side.element);
        const marker: ImportedDocxChangeMarker = {
          id,
          kind: 'move',
          moveRole: side.role,
          author: side.author,
          date: side.date,
          start: `__A3S_WORK_CHANGE_START_${markerIndex}__`,
          end: `__A3S_WORK_CHANGE_END_${markerIndex}__`,
          ...(companion
            ? {
                moveRangeId: companion.rangeId,
                moveRangeName: companion.rangeName,
              }
            : {}),
        };
        unwrapRevision(document, side.element, marker.start, marker.end);
        changes.push(marker);
      }
      continue;
    }
    const kind: WorkDocumentChangeKind =
      revision.localName === 'del' ? 'deletion' : 'insertion';
    if (!revisionText(revision, kind)) continue;
    if (kind === 'deletion') convertDeletedText(document, revision);
    markerIndex += 1;
    const sourceId = attribute(revision, 'id')?.trim() ?? '';
    const id = uniqueChangeId(
      sourceId ? `docx-change-${sourceId}` : `docx-change-${markerIndex}`,
      changes,
    );
    const marker: ImportedDocxChangeMarker = {
      id,
      kind,
      author: attribute(revision, 'author')?.trim() || '未知审阅者',
      date: normalizeDate(attribute(revision, 'date')),
      start: `__A3S_WORK_CHANGE_START_${markerIndex}__`,
      end: `__A3S_WORK_CHANGE_END_${markerIndex}__`,
    };
    unwrapRevision(document, revision, marker.start, marker.end);
    changes.push(marker);
  }
  return { changes };
}

export function applyImportedDocxChangeMarkers(
  document: Document,
  markers: ImportedDocxChangeMarkers,
): void {
  for (const marker of markers.changes) {
    const isMoveFrom = marker.kind === 'move' && marker.moveRole === 'from';
    const element = document.createElement(
      marker.kind === 'deletion' || isMoveFrom ? 'del' : 'ins',
    );
    element.dataset.documentChange = 'true';
    element.dataset.changeKind = marker.kind;
    if (marker.kind === 'move' && marker.moveRole) {
      element.dataset.changeMoveRole = marker.moveRole;
    }
    if (marker.kind === 'move' && marker.moveRangeId && marker.moveRangeName) {
      element.dataset.changeMoveRangeId = marker.moveRangeId;
      element.dataset.changeMoveRangeName = marker.moveRangeName;
    }
    element.dataset.changeId = marker.id;
    element.dataset.changeAuthor = marker.author;
    element.dataset.changeDate = marker.date;
    wrapMarkerRange(document.body, marker.start, marker.end, element);
  }
}

export function hasImportedDocxChangeMarkers(
  markers: ImportedDocxChangeMarkers,
): boolean {
  return markers.changes.length > 0;
}

function revisionText(revision: Element, kind: WorkDocumentChangeKind): string {
  const names = kind === 'deletion' ? ['delText', 't'] : ['t'];
  return names
    .flatMap((name) => descendants(revision, name))
    .map((element) => element.textContent ?? '')
    .join('');
}

/**
 * Returns whether one move wrapper uses the bounded, text-only Word shape.
 * Unpaired or non-companion range-marker moves and relationship-bearing
 * children intentionally remain on the compatibility path.
 */
export function isSupportedDocxMoveChange(element: Element): boolean {
  return supportedMoveRevision(element) !== null;
}

/** Counts complete, paired move revisions in a WordprocessingML story. */
export function supportedDocxMovePairCount(document: Document): number {
  return supportedMovePairs(document).length;
}

/** Companion range bookmarks paired with supported move wrappers. */
export function inspectDocxMoveRangeCompanions(document: Document): {
  companionCount: number;
  unpairedMarkerCount: number;
} {
  const pairs = supportedMovePairs(document);
  const companions = companionDocxMoveRangeBookmarks(
    document,
    pairs.map((pair) => ({
      from: pair.from.element,
      to: pair.to.element,
    })),
  );
  return {
    companionCount: companions.length,
    unpairedMarkerCount: inspectUnpairedDocxMoveRangeMarkers(
      document,
      companions,
    ),
  };
}

function supportedMovePairs(document: Document): SupportedMovePair[] {
  const from = descendants(document, MOVE_FROM)
    .map((element) => supportedMoveRevision(element, 'from'))
    .filter((value): value is SupportedMoveRevision => value !== null);
  const to = descendants(document, MOVE_TO)
    .map((element) => supportedMoveRevision(element, 'to'))
    .filter((value): value is SupportedMoveRevision => value !== null);
  const fromById = uniqueMoveById(from);
  const toById = uniqueMoveById(to);
  const pairs: SupportedMovePair[] = [];
  for (const [id, source] of fromById) {
    const destination = toById.get(id);
    if (
      !destination ||
      source.author !== destination.author ||
      source.date !== destination.date ||
      source.text !== destination.text ||
      closestRevision(source.element.parentElement) ||
      closestRevision(destination.element.parentElement)
    ) {
      continue;
    }
    pairs.push({ from: source, to: destination });
  }
  return pairs.slice(0, MAX_MOVE_REVISIONS);
}

function uniqueMoveById(
  revisions: readonly SupportedMoveRevision[],
): Map<string, SupportedMoveRevision> {
  const result = new Map<string, SupportedMoveRevision>();
  const duplicates = new Set<string>();
  for (const revision of revisions) {
    if (duplicates.has(revision.id)) continue;
    if (result.has(revision.id)) {
      result.delete(revision.id);
      duplicates.add(revision.id);
      continue;
    }
    result.set(revision.id, revision);
  }
  return result;
}

function supportedMoveRevision(
  element: Element,
  role?: WorkDocumentMoveRole,
): SupportedMoveRevision | null {
  const resolvedRole =
    role ??
    (element.localName === MOVE_FROM
      ? 'from'
      : element.localName === MOVE_TO
        ? 'to'
        : null);
  if (
    !resolvedRole ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') ||
    (element.localName !== MOVE_FROM && element.localName !== MOVE_TO) ||
    (element.parentElement !== null &&
      element.namespaceURI !== element.parentElement.namespaceURI)
  ) {
    return null;
  }
  const id = wordAttribute(element, 'id')?.trim() ?? '';
  const author = wordAttribute(element, 'author')?.trim() ?? '';
  const rawDate = wordAttribute(element, 'date');
  const text = moveRevisionText(element, resolvedRole);
  if (
    !/^\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (rawDate !== null && rawDate.length > MAX_MOVE_DATE_LENGTH) ||
    (rawDate !== null && !Number.isFinite(Date.parse(rawDate))) ||
    !text ||
    text.length > MAX_MOVE_TEXT_LENGTH ||
    hasUnsupportedWordAttribute(element) ||
    !moveChildrenAreTextOnly(element, resolvedRole)
  ) {
    return null;
  }
  return {
    element,
    role: resolvedRole,
    id,
    author,
    date: normalizeDate(rawDate),
    text,
  };
}

function moveRevisionText(
  revision: Element,
  role: WorkDocumentMoveRole,
): string {
  const names = role === 'from' ? ['delText'] : ['t'];
  return names
    .flatMap((name) => descendants(revision, name))
    .map((element) => element.textContent ?? '')
    .join('');
}

function moveChildrenAreTextOnly(
  revision: Element,
  role: WorkDocumentMoveRole,
): boolean {
  const allowedText = role === 'from' ? new Set(['delText']) : new Set(['t']);
  const children = Array.from(revision.children);
  if (!children.length) return false;
  let hasText = false;
  for (const child of children) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return false;
    }
    if (child.localName === 'r') {
      if (!moveRunIsTextOnly(child, allowedText)) return false;
      hasText ||= moveRunHasVisibleContent(child, allowedText);
      continue;
    }
    if (child.localName === 'hyperlink') {
      if (!isRelationshipFreeInternalHyperlink(child)) return false;
      const runs = Array.from(child.children);
      if (!runs.length) return false;
      for (const run of runs) {
        if (
          run.localName !== 'r' ||
          !DOCX_WORDPROCESSING_NAMESPACES.has(run.namespaceURI ?? '') ||
          !moveRunIsTextOnly(run, allowedText)
        ) {
          return false;
        }
        hasText ||= moveRunHasVisibleContent(run, allowedText);
      }
      continue;
    }
    if (
      child.localName === 'bookmarkStart' ||
      child.localName === 'bookmarkEnd'
    ) {
      if (!isRelationshipFreeBookmarkMarker(child)) return false;
      continue;
    }
    return false;
  }
  return hasText;
}

function moveRunIsTextOnly(
  run: Element,
  allowedText: ReadonlySet<string>,
): boolean {
  const properties = Array.from(run.children).filter(
    (child) => child.localName === 'rPr',
  );
  if (properties.length > 1) return false;
  for (const child of Array.from(run.children)) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')) {
      return false;
    }
    if (child.localName === 'rPr') {
      if (
        Array.from(child.querySelectorAll('*')).some(
          (descendant) =>
            !DOCX_WORDPROCESSING_NAMESPACES.has(descendant.namespaceURI ?? ''),
        )
      ) {
        return false;
      }
      continue;
    }
    if (child.localName === 'br') {
      if (!isTextWrappingBreak(child)) return false;
      continue;
    }
    if (isAdmittedRunGlyph(child)) continue;
    if (!allowedText.has(child.localName) || child.querySelector('*')) {
      return false;
    }
  }
  return true;
}

function moveRunHasVisibleContent(
  run: Element,
  allowedText: ReadonlySet<string>,
): boolean {
  return Array.from(run.children).some((node) => {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(node.namespaceURI ?? '')) {
      return false;
    }
    if (isAdmittedRunGlyph(node)) return true;
    return allowedText.has(node.localName) && Boolean(node.textContent);
  });
}

const HYPERLINK_ATTRIBUTES = new Set([
  'anchor',
  'docLocation',
  'history',
  'tgtFrame',
  'tooltip',
]);
const BOOKMARK_START_ATTRIBUTES = new Set(['id', 'name']);
const BOOKMARK_END_ATTRIBUTES = new Set(['id']);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);

function isRelationshipFreeInternalHyperlink(element: Element): boolean {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
    return false;
  }
  let hasAnchor = false;
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return false;
    if (namespace && namespace !== element.namespaceURI) return false;
    const localName = xmlAttributeLocalName(attribute);
    if (namespace === element.namespaceURI || !namespace) {
      if (!HYPERLINK_ATTRIBUTES.has(localName)) return false;
      if (localName === 'anchor') {
        const value = attribute.value.trim();
        if (
          !value ||
          value.length > 255 ||
          /[\u0000-\u001f\u007f]/.test(value)
        ) {
          return false;
        }
        hasAnchor = true;
      }
    }
  }
  return hasAnchor;
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

function isTextWrappingBreak(element: Element): boolean {
  if (element.children.length) return false;
  const type = wordAttribute(element, 'type');
  return type === null || type === 'textWrapping';
}

/** Empty CT_Empty run glyphs already projected elsewhere in Writer import. */
const ADMITTED_EMPTY_RUN_GLYPHS = new Set([
  'tab',
  'noBreakHyphen',
  'softHyphen',
  'cr',
  'lastRenderedPageBreak',
  'pgNum',
  'dayLong',
  'monthLong',
  'yearLong',
  'dayShort',
  'monthShort',
  'yearShort',
  'footnoteRef',
  'endnoteRef',
  'annotationRef',
  'separator',
  'continuationSeparator',
]);

function isAdmittedEmptyRunGlyph(element: Element): boolean {
  return (
    ADMITTED_EMPTY_RUN_GLYPHS.has(element.localName) &&
    element.children.length === 0 &&
    element.attributes.length === 0
  );
}

/**
 * Bounded body footnote markers: empty element with exactly one Word-ns w:id.
 * Attributed footnoteRef, endnoteReference, and malformed ids stay fail-closed.
 */
function isAdmittedFootnoteReferenceGlyph(element: Element): boolean {
  if (
    element.localName !== 'footnoteReference' ||
    element.children.length !== 0 ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')
  ) {
    return false;
  }
  const namespace = element.namespaceURI;
  if (!namespace) return false;
  const wordAttributes = Array.from(element.attributes).filter(
    (candidate) => xmlAttributeNamespace(element, candidate) === namespace,
  );
  if (wordAttributes.length !== 1) return false;
  const only = wordAttributes[0];
  if (!only || xmlAttributeLocalName(only) !== 'id') return false;
  const id = only.value.trim();
  return /^\+?\d{1,10}$/.test(id);
}

function isAdmittedRunGlyph(element: Element): boolean {
  return (
    isAdmittedEmptyRunGlyph(element) ||
    isAdmittedFootnoteReferenceGlyph(element)
  );
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

function hasUnsupportedWordAttribute(element: Element): boolean {
  const namespace = element.namespaceURI;
  const supported = new Set(['id', 'author', 'date']);
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !supported.has(xmlAttributeLocalName(candidate)),
  );
}

function convertDeletedText(document: Document, revision: Element): void {
  for (const deleted of descendants(revision, 'delText')) {
    const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
    const preserve =
      deleted.getAttributeNS(XML_NAMESPACE, 'space') ??
      deleted.getAttribute('xml:space');
    if (preserve) text.setAttributeNS(XML_NAMESPACE, 'xml:space', preserve);
    text.textContent = deleted.textContent;
    deleted.replaceWith(text);
  }
}

function unwrapRevision(
  document: Document,
  revision: Element,
  start: string,
  end: string,
): void {
  const parent = revision.parentNode;
  if (!parent) return;
  parent.insertBefore(markerRun(document, start), revision);
  while (revision.firstChild)
    parent.insertBefore(revision.firstChild, revision);
  parent.insertBefore(markerRun(document, end), revision);
  revision.remove();
}

function markerRun(document: Document, value: string): Element {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = value;
  run.append(text);
  return run;
}

function wrapMarkerRange(
  root: HTMLElement,
  start: string,
  end: string,
  wrapper: HTMLElement,
): boolean {
  const html = root.innerHTML;
  const startIndex = html.indexOf(start);
  const endIndex = html.indexOf(end, startIndex + start.length);
  if (startIndex < 0 || endIndex < 0) return false;
  const closingTag = `</${wrapper.localName}>`;
  const serializedWrapper = wrapper.outerHTML;
  if (!serializedWrapper.endsWith(closingTag)) return false;
  const openingTag = serializedWrapper.slice(0, -closingTag.length);
  root.innerHTML = `${html.slice(0, startIndex)}${openingTag}${html.slice(
    startIndex + start.length,
    endIndex,
  )}${closingTag}${html.slice(endIndex + end.length)}`;
  return true;
}

function closestRevision(element: Element | null): Element | null {
  let current = element;
  while (current) {
    if (
      current.localName === 'ins' ||
      current.localName === 'del' ||
      current.localName === MOVE_FROM ||
      current.localName === MOVE_TO
    )
      return current;
    current = current.parentElement;
  }
  return null;
}

function uniqueChangeId(
  base: string,
  changes: ImportedDocxChangeMarker[],
): string {
  const ids = new Set(changes.map((change) => change.id));
  if (!ids.has(base)) return base;
  let suffix = 2;
  while (ids.has(`${base}-${suffix}`)) suffix += 1;
  return `${base}-${suffix}`;
}

function normalizeDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

function compareDocumentOrder(left: Element, right: Element): number {
  if (left === right) return 0;
  return left.compareDocumentPosition(right) & Node.DOCUMENT_POSITION_FOLLOWING
    ? -1
    : 1;
}
