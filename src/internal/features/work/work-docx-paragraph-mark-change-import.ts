import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, directChildren } from './work-ooxml-package';

export type DocxParagraphMarkChangeKind = 'insertion' | 'deletion';

export interface ImportedDocxParagraphMarkChangeMarker {
  marker: string;
  id: string;
  kind: DocxParagraphMarkChangeKind;
  author: string;
  date: string;
}

export interface ImportedDocxParagraphMarkChangeMarkers {
  paragraphs: ImportedDocxParagraphMarkChangeMarker[];
}

interface SupportedDocxParagraphMarkChange
  extends Omit<ImportedDocxParagraphMarkChangeMarker, 'marker'> {
  element: Element;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MAX_PARAGRAPH_MARK_CHANGES = 65_536;
const MAX_REVISION_DATE_LENGTH = 64;
const PARAGRAPH_MARK_CHANGE_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_MARK_CHANGE_\d+__/g;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);

/**
 * Finds the bounded Word paragraph-mark insertion/deletion shape. Word and
 * WPS store the revision on the paragraph mark's run-properties node while
 * keeping the paragraph's text revisions in ordinary w:ins/w:del wrappers.
 */
export function markDocxParagraphMarkChanges(
  document: Document,
): ImportedDocxParagraphMarkChangeMarkers {
  const paragraphs: ImportedDocxParagraphMarkChangeMarker[] = [];
  const changeIds = new Set<string>();
  for (const paragraph of descendants(document, 'p')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(paragraph.namespaceURI ?? '')) {
      continue;
    }
    const change = supportedParagraphMarkChange(paragraph);
    if (!change) continue;
    if (paragraphs.length >= MAX_PARAGRAPH_MARK_CHANGES) {
      throw new Error('Document exceeds the paragraph-mark revision limit.');
    }
    const marker = `__A3S_WORK_PARAGRAPH_MARK_CHANGE_${paragraphs.length + 1}__`;
    insertParagraphMarker(paragraph, marker);
    // Mammoth applies w:del on a paragraph mark by joining the following
    // paragraph.  The bounded Work model reviews the complete text-only
    // paragraph atomically, so remove the native boundary flag after capturing
    // it and let the marker retain the original block boundary through HTML.
    change.element.remove();
    paragraphs.push({
      marker,
      id: uniqueChangeId(change.id, changeIds),
      kind: change.kind,
      author: change.author,
      date: change.date,
    });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphMarkChangeMarkers(
  document: Document,
  markers: ImportedDocxParagraphMarkChangeMarkers,
): void {
  const changes = new Map(
    markers.paragraphs.map((change) => [change.marker, change]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_MARK_CHANGE_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_MARK_CHANGE_MARKER_PATTERN,
      (marker) => {
        const change = changes.get(marker);
        const block = change
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && change) {
          block.dataset.documentBlockChange = 'true';
          block.dataset.blockChangeKind = change.kind;
          block.dataset.blockChangeId = change.id;
          block.dataset.blockChangeAuthor = change.author;
          block.dataset.blockChangeDate = change.date;
        }
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphMarkChangeMarkers(
  markers: ImportedDocxParagraphMarkChangeMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

/** Returns true only for a native, unambiguous paragraph-mark revision. */
export function isSupportedDocxParagraphMarkChange(change: Element): boolean {
  const runProperties = change.parentElement;
  const properties = runProperties?.parentElement;
  const paragraph = properties?.parentElement;
  if (
    runProperties?.localName === 'rPr' &&
    properties?.localName === 'pPr' &&
    paragraph?.localName === 'p' &&
    runProperties.namespaceURI === change.namespaceURI &&
    properties.namespaceURI === change.namespaceURI &&
    paragraph.namespaceURI === change.namespaceURI
  ) {
    return supportedParagraphMarkChange(paragraph)?.element === change;
  }
  return false;
}

export type DocxParagraphBreakMarkKind = 'merge' | 'split';

export interface InspectedDocxParagraphBreakMark {
  author: string;
  date: string;
  id: string;
  kind: DocxParagraphBreakMarkKind;
}

export interface ImportedDocxParagraphBreakChangeMarker {
  marker: string;
  id: string;
  kind: DocxParagraphBreakMarkKind;
  author: string;
  date: string;
}

export interface ImportedDocxParagraphBreakChangeMarkers {
  paragraphs: ImportedDocxParagraphBreakChangeMarker[];
}

const PARAGRAPH_BREAK_CHANGE_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_BREAK_CHANGE_\d+__/g;

/**
 * Mark-only paragraph revisions (no matching body wrap) are paragraph-break
 * merge/split candidates. Eligible neighbors promote them into reviewable
 * Work changes; others stay diagnostics-only.
 */
export function isIsolatedDocxParagraphBreakMarkChange(
  change: Element,
): boolean {
  return isolatedParagraphBreakMarkChange(change) !== null;
}

export function inspectDocxParagraphBreakMarkChanges(
  document: Document,
): InspectedDocxParagraphBreakMark[] {
  const inspected: InspectedDocxParagraphBreakMark[] = [];
  for (const change of [
    ...descendants(document, 'ins'),
    ...descendants(document, 'del'),
  ]) {
    const breakMark = isolatedParagraphBreakMarkChange(change);
    if (breakMark) inspected.push(breakMark);
  }
  return inspected;
}

/**
 * Promotes eligible isolated paragraph-break marks into Work review markers.
 * Mammoth would join deleted marks; remove the native flag after capture so
 * both paragraphs survive for accept/reject.
 */
export function markDocxParagraphBreakChanges(
  document: Document,
): ImportedDocxParagraphBreakChangeMarkers {
  const paragraphs: ImportedDocxParagraphBreakChangeMarker[] = [];
  const changeIds = new Set<string>();
  for (const paragraph of descendants(document, 'p')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(paragraph.namespaceURI ?? '')) {
      continue;
    }
    const change = reviewableParagraphBreakMarkChange(paragraph);
    if (!change) continue;
    if (paragraphs.length >= MAX_PARAGRAPH_MARK_CHANGES) {
      throw new Error('Document exceeds the paragraph-break revision limit.');
    }
    const marker = `__A3S_WORK_PARAGRAPH_BREAK_CHANGE_${paragraphs.length + 1}__`;
    insertParagraphMarker(paragraph, marker);
    change.element.remove();
    paragraphs.push({
      marker,
      id: uniqueChangeId(change.id, changeIds),
      kind: change.kind,
      author: change.author,
      date: change.date,
    });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphBreakChangeMarkers(
  document: Document,
  markers: ImportedDocxParagraphBreakChangeMarkers,
): void {
  const changes = new Map(
    markers.paragraphs.map((change) => [change.marker, change]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_BREAK_CHANGE_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_BREAK_CHANGE_MARKER_PATTERN,
      (marker) => {
        const change = changes.get(marker);
        const block = change
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && change) {
          block.dataset.paragraphBreakChange = 'true';
          block.dataset.paragraphBreakKind = change.kind;
          block.dataset.paragraphBreakId = change.id;
          block.dataset.paragraphBreakAuthor = change.author;
          block.dataset.paragraphBreakDate = change.date;
        }
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphBreakChangeMarkers(
  markers: ImportedDocxParagraphBreakChangeMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function reviewableParagraphBreakMarkChange(
  paragraph: Element,
): (InspectedDocxParagraphBreakMark & { element: Element }) | null {
  const properties = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (properties.length !== 1) return null;
  const runProperties = directChildren(properties[0] as Element, 'rPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (runProperties.length !== 1) return null;
  const revision = Array.from((runProperties[0] as Element).children).find(
    (element) =>
      (element.localName === 'ins' || element.localName === 'del') &&
      element.namespaceURI === paragraph.namespaceURI,
  );
  if (!revision) return null;
  const inspected = isolatedParagraphBreakMarkChange(revision);
  if (!inspected) return null;
  if (!paragraphBreakNeighborIsEligible(paragraph, inspected.kind)) {
    return null;
  }
  return { ...inspected, element: revision };
}

function paragraphBreakNeighborIsEligible(
  paragraph: Element,
  kind: DocxParagraphBreakMarkKind,
): boolean {
  const sibling =
    kind === 'merge'
      ? nextWordParagraphSibling(paragraph)
      : previousWordParagraphSibling(paragraph);
  if (!sibling) return false;
  const properties = directChildren(sibling, 'pPr').find(
    (element) => element.namespaceURI === sibling.namespaceURI,
  );
  return paragraphBodyIsUntrackedTextOnly(sibling, properties);
}

function nextWordParagraphSibling(paragraph: Element): Element | null {
  let sibling = paragraph.nextElementSibling;
  while (sibling) {
    if (
      sibling.localName === 'p' &&
      sibling.namespaceURI === paragraph.namespaceURI
    ) {
      return sibling;
    }
    if (
      sibling.localName === 'tbl' ||
      sibling.localName === 'sectPr' ||
      sibling.localName === 'sdt'
    ) {
      return null;
    }
    sibling = sibling.nextElementSibling;
  }
  return null;
}

function previousWordParagraphSibling(paragraph: Element): Element | null {
  let sibling = paragraph.previousElementSibling;
  while (sibling) {
    if (
      sibling.localName === 'p' &&
      sibling.namespaceURI === paragraph.namespaceURI
    ) {
      return sibling;
    }
    if (
      sibling.localName === 'tbl' ||
      sibling.localName === 'sectPr' ||
      sibling.localName === 'sdt'
    ) {
      return null;
    }
    sibling = sibling.previousElementSibling;
  }
  return null;
}

function isolatedParagraphBreakMarkChange(
  change: Element,
): InspectedDocxParagraphBreakMark | null {
  const runProperties = change.parentElement;
  const properties = runProperties?.parentElement;
  const paragraph = properties?.parentElement;
  if (
    runProperties?.localName !== 'rPr' ||
    properties?.localName !== 'pPr' ||
    paragraph?.localName !== 'p' ||
    runProperties.namespaceURI !== change.namespaceURI ||
    properties.namespaceURI !== change.namespaceURI ||
    paragraph.namespaceURI !== change.namespaceURI ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  if (supportedParagraphMarkChange(paragraph)?.element === change) {
    return null;
  }
  const markChanges = Array.from(runProperties.children).filter(
    (element) =>
      (element.localName === 'ins' || element.localName === 'del') &&
      element.namespaceURI === paragraph.namespaceURI,
  );
  if (markChanges.length !== 1 || markChanges[0] !== change) return null;
  const parsed = paragraphMarkChangeFromElement(change);
  if (!parsed) return null;
  if (!paragraphBodyIsUntrackedTextOnly(paragraph, properties)) return null;
  return {
    author: parsed.author,
    date: parsed.date,
    id: parsed.id,
    kind: parsed.kind === 'deletion' ? 'merge' : 'split',
  };
}

function paragraphBodyIsUntrackedTextOnly(
  paragraph: Element,
  properties: Element | undefined,
): boolean {
  const body = directChildren(paragraph).filter(
    (element) => element !== properties,
  );
  if (!body.length) return true;
  for (const run of body) {
    if (run.localName !== 'r' || run.namespaceURI !== paragraph.namespaceURI) {
      return false;
    }
    const children = directChildren(run);
    const runProperties = children.filter(
      (child) =>
        child.localName === 'rPr' &&
        child.namespaceURI === paragraph.namespaceURI,
    );
    if (runProperties.length > 1) return false;
    for (const child of children) {
      if (child.namespaceURI !== paragraph.namespaceURI) return false;
      if (child.localName === 'rPr') {
        if (
          Array.from(child.querySelectorAll('*')).some(
            (descendant) =>
              descendant.localName === 'ins' ||
              descendant.localName === 'del' ||
              descendant.namespaceURI !== paragraph.namespaceURI,
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
      if (child.localName !== 't' || child.children.length) return false;
    }
  }
  return true;
}

function supportedParagraphMarkChange(
  paragraph: Element,
): SupportedDocxParagraphMarkChange | null {
  const properties = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (properties.length !== 1) return null;
  const runProperties = directChildren(properties[0] as Element, 'rPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (runProperties.length !== 1) return null;
  const changes = (runProperties[0] as Element).children;
  const revision = Array.from(changes).find(
    (element) =>
      (element.localName === 'ins' || element.localName === 'del') &&
      element.namespaceURI === paragraph.namespaceURI,
  );
  if (
    !revision ||
    Array.from(changes).filter(
      (element) =>
        (element.localName === 'ins' || element.localName === 'del') &&
        element.namespaceURI === paragraph.namespaceURI,
    ).length !== 1
  ) {
    return null;
  }
  const change = paragraphMarkChangeFromElement(revision);
  return change && paragraphBodyMatchesChange(paragraph, properties[0], change)
    ? { element: revision, ...change }
    : null;
}

function paragraphMarkChangeFromElement(
  change: Element,
): Omit<ImportedDocxParagraphMarkChangeMarker, 'marker'> | null {
  if (
    (change.localName !== 'ins' && change.localName !== 'del') ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '') ||
    change.children.length > 0 ||
    hasUnsupportedWordAttributes(change)
  ) {
    return null;
  }
  const id = wordAttribute(change, 'id')?.trim() ?? '';
  const author = wordAttribute(change, 'author')?.trim() ?? '';
  const date = wordAttribute(change, 'date');
  if (
    !/^\+?\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (date !== null && date.length > MAX_REVISION_DATE_LENGTH) ||
    (date !== null && !Number.isFinite(Date.parse(date)))
  ) {
    return null;
  }
  return {
    id: `docx-paragraph-mark-change-${id}`,
    kind: change.localName === 'del' ? 'deletion' : 'insertion',
    author,
    date: normalizeRevisionDate(date),
  };
}

function paragraphBodyMatchesChange(
  paragraph: Element,
  properties: Element | undefined,
  change: Omit<ImportedDocxParagraphMarkChangeMarker, 'marker'>,
): boolean {
  const body = directChildren(paragraph).filter(
    (element) => element !== properties,
  );
  if (!body.length) return false;
  const expectedName = change.kind === 'deletion' ? 'del' : 'ins';
  // Admit one or more consecutive matching body wrappers (Word/WPS often
  // split a whole-paragraph revision across run formatting siblings). Mixed
  // untracked text, drawings, or mismatched authors stay excluded.
  // Relationship-free bookmarkStart/End markers and empty/rPr-only untracked
  // runs may appear as siblings of the revision wrappers without blocking
  // admission.
  let sawRevision = false;
  for (const revision of body) {
    if (
      revision.localName === 'bookmarkStart' ||
      revision.localName === 'bookmarkEnd'
    ) {
      if (!isRelationshipFreeBookmarkMarker(revision)) return false;
      continue;
    }
    if (revision.localName === 'r') {
      if (
        revision.namespaceURI !== paragraph.namespaceURI ||
        !isPropertiesOnlyUntrackedRun(revision)
      ) {
        return false;
      }
      continue;
    }
    if (
      revision.localName !== expectedName ||
      revision.namespaceURI !== paragraph.namespaceURI ||
      hasUnsupportedWordAttributes(revision) ||
      !revisionBodyIsTextOnly(revision, change.kind)
    ) {
      return false;
    }
    const id = wordAttribute(revision, 'id')?.trim() ?? '';
    const author = wordAttribute(revision, 'author')?.trim() ?? '';
    const rawDate = wordAttribute(revision, 'date');
    if (
      !/^\+?\d{1,10}$/.test(id) ||
      author !== change.author ||
      normalizeRevisionDate(rawDate) !== change.date ||
      (rawDate !== null && !Number.isFinite(Date.parse(rawDate)))
    ) {
      return false;
    }
    sawRevision = true;
  }
  return sawRevision;
}

function revisionBodyIsTextOnly(
  revision: Element,
  kind: DocxParagraphMarkChangeKind,
): boolean {
  const children = directChildren(revision);
  if (!children.length) return false;
  let hasText = false;
  for (const child of children) {
    if (child.namespaceURI !== revision.namespaceURI) return false;
    if (child.localName === 'r') {
      if (!runIsTextOnly(child, kind)) return false;
      hasText ||= runHasVisibleText(child, kind);
      continue;
    }
    if (child.localName === 'hyperlink') {
      if (!isRelationshipFreeInternalHyperlink(child)) return false;
      const runs = directChildren(child);
      if (!runs.length) return false;
      for (const run of runs) {
        if (
          run.localName !== 'r' ||
          run.namespaceURI !== revision.namespaceURI ||
          !runIsTextOnly(run, kind)
        ) {
          return false;
        }
        hasText ||= runHasVisibleText(run, kind);
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

function runIsTextOnly(
  run: Element,
  kind: DocxParagraphMarkChangeKind,
): boolean {
  const textName = kind === 'deletion' ? 'delText' : 't';
  const children = directChildren(run);
  const properties = children.filter(
    (child) =>
      child.localName === 'rPr' && child.namespaceURI === run.namespaceURI,
  );
  if (properties.length > 1) return false;
  for (const child of children) {
    if (child.namespaceURI !== run.namespaceURI) return false;
    if (child.localName === 'rPr') {
      if (
        Array.from(child.querySelectorAll('*')).some(
          (descendant) => descendant.namespaceURI !== run.namespaceURI,
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
    if (child.localName !== textName || child.children.length) return false;
  }
  return true;
}

function runHasVisibleText(
  run: Element,
  kind: DocxParagraphMarkChangeKind,
): boolean {
  const textName = kind === 'deletion' ? 'delText' : 't';
  return directChildren(run).some(
    (child) =>
      child.localName === textName &&
      child.namespaceURI === run.namespaceURI &&
      Boolean(child.textContent),
  );
}

/**
 * Word often emits empty or properties-only runs beside a whole-paragraph
 * mark revision. Those carry no visible body text and must not block
 * admission; runs with any text or break content stay fail-closed.
 */
function isPropertiesOnlyUntrackedRun(run: Element): boolean {
  const children = directChildren(run);
  const properties = children.filter(
    (child) =>
      child.localName === 'rPr' && child.namespaceURI === run.namespaceURI,
  );
  if (properties.length > 1) return false;
  for (const child of children) {
    if (child.namespaceURI !== run.namespaceURI) return false;
    if (child.localName === 'rPr') {
      if (
        Array.from(child.querySelectorAll('*')).some(
          (descendant) => descendant.namespaceURI !== run.namespaceURI,
        )
      ) {
        return false;
      }
      continue;
    }
    return false;
  }
  return true;
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

function hasUnsupportedWordAttributes(element: Element): boolean {
  const namespace = element.namespaceURI;
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !REVISION_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
  );
}

function uniqueChangeId(base: string, ids: Set<string>): string {
  if (!ids.has(base)) {
    ids.add(base);
    return base;
  }
  let suffix = 2;
  while (ids.has(`${base}-${suffix}`)) suffix += 1;
  const id = `${base}-${suffix}`;
  ids.add(id);
  return id;
}

function insertParagraphMarker(paragraph: Element, marker: string): void {
  const document = paragraph.ownerDocument;
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChildren(paragraph, 'pPr').find(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
}

function closestParagraphBlock(
  element: Element | null,
  markerNode: Node,
): HTMLElement | null {
  const explicit = element?.closest('p, h1, h2, h3, h4, h5, h6, figcaption');
  if (explicit instanceof HTMLElement) return explicit;
  const container = element?.closest('li, blockquote, td, th, div');
  return container instanceof HTMLElement
    ? wrapParagraphSegment(container, markerNode)
    : null;
}

function wrapParagraphSegment(
  container: HTMLElement,
  markerNode: Node,
): HTMLElement {
  let anchor: Node = markerNode;
  while (anchor.parentNode && anchor.parentNode !== container) {
    anchor = anchor.parentNode;
  }
  const children = Array.from(container.childNodes);
  const anchorIndex = children.indexOf(anchor as ChildNode);
  if (anchorIndex < 0) return container;
  let start = anchorIndex;
  let end = anchorIndex;
  while (start > 0 && !isParagraphBoundary(children[start - 1])) start -= 1;
  while (end + 1 < children.length && !isParagraphBoundary(children[end + 1])) {
    end += 1;
  }
  const paragraph = container.ownerDocument.createElement('p');
  const grouped = children.slice(start, end + 1);
  container.insertBefore(paragraph, grouped[0] ?? null);
  paragraph.append(...grouped);
  return paragraph;
}

function isParagraphBoundary(node: Node | undefined): boolean {
  if (!(node instanceof HTMLElement)) return false;
  return [
    'blockquote',
    'div',
    'figcaption',
    'h1',
    'h2',
    'h3',
    'h4',
    'h5',
    'h6',
    'hr',
    'ol',
    'p',
    'pre',
    'table',
    'ul',
  ].includes(node.tagName.toLowerCase());
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

function textNodes(root: ParentNode): Text[] {
  const walker = root.ownerDocument?.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
  );
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
