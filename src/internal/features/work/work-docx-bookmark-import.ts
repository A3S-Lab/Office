import {
  normalizeDocumentBookmarkName,
  normalizeDocumentBookmarkNativeId,
} from './work-document-bookmarks';
import {
  docxCaptionBookmark,
  docxCaptionSequenceKind,
} from './work-docx-caption-fields';
import { docxFieldOccurrences } from './work-docx-field-instructions';
import { attribute, descendants } from './work-ooxml-package';

export interface ImportedDocxBookmarkMarkers {
  bookmarks: ImportedDocxBookmarkMarker[];
}

interface ImportedDocxBookmarkMarker {
  start: string;
  end: string;
  id: string;
  name: string;
  sourceName: string;
  nativeId: number;
}

interface DocxBookmarkPair {
  start: Element;
  end: Element;
  sourceName: string;
  sourceNativeId: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MAX_BOOKMARK_NAME_LENGTH = 40;
const MAX_BOOKMARK_NATIVE_ID = 0x7fff_ffff;

export function markDocxBookmarks(
  document: Document,
): ImportedDocxBookmarkMarkers {
  const captionBookmarks = captionBookmarkStarts(document);
  const pairs = docxBookmarkPairs(document).filter(
    (pair) =>
      !captionBookmarks.has(pair.start) && pair.sourceName !== '_GoBack',
  );
  const names = new Set<string>();
  const normalizedNames = new Map<DocxBookmarkPair, string>();
  for (let index = pairs.length - 1; index >= 0; index -= 1) {
    const pair = pairs[index];
    if (!pair) continue;
    const base =
      normalizeDocumentBookmarkName(pair.sourceName) ??
      safeImportedBookmarkName(pair.sourceName, index + 1);
    const name = uniqueImportedBookmarkName(base, names);
    names.add(name.toLowerCase());
    normalizedNames.set(pair, name);
  }

  const nativeIds = new Set<number>();
  const internalIds = new Set<string>();
  const bookmarks: ImportedDocxBookmarkMarker[] = [];
  for (const [index, pair] of pairs.entries()) {
    const preferredNativeId = normalizeDocumentBookmarkNativeId(
      pair.sourceNativeId,
    );
    const nativeId =
      preferredNativeId !== null && !nativeIds.has(preferredNativeId)
        ? preferredNativeId
        : nextNativeId(nativeIds);
    nativeIds.add(nativeId);
    const name = normalizedNames.get(pair) ?? `Bookmark_${index + 1}`;
    const id = uniqueImportedBookmarkId(name, nativeId, index + 1, internalIds);
    internalIds.add(id);
    const marker = importedBookmarkMarker(
      index + 1,
      id,
      name,
      pair.sourceName,
      nativeId,
      document.documentElement.textContent ?? '',
    );
    insertBoundaryMarker(document, pair.start, marker.start);
    insertBoundaryMarker(document, pair.end, marker.end);
    pair.start.remove();
    pair.end.remove();
    bookmarks.push(marker);
  }
  return { bookmarks };
}

export function applyImportedDocxBookmarkMarkers(
  document: Document,
  markers: ImportedDocxBookmarkMarkers,
): void {
  const targetNames = new Map<string, string>();
  for (const bookmark of markers.bookmarks) {
    targetNames.set(bookmark.sourceName.toLowerCase(), bookmark.name);
    replaceBookmarkMarker(document.body, bookmark.start, {
      ...bookmark,
      kind: 'start',
    });
    replaceBookmarkMarker(document.body, bookmark.end, {
      ...bookmark,
      kind: 'end',
    });
  }
  for (const link of document.body.querySelectorAll<HTMLAnchorElement>(
    'a[href^="#"]',
  )) {
    const sourceName = (link.getAttribute('href') ?? '').slice(1);
    const name = targetNames.get(sourceName.toLowerCase());
    if (name) link.setAttribute('href', `#${name}`);
  }
}

export function hasImportedDocxBookmarkMarkers(
  markers: ImportedDocxBookmarkMarkers,
): boolean {
  return markers.bookmarks.length > 0;
}

function docxBookmarkPairs(document: Document): DocxBookmarkPair[] {
  const open = new Map<string, Element[]>();
  const pairs: DocxBookmarkPair[] = [];
  for (const element of Array.from(document.getElementsByTagName('*'))) {
    if (element.localName === 'bookmarkStart') {
      const id = attribute(element, 'id')?.trim() ?? '';
      const stack = open.get(id) ?? [];
      stack.push(element);
      open.set(id, stack);
      continue;
    }
    if (element.localName !== 'bookmarkEnd') continue;
    const id = attribute(element, 'id')?.trim() ?? '';
    const start = open.get(id)?.pop();
    if (!start) continue;
    pairs.push({
      start,
      end: element,
      sourceName: attribute(start, 'name')?.trim() ?? '',
      sourceNativeId: id,
    });
  }
  pairs.sort((left, right) => {
    const position = left.start.compareDocumentPosition(right.start);
    if (position & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
    if (position & Node.DOCUMENT_POSITION_PRECEDING) return 1;
    return 0;
  });
  return pairs;
}

function captionBookmarkStarts(document: Document): Set<Element> {
  const bookmarks = new Set<Element>();
  for (const paragraph of descendants(document, 'p')) {
    for (const field of docxFieldOccurrences(paragraph)) {
      if (!docxCaptionSequenceKind(field.instruction)) continue;
      const bookmark = docxCaptionBookmark(paragraph, field);
      if (bookmark) bookmarks.add(bookmark);
    }
  }
  return bookmarks;
}

function importedBookmarkMarker(
  index: number,
  id: string,
  name: string,
  sourceName: string,
  nativeId: number,
  occupiedText: string,
): ImportedDocxBookmarkMarker {
  const base = `${index}_${stableHash(`${id}:${name}`)}`;
  let suffix = base;
  let collision = 2;
  while (
    occupiedText.includes(`__A3S_WORK_BOOKMARK_START_${suffix}__`) ||
    occupiedText.includes(`__A3S_WORK_BOOKMARK_END_${suffix}__`)
  ) {
    suffix = `${base}_${collision}`;
    collision += 1;
  }
  return {
    start: `__A3S_WORK_BOOKMARK_START_${suffix}__`,
    end: `__A3S_WORK_BOOKMARK_END_${suffix}__`,
    id,
    name,
    sourceName,
    nativeId,
  };
}

function insertBoundaryMarker(
  document: Document,
  boundary: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  boundary.parentNode?.insertBefore(run, boundary);
}

function replaceBookmarkMarker(
  root: HTMLElement,
  marker: string,
  bookmark: ImportedDocxBookmarkMarker & { kind: 'start' | 'end' },
): boolean {
  const text = textNodes(root).find((node) => node.data.includes(marker));
  if (!text) return false;
  const offset = text.data.indexOf(marker);
  const element = root.ownerDocument.createElement('span');
  element.dataset.documentBookmarkBoundary = 'true';
  element.dataset.bookmarkKind = bookmark.kind;
  element.dataset.bookmarkId = bookmark.id;
  element.dataset.bookmarkName = bookmark.name;
  element.dataset.officeBookmarkId = String(bookmark.nativeId);
  element.className = `work-document-bookmark-boundary ${bookmark.kind}`;
  element.contentEditable = 'false';
  element.setAttribute('aria-hidden', 'true');
  if (bookmark.kind === 'start') element.id = bookmark.name;
  const before = text.data.slice(0, offset);
  const after = text.data.slice(offset + marker.length);
  const fragment = root.ownerDocument.createDocumentFragment();
  if (before) fragment.append(root.ownerDocument.createTextNode(before));
  fragment.append(element);
  if (after) fragment.append(root.ownerDocument.createTextNode(after));
  text.replaceWith(fragment);
  return true;
}

function safeImportedBookmarkName(source: string, index: number): string {
  const characters = Array.from(source.trim())
    .map((character) => (/^[\p{L}\p{N}_]$/u.test(character) ? character : '_'))
    .join('');
  const leading = /^[\p{L}_]/u.test(characters)
    ? characters
    : `Bookmark_${characters}`;
  return (
    Array.from(leading || `Bookmark_${index}`)
      .slice(0, MAX_BOOKMARK_NAME_LENGTH)
      .join('') || `Bookmark_${index}`
  );
}

function uniqueImportedBookmarkName(
  base: string,
  names: ReadonlySet<string>,
): string {
  if (!names.has(base.toLowerCase())) return base;
  let suffix = 2;
  while (suffix <= MAX_BOOKMARK_NATIVE_ID) {
    const ending = `_${suffix}`;
    const prefix = Array.from(base)
      .slice(0, MAX_BOOKMARK_NAME_LENGTH - ending.length)
      .join('');
    const candidate = `${prefix}${ending}`;
    if (!names.has(candidate.toLowerCase())) return candidate;
    suffix += 1;
  }
  throw new Error('No unique imported Word bookmark name is available.');
}

function uniqueImportedBookmarkId(
  name: string,
  nativeId: number,
  index: number,
  ids: ReadonlySet<string>,
): string {
  const base = `docx-bookmark-${nativeId}-${stableHash(name)}`;
  if (!ids.has(base)) return base;
  let suffix = index;
  while (ids.has(`${base}-${suffix}`)) suffix += 1;
  return `${base}-${suffix}`;
}

function nextNativeId(ids: ReadonlySet<number>): number {
  for (let id = 0; id <= MAX_BOOKMARK_NATIVE_ID; id += 1) {
    if (!ids.has(id)) return id;
  }
  throw new Error('No unique imported Word bookmark identifier is available.');
}

function stableHash(source: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < source.length; index += 1) {
    hash ^= source.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(36);
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
