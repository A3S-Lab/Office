import { attribute, descendants } from './work-ooxml-package';
import type {
  WorkDocumentChangeIdentity,
  WorkDocumentChangeKind,
} from './work-document-changes';

export interface ImportedDocxChangeMarker extends WorkDocumentChangeIdentity {
  kind: WorkDocumentChangeKind;
  start: string;
  end: string;
}

export interface ImportedDocxChangeMarkers {
  changes: ImportedDocxChangeMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export function markDocxTextChanges(
  document: Document,
): ImportedDocxChangeMarkers {
  const changes: ImportedDocxChangeMarker[] = [];
  const revisions = [
    ...descendants(document, 'ins'),
    ...descendants(document, 'del'),
  ].sort(compareDocumentOrder);
  for (const revision of revisions) {
    if (!revision.parentNode || closestRevision(revision.parentElement))
      continue;
    const kind: WorkDocumentChangeKind =
      revision.localName === 'del' ? 'deletion' : 'insertion';
    if (!revisionText(revision, kind)) continue;
    if (kind === 'deletion') convertDeletedText(document, revision);
    const index = changes.length + 1;
    const sourceId = attribute(revision, 'id')?.trim() ?? '';
    const id = uniqueChangeId(
      sourceId ? `docx-change-${sourceId}` : `docx-change-${index}`,
      changes,
    );
    const marker: ImportedDocxChangeMarker = {
      id,
      kind,
      author: attribute(revision, 'author')?.trim() || '未知审阅者',
      date: normalizeDate(attribute(revision, 'date')),
      start: `__A3S_WORK_CHANGE_START_${index}__`,
      end: `__A3S_WORK_CHANGE_END_${index}__`,
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
    const element = document.createElement(
      marker.kind === 'deletion' ? 'del' : 'ins',
    );
    element.dataset.documentChange = 'true';
    element.dataset.changeKind = marker.kind;
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
    if (current.localName === 'ins' || current.localName === 'del')
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
