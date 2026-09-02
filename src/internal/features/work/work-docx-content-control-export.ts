import type { ParagraphChild } from 'docx';
import JSZip from 'jszip';
import {
  documentContentControlPropertiesFromElement,
  normalizeDocumentContentControlProperties,
  type WorkDocumentContentControlProperties,
} from './work-document-content-control';
import { ensureIgnorableContentControlNamespace } from './work-docx-note-comment-content-control-xml';
import { XML_NAMESPACE, xmlAttributeNamespace } from './work-docx-settings-xml';
import {
  descendants,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_NAMESPACES = new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE]);
const STORY_PATTERN =
  /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/;
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

export interface DocxContentControlPatch {
  startMarker: string;
  endMarker: string;
  properties: WorkDocumentContentControlProperties;
}

export interface DocxContentControlMarkerRegistration {
  startMarker: string;
  endMarker: string;
}

/** Collects native content-control patches while the docx package is built. */
export class DocxContentControlPatchCollector {
  readonly patches: DocxContentControlPatch[] = [];
  private nextMarker = 1;

  register(
    properties: WorkDocumentContentControlProperties,
  ): DocxContentControlMarkerRegistration {
    const index = this.nextMarker;
    this.nextMarker += 1;
    const patch: DocxContentControlPatch = {
      startMarker: `__A3S_WORK_CONTENT_CONTROL_START_${index}__`,
      endMarker: `__A3S_WORK_CONTENT_CONTROL_END_${index}__`,
      properties: normalizeDocumentContentControlProperties(properties),
    };
    this.patches.push(patch);
    return {
      startMarker: patch.startMarker,
      endMarker: patch.endMarker,
    };
  }
}

export function docxContentControlRuns(
  element: HTMLElement,
  docx: typeof import('docx'),
  collector: DocxContentControlPatchCollector,
  children: ParagraphChild[],
): ParagraphChild[] {
  const properties = documentContentControlPropertiesFromElement(element);
  if (containsUnsupportedContentControlSemantics(element)) {
    throw new Error(
      'Document content controls support only inline text and rich-text formatting.',
    );
  }
  const marker = collector.register(properties);
  return [
    new docx.TextRun(marker.startMarker),
    ...children,
    new docx.TextRun(marker.endMarker),
  ];
}

/**
 * Replaces generated marker runs with native inline `w:sdt` controls in every
 * Word story. The patcher is intentionally conservative: a marker pair must
 * be in one paragraph and may contain only generated runs.
 */
export async function patchDocxContentControls(
  buffer: ArrayBuffer,
  patches: readonly DocxContentControlPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  const archive = await JSZip.loadAsync(buffer);
  const entries = Object.values(archive.files).filter(
    (entry) => !entry.dir && STORY_PATTERN.test(entry.name),
  );
  const documents = await Promise.all(
    entries.map(async (entry) => ({
      entry,
      document: parseXml(
        await entry.async('text'),
        `generated DOCX ${entry.name}`,
      ),
    })),
  );
  const usedIds = new Set<number>();
  for (const { document } of documents) {
    for (const id of descendants(document, 'id')) {
      const value = wordAttribute(id, 'val');
      const parsed = value === null ? null : Number(value);
      if (
        parsed !== null &&
        Number.isSafeInteger(parsed) &&
        parsed >= -2_147_483_648 &&
        parsed <= 2_147_483_647
      ) {
        usedIds.add(parsed);
      }
    }
  }
  let nextId = 1;
  const changed = new Set<Document>();
  for (const patch of patches) {
    const match = findMarkerPair(documents, patch);
    if (!match) {
      throw new Error(
        `Generated DOCX did not emit both content-control markers ${patch.startMarker} and ${patch.endMarker}.`,
      );
    }
    const nativeId = allocateNativeId(
      usedIds,
      patch.properties.nativeId,
      nextId,
    );
    usedIds.add(nativeId);
    nextId = nativeId + 1;
    wrapMarkerPair(match, patch.properties, nativeId);
    changed.add(match.document);
  }
  for (const { entry, document } of documents) {
    if (changed.has(document)) {
      archive.file(entry.name, new XMLSerializer().serializeToString(document));
    }
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

interface MarkerPair {
  document: Document;
  paragraph: Element;
  start: Element;
  end: Element;
}

function findMarkerPair(
  documents: readonly { document: Document }[],
  patch: DocxContentControlPatch,
): MarkerPair | null {
  for (const { document } of documents) {
    const start = markerRun(document, patch.startMarker);
    const end = markerRun(document, patch.endMarker);
    if (!start || !end) continue;
    const startParagraph = closestAncestor(start, 'p');
    const endParagraph = closestAncestor(end, 'p');
    if (!startParagraph || startParagraph !== endParagraph) return null;
    return { document, paragraph: startParagraph, start, end };
  }
  return null;
}

function wrapMarkerPair(
  match: MarkerPair,
  properties: WorkDocumentContentControlProperties,
  nativeId: number,
): void {
  const { document, paragraph, start, end } = match;
  if (start.parentNode !== paragraph || end.parentNode !== paragraph) {
    throw new Error(
      'Generated DOCX content-control markers must be direct paragraph runs.',
    );
  }
  const contentRuns: Element[] = [];
  let current = start.nextSibling;
  while (current && current !== end) {
    const next = current.nextSibling;
    if (!isWordElement(current) || current.localName !== 'r') {
      throw new Error(
        'Generated DOCX content controls cannot contain non-run semantic nodes.',
      );
    }
    contentRuns.push(current);
    current = next;
  }
  if (current !== end) {
    throw new Error('Generated DOCX content-control marker order is invalid.');
  }
  const control = wordElement(document, paragraph, 'sdt');
  const propertiesElement = wordElement(document, paragraph, 'sdtPr');
  appendContentControlProperties(
    document,
    paragraph,
    propertiesElement,
    properties,
    nativeId,
  );
  const content = wordElement(document, paragraph, 'sdtContent');
  for (const run of contentRuns) content.append(run);
  if (!contentRuns.length) content.append(emptyRun(document, paragraph));
  control.append(propertiesElement, content);
  paragraph.insertBefore(control, start);
  start.remove();
  end.remove();
  const root = document.documentElement;
  if (properties.appearance !== 'boundingBox' || properties.color) {
    ensureIgnorableContentControlNamespace(root, WORD_2012_NAMESPACE, 'w15');
  }
}

function appendContentControlProperties(
  document: Document,
  context: Element,
  propertiesElement: Element,
  properties: WorkDocumentContentControlProperties,
  nativeId: number,
): void {
  const id = wordElement(document, context, 'id');
  setWordAttribute(id, context, 'val', String(nativeId));
  propertiesElement.append(id);
  if (properties.alias) {
    const alias = wordElement(document, context, 'alias');
    setWordAttribute(alias, context, 'val', properties.alias);
    propertiesElement.append(alias);
  }
  if (properties.tag) {
    const tag = wordElement(document, context, 'tag');
    setWordAttribute(tag, context, 'val', properties.tag);
    propertiesElement.append(tag);
  }
  if (properties.lock !== 'unlocked') {
    const lock = wordElement(document, context, 'lock');
    setWordAttribute(lock, context, 'val', properties.lock);
    propertiesElement.append(lock);
  }
  if (properties.type === 'text') {
    const text = wordElement(document, context, 'text');
    if (properties.multiLine) setWordAttribute(text, context, 'multiLine', '1');
    propertiesElement.append(text);
  } else {
    propertiesElement.append(wordElement(document, context, 'richText'));
  }
  if (properties.appearance !== 'boundingBox') {
    const appearance = namespacedElement(
      document,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'appearance',
    );
    setNamespacedAttribute(
      appearance,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'val',
      properties.appearance,
    );
    propertiesElement.append(appearance);
  }
  if (properties.color) {
    const color = namespacedElement(
      document,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'color',
    );
    setNamespacedAttribute(
      color,
      context,
      WORD_2012_NAMESPACE,
      'w15',
      'val',
      properties.color.slice(1).toUpperCase(),
    );
    propertiesElement.append(color);
  }
}

function emptyRun(document: Document, context: Element): Element {
  const run = wordElement(document, context, 'r');
  const text = wordElement(document, context, 't');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = '';
  run.append(text);
  return run;
}

function markerRun(document: Document, marker: string): Element | null {
  const texts = descendants(document, 't');
  let occurrences = 0;
  for (const text of texts) {
    occurrences += markerOccurrences(text.textContent ?? '', marker);
  }
  if (occurrences > 1) {
    throw new Error(
      `Generated DOCX content-control marker ${marker} occurs more than once.`,
    );
  }
  for (const text of texts) {
    if (!(text.textContent ?? '').includes(marker)) continue;
    const run = closestAncestor(text, 'r');
    if (!run) continue;
    isolateMarkerRun(text, run, marker);
    return findRunContainingMarker(document, marker);
  }
  return null;
}

function markerOccurrences(value: string, marker: string): number {
  if (!marker) return 0;
  let count = 0;
  let offset = 0;
  while (true) {
    const index = value.indexOf(marker, offset);
    if (index < 0) return count;
    count += 1;
    offset = index + marker.length;
  }
}

function isolateMarkerRun(text: Element, run: Element, marker: string): void {
  const value = text.textContent ?? '';
  const index = value.indexOf(marker);
  if (index < 0 || !run.parentNode) return;
  if (
    Array.from(run.children).some(
      (child) => child !== text && child.localName !== 'rPr',
    )
  ) {
    throw new Error('Generated DOCX content-control marker run is ambiguous.');
  }
  const before = value.slice(0, index);
  const after = value.slice(index + marker.length);
  const parent = run.parentNode;
  if (before) parent.insertBefore(cloneRunWithText(run, before), run);
  parent.insertBefore(cloneRunWithText(run, marker), run);
  if (after) parent.insertBefore(cloneRunWithText(run, after), run);
  run.remove();
}

function cloneRunWithText(run: Element, value: string): Element {
  const clone = run.cloneNode(true) as Element;
  const text = Array.from(clone.children).find(
    (child) => child.localName === 't',
  );
  if (!text) throw new Error('Generated DOCX marker run has no text node.');
  for (const child of Array.from(clone.children)) {
    if (child !== text && child.localName !== 'rPr') child.remove();
  }
  text.textContent = value;
  if (/^\s|\s$/u.test(value))
    text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  return clone;
}

function findRunContainingMarker(
  document: Document,
  marker: string,
): Element | null {
  return (
    descendants(document, 'r').find((run) =>
      Array.from(run.children).some(
        (child) =>
          child.localName === 't' && (child.textContent ?? '').includes(marker),
      ),
    ) ?? null
  );
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName && isWordElement(current))
      return current;
    current = current.parentElement;
  }
  return null;
}

function allocateNativeId(
  used: Set<number>,
  preferred: number | null,
  start: number,
): number {
  if (
    preferred !== null &&
    Number.isSafeInteger(preferred) &&
    preferred >= -2_147_483_648 &&
    preferred <= 2_147_483_647 &&
    !used.has(preferred)
  ) {
    return preferred;
  }
  let candidate = Math.max(1, start);
  while (candidate <= 2_147_483_647 && used.has(candidate)) candidate += 1;
  if (candidate > 2_147_483_647) {
    throw new Error('DOCX content-control IDs are exhausted.');
  }
  return candidate;
}

function containsUnsupportedContentControlSemantics(
  element: HTMLElement,
): boolean {
  return Boolean(
    element.querySelector(
      'a, img, [data-document-field], [data-document-equation], [data-document-note-reference], [data-document-content-control]',
    ),
  );
}

function isWordElement(element: Node): element is Element {
  return (
    element instanceof Element &&
    WORD_NAMESPACES.has(element.namespaceURI ?? '')
  );
}

function wordElement(
  document: Document,
  context: Element,
  localName: string,
): Element {
  const namespace = context.namespaceURI ?? WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  return document.createElementNS(namespace, `${prefix}:${localName}`);
}

function setWordAttribute(
  element: Element,
  context: Element,
  localName: string,
  value: string,
): void {
  const namespace = context.namespaceURI ?? WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

function namespacedElement(
  document: Document,
  context: Element,
  namespace: string,
  preferredPrefix: string,
  localName: string,
): Element {
  const root = context.ownerDocument.documentElement;
  const prefix =
    root.lookupPrefix?.(namespace) ??
    Array.from(root.attributes)
      .find(
        (item) => item.name.startsWith('xmlns:') && item.value === namespace,
      )
      ?.name.slice('xmlns:'.length) ??
    preferredPrefix;
  return document.createElementNS(namespace, `${prefix}:${localName}`);
}

function setNamespacedAttribute(
  element: Element,
  context: Element,
  namespace: string,
  preferredPrefix: string,
  localName: string,
  value: string,
): void {
  const prefix =
    context.ownerDocument.documentElement.lookupPrefix?.(namespace) ??
    preferredPrefix;
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

function wordAttribute(element: Element, localName: string): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      item.localName === localName &&
      WORD_NAMESPACES.has(xmlAttributeNamespace(element, item) ?? ''),
  );
  return matches.length === 1 ? matches[0].value : null;
}
