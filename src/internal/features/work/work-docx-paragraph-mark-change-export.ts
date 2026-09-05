import JSZip from 'jszip';
import {
  descendants,
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxParagraphMarkChangePatch {
  marker: string;
  id: number;
  kind: 'insertion' | 'deletion';
  author: string;
  date: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_PARAGRAPH_MARK_CHANGE_PATCHES = 65_536;
const PARAGRAPH_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export class DocxParagraphMarkChangePatchCollector {
  readonly patches: DocxParagraphMarkChangePatch[] = [];

  register(element: HTMLElement, id: number): string | null {
    if (
      element.dataset.documentBlockChange !== 'true' ||
      (element.dataset.blockChangeKind !== 'insertion' &&
        element.dataset.blockChangeKind !== 'deletion')
    ) {
      return null;
    }
    const key = element.dataset.blockChangeId?.trim() ?? '';
    const author = element.dataset.blockChangeAuthor?.trim() ?? '';
    const date = normalizedRevisionDate(element.dataset.blockChangeDate);
    if (
      !key ||
      !author ||
      author.length > 255 ||
      /[\u0000-\u001f\u007f]/.test(author) ||
      !Number.isSafeInteger(id) ||
      id < 1 ||
      !browserParagraphBodyMatchesChange(
        element,
        element.dataset.blockChangeKind,
        author,
        element.dataset.blockChangeDate,
      )
    ) {
      throw new Error('Document contains an invalid paragraph-mark revision.');
    }
    if (this.patches.length >= MAX_PARAGRAPH_MARK_CHANGE_PATCHES) {
      throw new Error('Document exceeds the paragraph-mark revision limit.');
    }
    const marker = `__A3S_WORK_PARAGRAPH_MARK_CHANGE_EXPORT_${this.patches.length + 1}__`;
    this.patches.push({
      marker,
      id,
      kind: element.dataset.blockChangeKind,
      author,
      date,
    });
    return marker;
  }
}

export async function patchDocxParagraphMarkChanges(
  buffer: ArrayBuffer,
  patches: readonly DocxParagraphMarkChangePatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_PARAGRAPH_MARK_CHANGE_PATCHES) {
    throw new Error('Document exceeds the paragraph-mark revision limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !PARAGRAPH_PART_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    let changed = false;
    for (const paragraph of descendants(document, 'p').filter(
      (element) => element.namespaceURI === WORD_NAMESPACE,
    )) {
      const match = paragraphMarkChangeMarker(paragraph, byMarker);
      if (!match) continue;
      if (applied.has(match.patch.marker)) {
        throw new Error(
          `Generated DOCX contains a duplicate paragraph-mark revision marker: ${match.patch.marker}.`,
        );
      }
      setParagraphMarkChange(document, paragraph, match.patch);
      match.run.remove();
      applied.add(match.patch.marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches
    .map((patch) => patch.marker)
    .filter((marker) => !applied.has(marker));
  if (missing.length) {
    throw new Error(
      `DOCX paragraph-mark revision markers were not emitted: ${missing.join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function paragraphMarkChangeMarker(
  paragraph: Element,
  patches: ReadonlyMap<string, DocxParagraphMarkChangePatch>,
): { patch: DocxParagraphMarkChangePatch; run: Element } | null {
  const matches = directChildren(paragraph, 'r').flatMap((run) => {
    if (run.namespaceURI !== WORD_NAMESPACE) return [];
    const texts = directChildren(run, 't').filter(
      (text) => text.namespaceURI === run.namespaceURI,
    );
    if (texts.length !== 1 || !runHasOnlyMarkerText(run, texts[0])) return [];
    const patch = patches.get(texts[0]?.textContent ?? '');
    return patch ? [{ patch, run }] : [];
  });
  if (matches.length > 1) {
    throw new Error(
      'Generated DOCX paragraph has duplicate paragraph-mark markers.',
    );
  }
  return matches[0] ?? null;
}

function runHasOnlyMarkerText(
  run: Element,
  markerText: Element | undefined,
): boolean {
  return Boolean(
    markerText &&
      directChildren(run).every(
        (child) =>
          child === markerText ||
          (child.localName === 'rPr' && child.namespaceURI === WORD_NAMESPACE),
      ),
  );
}

function setParagraphMarkChange(
  document: Document,
  paragraph: Element,
  patch: DocxParagraphMarkChangePatch,
): void {
  const properties = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  if (properties.length > 1) {
    throw new Error('Generated DOCX paragraph contains duplicate properties.');
  }
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  const paragraphProperties =
    properties[0] ?? insertParagraphProperties(document, paragraph, prefix);
  const runProperties = directChildren(paragraphProperties, 'rPr').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  if (runProperties.length > 1) {
    throw new Error(
      'Generated DOCX paragraph contains duplicate paragraph-mark properties.',
    );
  }
  const markProperties =
    runProperties[0] ??
    insertRunProperties(document, paragraphProperties, prefix);
  if (
    directChildren(markProperties).some(
      (element) =>
        (element.localName === 'ins' || element.localName === 'del') &&
        element.namespaceURI === WORD_NAMESPACE,
    )
  ) {
    throw new Error(
      'Generated DOCX paragraph already contains a paragraph-mark revision.',
    );
  }
  const revision = document.createElementNS(
    WORD_NAMESPACE,
    `${prefix}:${patch.kind === 'deletion' ? 'del' : 'ins'}`,
  );
  setWordAttribute(revision, prefix, 'id', String(patch.id));
  setWordAttribute(revision, prefix, 'author', patch.author);
  setWordAttribute(revision, prefix, 'date', patch.date);
  markProperties.insertBefore(revision, markProperties.firstChild);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
  prefix: string,
): Element {
  const properties = document.createElementNS(WORD_NAMESPACE, `${prefix}:pPr`);
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function insertRunProperties(
  document: Document,
  properties: Element,
  prefix: string,
): Element {
  const runProperties = document.createElementNS(
    WORD_NAMESPACE,
    `${prefix}:rPr`,
  );
  const following = directChildren(properties).find(
    (child) =>
      child.namespaceURI === WORD_NAMESPACE &&
      (child.localName === 'sectPr' || child.localName === 'pPrChange'),
  );
  properties.insertBefore(runProperties, following ?? null);
  return runProperties;
}

function browserParagraphBodyMatchesChange(
  paragraph: HTMLElement,
  kind: 'insertion' | 'deletion',
  author: string,
  date: string | undefined,
): boolean {
  const tag = paragraph.tagName.toLowerCase();
  if (!['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tag)) {
    return false;
  }
  if (
    paragraph.querySelector(
      'br, img, svg, math, audio, video, canvas, iframe, object, embed, [contenteditable="false"]',
    )
  ) {
    return false;
  }
  const revisions = Array.from(
    paragraph.querySelectorAll<HTMLElement>(
      'ins[data-document-change="true"], del[data-document-change="true"]',
    ),
  );
  if (!revisions.length) return false;
  const expectedTag = kind === 'deletion' ? 'del' : 'ins';
  if (
    revisions.some(
      (revision) =>
        revision.tagName.toLowerCase() !== expectedTag ||
        revision.dataset.changeKind !== kind ||
        revision.dataset.changeAuthor?.trim() !== author ||
        !sameRevisionDate(revision.dataset.changeDate, date) ||
        revision.parentElement?.closest('[data-document-change="true"]'),
    ) ||
    paragraph.querySelector(
      '[data-document-change="true"]:not(ins):not(del), ins:not([data-document-change="true"]), del:not([data-document-change="true"])',
    )
  ) {
    return false;
  }
  const walker = paragraph.ownerDocument.createTreeWalker(
    paragraph,
    NodeFilter.SHOW_TEXT,
  );
  let hasText = false;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.data) continue;
    hasText = true;
    const revision = node.parentElement?.closest<HTMLElement>(
      '[data-document-change="true"]',
    );
    if (!revision || !paragraph.contains(revision)) return false;
  }
  return hasText;
}

function sameRevisionDate(
  left: string | undefined,
  right: string | undefined,
): boolean {
  const leftValue = left?.trim() ?? '';
  const rightValue = right?.trim() ?? '';
  if (!leftValue || !rightValue) return leftValue === rightValue;
  const leftTime = Date.parse(leftValue);
  const rightTime = Date.parse(rightValue);
  return (
    Number.isFinite(leftTime) &&
    Number.isFinite(rightTime) &&
    leftTime === rightTime
  );
}

function setWordAttribute(
  element: Element,
  prefix: string,
  name: string,
  value: string,
): void {
  element.setAttributeNS(WORD_NAMESPACE, `${prefix}:${name}`, value);
}

function normalizedRevisionDate(value: string | undefined): string {
  const time = Date.parse(value?.trim() ?? '');
  return Number.isFinite(time)
    ? new Date(time).toISOString()
    : new Date().toISOString();
}
