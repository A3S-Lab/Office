import JSZip from 'jszip';
import { MAX_DOCUMENT_NUMBERING_START } from './work-document-lists';
import { parseDocumentNumberingChange } from './work-document-numbering-changes';
import {
  descendants,
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxNumberingChangePatch {
  marker: string;
  id: number;
  author: string;
  date: string;
  original: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_NUMBERING_CHANGE_PATCHES = 65_536;
const NUMBERING_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export class DocxNumberingChangePatchCollector {
  readonly patches: DocxNumberingChangePatch[] = [];

  register(list: HTMLElement, itemIndex: number, id: number): string | null {
    if (
      list.tagName.toLowerCase() !== 'ol' ||
      list.dataset.changeKind !== 'numbering' ||
      !list.hasAttribute('data-document-change')
    ) {
      return null;
    }
    const key = list.dataset.changeId?.trim() ?? '';
    const author = list.dataset.changeAuthor?.trim() ?? '';
    const date = normalizedRevisionDate(list.dataset.changeDate);
    const snapshot = parseDocumentNumberingChange(
      list.dataset.changeBefore ?? '',
    );
    const value = snapshot ? snapshot.start + itemIndex : 0;
    if (
      !key ||
      !author ||
      author.length > 255 ||
      /[\u0000-\u001f\u007f]/.test(author) ||
      !snapshot ||
      !Number.isSafeInteger(itemIndex) ||
      itemIndex < 0 ||
      value < 1 ||
      value > MAX_DOCUMENT_NUMBERING_START
    ) {
      throw new Error('Document contains an invalid numbering revision.');
    }
    if (this.patches.length >= MAX_NUMBERING_CHANGE_PATCHES) {
      throw new Error('Document exceeds the numbering revision limit.');
    }
    const marker = `__A3S_WORK_NUMBERING_CHANGE_EXPORT_${this.patches.length + 1}__`;
    this.patches.push({
      marker,
      id,
      author,
      date,
      original: `%${snapshot.level + 1}:${value}:${snapshot.originalFormat}:${snapshot.originalSuffix}`,
    });
    return marker;
  }
}

export async function patchDocxNumberingChanges(
  buffer: ArrayBuffer,
  patches: readonly DocxNumberingChangePatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_NUMBERING_CHANGE_PATCHES) {
    throw new Error('Document exceeds the numbering revision limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !NUMBERING_PART_PATTERN.test(entry.name)) continue;
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
      for (const marker of numberingMarkers(paragraph, byMarker)) {
        if (applied.has(marker.patch.marker)) {
          throw new Error(
            `Generated DOCX contains a duplicate numbering revision marker: ${marker.patch.marker}.`,
          );
        }
        setNumberingChange(document, paragraph, marker.patch);
        marker.run.remove();
        applied.add(marker.patch.marker);
        changed = true;
      }
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch.marker));
  if (missing.length) {
    throw new Error(
      `DOCX numbering revision markers were not emitted: ${missing
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function numberingMarkers(
  paragraph: Element,
  patches: ReadonlyMap<string, DocxNumberingChangePatch>,
): Array<{ patch: DocxNumberingChangePatch; run: Element }> {
  return directChildren(paragraph, 'r').flatMap((run) => {
    if (run.namespaceURI !== WORD_NAMESPACE) return [];
    const texts = directChildren(run, 't').filter(
      (text) => text.namespaceURI === WORD_NAMESPACE,
    );
    if (texts.length !== 1 || !runHasOnlyMarkerText(run, texts[0])) return [];
    const patch = patches.get(texts[0]?.textContent ?? '');
    return patch ? [{ patch, run }] : [];
  });
}

function runHasOnlyMarkerText(run: Element, markerText: Element | undefined) {
  return Boolean(
    markerText &&
      directChildren(run).every(
        (child) =>
          child === markerText ||
          (child.localName === 'rPr' && child.namespaceURI === WORD_NAMESPACE),
      ),
  );
}

function setNumberingChange(
  document: Document,
  paragraph: Element,
  patch: DocxNumberingChangePatch,
): void {
  const propertyNodes = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  if (propertyNodes.length !== 1) {
    throw new Error(
      'Generated DOCX numbered paragraph must contain one property node.',
    );
  }
  const numberingNodes = directChildren(propertyNodes[0], 'numPr').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  if (numberingNodes.length !== 1) {
    throw new Error(
      'Generated DOCX numbered paragraph must contain one numbering property node.',
    );
  }
  const numbering = numberingNodes[0];
  if (
    directChildren(numbering, 'numberingChange').some(
      (element) => element.namespaceURI === WORD_NAMESPACE,
    )
  ) {
    throw new Error(
      'Generated DOCX paragraph already contains a numbering revision.',
    );
  }
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  const change = document.createElementNS(
    WORD_NAMESPACE,
    `${prefix}:numberingChange`,
  );
  setWordAttribute(change, prefix, 'id', String(patch.id));
  setWordAttribute(change, prefix, 'author', patch.author);
  setWordAttribute(change, prefix, 'date', patch.date);
  setWordAttribute(change, prefix, 'original', patch.original);
  numbering.append(change);
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
