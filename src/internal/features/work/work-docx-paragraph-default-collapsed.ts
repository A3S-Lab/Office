import JSZip from 'jszip';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { dataBoolean } from './work-docx-export-formatting';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  hasNonWhitespaceXmlText,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export type ParsedDocxParagraphDefaultCollapsed =
  | { status: 'absent' }
  | { status: 'invalid' }
  | { status: 'valid'; value: boolean };

interface DocxParagraphDefaultCollapsedPatch {
  marker: string;
  value: boolean;
}

const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_PARAGRAPH_DEFAULT_COLLAPSED_PATCHES = 65_536;
const PARAGRAPH_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export class DocxParagraphDefaultCollapsedPatchCollector {
  readonly patches: DocxParagraphDefaultCollapsedPatch[] = [];

  register(marker: string, element: HTMLElement): void {
    const value = dataBoolean(element.dataset.officeDefaultCollapsed);
    if (value === undefined) return;
    if (this.patches.length >= MAX_PARAGRAPH_DEFAULT_COLLAPSED_PATCHES) {
      throw new Error(
        'Document exceeds the default-collapsed paragraph limit.',
      );
    }
    this.patches.push({ marker, value });
  }
}

export function parseDocxParagraphDefaultCollapsed(
  properties: Element,
): ParsedDocxParagraphDefaultCollapsed {
  if (
    properties.localName !== 'pPr' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(properties.namespaceURI ?? '')
  ) {
    return { status: 'absent' };
  }
  const collapsed = directChildren(properties, 'collapsed').filter(
    (element) => element.namespaceURI === WORD_2012_NAMESPACE,
  );
  if (!collapsed.length) return { status: 'absent' };
  if (collapsed.length !== 1) return { status: 'invalid' };

  const element = collapsed[0];
  if (directChildren(element).length || hasNonWhitespaceXmlText(element)) {
    return { status: 'invalid' };
  }
  const attributes = Array.from(element.attributes).filter(
    (item) => xmlAttributeNamespace(element, item) !== XMLNS_NAMESPACE,
  );
  if (!attributes.length) return { status: 'valid', value: true };
  if (attributes.length !== 1) return { status: 'invalid' };

  const valueAttribute = attributes[0];
  if (
    xmlAttributeLocalName(valueAttribute) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(element, valueAttribute) ?? '',
    )
  ) {
    return { status: 'invalid' };
  }
  if (
    valueAttribute.value === 'true' ||
    valueAttribute.value === 'on' ||
    valueAttribute.value === '1'
  ) {
    return { status: 'valid', value: true };
  }
  if (
    valueAttribute.value === 'false' ||
    valueAttribute.value === 'off' ||
    valueAttribute.value === '0'
  ) {
    return { status: 'valid', value: false };
  }
  return { status: 'invalid' };
}

export async function patchDocxParagraphDefaultCollapsed(
  buffer: ArrayBuffer,
  patches: readonly DocxParagraphDefaultCollapsedPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_PARAGRAPH_DEFAULT_COLLAPSED_PATCHES) {
    throw new Error('Document exceeds the default-collapsed paragraph limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Set<string>();
  const entries = Object.values(archive.files).filter(
    (entry) => !entry.dir && PARAGRAPH_PART_PATTERN.test(entry.name),
  );
  for (const entry of entries) {
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    if (!isParagraphPartRoot(document.documentElement, entry.name)) continue;
    let changed = false;
    for (const paragraph of descendants(document, 'p').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    )) {
      const patch = paragraphDefaultCollapsedPatch(paragraph, byMarker);
      if (!patch) continue;
      setParagraphDefaultCollapsed(document, paragraph, patch.value);
      applied.add(patch.marker);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches
    .map((patch) => patch.marker)
    .filter((marker) => !applied.has(marker));
  if (missing.length) {
    throw new Error(
      `DOCX default-collapsed paragraph markers were not emitted: ${missing.join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function paragraphDefaultCollapsedPatch(
  paragraph: Element,
  patches: ReadonlyMap<string, DocxParagraphDefaultCollapsedPatch>,
): DocxParagraphDefaultCollapsedPatch | null {
  const matches = directChildren(paragraph, 'r').flatMap((run) => {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(run.namespaceURI ?? '')) return [];
    const text = directChildren(run, 't').find((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    const patch = text ? patches.get(text.textContent ?? '') : undefined;
    return text && patch && runHasOnlyMarkerText(run, text) ? [patch] : [];
  });
  return matches.length === 1 ? matches[0] : null;
}

function runHasOnlyMarkerText(run: Element, markerText: Element): boolean {
  return directChildren(run).every(
    (child) =>
      child === markerText ||
      (child.localName === 'rPr' &&
        DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')),
  );
}

function setParagraphDefaultCollapsed(
  document: Document,
  paragraph: Element,
  value: boolean,
): void {
  const wordNamespace = paragraph.namespaceURI ?? '';
  const root = document.documentElement;
  const wordPrefix = ensureNamespacePrefix(root, 'w', wordNamespace);
  const word2012Prefix = ensureNamespacePrefix(
    root,
    'w15',
    WORD_2012_NAMESPACE,
  );
  ensureIgnorableNamespace(root, word2012Prefix, WORD_2012_NAMESPACE);

  const propertyNodes = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === wordNamespace,
  );
  if (propertyNodes.length > 1) {
    throw new Error('Generated DOCX paragraph contains duplicate properties.');
  }
  const properties =
    propertyNodes[0] ??
    insertParagraphProperties(document, paragraph, wordPrefix, wordNamespace);
  for (const child of directChildren(properties, 'collapsed')) {
    if (child.namespaceURI === WORD_2012_NAMESPACE) child.remove();
  }
  const collapsed = document.createElementNS(
    WORD_2012_NAMESPACE,
    `${word2012Prefix}:collapsed`,
  );
  collapsed.setAttributeNS(
    wordNamespace,
    `${wordPrefix}:val`,
    value ? '1' : '0',
  );
  properties.append(collapsed);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
  wordPrefix: string,
  wordNamespace: string,
): Element {
  const properties = document.createElementNS(
    wordNamespace,
    `${wordPrefix}:pPr`,
  );
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function ensureIgnorableNamespace(
  root: Element,
  prefix: string,
  namespace: string,
): void {
  const attribute = Array.from(root.attributes).find(
    (item) =>
      xmlAttributeLocalName(item) === 'Ignorable' &&
      xmlAttributeNamespace(root, item) === MARKUP_COMPATIBILITY_NAMESPACE,
  );
  const tokens = (attribute?.value ?? '').trim().split(/\s+/).filter(Boolean);
  if (!tokens.some((token) => xmlNamespaceUri(root, token) === namespace)) {
    tokens.push(prefix);
  }
  const compatibilityPrefix = ensureNamespacePrefix(
    root,
    'mc',
    MARKUP_COMPATIBILITY_NAMESPACE,
  );
  attribute?.ownerElement?.removeAttributeNode(attribute);
  root.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    `${compatibilityPrefix}:Ignorable`,
    tokens.join(' '),
  );
}

function ensureNamespacePrefix(
  root: Element,
  preferred: string,
  namespace: string,
): string {
  const existing = xmlDeclaredPrefix(root, namespace);
  if (existing) return existing;
  if (!xmlNamespaceUri(root, preferred)) {
    root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${preferred}`, namespace);
    return preferred;
  }
  let suffix = 1;
  let prefix = '';
  do {
    prefix = `a3s${suffix}`;
    suffix += 1;
  } while (xmlNamespaceUri(root, prefix));
  root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, namespace);
  return prefix;
}

function isParagraphPartRoot(root: Element, path: string): boolean {
  const expected = /^word\/header\d*\.xml$/i.test(path)
    ? 'hdr'
    : /^word\/footer\d*\.xml$/i.test(path)
      ? 'ftr'
      : /^word\/footnotes\.xml$/i.test(path)
        ? 'footnotes'
        : /^word\/endnotes\.xml$/i.test(path)
          ? 'endnotes'
          : 'document';
  return (
    root.localName === expected &&
    DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
  );
}
