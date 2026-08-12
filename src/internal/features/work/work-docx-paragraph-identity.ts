import JSZip from 'jszip';
import {
  applyDocumentParagraphIdentityToElement,
  createDocumentParagraphIdentityRegistry,
  documentParagraphIdentityFromElement,
  uniqueDocumentParagraphIdentity,
  type WorkDocumentParagraphIdentity,
  type WorkDocumentParagraphIdentityRegistry,
} from './work-document-paragraph-identity';
import {
  applyDocumentTableRowIdentityToElement,
  documentTableRowIdentityFromElement,
} from './work-document-table-row-identity';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxParagraphIdentityPatch {
  marker: string;
  identity: WorkDocumentParagraphIdentity;
  target: 'paragraph' | 'tableRow';
}

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_PARAGRAPH_IDENTITY_PATCHES = 65_536;
const PARAGRAPH_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export class DocxParagraphIdentityPatchCollector {
  readonly patches: DocxParagraphIdentityPatch[] = [];
  private readonly registry: WorkDocumentParagraphIdentityRegistry =
    createDocumentParagraphIdentityRegistry();
  private nextMarker = 1;

  constructor(private readonly source: string) {}

  marker(element: HTMLElement): string {
    const identity = uniqueDocumentParagraphIdentity(
      documentParagraphIdentityFromElement(element) ?? {},
      this.registry,
    );
    applyDocumentParagraphIdentityToElement(element, identity);
    return this.registerPatch(identity, 'paragraph');
  }

  rowMarker(element: HTMLTableRowElement): string {
    const current = documentTableRowIdentityFromElement(element);
    const identity = uniqueDocumentParagraphIdentity(
      {
        paragraphId: current?.rowId,
        textId: current?.rowTextId,
      },
      this.registry,
    );
    applyDocumentTableRowIdentityToElement(element, {
      rowId: identity.paragraphId,
      rowTextId: identity.textId,
    });
    return this.registerPatch(identity, 'tableRow');
  }

  private registerPatch(
    identity: WorkDocumentParagraphIdentity,
    target: DocxParagraphIdentityPatch['target'],
  ): string {
    if (this.patches.length >= MAX_PARAGRAPH_IDENTITY_PATCHES) {
      throw new Error('Document exceeds the paragraph-identity limit.');
    }
    let marker = '';
    do {
      marker = `__A3S_PARAGRAPH_IDENTITY_${this.nextMarker}__`;
      this.nextMarker += 1;
    } while (this.source.includes(marker));
    this.patches.push({ marker, identity, target });
    return marker;
  }
}

export async function patchDocxParagraphIdentities(
  buffer: ArrayBuffer,
  patches: readonly DocxParagraphIdentityPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_PARAGRAPH_IDENTITY_PATCHES) {
    throw new Error('Document exceeds the paragraph-identity limit.');
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
      const match = paragraphIdentityPatch(paragraph, byMarker);
      if (!match) continue;
      if (match.patch.target === 'tableRow') {
        const row = closestWordAncestor(paragraph.parentElement, 'tr');
        if (!row || !paragraphHasOnlyMarkerRun(paragraph, match.run)) continue;
        paragraph.remove();
        setWord2010Identity(document, row, match.patch.identity);
      } else {
        match.run.remove();
        setWord2010Identity(document, paragraph, match.patch.identity);
      }
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
      `DOCX paragraph identity markers were not emitted: ${missing.join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function paragraphIdentityPatch(
  paragraph: Element,
  patches: ReadonlyMap<string, DocxParagraphIdentityPatch>,
): { patch: DocxParagraphIdentityPatch; run: Element } | null {
  const matches = directChildren(paragraph, 'r').flatMap((run) => {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(run.namespaceURI ?? '')) return [];
    const text = directChildren(run, 't').find((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    const patch = text ? patches.get(text.textContent ?? '') : undefined;
    return text && patch && runHasOnlyMarkerText(run, text)
      ? [{ patch, run }]
      : [];
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

function paragraphHasOnlyMarkerRun(paragraph: Element, markerRun: Element) {
  return directChildren(paragraph).every(
    (child) =>
      child === markerRun ||
      (child.localName === 'pPr' &&
        DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')),
  );
}

function closestWordAncestor(
  element: Element | null,
  localName: string,
): Element | null {
  let current = element;
  while (current) {
    if (
      current.localName === localName &&
      DOCX_WORDPROCESSING_NAMESPACES.has(current.namespaceURI ?? '')
    ) {
      return current;
    }
    current = current.parentElement;
  }
  return null;
}

function setWord2010Identity(
  document: Document,
  element: Element,
  identity: WorkDocumentParagraphIdentity,
): void {
  const root = document.documentElement;
  const prefix = ensureNamespacePrefix(root, 'w14', WORD_2010_NAMESPACE);
  element.setAttributeNS(
    WORD_2010_NAMESPACE,
    `${prefix}:paraId`,
    identity.paragraphId,
  );
  element.setAttributeNS(
    WORD_2010_NAMESPACE,
    `${prefix}:textId`,
    identity.textId,
  );
  ensureIgnorableNamespace(root, prefix, WORD_2010_NAMESPACE);
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
  if (tokens.some((token) => xmlNamespaceUri(root, token) === namespace)) {
    return;
  }
  tokens.push(prefix);
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
