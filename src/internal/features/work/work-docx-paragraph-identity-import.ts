import {
  applyDocumentParagraphIdentityToElement,
  normalizeDocumentParagraphId,
  normalizeDocumentParagraphIdentity,
  type WorkDocumentParagraphIdentity,
} from './work-document-paragraph-identity';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChildren } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

export interface ImportedDocxParagraphIdentityMarker
  extends WorkDocumentParagraphIdentity {
  marker: string;
}

export interface ImportedDocxParagraphIdentityMarkers {
  paragraphs: ImportedDocxParagraphIdentityMarker[];
}

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const PARAGRAPH_IDENTITY_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_IDENTITY_\d+__/g;
const MAX_PARAGRAPH_IDENTITIES = 65_536;

export function markDocxParagraphIdentities(
  document: Document,
): ImportedDocxParagraphIdentityMarkers {
  const counts = wordIdentityIdCounts(document);
  const candidates = descendants(document, 'p').flatMap((element) => {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
      return [];
    }
    const identity = wordParagraphIdentity(element);
    return identity ? [{ element, identity }] : [];
  });
  if (candidates.length > MAX_PARAGRAPH_IDENTITIES) {
    throw new Error('DOCX exceeds the paragraph-identity limit.');
  }
  const paragraphs: ImportedDocxParagraphIdentityMarker[] = [];
  for (const candidate of candidates) {
    if (counts.get(candidate.identity.paragraphId) !== 1) continue;
    const marker = `__A3S_WORK_PARAGRAPH_IDENTITY_${paragraphs.length + 1}__`;
    insertParagraphMarker(candidate.element, marker);
    paragraphs.push({ marker, ...candidate.identity });
  }
  return { paragraphs };
}

function wordIdentityIdCounts(document: Document): Map<string, number> {
  const counts = new Map<string, number>();
  for (const localName of ['p', 'tr']) {
    for (const element of descendants(document, localName)) {
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
        continue;
      }
      const paragraphId = normalizeDocumentParagraphId(
        word2010Attribute(element, 'paraId'),
      );
      if (!paragraphId) continue;
      counts.set(paragraphId, (counts.get(paragraphId) ?? 0) + 1);
    }
  }
  return counts;
}

export function applyImportedDocxParagraphIdentityMarkers(
  document: Document,
  markers: ImportedDocxParagraphIdentityMarkers,
): void {
  const identities = new Map(
    markers.paragraphs.map(({ marker, paragraphId, textId }) => [
      marker,
      { paragraphId, textId },
    ]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_IDENTITY_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_IDENTITY_MARKER_PATTERN,
      (marker) => {
        const identity = identities.get(marker);
        const block = identity
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (identity && block) {
          applyDocumentParagraphIdentityToElement(block, identity);
        }
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphIdentityMarkers(
  markers: ImportedDocxParagraphIdentityMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function wordParagraphIdentity(
  paragraph: Element,
): WorkDocumentParagraphIdentity | null {
  return normalizeDocumentParagraphIdentity({
    paragraphId: word2010Attribute(paragraph, 'paraId'),
    textId: word2010Attribute(paragraph, 'textId'),
  });
}

function word2010Attribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

function insertParagraphMarker(paragraph: Element, marker: string): void {
  const document = paragraph.ownerDocument;
  const namespace = paragraph.namespaceURI ?? '';
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChildren(paragraph, 'pPr').find((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
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
