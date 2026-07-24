import type { DocumentParagraphDirection } from './work-document-paragraph-formatting';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxParagraphDirectionMarker {
  marker: string;
  direction: DocumentParagraphDirection;
}

export interface ImportedDocxParagraphDirectionMarkers {
  paragraphs: ImportedDocxParagraphDirectionMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const PARAGRAPH_DIRECTION_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_DIRECTION_\d+__/g;

export function markDocxParagraphDirections(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
): ImportedDocxParagraphDirectionMarkers {
  const paragraphs: ImportedDocxParagraphDirectionMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  for (const paragraph of descendants(document, 'p')) {
    let properties = directChild(paragraph, 'pPr');
    const sources = docxParagraphPropertySources(properties, styles);
    const direction = paragraphDirection(sources);
    if (!direction) continue;
    properties ??= insertParagraphProperties(document, paragraph);
    const marker = `__A3S_WORK_PARAGRAPH_DIRECTION_${paragraphs.length + 1}__`;
    insertParagraphMarker(document, paragraph, properties, marker);
    paragraphs.push({ marker, direction });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphDirectionMarkers(
  document: Document,
  markers: ImportedDocxParagraphDirectionMarkers,
): void {
  const directionByMarker = new Map(
    markers.paragraphs.map(({ marker, direction }) => [marker, direction]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_DIRECTION_')) continue;
    const block = closestParagraphBlock(node.parentElement);
    node.data = node.data.replace(
      PARAGRAPH_DIRECTION_MARKER_PATTERN,
      (marker) => {
        const direction = directionByMarker.get(marker);
        if (block && direction) block.setAttribute('dir', direction);
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphDirectionMarkers(
  markers: ImportedDocxParagraphDirectionMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function paragraphDirection(
  propertySources: readonly Element[],
): DocumentParagraphDirection | null {
  let direction: DocumentParagraphDirection | null = null;
  for (const properties of propertySources) {
    const bidi = directChild(properties, 'bidi');
    if (bidi) direction = onOffValue(bidi) ? 'rtl' : 'ltr';
  }
  return direction;
}

function onOffValue(element: Element): boolean {
  const value = (attribute(element, 'val') ?? attribute(element, 'w:val'))
    ?.trim()
    .toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  properties: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.insertBefore(run, properties.nextSibling);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
): Element {
  const properties = document.createElementNS(WORD_NAMESPACE, 'w:pPr');
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function closestParagraphBlock(element: Element | null): HTMLElement | null {
  const block = element?.closest('p, h1, h2, h3, h4, h5, h6, blockquote, li');
  return block instanceof HTMLElement ? block : null;
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
