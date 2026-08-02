import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import {
  type DocxTableStyleSource,
  docxTableParagraphPropertySources,
  resolveDocxTableStyleResolver,
} from './work-docx-table-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';

export type ImportedDocxParagraphAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'justify';

export interface ImportedDocxParagraphAlignmentMarker {
  marker: string;
  alignment: ImportedDocxParagraphAlignment;
}

export interface ImportedDocxParagraphAlignmentMarkers {
  paragraphs: ImportedDocxParagraphAlignmentMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const PARAGRAPH_ALIGNMENT_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_ALIGNMENT_\d+__/g;

export function markDocxParagraphAlignments(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxParagraphAlignmentMarkers {
  const paragraphs: ImportedDocxParagraphAlignmentMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  for (const paragraph of descendants(document, 'p')) {
    let properties = directChild(paragraph, 'pPr');
    const sources = docxParagraphPropertySources(
      properties,
      styles,
      docxTableParagraphPropertySources(paragraph, tableStyles),
    );
    const alignment = paragraphAlignment(sources);
    if (!alignment) continue;
    properties ??= insertParagraphProperties(document, paragraph);
    const marker = `__A3S_WORK_PARAGRAPH_ALIGNMENT_${paragraphs.length + 1}__`;
    insertParagraphMarker(document, paragraph, properties, marker);
    paragraphs.push({ marker, alignment });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphAlignmentMarkers(
  document: Document,
  markers: ImportedDocxParagraphAlignmentMarkers,
): void {
  const alignmentByMarker = new Map(
    markers.paragraphs.map(({ marker, alignment }) => [marker, alignment]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_ALIGNMENT_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_ALIGNMENT_MARKER_PATTERN,
      (marker) => {
        const alignment = alignmentByMarker.get(marker);
        const block = alignment
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && alignment) block.style.textAlign = alignment;
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphAlignmentMarkers(
  markers: ImportedDocxParagraphAlignmentMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function paragraphAlignment(
  propertySources: readonly Element[],
): ImportedDocxParagraphAlignment | null {
  let direction: 'ltr' | 'rtl' = 'ltr';
  let justification: string | null = null;
  let hasJustification = false;
  for (const properties of propertySources) {
    const bidi = directChild(properties, 'bidi');
    if (bidi) direction = onOffValue(bidi) ? 'rtl' : 'ltr';
    const source = directChild(properties, 'jc');
    if (!source) continue;
    hasJustification = true;
    justification = wordAttribute(source, 'val')?.trim() ?? null;
  }
  if (!hasJustification || !justification) return null;

  if (justification === 'center') return 'center';
  if (justification === 'left') return 'left';
  if (justification === 'right') return 'right';
  if (justification === 'start') return direction === 'rtl' ? 'right' : 'left';
  if (justification === 'end') return direction === 'rtl' ? 'left' : 'right';
  if (
    justification === 'both' ||
    justification === 'distribute' ||
    justification === 'highKashida' ||
    justification === 'lowKashida' ||
    justification === 'mediumKashida' ||
    justification === 'thaiDistribute'
  ) {
    return 'justify';
  }
  return null;
}

function onOffValue(element: Element): boolean {
  const value = wordAttribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
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

function closestParagraphBlock(
  element: Element | null,
  markerNode: Node,
): HTMLElement | null {
  const block = element?.closest('p, h1, h2, h3, h4, h5, h6, blockquote, li');
  if (!(block instanceof HTMLElement)) return null;
  return block.tagName.toLowerCase() === 'li'
    ? wrapListItemParagraph(block, markerNode)
    : block;
}

function wrapListItemParagraph(
  item: HTMLElement,
  markerNode: Node,
): HTMLElement {
  let anchor: Node = markerNode;
  while (anchor.parentNode && anchor.parentNode !== item) {
    anchor = anchor.parentNode;
  }
  const children = Array.from(item.childNodes);
  const anchorIndex = children.indexOf(anchor as ChildNode);
  if (anchorIndex < 0) return item;
  let start = anchorIndex;
  let end = anchorIndex;
  while (start > 0 && !isListItemBlock(children[start - 1])) start -= 1;
  while (end + 1 < children.length && !isListItemBlock(children[end + 1])) {
    end += 1;
  }
  const paragraph = item.ownerDocument.createElement('p');
  const grouped = children.slice(start, end + 1);
  item.insertBefore(paragraph, grouped[0] ?? null);
  paragraph.append(...grouped);
  return paragraph;
}

function isListItemBlock(node: Node | undefined): boolean {
  if (!(node instanceof HTMLElement)) return false;
  return [
    'blockquote',
    'div',
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
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
