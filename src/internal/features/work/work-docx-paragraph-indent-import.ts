import {
  normalizeDocumentParagraphIndent,
  type DocumentParagraphIndent,
} from './work-document-paragraph-formatting';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxParagraphIndentMarker {
  marker: string;
  indent: DocumentParagraphIndent;
}

export interface ImportedDocxParagraphIndentMarkers {
  paragraphs: ImportedDocxParagraphIndentMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TWIPS_PER_CSS_PIXEL = 15;
const PARAGRAPH_INDENT_MARKER_PATTERN = /__A3S_WORK_PARAGRAPH_INDENT_\d+__/g;

export function markDocxParagraphIndents(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
): ImportedDocxParagraphIndentMarkers {
  const paragraphs: ImportedDocxParagraphIndentMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  for (const paragraph of descendants(document, 'p')) {
    let properties = directChild(paragraph, 'pPr');
    const sources = docxParagraphPropertySources(properties, styles);
    if (sources.some((source) => directChild(source, 'numPr'))) continue;
    const indent = paragraphIndent(sources);
    if (!indent.left && !indent.right && !indent.firstLine) continue;
    properties ??= insertParagraphProperties(document, paragraph);
    const marker = `__A3S_WORK_PARAGRAPH_INDENT_${paragraphs.length + 1}__`;
    insertParagraphMarker(document, paragraph, properties, marker);
    paragraphs.push({ marker, indent });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphIndentMarkers(
  document: Document,
  markers: ImportedDocxParagraphIndentMarkers,
): void {
  const indentByMarker = new Map(
    markers.paragraphs.map(({ marker, indent }) => [marker, indent]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_INDENT_')) continue;
    const block = closestParagraphBlock(node.parentElement);
    node.data = node.data.replace(PARAGRAPH_INDENT_MARKER_PATTERN, (marker) => {
      const indent = indentByMarker.get(marker);
      if (block && indent) applyParagraphIndent(block, indent);
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphIndentMarkers(
  markers: ImportedDocxParagraphIndentMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function paragraphIndent(
  propertySources: readonly Element[],
): DocumentParagraphIndent {
  let left: number | undefined;
  let right: number | undefined;
  let firstLine: number | undefined;
  for (const properties of propertySources) {
    const indent = directChild(properties, 'ind');
    if (!indent) continue;
    const sourceLeft =
      numericAttribute(indent, 'start') ?? numericAttribute(indent, 'left');
    const sourceRight =
      numericAttribute(indent, 'end') ?? numericAttribute(indent, 'right');
    const sourceFirstLine = numericAttribute(indent, 'firstLine');
    const sourceHanging = numericAttribute(indent, 'hanging');
    if (sourceLeft !== undefined) left = sourceLeft;
    if (sourceRight !== undefined) right = sourceRight;
    if (sourceFirstLine !== undefined) firstLine = sourceFirstLine;
    else if (sourceHanging !== undefined) firstLine = -sourceHanging;
  }
  return normalizeDocumentParagraphIndent({
    left: twipsToPixels(left),
    right: twipsToPixels(right),
    firstLine: twipsToPixels(firstLine),
  });
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
  const block = element?.closest('p, h1, h2, h3, h4, h5, h6, blockquote');
  return block instanceof HTMLElement ? block : null;
}

function applyParagraphIndent(
  element: HTMLElement,
  value: DocumentParagraphIndent,
): void {
  const indent = normalizeDocumentParagraphIndent(value);
  if (indent.left) {
    element.dataset.officeIndentLevel = formatNumber(indent.left / 24);
    element.style.marginLeft = `${formatNumber(indent.left)}px`;
  }
  if (indent.right) {
    element.dataset.officeIndentRight = formatNumber(indent.right);
    element.style.marginRight = `${formatNumber(indent.right)}px`;
  }
  if (indent.firstLine) {
    element.dataset.officeIndentFirstLine = formatNumber(indent.firstLine);
    element.style.textIndent = `${formatNumber(indent.firstLine)}px`;
  }
}

function numericAttribute(element: Element, name: string): number | undefined {
  const raw = attribute(element, name) ?? attribute(element, `w:${name}`);
  if (raw === null || !raw.trim()) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function twipsToPixels(value: number | undefined): number {
  if (value === undefined) return 0;
  return Math.round(value / TWIPS_PER_CSS_PIXEL);
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
