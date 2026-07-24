import type {
  DocumentParagraphLineRule,
  DocumentParagraphSpacing,
} from './work-document-paragraph-formatting';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxParagraphSpacingMarker {
  marker: string;
  spacing: Partial<DocumentParagraphSpacing>;
}

export interface ImportedDocxParagraphSpacingMarkers {
  paragraphs: ImportedDocxParagraphSpacingMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const PARAGRAPH_SPACING_MARKER_PATTERN = /__A3S_WORK_PARAGRAPH_SPACING_\d+__/g;

export function markDocxParagraphSpacing(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
): ImportedDocxParagraphSpacingMarkers {
  const paragraphs: ImportedDocxParagraphSpacingMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  for (const paragraph of descendants(document, 'p')) {
    let properties = directChild(paragraph, 'pPr');
    const sources = docxParagraphPropertySources(properties, styles);
    if (sources.some((source) => directChild(source, 'numPr'))) continue;
    const spacing = paragraphSpacing(sources);
    if (!Object.keys(spacing).length) continue;
    properties ??= insertParagraphProperties(document, paragraph);
    const marker = `__A3S_WORK_PARAGRAPH_SPACING_${paragraphs.length + 1}__`;
    insertParagraphMarker(document, paragraph, properties, marker);
    paragraphs.push({ marker, spacing });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphSpacingMarkers(
  document: Document,
  markers: ImportedDocxParagraphSpacingMarkers,
): void {
  const spacingByMarker = new Map(
    markers.paragraphs.map(({ marker, spacing }) => [marker, spacing]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_SPACING_')) continue;
    const block = closestParagraphBlock(node.parentElement);
    node.data = node.data.replace(
      PARAGRAPH_SPACING_MARKER_PATTERN,
      (marker) => {
        const spacing = spacingByMarker.get(marker);
        if (block && spacing) applyParagraphSpacing(block, spacing);
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphSpacingMarkers(
  markers: ImportedDocxParagraphSpacingMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function paragraphSpacing(
  propertySources: readonly Element[],
): Partial<DocumentParagraphSpacing> {
  let before: number | undefined;
  let after: number | undefined;
  let line: number | undefined;
  let lineRule: DocumentParagraphLineRule | undefined;
  for (const properties of propertySources) {
    const spacing = directChild(properties, 'spacing');
    if (!spacing) continue;
    const sourceBefore = numericAttribute(spacing, 'before');
    const sourceAfter = numericAttribute(spacing, 'after');
    const sourceLine = numericAttribute(spacing, 'line');
    const sourceLineRule = normalizedLineRule(
      wordAttribute(spacing, 'lineRule'),
    );
    if (sourceBefore !== undefined) before = sourceBefore;
    if (sourceAfter !== undefined) after = sourceAfter;
    if (sourceLine !== undefined) {
      line = sourceLine;
      lineRule = sourceLineRule ?? 'auto';
    } else if (sourceLineRule) {
      lineRule = sourceLineRule;
    }
  }

  const result: Partial<DocumentParagraphSpacing> = {};
  if (before !== undefined) result.before = twipsToPoints(before);
  if (after !== undefined) result.after = twipsToPoints(after);
  if (line !== undefined) {
    const rule = lineRule ?? 'auto';
    result.lineHeight =
      rule === 'auto'
        ? formatNumber(line / 240)
        : `${formatNumber(line / 20)}pt`;
    result.lineRule = rule;
  }
  return result;
}

function applyParagraphSpacing(
  element: HTMLElement,
  spacing: Partial<DocumentParagraphSpacing>,
): void {
  if (spacing.before !== undefined && spacing.before !== null) {
    const value = formatNumber(spacing.before);
    element.dataset.officeSpaceBefore = value;
    element.style.marginTop = `${value}pt`;
  }
  if (spacing.after !== undefined && spacing.after !== null) {
    const value = formatNumber(spacing.after);
    element.dataset.officeSpaceAfter = value;
    element.style.marginBottom = `${value}pt`;
  }
  if (spacing.lineHeight) element.style.lineHeight = spacing.lineHeight;
  if (spacing.lineRule) element.dataset.officeLineRule = spacing.lineRule;
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

function numericAttribute(element: Element, name: string): number | undefined {
  const raw = wordAttribute(element, name);
  if (raw === null || !raw.trim()) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}

function normalizedLineRule(
  value: string | null,
): DocumentParagraphLineRule | undefined {
  if (value === 'auto' || value === 'atLeast') return value;
  if (value === 'exact' || value === 'exactly') return 'exact';
  return undefined;
}

function twipsToPoints(value: number): number {
  return Number((value / 20).toFixed(2));
}

function formatNumber(value: number): string {
  return Number(value.toFixed(3)).toString();
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
