import {
  DOCUMENT_PARAGRAPH_BORDER_EDGES,
  DOCUMENT_PARAGRAPH_BORDER_STYLES,
  type DocumentParagraphBorder,
  type DocumentParagraphBorderColor,
  type DocumentParagraphBorderEdge,
  type DocumentParagraphBorders,
  documentParagraphBordersDomAttributes,
  normalizeDocumentParagraphBorder,
  normalizeDocumentParagraphBorders,
} from './work-document-paragraph-borders';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  type DocxParagraphStyleSource,
  docxParagraphPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  type DocxTableStyleSource,
  docxTableParagraphPropertySources,
  resolveDocxTableStyleResolver,
} from './work-docx-table-styles';
import {
  type DocxThemeResolver,
  type DocxThemeSource,
  docxThemeColor,
  resolveDocxThemeResolver,
} from './work-docx-theme';
import type { DocxThemeColorReference } from './work-docx-theme-reference';
import { descendants, directChildren } from './work-ooxml-package';

export interface ImportedDocxParagraphBorderMarker {
  marker: string;
  borders: DocumentParagraphBorders;
}

export interface ImportedDocxParagraphBorderMarkers {
  paragraphs: ImportedDocxParagraphBorderMarker[];
  invalidCount: number;
  spoofedCount: number;
}

export interface ResolvedDocxParagraphBorders {
  borders: DocumentParagraphBorders | null;
  invalidCount: number;
  spoofedCount: number;
}

interface ParsedBorderContainer {
  borders: DocumentParagraphBorders;
  invalidCount: number;
  spoofedCount: number;
  valid: boolean;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const PARAGRAPH_BORDER_MARKER_PATTERN = /__A3S_WORK_PARAGRAPH_BORDER_\d+__/g;
const BORDER_EDGE_SET = new Set<string>(DOCUMENT_PARAGRAPH_BORDER_EDGES);
const BORDER_STYLE_SET = new Set<string>(DOCUMENT_PARAGRAPH_BORDER_STYLES);
const WORD_THEME_COLORS = new Set([
  'dark1',
  'light1',
  'dark2',
  'light2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hyperlink',
  'followedHyperlink',
  'none',
  'background1',
  'text1',
  'background2',
  'text2',
]);
const BORDER_ATTRIBUTES = new Set([
  'val',
  'color',
  'themeColor',
  'themeTint',
  'themeShade',
  'sz',
  'space',
  'shadow',
  'frame',
]);

export function markDocxParagraphBorders(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxParagraphBorderMarkers {
  const paragraphs: ImportedDocxParagraphBorderMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  const theme = resolveDocxThemeResolver(themeSource);
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const paragraph of descendants(document, 'p')) {
    if (!isWordElement(paragraph)) continue;
    const properties = wordDirectChildren(paragraph, 'pPr')[0];
    const resolved = resolveDocxParagraphBordersForParagraph(
      paragraph,
      styles,
      theme,
      tableStyles,
    );
    invalidCount += resolved.invalidCount;
    spoofedCount += resolved.spoofedCount;
    const borders = resolved.borders;
    if (!borders) continue;
    const marker = `__A3S_WORK_PARAGRAPH_BORDER_${paragraphs.length + 1}__`;
    const writableProperties =
      properties ?? insertParagraphProperties(document, paragraph);
    insertParagraphMarker(document, paragraph, writableProperties, marker);
    paragraphs.push({ marker, borders });
  }
  return { paragraphs, invalidCount, spoofedCount };
}

export function resolveDocxParagraphBordersForParagraph(
  paragraph: Element,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ResolvedDocxParagraphBorders {
  if (!isWordElement(paragraph) || paragraph.localName !== 'p') {
    return { borders: null, invalidCount: 0, spoofedCount: 0 };
  }
  const properties = wordDirectChildren(paragraph, 'pPr')[0];
  const resolved = resolveParagraphBorderSources(
    docxParagraphPropertySources(
      properties,
      resolveDocxParagraphStyleResolver(styleSource),
      docxTableParagraphPropertySources(
        paragraph,
        resolveDocxTableStyleResolver(tableStyleSource),
      ),
    ),
    resolveDocxThemeResolver(themeSource),
  );
  return {
    borders: normalizeDocumentParagraphBorders(resolved.borders),
    invalidCount: resolved.invalidCount,
    spoofedCount: resolved.spoofedCount,
  };
}

export function applyImportedDocxParagraphBorderMarkers(
  document: Document,
  markers: ImportedDocxParagraphBorderMarkers,
): void {
  const bordersByMarker = new Map(
    markers.paragraphs.map(({ marker, borders }) => [marker, borders]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_BORDER_')) continue;
    node.data = node.data.replace(PARAGRAPH_BORDER_MARKER_PATTERN, (marker) => {
      const borders = bordersByMarker.get(marker);
      const block = borders
        ? closestParagraphBlock(node.parentElement, node)
        : null;
      if (block && borders) applyParagraphBorders(block, borders);
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphBorderMarkers(
  markers: ImportedDocxParagraphBorderMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

export function parseDirectDocxParagraphBorders(
  properties: Element,
  themeSource?: DocxThemeSource,
): DocumentParagraphBorders | null | undefined {
  const named = directChildren(properties, 'pBdr');
  const candidates = named.filter(
    (element) => element.namespaceURI === properties.namespaceURI,
  );
  if (!candidates.length) return undefined;
  if (candidates.length !== 1) return null;
  const parsed = parseBorderContainer(
    candidates[0],
    resolveDocxThemeResolver(themeSource),
  );
  if (!parsed.valid) return null;
  return normalizeDocumentParagraphBorders(parsed.borders) ?? undefined;
}

function resolveParagraphBorderSources(
  propertySources: readonly Element[],
  theme: DocxThemeResolver,
): {
  borders: DocumentParagraphBorders;
  invalidCount: number;
  spoofedCount: number;
} {
  const borders: DocumentParagraphBorders = {};
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const properties of propertySources) {
    const named = directChildren(properties, 'pBdr');
    spoofedCount += named.filter(
      (element) => element.namespaceURI !== properties.namespaceURI,
    ).length;
    const candidates = named.filter(
      (element) => element.namespaceURI === properties.namespaceURI,
    );
    if (!candidates.length) continue;
    if (candidates.length !== 1) {
      invalidCount += candidates.length;
      Object.assign(borders, resetParagraphBorders());
      continue;
    }
    const parsed = parseBorderContainer(candidates[0], theme);
    invalidCount += parsed.invalidCount;
    spoofedCount += parsed.spoofedCount;
    if (!parsed.valid) {
      invalidCount += 1;
      Object.assign(borders, resetParagraphBorders());
      continue;
    }
    Object.assign(borders, parsed.borders);
  }
  return { borders, invalidCount, spoofedCount };
}

function parseBorderContainer(
  container: Element,
  theme: DocxThemeResolver,
): ParsedBorderContainer {
  const borders: DocumentParagraphBorders = {};
  let invalidCount = 0;
  let spoofedCount = 0;
  if (
    !isWordElement(container) ||
    hasSemanticAttributes(container) ||
    Array.from(container.childNodes).some(
      (node) => node.nodeType === Node.TEXT_NODE && node.textContent?.trim(),
    )
  ) {
    return { borders, invalidCount, spoofedCount, valid: false };
  }
  let previousIndex = -1;
  const seen = new Set<DocumentParagraphBorderEdge>();
  for (const child of directChildren(container)) {
    const edge = child.localName as DocumentParagraphBorderEdge;
    if (child.namespaceURI !== container.namespaceURI) {
      spoofedCount += 1;
      continue;
    }
    const index = DOCUMENT_PARAGRAPH_BORDER_EDGES.indexOf(edge);
    if (!BORDER_EDGE_SET.has(edge) || index < previousIndex || seen.has(edge)) {
      return { borders, invalidCount, spoofedCount, valid: false };
    }
    previousIndex = index;
    seen.add(edge);
    const border = parseDocxBorderElement(child, theme);
    if (!border) {
      invalidCount += 1;
      borders[edge] = { style: 'nil' };
    } else {
      borders[edge] = border;
    }
  }
  return { borders, invalidCount, spoofedCount, valid: true };
}

export function parseDocxBorderElement(
  element: Element,
  theme: DocxThemeResolver,
): DocumentParagraphBorder | null {
  const attributes = wordLeafAttributes(element, BORDER_ATTRIBUTES);
  const style = attributes?.get('val')?.trim();
  if (!attributes || !style || !BORDER_STYLE_SET.has(style)) return null;
  const color = parseBorderColor(attributes, theme);
  if (color === null) return null;
  const size = unsignedInteger(attributes.get('sz'));
  const space = unsignedInteger(attributes.get('space'));
  const shadow = onOffAttribute(attributes.get('shadow'));
  const frame = onOffAttribute(attributes.get('frame'));
  if (size === null || space === null || shadow === null || frame === null)
    return null;
  return normalizeDocumentParagraphBorder({
    style,
    ...(color ? { color } : {}),
    ...(size !== undefined ? { size } : {}),
    ...(space !== undefined ? { space } : {}),
    ...(shadow !== undefined ? { shadow } : {}),
    ...(frame !== undefined ? { frame } : {}),
  });
}

function parseBorderColor(
  attributes: ReadonlyMap<string, string>,
  theme: DocxThemeResolver,
): DocumentParagraphBorderColor | null | undefined {
  const rawColor = attributes.get('color')?.trim();
  const rawTheme = attributes.get('themeColor')?.trim();
  const rawTint = attributes.get('themeTint')?.trim();
  const rawShade = attributes.get('themeShade')?.trim();
  if (
    rawColor === undefined &&
    rawTheme === undefined &&
    rawTint === undefined &&
    rawShade === undefined
  ) {
    return undefined;
  }
  const color =
    rawColor === 'auto'
      ? ('auto' as const)
      : rawColor && /^[0-9a-f]{6}$/i.test(rawColor)
        ? (`#${rawColor.toLowerCase()}` as const)
        : null;
  const themeName =
    rawTheme && WORD_THEME_COLORS.has(rawTheme) ? rawTheme : null;
  const tint = byteHex(rawTint);
  const shade = byteHex(rawShade);
  if (
    (rawColor !== undefined && !color) ||
    (rawTheme !== undefined && !themeName) ||
    (rawTint !== undefined && !tint) ||
    (rawShade !== undefined && !shade) ||
    ((!themeName || themeName === 'none') && (tint || shade))
  ) {
    return null;
  }
  if (themeName && themeName !== 'none') {
    const resolved = docxThemeColor(
      theme,
      themeName,
      tint,
      tint ? undefined : shade,
    );
    const fallback = color && color !== 'auto' ? color : null;
    const semanticValue = resolved ? (`#${resolved}` as const) : fallback;
    if (!semanticValue) return null;
    return {
      value: semanticValue,
      theme: themeReference(themeName, semanticValue, tint, shade),
    };
  }
  if (themeName === 'none') {
    const value = color ?? ('auto' as const);
    return {
      value,
      theme: themeReference(
        themeName,
        value === 'auto' ? '#000000' : value,
        tint,
        shade,
      ),
    };
  }
  if (!color) return null;
  return { value: color };
}

function themeReference(
  theme: string,
  resolved: `#${string}`,
  tint: string | undefined,
  shade: string | undefined,
): DocxThemeColorReference {
  return {
    theme,
    resolved,
    ...(tint ? { tint } : {}),
    ...(shade ? { shade } : {}),
  };
}

function wordLeafAttributes(
  element: Element,
  allowed: ReadonlySet<string>,
): Map<string, string> | null {
  if (
    !isWordElement(element) ||
    directChildren(element).length ||
    Array.from(element.childNodes).some(
      (node) => node.nodeType === Node.TEXT_NODE && node.textContent?.trim(),
    )
  ) {
    return null;
  }
  const result = new Map<string, string>();
  for (const attribute of Array.from(element.attributes)) {
    if (isNamespaceDeclaration(attribute)) continue;
    const name = xmlAttributeLocalName(attribute);
    if (
      xmlAttributeNamespace(element, attribute) !== element.namespaceURI ||
      !allowed.has(name) ||
      result.has(name)
    ) {
      return null;
    }
    result.set(name, attribute.value);
  }
  return result;
}

function hasSemanticAttributes(element: Element): boolean {
  return Array.from(element.attributes).some(
    (attribute) => !isNamespaceDeclaration(attribute),
  );
}

function isNamespaceDeclaration(attribute: Attr): boolean {
  return (
    attribute.namespaceURI === XMLNS_NAMESPACE ||
    attribute.name === 'xmlns' ||
    attribute.name.startsWith('xmlns:')
  );
}

function unsignedInteger(value: string | undefined): number | null | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim();
  if (!/^\+?\d+$/.test(normalized)) return null;
  const parsed = Number(normalized);
  return Number.isSafeInteger(parsed) ? parsed : null;
}

function onOffAttribute(value: string | undefined): boolean | null | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === '1' || normalized === 'true' || normalized === 'on')
    return true;
  if (normalized === '0' || normalized === 'false' || normalized === 'off')
    return false;
  return null;
}

function byteHex(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return /^[0-9a-f]{2}$/i.test(value) ? value.toUpperCase() : undefined;
}

function resetParagraphBorders(): DocumentParagraphBorders {
  return Object.fromEntries(
    DOCUMENT_PARAGRAPH_BORDER_EDGES.map((edge) => [edge, { style: 'nil' }]),
  );
}

function applyParagraphBorders(
  element: HTMLElement,
  borders: DocumentParagraphBorders,
): void {
  for (const [name, value] of Object.entries(
    documentParagraphBordersDomAttributes(borders),
  )) {
    if (name !== 'style') {
      element.setAttribute(name, value);
      continue;
    }
    for (const declaration of value.split(';')) {
      const separator = declaration.indexOf(':');
      if (separator < 0) continue;
      element.style.setProperty(
        declaration.slice(0, separator).trim(),
        declaration.slice(separator + 1).trim(),
      );
    }
  }
}

function insertParagraphMarker(
  document: Document,
  paragraph: Element,
  properties: Element,
  marker: string,
): void {
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix ?? 'w';
  const run = document.createElementNS(namespace, `${prefix}:r`);
  const text = document.createElementNS(namespace, `${prefix}:t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.insertBefore(run, properties.nextSibling);
}

function insertParagraphProperties(
  document: Document,
  paragraph: Element,
): Element {
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix ?? 'w';
  const properties = document.createElementNS(namespace, `${prefix}:pPr`);
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function wordDirectChildren(parent: ParentNode, name: string): Element[] {
  return directChildren(parent, name).filter(isWordElement);
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
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
