import {
  DOCUMENT_PARAGRAPH_SHADING_PATTERNS,
  type DocumentParagraphShading,
  type DocumentParagraphShadingColor,
  type DocumentParagraphShadingPattern,
  documentParagraphShadingDomAttributes,
} from './work-document-paragraph-shading';
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

export interface ImportedDocxParagraphShadingMarker {
  marker: string;
  shading: DocumentParagraphShading;
}

export interface ImportedDocxParagraphShadingMarkers {
  paragraphs: ImportedDocxParagraphShadingMarker[];
  invalidCount: number;
  spoofedCount: number;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const PARAGRAPH_SHADING_MARKER_PATTERN = /__A3S_WORK_PARAGRAPH_SHADING_\d+__/g;
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
const SHADING_ATTRIBUTES = new Set([
  'val',
  'color',
  'themeColor',
  'themeTint',
  'themeShade',
  'fill',
  'themeFill',
  'themeFillTint',
  'themeFillShade',
]);

export function markDocxParagraphShading(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxParagraphShadingMarkers {
  const paragraphs: ImportedDocxParagraphShadingMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  const theme = resolveDocxThemeResolver(themeSource);
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const paragraph of descendants(document, 'p')) {
    if (!isWordElement(paragraph)) continue;
    const properties = wordDirectChildren(paragraph, 'pPr')[0];
    const sources = docxParagraphPropertySources(
      properties,
      styles,
      docxTableParagraphPropertySources(paragraph, tableStyles),
    );
    const resolved = resolveParagraphShading(sources, theme);
    invalidCount += resolved.invalidCount;
    spoofedCount += resolved.spoofedCount;
    if (!resolved.shading) continue;
    const marker = `__A3S_WORK_PARAGRAPH_SHADING_${paragraphs.length + 1}__`;
    const writableProperties =
      properties ?? insertParagraphProperties(document, paragraph);
    insertParagraphMarker(document, paragraph, writableProperties, marker);
    paragraphs.push({ marker, shading: resolved.shading });
  }
  return { paragraphs, invalidCount, spoofedCount };
}

export function applyImportedDocxParagraphShadingMarkers(
  document: Document,
  markers: ImportedDocxParagraphShadingMarkers,
): void {
  const shadingByMarker = new Map(
    markers.paragraphs.map(({ marker, shading }) => [marker, shading]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_SHADING_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_SHADING_MARKER_PATTERN,
      (marker) => {
        const shading = shadingByMarker.get(marker);
        const block = shading
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && shading) applyParagraphShading(block, shading);
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphShadingMarkers(
  markers: ImportedDocxParagraphShadingMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

export function parseDocxParagraphShadingElement(
  element: Element,
  themeSource?: DocxThemeSource,
): DocumentParagraphShading | null {
  const attributes = wordLeafAttributes(element, SHADING_ATTRIBUTES);
  const pattern = attributes?.get('val')?.trim();
  if (
    !attributes ||
    !DOCUMENT_PARAGRAPH_SHADING_PATTERNS.has(
      pattern as DocumentParagraphShadingPattern,
    )
  ) {
    return null;
  }
  const theme = resolveDocxThemeResolver(themeSource);
  const color = parseShadingColor(attributes, theme, {
    value: 'color',
    theme: 'themeColor',
    tint: 'themeTint',
    shade: 'themeShade',
  });
  const fill = parseShadingColor(attributes, theme, {
    value: 'fill',
    theme: 'themeFill',
    tint: 'themeFillTint',
    shade: 'themeFillShade',
  });
  if (color === null || fill === null) return null;
  return {
    pattern: pattern as DocumentParagraphShadingPattern,
    ...(color ? { color } : {}),
    ...(fill ? { fill } : {}),
  };
}

export function parseDirectDocxParagraphShading(
  properties: Element,
  themeSource?: DocxThemeSource,
): DocumentParagraphShading | null | undefined {
  const candidates = directChildren(properties, 'shd').filter(isWordElement);
  if (!candidates.length) return undefined;
  if (candidates.length !== 1) return null;
  const candidate = candidates[0];
  return candidate
    ? parseDocxParagraphShadingElement(candidate, themeSource)
    : undefined;
}

function resolveParagraphShading(
  propertySources: readonly Element[],
  theme: DocxThemeResolver,
): {
  shading: DocumentParagraphShading | null;
  invalidCount: number;
  spoofedCount: number;
} {
  let shading: DocumentParagraphShading | null = null;
  let specified = false;
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const properties of propertySources) {
    const named = directChildren(properties, 'shd');
    spoofedCount += named.filter((child) => !isWordElement(child)).length;
    const candidates = named.filter(isWordElement);
    if (!candidates.length) continue;
    specified = true;
    if (candidates.length !== 1) {
      invalidCount += candidates.length;
      shading = { pattern: 'nil' };
      continue;
    }
    const candidate = candidates[0];
    const parsed = candidate
      ? parseDocxParagraphShadingElement(candidate, theme)
      : null;
    if (!parsed) {
      invalidCount += 1;
      shading = { pattern: 'nil' };
      continue;
    }
    shading = parsed;
  }
  return {
    shading: specified ? shading : null,
    invalidCount,
    spoofedCount,
  };
}

function parseShadingColor(
  attributes: ReadonlyMap<string, string>,
  theme: DocxThemeResolver,
  names: {
    value: 'color' | 'fill';
    theme: 'themeColor' | 'themeFill';
    tint: 'themeTint' | 'themeFillTint';
    shade: 'themeShade' | 'themeFillShade';
  },
): DocumentParagraphShadingColor | null | undefined {
  const rawValue = attributes.get(names.value)?.trim();
  const rawTheme = attributes.get(names.theme)?.trim();
  const rawTint = attributes.get(names.tint)?.trim();
  const rawShade = attributes.get(names.shade)?.trim();
  if (
    rawValue === undefined &&
    rawTheme === undefined &&
    rawTint === undefined &&
    rawShade === undefined
  ) {
    return undefined;
  }
  const value =
    rawValue === 'auto'
      ? ('auto' as const)
      : rawValue && /^[0-9a-f]{6}$/i.test(rawValue)
        ? (`#${rawValue.toLowerCase()}` as const)
        : null;
  const themeName =
    rawTheme && WORD_THEME_COLORS.has(rawTheme) ? rawTheme : null;
  const tint = byteHex(rawTint);
  const shade = byteHex(rawShade);
  if (
    (rawValue !== undefined && !value) ||
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
    const fallback = value && value !== 'auto' ? value : null;
    const semanticValue = resolved ? (`#${resolved}` as const) : fallback;
    if (!semanticValue) return null;
    return {
      value: semanticValue,
      theme: themeReference(themeName, semanticValue, tint, shade),
    };
  }
  if (!value) return null;
  return {
    value,
    ...(themeName === 'none' && value !== 'auto'
      ? {
          theme: themeReference(themeName, value, tint, shade),
        }
      : {}),
  };
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

function byteHex(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  return /^[0-9a-f]{2}$/i.test(value) ? value.toUpperCase() : undefined;
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
    if (
      attribute.namespaceURI === XMLNS_NAMESPACE ||
      attribute.name === 'xmlns' ||
      attribute.name.startsWith('xmlns:')
    ) {
      continue;
    }
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

function applyParagraphShading(
  element: HTMLElement,
  shading: DocumentParagraphShading,
): void {
  for (const [name, value] of Object.entries(
    documentParagraphShadingDomAttributes(shading),
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
