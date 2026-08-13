import {
  attribute,
  directChild,
  directChildren,
  firstDescendant,
} from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

export interface DocxThemeResolver {
  colors: ReadonlyMap<string, string>;
  colorMapping?: ReadonlyMap<string, string>;
  fonts: ReadonlyMap<string, string>;
}

export type DocxThemeSource = Document | DocxThemeResolver | null | undefined;

const WORD_THEME_COLOR_KEYS = new Map([
  ['dark1', 'dk1'],
  ['light1', 'lt1'],
  ['dark2', 'dk2'],
  ['light2', 'lt2'],
  ['background1', 'lt1'],
  ['text1', 'dk1'],
  ['background2', 'lt2'],
  ['text2', 'dk2'],
  ['hyperlink', 'hlink'],
  ['followedhyperlink', 'folhlink'],
]);
export const DOCX_COLOR_SCHEME_MAPPING_ATTRIBUTES = new Map([
  ['background1', 'bg1'],
  ['text1', 't1'],
  ['background2', 'bg2'],
  ['text2', 't2'],
  ['accent1', 'accent1'],
  ['accent2', 'accent2'],
  ['accent3', 'accent3'],
  ['accent4', 'accent4'],
  ['accent5', 'accent5'],
  ['accent6', 'accent6'],
  ['hyperlink', 'hyperlink'],
  ['followedhyperlink', 'followedHyperlink'],
]);
const WORD_COLOR_SCHEME_INDEXES = new Set([
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
]);
const WORDPROCESSING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);
const MAX_COLOR_MAPPING_DEPTH = 16;

export function createDocxThemeResolver(
  themeDocument?: Document | null,
  settingsDocument?: Document | null,
): DocxThemeResolver {
  const colorMapping = docxColorSchemeMapping(settingsDocument);
  if (!themeDocument)
    return { colors: new Map(), colorMapping, fonts: new Map() };
  const fonts = new Map<string, string>();
  const fontScheme = firstDescendant(themeDocument, 'fontScheme');
  addThemeFontFamily(
    fonts,
    'major',
    fontScheme ? directChild(fontScheme, 'majorFont') : undefined,
  );
  addThemeFontFamily(
    fonts,
    'minor',
    fontScheme ? directChild(fontScheme, 'minorFont') : undefined,
  );

  const colors = new Map<string, string>();
  const colorScheme = firstDescendant(themeDocument, 'clrScheme');
  if (colorScheme) {
    for (const entry of directChildren(colorScheme)) {
      const color = themeColorValue(directChildren(entry)[0]);
      if (color) colors.set(entry.localName.toLowerCase(), color);
    }
  }
  return { colors, colorMapping, fonts };
}

export function resolveDocxThemeResolver(
  source: DocxThemeSource,
): DocxThemeResolver {
  return isDocxThemeResolver(source) ? source : createDocxThemeResolver(source);
}

export function docxThemeFont(
  resolver: DocxThemeResolver,
  key: string | null,
): string | undefined {
  return key ? resolver.fonts.get(key.trim().toLowerCase()) : undefined;
}

export function docxThemeColor(
  resolver: DocxThemeResolver,
  key: string | null,
  tint?: string | null,
  shade?: string | null,
): string | undefined {
  const sourceKey = key?.trim().toLowerCase();
  const mappedKey = sourceKey
    ? mappedWordThemeColorKey(sourceKey, resolver.colorMapping)
    : undefined;
  const colorKey = mappedKey
    ? (WORD_THEME_COLOR_KEYS.get(mappedKey) ?? mappedKey)
    : undefined;
  const color = colorKey ? resolver.colors.get(colorKey) : undefined;
  if (!color) return undefined;
  const channels = [
    Number.parseInt(color.slice(0, 2), 16),
    Number.parseInt(color.slice(2, 4), 16),
    Number.parseInt(color.slice(4, 6), 16),
  ];
  const shadeFactor = hexFactor(shade);
  const tintFactor = hexFactor(tint);
  const transformed = channels.map((channel) => {
    const shaded = shadeFactor === undefined ? channel : channel * shadeFactor;
    const tinted =
      tintFactor === undefined ? shaded : shaded + (255 - shaded) * tintFactor;
    return Math.max(0, Math.min(255, Math.round(tinted)));
  });
  return transformed
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('');
}

function docxColorSchemeMapping(
  settingsDocument: Document | null | undefined,
): ReadonlyMap<string, string> {
  const root = settingsDocument?.documentElement;
  if (
    !root ||
    root.localName !== 'settings' ||
    !WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
  ) {
    return new Map();
  }
  const candidates = directChildren(root, 'clrSchemeMapping').filter(
    (element) => element.namespaceURI === root.namespaceURI,
  );
  if (candidates.length !== 1) return new Map();
  return parseDocxColorSchemeMappingElement(candidates[0]) ?? new Map();
}

export function parseDocxColorSchemeMappingElement(
  element: Element | undefined,
): ReadonlyMap<string, string> | null {
  const namespace = element?.namespaceURI ?? '';
  if (
    !element ||
    element.localName !== 'clrSchemeMapping' ||
    !WORDPROCESSING_NAMESPACES.has(namespace) ||
    directChildren(element).length ||
    Array.from(element.childNodes).some(
      (node) => node.nodeType === Node.TEXT_NODE && node.textContent?.trim(),
    )
  ) {
    return null;
  }
  const semanticsByAttribute = new Map(
    Array.from(DOCX_COLOR_SCHEME_MAPPING_ATTRIBUTES).map(
      ([semantic, attributeName]) => [attributeName, semantic],
    ),
  );
  const mapping = new Map<string, string>();
  for (const item of Array.from(element.attributes)) {
    if (
      item.namespaceURI === XMLNS_NAMESPACE ||
      item.name === 'xmlns' ||
      item.name.startsWith('xmlns:')
    ) {
      continue;
    }
    const attributeName = xmlAttributeLocalName(item);
    const semantic = semanticsByAttribute.get(attributeName);
    const value = item.value.trim();
    if (
      xmlAttributeNamespace(element, item) !== namespace ||
      !semantic ||
      mapping.has(semantic) ||
      !WORD_COLOR_SCHEME_INDEXES.has(value)
    ) {
      return null;
    }
    mapping.set(semantic, value);
  }
  return mapping;
}

function mappedWordThemeColorKey(
  source: string,
  mapping: ReadonlyMap<string, string> | undefined,
): string | undefined {
  let current = source;
  const visited = new Set<string>();
  for (let depth = 0; depth < MAX_COLOR_MAPPING_DEPTH; depth += 1) {
    if (visited.has(current)) return undefined;
    visited.add(current);
    const next = mapping?.get(current)?.toLowerCase();
    if (!next || next === current) return current;
    current = next;
  }
  return undefined;
}

function addThemeFontFamily(
  fonts: Map<string, string>,
  prefix: 'major' | 'minor',
  family: Element | undefined,
): void {
  if (!family) return;
  const latin = typeface(directChild(family, 'latin'));
  const eastAsia =
    typeface(directChild(family, 'ea')) ||
    scriptTypeface(family, ['Hans', 'Hant', 'Jpan', 'Hang']) ||
    latin;
  const complex = typeface(directChild(family, 'cs')) || latin;
  if (latin) {
    fonts.set(`${prefix}ascii`, latin);
    fonts.set(`${prefix}hansi`, latin);
  }
  if (eastAsia) fonts.set(`${prefix}eastasia`, eastAsia);
  if (complex) fonts.set(`${prefix}bidi`, complex);
}

function scriptTypeface(
  family: Element,
  scripts: readonly string[],
): string | undefined {
  const fonts = directChildren(family, 'font');
  for (const script of scripts) {
    const font = fonts.find(
      (candidate) => drawingAttribute(candidate, 'script') === script,
    );
    const value = font ? typeface(font) : undefined;
    if (value) return value;
  }
  return undefined;
}

function typeface(element: Element | undefined): string | undefined {
  const value = element
    ? drawingAttribute(element, 'typeface')?.trim()
    : undefined;
  return value || undefined;
}

function themeColorValue(element: Element | undefined): string | undefined {
  if (!element) return undefined;
  const value =
    element.localName === 'sysClr'
      ? drawingAttribute(element, 'lastClr')
      : drawingAttribute(element, 'val');
  return value && /^[0-9a-f]{6}$/i.test(value)
    ? value.toLowerCase()
    : undefined;
}

function drawingAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `a:${name}`);
}

function hexFactor(value: string | null | undefined): number | undefined {
  if (!value || !/^[0-9a-f]{2}$/i.test(value)) return undefined;
  return Number.parseInt(value, 16) / 255;
}

function isDocxThemeResolver(
  source: DocxThemeSource,
): source is DocxThemeResolver {
  return Boolean(
    source &&
      'colors' in source &&
      source.colors instanceof Map &&
      'fonts' in source &&
      source.fonts instanceof Map,
  );
}
