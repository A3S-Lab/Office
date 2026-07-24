import {
  attribute,
  directChild,
  directChildren,
  firstDescendant,
} from './work-ooxml-package';

export interface DocxThemeResolver {
  colors: ReadonlyMap<string, string>;
  fonts: ReadonlyMap<string, string>;
}

export type DocxThemeSource = Document | DocxThemeResolver | null | undefined;

export function createDocxThemeResolver(
  themeDocument?: Document | null,
): DocxThemeResolver {
  if (!themeDocument) return { colors: new Map(), fonts: new Map() };
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
  return { colors, fonts };
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
  const color = key ? resolver.colors.get(key.trim().toLowerCase()) : undefined;
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
