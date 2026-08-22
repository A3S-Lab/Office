import type { Cell } from '@fortune-sheet/core';
import {
  attribute,
  directChild,
  directChildren,
  firstDescendant,
} from './work-ooxml-package';
import {
  ensureXlsxStyleCollection,
  xlsxRgbColor,
} from './work-xlsx-cell-style-xml';
import {
  createXlsxColorResolver,
  resolveXlsxColor,
  type XlsxColorResolver,
} from './work-xlsx-colors';

export type XlsxSemanticColorOrigin =
  | {
      kind: 'automatic';
      baseColor: string;
      renderedColor: string;
      tint?: number;
    }
  | {
      kind: 'indexed';
      baseColor: string;
      index: number;
      renderedColor: string;
      tint?: number;
    }
  | {
      kind: 'theme';
      baseColor: string;
      index: number;
      renderedColor: string;
      tint?: number;
    };

export type XlsxBorderColorOrigins = Partial<
  Record<
    'bottom' | 'diagonal' | 'left' | 'right' | 'top',
    XlsxSemanticColorOrigin
  >
>;

export interface XlsxCellStyleOrigin {
  borderColors?: XlsxBorderColorOrigins;
  fillColor?: XlsxSemanticColorOrigin;
  fontColor?: XlsxSemanticColorOrigin;
}

export interface XlsxSemanticPalette {
  indexed: ReadonlyMap<number, string>;
  theme: ReadonlyMap<number, string>;
}

export interface PreparedXlsxSemanticPalette {
  palette: XlsxSemanticPalette;
  stylesChanged: boolean;
  themeChanged: boolean;
}

type CellWithXlsxStyleOrigin = Cell & {
  a3sXlsxStyleOrigin?: unknown;
};

const THEME_COLOR_NAMES = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

const MAX_INDEXED_COLOR = 255;

export function readXlsxSemanticColorOrigin(
  element: Element | undefined,
  colors: XlsxColorResolver,
): XlsxSemanticColorOrigin | undefined {
  if (!element || attribute(element, 'rgb')) return undefined;
  const renderedColor = resolveXlsxColor(element, colors);
  if (!renderedColor) return undefined;
  const tint = boundedTint(attribute(element, 'tint'));
  const theme = boundedInteger(
    attribute(element, 'theme'),
    THEME_COLOR_NAMES.length - 1,
  );
  if (theme !== null && colors.theme[theme]) {
    return {
      kind: 'theme',
      baseColor: `#${colors.theme[theme]}`,
      index: theme,
      renderedColor,
      ...(tint === null ? {} : { tint }),
    };
  }
  const indexed = boundedInteger(
    attribute(element, 'indexed'),
    MAX_INDEXED_COLOR,
  );
  if (indexed !== null && colors.indexed[indexed]) {
    return {
      kind: 'indexed',
      baseColor: `#${colors.indexed[indexed]}`,
      index: indexed,
      renderedColor,
      ...(tint === null ? {} : { tint }),
    };
  }
  if (xlsxBooleanAttribute(element, 'auto')) {
    return {
      kind: 'automatic',
      baseColor: '#000000',
      renderedColor,
      ...(tint === null ? {} : { tint }),
    };
  }
  return undefined;
}

export function withXlsxCellStyleOrigin(
  cell: Cell,
  origin: XlsxCellStyleOrigin | undefined,
): Cell {
  return origin && xlsxCellStyleOriginHasValues(origin)
    ? ({ ...cell, a3sXlsxStyleOrigin: origin } as CellWithXlsxStyleOrigin)
    : cell;
}

export function xlsxCellStyleOrigin(
  cell: Cell | null | undefined,
): XlsxCellStyleOrigin | undefined {
  const candidate = (cell as CellWithXlsxStyleOrigin | null | undefined)
    ?.a3sXlsxStyleOrigin;
  if (!isRecord(candidate)) return undefined;
  const fontColor = normalizeXlsxSemanticColorOrigin(candidate.fontColor);
  const fillColor = normalizeXlsxSemanticColorOrigin(candidate.fillColor);
  const borderColors = normalizedBorderColors(candidate.borderColors);
  const origin = {
    ...(fontColor ? { fontColor } : {}),
    ...(fillColor ? { fillColor } : {}),
    ...(borderColors ? { borderColors } : {}),
  };
  return xlsxCellStyleOriginHasValues(origin) ? origin : undefined;
}

export function xlsxSemanticColorMatchesValue(
  origin: XlsxSemanticColorOrigin,
  value: unknown,
): boolean {
  return normalizedColor(value) === origin.renderedColor.toLowerCase();
}

export function xlsxColorElementMatchesOrigin(
  element: Element | undefined,
  origin: XlsxSemanticColorOrigin,
): boolean {
  if (!element || attribute(element, 'rgb')) return false;
  if (origin.kind === 'theme') {
    if (attribute(element, 'theme') !== String(origin.index)) return false;
  } else if (origin.kind === 'indexed') {
    if (attribute(element, 'indexed') !== String(origin.index)) return false;
  } else if (!xlsxBooleanAttribute(element, 'auto')) {
    return false;
  }
  const tint = boundedTint(attribute(element, 'tint'));
  return tint === (origin.tint ?? null);
}

export function applyXlsxSemanticColorOrigin(
  element: Element,
  origin: XlsxSemanticColorOrigin,
): void {
  for (const name of ['rgb', 'theme', 'indexed', 'auto', 'tint']) {
    element.removeAttribute(name);
  }
  if (origin.kind === 'theme')
    element.setAttribute('theme', String(origin.index));
  else if (origin.kind === 'indexed')
    element.setAttribute('indexed', String(origin.index));
  else element.setAttribute('auto', '1');
  if (origin.tint !== undefined)
    element.setAttribute('tint', String(origin.tint));
}

export function xlsxSemanticColorOriginKey(
  origin: XlsxSemanticColorOrigin | undefined,
): string {
  return origin
    ? `${origin.kind}:${'index' in origin ? origin.index : ''}:${origin.baseColor}:${origin.renderedColor}:${origin.tint ?? ''}`
    : '';
}

export function xlsxSemanticColorOriginSupported(
  origin: XlsxSemanticColorOrigin,
  palette: XlsxSemanticPalette | undefined,
): boolean {
  if (origin.kind === 'automatic') return true;
  return (
    palette?.[origin.kind].get(origin.index)?.toLowerCase() ===
    origin.baseColor.toLowerCase()
  );
}

export function prepareXlsxSemanticPalette(
  styles: Document,
  theme: Document | null,
  colors: Iterable<XlsxSemanticColorOrigin>,
): PreparedXlsxSemanticPalette {
  const candidates = semanticPaletteCandidates(colors);
  const supportedTheme = new Map<number, string>();
  let themeChanged = false;
  const scheme = theme ? firstDescendant(theme, 'clrScheme') : undefined;
  for (const [index, color] of candidates.theme) {
    const name = THEME_COLOR_NAMES[index];
    const entry = name && scheme ? directChild(scheme, name) : undefined;
    if (!entry) continue;
    themeChanged = writeThemeColor(entry, color) || themeChanged;
    supportedTheme.set(index, color);
  }

  const supportedIndexed = new Map<number, string>();
  let stylesChanged = false;
  if (candidates.indexed.size) {
    const colors = ensureXlsxStyleCollection(styles, 'colors', ['extLst']);
    let indexedColors = directChild(colors, 'indexedColors');
    if (!indexedColors) {
      indexedColors = styles.createElementNS(
        styles.documentElement.namespaceURI,
        'indexedColors',
      );
      colors.prepend(indexedColors);
      stylesChanged = true;
    }
    const resolver = createXlsxColorResolver(styles, theme);
    const highestIndex = Math.max(...candidates.indexed.keys());
    while (directChildren(indexedColors, 'rgbColor').length <= highestIndex) {
      const index = directChildren(indexedColors, 'rgbColor').length;
      const element = styles.createElementNS(
        styles.documentElement.namespaceURI,
        'rgbColor',
      );
      element.setAttribute(
        'rgb',
        `FF${(resolver.indexed[index] ?? '000000').toUpperCase()}`,
      );
      indexedColors.append(element);
      stylesChanged = true;
    }
    const entries = directChildren(indexedColors, 'rgbColor');
    for (const [index, color] of candidates.indexed) {
      const expected = `FF${color.slice(1).toUpperCase()}`;
      const entry = entries[index];
      if (!entry) continue;
      if (attribute(entry, 'rgb')?.toUpperCase() !== expected) {
        entry.setAttribute('rgb', expected);
        stylesChanged = true;
      }
      supportedIndexed.set(index, color);
    }
  }

  return {
    palette: { indexed: supportedIndexed, theme: supportedTheme },
    stylesChanged,
    themeChanged,
  };
}

function semanticPaletteCandidates(
  colors: Iterable<XlsxSemanticColorOrigin>,
): XlsxSemanticPalette {
  const theme = new Map<number, string>();
  const indexed = new Map<number, string>();
  const themeConflicts = new Set<number>();
  const indexedConflicts = new Set<number>();
  for (const color of colors) {
    if (color.kind === 'automatic') continue;
    const target = color.kind === 'theme' ? theme : indexed;
    const conflicts =
      color.kind === 'theme' ? themeConflicts : indexedConflicts;
    const current = target.get(color.index);
    if (current && current.toLowerCase() !== color.baseColor.toLowerCase()) {
      conflicts.add(color.index);
    } else if (!current) {
      target.set(color.index, color.baseColor);
    }
  }
  for (const index of themeConflicts) theme.delete(index);
  for (const index of indexedConflicts) indexed.delete(index);
  return { indexed, theme };
}

export function xlsxCellStyleOriginSemanticColors(
  origin: XlsxCellStyleOrigin,
): XlsxSemanticColorOrigin[] {
  return [
    origin.fontColor,
    origin.fillColor,
    ...Object.values(origin.borderColors ?? {}),
  ].filter((value): value is XlsxSemanticColorOrigin => Boolean(value));
}

function writeThemeColor(entry: Element, color: string): boolean {
  const current = directChildren(entry)[0];
  if (
    current?.localName === 'srgbClr' &&
    attribute(current, 'val')?.toLowerCase() === color.slice(1).toLowerCase()
  ) {
    return false;
  }
  for (const child of directChildren(entry)) child.remove();
  const qualifiedName = entry.prefix ? `${entry.prefix}:srgbClr` : 'srgbClr';
  const replacement = entry.ownerDocument.createElementNS(
    entry.namespaceURI,
    qualifiedName,
  );
  replacement.setAttribute('val', color.slice(1).toUpperCase());
  entry.append(replacement);
  return true;
}

function normalizedBorderColors(
  value: unknown,
): XlsxBorderColorOrigins | undefined {
  if (!isRecord(value)) return undefined;
  const result: XlsxBorderColorOrigins = {};
  for (const side of ['bottom', 'diagonal', 'left', 'right', 'top'] as const) {
    const color = normalizeXlsxSemanticColorOrigin(value[side]);
    if (color) result[side] = color;
  }
  return Object.keys(result).length ? result : undefined;
}

export function normalizeXlsxSemanticColorOrigin(
  value: unknown,
): XlsxSemanticColorOrigin | undefined {
  if (!isRecord(value)) return undefined;
  const baseColor = normalizedColor(value.baseColor);
  const renderedColor = normalizedColor(value.renderedColor);
  const tint = normalizedTint(value.tint);
  if (!baseColor || !renderedColor || tint === undefined) return undefined;
  if (value.kind === 'automatic') {
    return {
      kind: 'automatic',
      baseColor,
      renderedColor,
      ...(tint === null ? {} : { tint }),
    };
  }
  const maximum =
    value.kind === 'theme' ? THEME_COLOR_NAMES.length - 1 : MAX_INDEXED_COLOR;
  const index = boundedNumber(value.index, maximum);
  if ((value.kind !== 'theme' && value.kind !== 'indexed') || index === null)
    return undefined;
  return {
    kind: value.kind,
    baseColor,
    index,
    renderedColor,
    ...(tint === null ? {} : { tint }),
  };
}

function xlsxCellStyleOriginHasValues(origin: XlsxCellStyleOrigin): boolean {
  return Boolean(
    origin.fontColor ||
      origin.fillColor ||
      Object.keys(origin.borderColors ?? {}).length,
  );
}

function normalizedColor(value: unknown): string | null {
  const rgb = xlsxRgbColor(value);
  return rgb ? `#${rgb.slice(-6).toLowerCase()}` : null;
}

function normalizedTint(value: unknown): number | null | undefined {
  if (value === undefined) return null;
  return typeof value === 'number' &&
    Number.isFinite(value) &&
    value >= -1 &&
    value <= 1
    ? value
    : undefined;
}

function boundedTint(value: string | null): number | null {
  if (value === null) return null;
  const tint = Number(value);
  return Number.isFinite(tint) && tint >= -1 && tint <= 1 ? tint : null;
}

function boundedInteger(value: string | null, maximum: number): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  return boundedNumber(Number(value), maximum);
}

function boundedNumber(value: unknown, maximum: number): number | null {
  return typeof value === 'number' &&
    Number.isSafeInteger(value) &&
    value >= 0 &&
    value <= maximum
    ? value
    : null;
}

function xlsxBooleanAttribute(element: Element, name: string): boolean {
  const value = attribute(element, name)?.trim().toLowerCase();
  return value === '1' || value === 'true' || value === 'on';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
