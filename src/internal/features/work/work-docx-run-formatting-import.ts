import {
  type DocxParagraphStyleSource,
  docxRunPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import {
  type DocxThemeSource,
  docxThemeColor,
  docxThemeFont,
  resolveDocxThemeResolver,
  type DocxThemeResolver,
} from './work-docx-theme';
import {
  type DocxTableStyleSource,
  docxTableRunPropertySources,
  resolveDocxTableStyleResolver,
} from './work-docx-table-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';
import { documentWordLineHeightFactor } from './work-document-word-line-metrics';

export interface ImportedDocxRunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  wordLineHeightFactor?: number;
  wordSnapToGrid?: boolean;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
}

export interface ImportedDocxRunFormattingMarker {
  startMarker: string;
  endMarker: string;
  formatting: ImportedDocxRunFormatting;
}

export interface ImportedDocxRunFormattingMarkers {
  runs: ImportedDocxRunFormattingMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const RUN_FORMATTING_MARKER_PATTERN = /__A3S_WORK_RUN_(?:START|END)_\d+__/g;

type DocxFontSlot = 'ascii' | 'hAnsi' | 'eastAsia' | 'complex';

interface DocxRunFonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  complex?: string;
}

export function markDocxRunFormatting(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxRunFormattingMarkers {
  const runs: ImportedDocxRunFormattingMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const theme = resolveDocxThemeResolver(themeSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  for (const run of descendants(document, 'r')) {
    const runText = directRunText(run);
    if (!runText || runText.includes('__A3S_')) continue;
    const paragraph = closestAncestor(run, 'p');
    if (!paragraph) continue;
    const paragraphProperties = directChild(paragraph, 'pPr');
    const runProperties = directChild(run, 'rPr');
    const formatting = resolvedRunFormatting(
      docxRunPropertySources(
        paragraphProperties,
        runProperties,
        styles,
        docxTableRunPropertySources(run, tableStyles),
      ),
      theme,
      runText,
    );
    if (!Object.keys(formatting).length) continue;
    const index = runs.length + 1;
    const startMarker = `__A3S_WORK_RUN_START_${index}__`;
    const endMarker = `__A3S_WORK_RUN_END_${index}__`;
    insertRunMarkers(document, run, runProperties, startMarker, endMarker);
    runs.push({ startMarker, endMarker, formatting });
  }
  return { runs };
}

export function applyImportedDocxRunFormattingMarkers(
  document: Document,
  markers: ImportedDocxRunFormattingMarkers,
): void {
  const replacements = new Map<string, string>();
  for (const marker of markers.runs) {
    const markup = formattingMarkup(document, marker.formatting);
    replacements.set(marker.startMarker, markup.start);
    replacements.set(marker.endMarker, markup.end);
  }
  document.body.innerHTML = document.body.innerHTML.replace(
    RUN_FORMATTING_MARKER_PATTERN,
    (marker) => replacements.get(marker) ?? '',
  );
}

export function hasImportedDocxRunFormattingMarkers(
  markers: ImportedDocxRunFormattingMarkers,
): boolean {
  return markers.runs.length > 0;
}

function resolvedRunFormatting(
  propertySources: readonly Element[],
  theme: DocxThemeResolver,
  runText: string,
): ImportedDocxRunFormatting {
  const fonts: DocxRunFonts = {};
  let fontSize: number | undefined;
  let complexFontSize: number | undefined;
  let color: string | undefined;
  let backgroundColor: string | undefined;
  let bold: boolean | undefined;
  let complexBold: boolean | undefined;
  let italic: boolean | undefined;
  let complexItalic: boolean | undefined;
  let underline: boolean | undefined;
  let strike: boolean | undefined;
  let snapToGrid: boolean | undefined;
  let complexScriptFormatting: boolean | undefined;
  let rightToLeft: boolean | undefined;
  let fontHint: DocxFontSlot | undefined;

  for (const properties of propertySources) {
    bold = overriddenBoolean(bold, onOffProperty(properties, 'b'));
    complexBold = overriddenBoolean(
      complexBold,
      onOffProperty(properties, 'bCs'),
    );
    italic = overriddenBoolean(italic, onOffProperty(properties, 'i'));
    complexItalic = overriddenBoolean(
      complexItalic,
      onOffProperty(properties, 'iCs'),
    );
    underline = overriddenBoolean(underline, underlineProperty(properties));
    strike = overriddenBoolean(strike, onOffProperty(properties, 'strike'));
    strike = overriddenBoolean(strike, onOffProperty(properties, 'dstrike'));
    snapToGrid = overriddenBoolean(
      snapToGrid,
      onOffProperty(properties, 'snapToGrid'),
    );
    complexScriptFormatting = overriddenBoolean(
      complexScriptFormatting,
      onOffProperty(properties, 'cs'),
    );
    rightToLeft = overriddenBoolean(
      rightToLeft,
      onOffProperty(properties, 'rtl'),
    );

    const runFonts = directChild(properties, 'rFonts');
    if (runFonts) {
      fontHint = resolvedFontHint(runFonts) ?? fontHint;
      assignFont(
        fonts,
        'ascii',
        resolvedFont(runFonts, 'ascii', 'asciiTheme', theme),
      );
      assignFont(
        fonts,
        'hAnsi',
        resolvedFont(runFonts, 'hAnsi', 'hAnsiTheme', theme),
      );
      assignFont(
        fonts,
        'eastAsia',
        resolvedFont(runFonts, 'eastAsia', 'eastAsiaTheme', theme),
      );
      assignFont(
        fonts,
        'complex',
        resolvedFont(runFonts, 'cs', 'cstheme', theme),
      );
    }

    const size = numericAttribute(directChild(properties, 'sz'), 'val');
    if (size !== undefined && size > 0) fontSize = Math.min(512, size / 2);
    const complexSize = numericAttribute(
      directChild(properties, 'szCs'),
      'val',
    );
    if (complexSize !== undefined && complexSize > 0) {
      complexFontSize = Math.min(512, complexSize / 2);
    }

    const colorElement = directChild(properties, 'color');
    if (colorElement) {
      const value = wordAttribute(colorElement, 'val')?.trim();
      const themed = docxThemeColor(
        theme,
        wordAttribute(colorElement, 'themeColor'),
        wordAttribute(colorElement, 'themeTint'),
        wordAttribute(colorElement, 'themeShade'),
      );
      if (themed) color = `#${themed}`;
      else if (value?.toLowerCase() === 'auto') color = 'inherit';
      else if (value && /^[0-9a-f]{6}$/i.test(value))
        color = `#${value.toLowerCase()}`;
    }

    const highlight = directChild(properties, 'highlight');
    if (highlight) {
      const value = wordAttribute(highlight, 'val')?.trim().toLowerCase();
      if (value) backgroundColor = wordHighlightColor(value);
    } else {
      const shading = directChild(properties, 'shd');
      const fill = shading ? wordAttribute(shading, 'fill')?.trim() : undefined;
      const themed = shading
        ? docxThemeColor(
            theme,
            wordAttribute(shading, 'themeFill'),
            wordAttribute(shading, 'themeFillTint'),
            wordAttribute(shading, 'themeFillShade'),
          )
        : undefined;
      if (themed) backgroundColor = `#${themed}`;
      else if (fill?.toLowerCase() === 'auto') backgroundColor = 'transparent';
      else if (fill && /^[0-9a-f]{6}$/i.test(fill))
        backgroundColor = `#${fill.toLowerCase()}`;
    }
  }

  const fontSlot =
    complexScriptFormatting === true || rightToLeft === true
      ? 'complex'
      : docxFontSlotForText(runText, fontHint);
  const usesComplexFormatting = fontSlot === 'complex';
  const resolvedBold = usesComplexFormatting ? (complexBold ?? bold) : bold;
  const resolvedItalic = usesComplexFormatting
    ? (complexItalic ?? italic)
    : italic;
  const resolvedFontSize = usesComplexFormatting
    ? (complexFontSize ?? fontSize)
    : fontSize;
  let fontFamily = uniqueFonts(orderedFonts(fonts, fontSlot));
  if (!fontFamily) {
    fontFamily = uniqueFonts(
      orderedFonts(
        {
          ascii: docxThemeFont(theme, 'minorAscii'),
          hAnsi: docxThemeFont(theme, 'minorHAnsi'),
          eastAsia: docxThemeFont(theme, 'minorEastAsia'),
          complex: docxThemeFont(theme, 'minorBidi'),
        },
        fontSlot,
      ),
    );
  }
  const hasRunPropertySource = propertySources.length > 0;
  return {
    ...(resolvedBold !== undefined || hasRunPropertySource
      ? { bold: resolvedBold ?? false }
      : {}),
    ...(resolvedItalic !== undefined || hasRunPropertySource
      ? { italic: resolvedItalic ?? false }
      : {}),
    ...(underline !== undefined || hasRunPropertySource
      ? { underline: underline ?? false }
      : {}),
    ...(strike !== undefined || hasRunPropertySource
      ? { strike: strike ?? false }
      : {}),
    ...(fontFamily
      ? {
          fontFamily,
          wordLineHeightFactor: documentWordLineHeightFactor(fontFamily),
        }
      : {}),
    ...(resolvedFontSize !== undefined ? { fontSize: resolvedFontSize } : {}),
    ...(snapToGrid !== undefined ? { wordSnapToGrid: snapToGrid } : {}),
    ...(color ? { color } : {}),
    ...(backgroundColor ? { backgroundColor } : {}),
  };
}

function insertRunMarkers(
  document: Document,
  run: Element,
  properties: Element | undefined,
  startMarker: string,
  endMarker: string,
): void {
  const start = markerText(document, startMarker);
  const end = markerText(document, endMarker);
  run.insertBefore(start, properties?.nextSibling ?? run.firstChild);
  run.append(end);
}

function markerText(document: Document, marker: string): Element {
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  return text;
}

function formattingMarkup(
  document: Document,
  formatting: ImportedDocxRunFormatting,
): { start: string; end: string } {
  const span = document.createElement('span');
  if (formatting.bold === false) span.style.fontWeight = 'normal';
  if (formatting.italic === false) span.style.fontStyle = 'normal';
  if (
    (formatting.underline === false || formatting.strike === false) &&
    !formatting.underline &&
    !formatting.strike
  ) {
    span.style.textDecorationLine = 'none';
  }
  if (formatting.fontFamily) span.style.fontFamily = formatting.fontFamily;
  if (formatting.wordLineHeightFactor !== undefined) {
    const factor = formatLineHeightFactor(formatting.wordLineHeightFactor);
    span.dataset.officeWordLineHeightFactor = factor;
    span.style.setProperty('--work-document-word-line-height-factor', factor);
  }
  if (formatting.wordSnapToGrid !== undefined) {
    span.dataset.officeWordSnapToGrid = String(formatting.wordSnapToGrid);
  }
  if (formatting.fontSize !== undefined)
    span.style.fontSize = `${formatNumber(formatting.fontSize)}pt`;
  if (formatting.color) span.style.color = formatting.color;
  if (formatting.backgroundColor)
    span.style.backgroundColor = formatting.backgroundColor;
  const html = span.outerHTML;
  const tags = [
    ...(formatting.bold ? ['strong'] : []),
    ...(formatting.italic ? ['em'] : []),
    ...(formatting.underline ? ['u'] : []),
    ...(formatting.strike ? ['s'] : []),
  ];
  return {
    start: `${html.slice(0, html.indexOf('>') + 1)}${tags
      .map((tag) => `<${tag}>`)
      .join('')}`,
    end: `${[...tags]
      .reverse()
      .map((tag) => `</${tag}>`)
      .join('')}</span>`,
  };
}

function overriddenBoolean(
  current: boolean | undefined,
  next: boolean | undefined,
): boolean | undefined {
  return next === undefined ? current : next;
}

function onOffProperty(
  properties: Element,
  propertyName: string,
): boolean | undefined {
  const element = directChild(properties, propertyName);
  if (!element) return undefined;
  const value = wordAttribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function underlineProperty(properties: Element): boolean | undefined {
  const element = directChild(properties, 'u');
  if (!element) return undefined;
  const value = wordAttribute(element, 'val')?.trim().toLowerCase();
  return (
    value !== '0' && value !== 'false' && value !== 'off' && value !== 'none'
  );
}

function directRunText(run: Element): string {
  return Array.from(run.children)
    .filter((child) => child.localName === 't' || child.localName === 'delText')
    .map((child) => child.textContent ?? '')
    .join('');
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current = element.parentElement;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function assignFont(
  fonts: DocxRunFonts,
  key: DocxFontSlot,
  value: string | null,
): void {
  const normalized = value?.trim();
  if (normalized) fonts[key] = normalized;
}

function resolvedFontHint(element: Element): DocxFontSlot | undefined {
  const hint = wordAttribute(element, 'hint')?.trim().toLowerCase();
  if (hint === 'eastasia') return 'eastAsia';
  if (hint === 'cs') return 'complex';
  if (hint === 'default') return 'ascii';
  return undefined;
}

function orderedFonts(
  fonts: DocxRunFonts,
  preferred: DocxFontSlot,
): Array<string | undefined> {
  const order: Record<DocxFontSlot, readonly DocxFontSlot[]> = {
    ascii: ['ascii', 'hAnsi', 'eastAsia', 'complex'],
    hAnsi: ['hAnsi', 'ascii', 'eastAsia', 'complex'],
    eastAsia: ['eastAsia', 'hAnsi', 'ascii', 'complex'],
    complex: ['complex', 'hAnsi', 'ascii', 'eastAsia'],
  };
  return order[preferred].map((slot) => fonts[slot]);
}

function docxFontSlotForText(
  text: string,
  hint: DocxFontSlot | undefined,
): DocxFontSlot {
  for (const character of text) {
    const codePoint = character.codePointAt(0);
    if (codePoint === undefined || isNeutralCodePoint(codePoint)) continue;
    if (isComplexScriptCodePoint(codePoint)) return 'complex';
    if (isEastAsianCodePoint(codePoint)) return 'eastAsia';
    return codePoint <= 0x7f ? 'ascii' : 'hAnsi';
  }
  return hint ?? 'ascii';
}

function isNeutralCodePoint(codePoint: number): boolean {
  return (
    codePoint <= 0x20 ||
    (codePoint <= 0x7f &&
      !(
        (codePoint >= 0x41 && codePoint <= 0x5a) ||
        (codePoint >= 0x61 && codePoint <= 0x7a)
      )) ||
    (codePoint >= 0x300 && codePoint <= 0x36f) ||
    (codePoint >= 0x2000 && codePoint <= 0x206f)
  );
}

function isComplexScriptCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x0590 && codePoint <= 0x08ff) ||
    (codePoint >= 0xfb1d && codePoint <= 0xfdff) ||
    (codePoint >= 0xfe70 && codePoint <= 0xfeff) ||
    (codePoint >= 0x1ee00 && codePoint <= 0x1eeff)
  );
}

function isEastAsianCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x1100 && codePoint <= 0x11ff) ||
    (codePoint >= 0x2e80 && codePoint <= 0xa4cf) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7af) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xfe10 && codePoint <= 0xfe6f) ||
    (codePoint >= 0xff00 && codePoint <= 0xffef) ||
    (codePoint >= 0x20000 && codePoint <= 0x323af)
  );
}

function resolvedFont(
  element: Element,
  explicitName: string,
  themeName: string,
  theme: DocxThemeResolver,
): string | null {
  const explicit = wordAttribute(element, explicitName);
  if (explicit?.trim()) return explicit;
  return docxThemeFont(theme, wordAttribute(element, themeName)) ?? null;
}

function uniqueFonts(values: Array<string | undefined>): string | undefined {
  const fonts = [
    ...new Set(values.filter((value): value is string => !!value)),
  ];
  return fonts.length ? fonts.map(cssFontFamily).join(', ') : undefined;
}

function cssFontFamily(value: string): string {
  return /[\s"',]/.test(value)
    ? `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
    : value;
}

function numericAttribute(
  element: Element | undefined,
  name: string,
): number | undefined {
  if (!element) return undefined;
  const raw = wordAttribute(element, name);
  if (raw === null || !raw.trim()) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}

function wordHighlightColor(value: string): string {
  const colors: Record<string, string> = {
    black: '#000000',
    blue: '#0000ff',
    cyan: '#00ffff',
    darkblue: '#000080',
    darkcyan: '#008080',
    darkgray: '#808080',
    darkgreen: '#008000',
    darkmagenta: '#800080',
    darkred: '#800000',
    darkyellow: '#808000',
    green: '#00ff00',
    lightgray: '#c0c0c0',
    magenta: '#ff00ff',
    none: 'transparent',
    red: '#ff0000',
    white: '#ffffff',
    yellow: '#ffff00',
  };
  return colors[value] ?? 'transparent';
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function formatLineHeightFactor(value: number): string {
  return Number(value.toFixed(4)).toString();
}
