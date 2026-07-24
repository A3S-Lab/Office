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
import { attribute, descendants, directChild } from './work-ooxml-package';

export interface ImportedDocxRunFormatting {
  fontFamily?: string;
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

export function markDocxRunFormatting(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
): ImportedDocxRunFormattingMarkers {
  const runs: ImportedDocxRunFormattingMarker[] = [];
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const theme = resolveDocxThemeResolver(themeSource);
  for (const run of descendants(document, 'r')) {
    const runText = directRunText(run);
    if (!runText || runText.includes('__A3S_')) continue;
    const paragraph = closestAncestor(run, 'p');
    if (!paragraph) continue;
    const paragraphProperties = directChild(paragraph, 'pPr');
    const runProperties = directChild(run, 'rPr');
    const formatting = resolvedRunFormatting(
      docxRunPropertySources(paragraphProperties, runProperties, styles),
      theme,
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
    replacements.set(
      marker.startMarker,
      formattingSpanStart(document, marker.formatting),
    );
    replacements.set(marker.endMarker, '</span>');
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
): ImportedDocxRunFormatting {
  const fonts: {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    complex?: string;
  } = {};
  let fontSize: number | undefined;
  let color: string | undefined;
  let backgroundColor: string | undefined;

  for (const properties of propertySources) {
    const runFonts = directChild(properties, 'rFonts');
    if (runFonts) {
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

  const fontFamily = uniqueFonts([
    fonts.eastAsia,
    fonts.hAnsi,
    fonts.ascii,
    fonts.complex,
  ]);
  return {
    ...(fontFamily ? { fontFamily } : {}),
    ...(fontSize !== undefined ? { fontSize } : {}),
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

function formattingSpanStart(
  document: Document,
  formatting: ImportedDocxRunFormatting,
): string {
  const span = document.createElement('span');
  if (formatting.fontFamily) span.style.fontFamily = formatting.fontFamily;
  if (formatting.fontSize !== undefined)
    span.style.fontSize = `${formatNumber(formatting.fontSize)}pt`;
  if (formatting.color) span.style.color = formatting.color;
  if (formatting.backgroundColor)
    span.style.backgroundColor = formatting.backgroundColor;
  const html = span.outerHTML;
  return html.slice(0, html.indexOf('>') + 1);
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
  fonts: {
    ascii?: string;
    hAnsi?: string;
    eastAsia?: string;
    complex?: string;
  },
  key: 'ascii' | 'hAnsi' | 'eastAsia' | 'complex',
  value: string | null,
): void {
  const normalized = value?.trim();
  if (normalized) fonts[key] = normalized;
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
