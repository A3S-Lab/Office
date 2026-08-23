import {
  normalizeDocumentFontName,
  normalizeDocumentScriptFontHint,
  normalizeDocumentScriptFonts,
  normalizeDocumentThemeFont,
  type WorkDocumentScriptFontFace,
  type WorkDocumentScriptFonts,
} from './work-document-script-fonts';
import { type DocxThemeResolver, docxThemeFont } from './work-docx-theme';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { directChildren } from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

export type InspectedDocxRunFonts =
  | { status: 'absent' }
  | { status: 'invalid' }
  | { status: 'valid'; value: WorkDocumentScriptFonts | null };

const FONT_ATTRIBUTES = new Set([
  'ascii',
  'hAnsi',
  'eastAsia',
  'cs',
  'asciiTheme',
  'hAnsiTheme',
  'eastAsiaTheme',
  'cstheme',
  'hint',
]);
const EMPTY_THEME: DocxThemeResolver = {
  colors: new Map(),
  fonts: new Map(),
};

export function inspectDocxRunFonts(
  properties: Element | null | undefined,
  theme: DocxThemeResolver = EMPTY_THEME,
): InspectedDocxRunFonts {
  if (!properties) return { status: 'absent' };
  const candidates = directChildren(properties, 'rFonts');
  if (!candidates.length) return { status: 'absent' };
  if (candidates.length !== 1) return { status: 'invalid' };
  const value = parseDocxRunFontsElement(candidates[0], theme);
  return value === undefined
    ? { status: 'invalid' }
    : { status: 'valid', value };
}

export function resolveDocxRunFonts(
  propertySources: readonly Element[],
  theme: DocxThemeResolver,
): WorkDocumentScriptFonts | null {
  const resolved: WorkDocumentScriptFonts = {};
  let found = false;
  for (const properties of propertySources) {
    const inspection = inspectDocxRunFonts(properties, theme);
    if (inspection.status !== 'valid' || !inspection.value) continue;
    found = true;
    const next = inspection.value;
    for (const slot of [
      'ascii',
      'highAnsi',
      'eastAsia',
      'complexScript',
    ] as const) {
      if (next[slot]) resolved[slot] = next[slot];
    }
    if (next.hint) resolved.hint = next.hint;
  }
  return found ? normalizeDocumentScriptFonts(resolved) : null;
}

export function parseDocxRunFontsElement(
  element: Element | undefined,
  theme: DocxThemeResolver = EMPTY_THEME,
): WorkDocumentScriptFonts | null | undefined {
  if (
    !element ||
    element.localName !== 'rFonts' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') ||
    directChildren(element).length ||
    Array.from(element.childNodes).some(
      (node) => node.nodeType === Node.TEXT_NODE && node.textContent?.trim(),
    )
  ) {
    return undefined;
  }
  const attributes = new Map<string, string>();
  for (const item of Array.from(element.attributes)) {
    if (
      item.namespaceURI === XMLNS_NAMESPACE ||
      item.name === 'xmlns' ||
      item.name.startsWith('xmlns:')
    ) {
      continue;
    }
    const name = xmlAttributeLocalName(item);
    if (
      !FONT_ATTRIBUTES.has(name) ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ) ||
      attributes.has(name)
    ) {
      return undefined;
    }
    attributes.set(name, item.value);
  }

  const ascii = fontFace(
    attributes.get('ascii'),
    attributes.get('asciiTheme'),
    theme,
  );
  const highAnsi = fontFace(
    attributes.get('hAnsi'),
    attributes.get('hAnsiTheme'),
    theme,
  );
  const eastAsia = fontFace(
    attributes.get('eastAsia'),
    attributes.get('eastAsiaTheme'),
    theme,
  );
  const complexScript = fontFace(
    attributes.get('cs'),
    attributes.get('cstheme'),
    theme,
  );
  const hintValue = attributes.get('hint');
  const hint =
    hintValue === undefined
      ? undefined
      : normalizeDocumentScriptFontHint(hintValue);
  if (
    ascii === undefined ||
    highAnsi === undefined ||
    eastAsia === undefined ||
    complexScript === undefined ||
    hint === null
  ) {
    return undefined;
  }
  return normalizeDocumentScriptFonts({
    ...(ascii ? { ascii } : {}),
    ...(highAnsi ? { highAnsi } : {}),
    ...(eastAsia ? { eastAsia } : {}),
    ...(complexScript ? { complexScript } : {}),
    ...(hint ? { hint } : {}),
  });
}

function fontFace(
  directSource: string | undefined,
  themeSource: string | undefined,
  theme: DocxThemeResolver,
): WorkDocumentScriptFontFace | null | undefined {
  const name =
    directSource === undefined
      ? undefined
      : normalizeDocumentFontName(directSource);
  const themeReference =
    themeSource === undefined
      ? undefined
      : normalizeDocumentThemeFont(themeSource);
  if ((directSource !== undefined && !name) || themeReference === null) {
    return undefined;
  }
  if (!name && !themeReference) return null;
  const themedFamily = themeReference
    ? docxThemeFont(theme, themeReference)
    : undefined;
  const resolved = themedFamily ?? name;
  return {
    ...(name ? { name } : {}),
    ...(themeReference ? { theme: themeReference } : {}),
    ...(resolved ? { resolved } : {}),
  };
}
