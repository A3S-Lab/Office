import {
  normalizeDocumentUnderlineColor,
  normalizeDocumentUnderlineStyle,
  type WorkDocumentUnderlineFormatting,
} from './work-document-underline';
import { docxThemeColor, type DocxThemeResolver } from './work-docx-theme';
import type { DocxThemeColorReference } from './work-docx-theme-reference';
import { attribute } from './work-ooxml-package';

export function importedDocxUnderline(
  element: Element,
  theme?: DocxThemeResolver,
): WorkDocumentUnderlineFormatting {
  const rawStyle = attribute(element, 'val');
  const style =
    rawStyle === null
      ? 'single'
      : (normalizeDocumentUnderlineStyle(rawStyle) ?? 'single');
  const themeName = attribute(element, 'themeColor')?.trim();
  const tint = normalizedByteHex(attribute(element, 'themeTint'));
  const shade = normalizedByteHex(attribute(element, 'themeShade'));
  const themed =
    theme && themeName
      ? docxThemeColor(theme, themeName, tint, shade)
      : undefined;
  const direct = normalizeDocumentUnderlineColor(attribute(element, 'color'));
  const color = themed ? `#${themed}` : direct;
  const themeColor =
    themeName && color
      ? underlineThemeReference(themeName, color, tint, shade)
      : undefined;
  return {
    style,
    ...(color ? { color } : {}),
    ...(themeColor ? { themeColor } : {}),
  };
}

function underlineThemeReference(
  theme: string,
  resolved: string,
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

function normalizedByteHex(value: string | null): string | undefined {
  const normalized = value?.trim().toUpperCase();
  return normalized && /^[0-9A-F]{2}$/.test(normalized)
    ? normalized
    : undefined;
}
