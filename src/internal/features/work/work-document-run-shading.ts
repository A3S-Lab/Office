import {
  type DocumentParagraphShading,
  type DocumentParagraphShadingColor,
  type DocumentParagraphShadingPattern,
  DOCUMENT_PARAGRAPH_SHADING_PATTERNS,
  documentParagraphShadingDomAttributes,
  documentParagraphShadingPresentation,
  normalizeDocumentParagraphShading,
  paragraphShadingBackgroundUsesForeground,
  serializeDocumentParagraphShading,
} from './work-document-paragraph-shading';
import { normalizeCssColor } from './work-css-color';
import { DOCUMENT_HIGHLIGHT_ATTRIBUTE } from './work-document-highlight';

export const DOCUMENT_RUN_SHADING_ATTRIBUTE = 'data-office-run-shading';

export type DocumentRunShading = DocumentParagraphShading;
export type DocumentRunShadingColor = DocumentParagraphShadingColor;
export type DocumentRunShadingPattern = DocumentParagraphShadingPattern;

export const DOCUMENT_RUN_SHADING_PATTERNS =
  DOCUMENT_PARAGRAPH_SHADING_PATTERNS;

const MAX_SERIALIZED_RUN_SHADING_BYTES = 4_096;
const RUN_SHADING_KEYS = new Set(['pattern', 'color', 'fill']);
const RUN_SHADING_COLOR_KEYS = new Set(['value', 'theme']);
const THEME_REFERENCE_KEYS = new Set(['theme', 'resolved', 'tint', 'shade']);

export function normalizeDocumentRunShading(
  source: unknown,
): DocumentRunShading | null {
  if (!isRecordWithKeys(source, RUN_SHADING_KEYS)) return null;
  for (const name of ['color', 'fill'] as const) {
    const color = source[name];
    if (color === undefined) continue;
    if (!isRecordWithKeys(color, RUN_SHADING_COLOR_KEYS)) return null;
    if (
      color.theme !== undefined &&
      !isRecordWithKeys(color.theme, THEME_REFERENCE_KEYS)
    ) {
      return null;
    }
  }
  return normalizeDocumentParagraphShading(source);
}

export function parseDocumentRunShading(
  source: unknown,
): DocumentRunShading | null {
  if (typeof source !== 'string') return normalizeDocumentRunShading(source);
  if (!source.trim() || source.length > MAX_SERIALIZED_RUN_SHADING_BYTES) {
    return null;
  }
  try {
    return normalizeDocumentRunShading(JSON.parse(source));
  } catch {
    return null;
  }
}

export function serializeDocumentRunShading(
  source: unknown,
): string | undefined {
  const shading = normalizeDocumentRunShading(source);
  const serialized = shading
    ? serializeDocumentParagraphShading(shading)
    : undefined;
  return serialized && serialized.length <= MAX_SERIALIZED_RUN_SHADING_BYTES
    ? serialized
    : undefined;
}

export function parseDocumentRunShadingElement(
  element: HTMLElement,
): DocumentRunShading | null {
  const semantic = parseDocumentRunShading(
    element.getAttribute(DOCUMENT_RUN_SHADING_ATTRIBUTE),
  );
  if (!semantic) return null;
  if (element.hasAttribute(DOCUMENT_HIGHLIGHT_ATTRIBUTE)) return semantic;
  const background = normalizeCssColor(element.style.backgroundColor);
  const expected = normalizeCssColor(
    documentParagraphShadingPresentation(semantic).backgroundColor,
  );
  if (!background || !expected || background === expected) return semantic;
  if (background === 'transparent') return { pattern: 'nil' };
  return paragraphShadingBackgroundUsesForeground(semantic)
    ? { ...semantic, color: { value: background } }
    : { ...semantic, fill: { value: background } };
}

export function documentRunShadingDomAttributes(
  source: unknown,
): Record<string, string> {
  const shading = normalizeDocumentRunShading(source);
  const serialized = serializeDocumentRunShading(shading);
  if (!shading || !serialized) return {};
  const paragraphAttributes = documentParagraphShadingDomAttributes(shading);
  const style = [
    paragraphAttributes.style,
    'box-decoration-break: clone',
    '-webkit-box-decoration-break: clone',
  ]
    .filter(Boolean)
    .join('; ');
  return {
    [DOCUMENT_RUN_SHADING_ATTRIBUTE]: serialized,
    style,
  };
}

export function documentRunShadingIsVisible(source: unknown): boolean {
  const shading = normalizeDocumentRunShading(source);
  return Boolean(shading && shading.pattern !== 'nil');
}

function isRecordWithKeys(
  source: unknown,
  allowed: ReadonlySet<string>,
): source is Record<string, unknown> {
  return (
    typeof source === 'object' &&
    source !== null &&
    !Array.isArray(source) &&
    Object.keys(source).every((key) => allowed.has(key))
  );
}
