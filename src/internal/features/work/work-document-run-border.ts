import {
  DOCUMENT_PARAGRAPH_BORDER_STYLES,
  type DocumentParagraphBorder,
  type DocumentParagraphBorderStyle,
  documentBorderPresentation,
  isDocumentParagraphArtBorderStyle,
  normalizeDocumentParagraphBorder,
} from './work-document-paragraph-borders';
import { normalizeCssColor } from './work-document-paragraph-shading';
import { serializeDocxThemeReference } from './work-docx-theme-reference';

export const DOCUMENT_RUN_BORDER_ATTRIBUTE = 'data-office-run-border';

export type DocumentRunBorder = DocumentParagraphBorder;
export type DocumentRunBorderStyle = DocumentParagraphBorderStyle;

export const DOCUMENT_RUN_BORDER_STYLES =
  DOCUMENT_PARAGRAPH_BORDER_STYLES.slice(
    0,
    27,
  ) as readonly DocumentRunBorderStyle[];

const MAX_SERIALIZED_RUN_BORDER_BYTES = 4_096;
const POINTS_TO_PIXELS = 96 / 72;

export function normalizeDocumentRunBorder(
  source: unknown,
): DocumentRunBorder | null {
  const border = normalizeDocumentParagraphBorder(source);
  return border && !isDocumentParagraphArtBorderStyle(border.style)
    ? border
    : null;
}

export function parseDocumentRunBorder(
  source: unknown,
): DocumentRunBorder | null {
  if (typeof source !== 'string') return normalizeDocumentRunBorder(source);
  if (!source.trim() || source.length > MAX_SERIALIZED_RUN_BORDER_BYTES) {
    return null;
  }
  try {
    return normalizeDocumentRunBorder(JSON.parse(source));
  } catch {
    return null;
  }
}

export function serializeDocumentRunBorder(
  source: unknown,
): string | undefined {
  const border = normalizeDocumentRunBorder(source);
  if (!border) return undefined;
  const theme = serializeDocxThemeReference(border.color?.theme ?? null);
  return JSON.stringify({
    style: border.style,
    ...(border.color
      ? {
          color: {
            value: border.color.value,
            ...(theme
              ? { theme: JSON.parse(theme) as Record<string, unknown> }
              : {}),
          },
        }
      : {}),
    ...(border.size !== undefined ? { size: border.size } : {}),
    ...(border.space !== undefined ? { space: border.space } : {}),
    ...(border.shadow !== undefined ? { shadow: border.shadow } : {}),
    ...(border.frame !== undefined ? { frame: border.frame } : {}),
  });
}

export function parseDocumentRunBorderElement(
  element: HTMLElement,
): DocumentRunBorder | null {
  const semantic = parseDocumentRunBorder(
    element.getAttribute(DOCUMENT_RUN_BORDER_ATTRIBUTE),
  );
  return semantic ?? documentRunBorderFromCss(element);
}

export function documentRunBorderDomAttributes(
  source: unknown,
): Record<string, string> {
  const border = normalizeDocumentRunBorder(source);
  const serialized = serializeDocumentRunBorder(border);
  if (!border || !serialized) return {};
  const presentation = documentBorderPresentation(border);
  const declarations = [
    `border: ${formatPixels(presentation.width)}px ${presentation.style} ${presentation.color}`,
    `padding: ${formatPixels((border.space ?? 0) * POINTS_TO_PIXELS)}px`,
    'box-decoration-break: clone',
    '-webkit-box-decoration-break: clone',
  ];
  if (border.shadow && presentation.width > 0) {
    declarations.push(`box-shadow: 2px 2px 0 ${presentation.color}`);
  }
  return {
    [DOCUMENT_RUN_BORDER_ATTRIBUTE]: serialized,
    style: declarations.join('; '),
  };
}

export function documentRunBorderIsVisible(source: unknown): boolean {
  const border = normalizeDocumentRunBorder(source);
  return Boolean(border && documentBorderPresentation(border).width > 0);
}

function documentRunBorderFromCss(
  element: HTMLElement,
): DocumentRunBorder | null {
  const style = element.style.borderStyle.trim();
  if (!style) return null;
  if (style === 'none' || style === 'hidden') return { style: 'none' };
  const width = Number.parseFloat(element.style.borderWidth);
  const color = normalizeCssColor(element.style.borderColor);
  if (
    !Number.isFinite(width) ||
    width <= 0 ||
    !color ||
    color === 'transparent'
  ) {
    return null;
  }
  const padding = Number.parseFloat(element.style.padding);
  return normalizeDocumentRunBorder({
    style: cssBorderStyle(style),
    color: { value: color },
    size: Math.max(2, Math.min(96, Math.round(width * 6))),
    ...(Number.isFinite(padding) && padding >= 0
      ? {
          space: Math.max(
            0,
            Math.min(31, Math.round(padding / POINTS_TO_PIXELS)),
          ),
        }
      : {}),
  });
}

function cssBorderStyle(style: string): DocumentRunBorderStyle {
  if (style === 'double') return 'double';
  if (style === 'dashed') return 'dashed';
  if (style === 'dotted') return 'dotted';
  if (style === 'inset' || style === 'groove') return 'inset';
  if (style === 'outset' || style === 'ridge') return 'outset';
  return 'single';
}

function formatPixels(value: number): string {
  return Number(value.toFixed(3)).toString();
}
