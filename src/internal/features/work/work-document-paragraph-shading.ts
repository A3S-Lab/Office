import {
  type DocxThemeColorReference,
  parseDocxThemeReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';

export type DocumentParagraphShadingPattern =
  | 'nil'
  | 'clear'
  | 'solid'
  | 'horzStripe'
  | 'vertStripe'
  | 'reverseDiagStripe'
  | 'diagStripe'
  | 'horzCross'
  | 'diagCross'
  | 'thinHorzStripe'
  | 'thinVertStripe'
  | 'thinReverseDiagStripe'
  | 'thinDiagStripe'
  | 'thinHorzCross'
  | 'thinDiagCross'
  | 'pct5'
  | 'pct10'
  | 'pct12'
  | 'pct15'
  | 'pct20'
  | 'pct25'
  | 'pct30'
  | 'pct35'
  | 'pct37'
  | 'pct40'
  | 'pct45'
  | 'pct50'
  | 'pct55'
  | 'pct60'
  | 'pct62'
  | 'pct65'
  | 'pct70'
  | 'pct75'
  | 'pct80'
  | 'pct85'
  | 'pct87'
  | 'pct90'
  | 'pct95';

export interface DocumentParagraphShadingColor {
  value: 'auto' | `#${string}`;
  theme?: DocxThemeColorReference;
}

export interface DocumentParagraphShading {
  pattern: DocumentParagraphShadingPattern;
  color?: DocumentParagraphShadingColor;
  fill?: DocumentParagraphShadingColor;
}

export const DOCUMENT_PARAGRAPH_SHADING_PATTERNS =
  new Set<DocumentParagraphShadingPattern>([
    'nil',
    'clear',
    'solid',
    'horzStripe',
    'vertStripe',
    'reverseDiagStripe',
    'diagStripe',
    'horzCross',
    'diagCross',
    'thinHorzStripe',
    'thinVertStripe',
    'thinReverseDiagStripe',
    'thinDiagStripe',
    'thinHorzCross',
    'thinDiagCross',
    'pct5',
    'pct10',
    'pct12',
    'pct15',
    'pct20',
    'pct25',
    'pct30',
    'pct35',
    'pct37',
    'pct40',
    'pct45',
    'pct50',
    'pct55',
    'pct60',
    'pct62',
    'pct65',
    'pct70',
    'pct75',
    'pct80',
    'pct85',
    'pct87',
    'pct90',
    'pct95',
  ]);

export function normalizeDocumentParagraphShading(
  source: unknown,
): DocumentParagraphShading | null {
  if (!source || typeof source !== 'object') return null;
  const value = source as Record<string, unknown>;
  const pattern = value.pattern;
  if (
    typeof pattern !== 'string' ||
    !DOCUMENT_PARAGRAPH_SHADING_PATTERNS.has(
      pattern as DocumentParagraphShadingPattern,
    )
  ) {
    return null;
  }
  const color = normalizeShadingColor(value.color);
  const fill = normalizeShadingColor(value.fill);
  if (
    (value.color !== undefined && !color) ||
    (value.fill !== undefined && !fill)
  ) {
    return null;
  }
  return {
    pattern: pattern as DocumentParagraphShadingPattern,
    ...(color ? { color } : {}),
    ...(fill ? { fill } : {}),
  };
}

export function parseDocumentParagraphShading(
  source: unknown,
): DocumentParagraphShading | null {
  if (typeof source !== 'string')
    return normalizeDocumentParagraphShading(source);
  if (!source.trim()) return null;
  try {
    return normalizeDocumentParagraphShading(JSON.parse(source));
  } catch {
    return null;
  }
}

export function serializeDocumentParagraphShading(
  source: unknown,
): string | undefined {
  const shading = normalizeDocumentParagraphShading(source);
  if (!shading) return undefined;
  return JSON.stringify({
    pattern: shading.pattern,
    ...(shading.color ? { color: serializedShadingColor(shading.color) } : {}),
    ...(shading.fill ? { fill: serializedShadingColor(shading.fill) } : {}),
  });
}

export function parseDocumentParagraphShadingElement(
  element: HTMLElement,
): DocumentParagraphShading | null {
  const semantic = parseDocumentParagraphShading(
    element.dataset.officeParagraphShading,
  );
  const background = normalizeCssColor(element.style.backgroundColor);
  if (semantic) {
    const presentation = documentParagraphShadingPresentation(semantic);
    const expected = normalizeCssColor(presentation.backgroundColor);
    if (background && expected && background !== expected) {
      if (background === 'transparent') return { pattern: 'nil' };
      if (paragraphShadingBackgroundUsesForeground(semantic)) {
        return { ...semantic, color: { value: background } };
      }
      return { ...semantic, fill: { value: background } };
    }
    return semantic;
  }
  return background && background !== 'transparent'
    ? { pattern: 'clear', fill: { value: background } }
    : null;
}

export function documentParagraphShadingDomAttributes(
  source: unknown,
): Record<string, string> {
  const shading = normalizeDocumentParagraphShading(source);
  const serialized = serializeDocumentParagraphShading(shading);
  if (!shading || !serialized) return {};
  const presentation = documentParagraphShadingPresentation(shading);
  const styles = [
    `background-color: ${presentation.backgroundColor}`,
    presentation.backgroundImage
      ? `background-image: ${presentation.backgroundImage}`
      : '',
    presentation.backgroundSize
      ? `background-size: ${presentation.backgroundSize}`
      : '',
  ].filter(Boolean);
  return {
    'data-office-paragraph-shading': serialized,
    style: styles.join('; '),
  };
}

export function normalizeCssColor(
  source: string | null | undefined,
): 'transparent' | `#${string}` | null {
  const value = source?.trim().toLowerCase();
  if (!value) return null;
  if (value === 'transparent') return 'transparent';
  const shortHex = /^#([0-9a-f]{3})$/i.exec(value);
  if (shortHex?.[1]) {
    return `#${Array.from(shortHex[1])
      .map((channel) => `${channel}${channel}`)
      .join('')}`;
  }
  if (/^#[0-9a-f]{6}$/i.test(value)) return value as `#${string}`;
  const rgb =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([\d.]+))?\s*\)$/i.exec(
      value,
    );
  if (!rgb) return null;
  const channels = rgb.slice(1, 4).map(Number);
  if (channels.some((channel) => channel < 0 || channel > 255)) return null;
  if (rgb[4] !== undefined && Number(rgb[4]) === 0) return 'transparent';
  if (rgb[4] !== undefined && Number(rgb[4]) !== 1) return null;
  return `#${channels
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}

function normalizeShadingColor(
  source: unknown,
): DocumentParagraphShadingColor | null {
  if (!source || typeof source !== 'object') return null;
  const value = source as Record<string, unknown>;
  const theme = parseDocxThemeReference(
    typeof value.theme === 'string'
      ? value.theme
      : value.theme
        ? JSON.stringify(value.theme)
        : undefined,
  );
  const direct =
    value.value === 'auto'
      ? ('auto' as const)
      : typeof value.value === 'string'
        ? normalizeCssColor(value.value)
        : null;
  const resolved = direct ?? theme?.resolved ?? null;
  if (!resolved || resolved === 'transparent') return null;
  if (theme && resolved !== theme.resolved) return null;
  return {
    value: resolved as DocumentParagraphShadingColor['value'],
    ...(theme ? { theme } : {}),
  };
}

function serializedShadingColor(
  color: DocumentParagraphShadingColor,
): Record<string, unknown> {
  const theme = serializeDocxThemeReference(color.theme ?? null);
  return {
    value: color.value,
    ...(theme ? { theme: JSON.parse(theme) as Record<string, unknown> } : {}),
  };
}

function documentParagraphShadingPresentation(
  shading: DocumentParagraphShading,
): {
  backgroundColor: string;
  backgroundImage?: string;
  backgroundSize?: string;
} {
  if (shading.pattern === 'nil') return { backgroundColor: 'transparent' };
  const foreground = shadingColor(shading.color, '#000000');
  const background = shadingColor(shading.fill, 'transparent');
  if (shading.pattern === 'clear') return { backgroundColor: background };
  if (shading.pattern === 'solid') return { backgroundColor: foreground };

  const percentage = shadingPercentage(shading.pattern);
  if (percentage !== null) {
    const inverted = percentage > 50;
    const dotColor = inverted ? background : foreground;
    const baseColor = inverted ? foreground : background;
    const density = Math.min(50, inverted ? 100 - percentage : percentage);
    const spacing = Math.max(2, Math.round(9 - density / 7));
    return {
      backgroundColor: baseColor,
      backgroundImage: `radial-gradient(circle, ${dotColor} 0 1px, transparent 1.2px)`,
      backgroundSize: `${spacing}px ${spacing}px`,
    };
  }

  const thin = shading.pattern.startsWith('thin');
  const width = thin ? 1 : 2;
  const period = thin ? 7 : 6;
  const stripe = (angle: number) =>
    `repeating-linear-gradient(${angle}deg, ${foreground} 0 ${width}px, transparent ${width}px ${period}px)`;
  const angles = shadingPatternAngles(shading.pattern);
  return {
    backgroundColor: background,
    backgroundImage: angles.map(stripe).join(', '),
  };
}

function shadingColor(
  color: DocumentParagraphShadingColor | undefined,
  fallback: string,
): string {
  return !color || color.value === 'auto' ? fallback : color.value;
}

function shadingPercentage(
  pattern: DocumentParagraphShadingPattern,
): number | null {
  const match = /^pct(\d+)$/.exec(pattern);
  return match?.[1] ? Number(match[1]) : null;
}

function paragraphShadingBackgroundUsesForeground(
  shading: DocumentParagraphShading,
): boolean {
  if (shading.pattern === 'solid') return true;
  const percentage = shadingPercentage(shading.pattern);
  return percentage !== null && percentage > 50;
}

function shadingPatternAngles(
  pattern: DocumentParagraphShadingPattern,
): number[] {
  if (pattern.includes('HorzCross')) return [0, 90];
  if (pattern.includes('DiagCross')) return [45, -45];
  if (pattern.includes('VertStripe')) return [90];
  if (pattern.includes('ReverseDiagStripe')) return [-45];
  if (pattern.includes('DiagStripe')) return [45];
  return [0];
}
