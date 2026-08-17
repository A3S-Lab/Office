import { serializeDocumentParagraphBorders } from './work-document-paragraph-borders';
import { serializeDocumentParagraphShading } from './work-document-paragraph-shading';
import {
  normalizeDocumentTabStops,
  serializeDocumentTabStops,
} from './work-document-tab-stops';

export const DOCUMENT_PARAGRAPH_CHANGE_ATTRIBUTES = [
  'paragraphChangeKind',
  'paragraphChangeId',
  'paragraphChangeActorId',
  'paragraphChangeAuthor',
  'paragraphChangeDate',
  'paragraphChangeBefore',
] as const;

export const DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES = [
  'textAlign',
  'paragraphDirection',
  'indentLevel',
  'rightIndent',
  'firstLineIndent',
  'spaceBefore',
  'spaceAfter',
  'lineHeight',
  'lineRule',
  'autoLineHeight',
  'keepLines',
  'keepWithNext',
  'pageBreakBefore',
  'widowControl',
  'contextualSpacing',
  'outlineLevel',
  'tabStops',
  'paragraphBorders',
  'paragraphShading',
  'defaultCollapsed',
] as const;

export type DocumentParagraphFormatAttribute =
  (typeof DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES)[number];

export type DocumentParagraphFormattingSnapshot = Record<
  DocumentParagraphFormatAttribute,
  unknown
>;

const MAX_PARAGRAPH_FORMAT_SNAPSHOT_BYTES = 65_536;
const MAX_DOCUMENT_INDENT_PX = 192;

export function serializeDocumentParagraphFormatting(
  attributes: Record<string, unknown>,
): string {
  return JSON.stringify(normalizeParagraphFormatting(attributes));
}

export function parseDocumentParagraphFormatting(
  value: unknown,
): DocumentParagraphFormattingSnapshot | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_PARAGRAPH_FORMAT_SNAPSHOT_BYTES
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (!isRecord(parsed)) return null;
  const keys = Object.keys(parsed);
  if (
    keys.length !== DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES.length ||
    keys.some(
      (key, index) => key !== DOCUMENT_PARAGRAPH_FORMAT_ATTRIBUTES[index],
    )
  ) {
    return null;
  }
  const normalized = normalizeParagraphFormatting(parsed);
  return JSON.stringify(normalized) === value ? normalized : null;
}

export function restoredDocumentParagraphAttributes(
  attributes: Record<string, unknown>,
  serialized: unknown,
): Record<string, unknown> | null {
  const formatting = parseDocumentParagraphFormatting(serialized);
  if (!formatting) return null;
  return clearDocumentParagraphChangeAttributes({
    ...attributes,
    ...formatting,
  });
}

export function clearDocumentParagraphChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  return {
    ...attributes,
    paragraphChangeKind: null,
    paragraphChangeId: '',
    paragraphChangeActorId: '',
    paragraphChangeAuthor: '',
    paragraphChangeDate: '',
    paragraphChangeBefore: '',
  };
}

function normalizeParagraphFormatting(
  source: Record<string, unknown>,
): DocumentParagraphFormattingSnapshot {
  const indentLevel = quarterNumber(source.indentLevel, 0, 8, 0);
  const leftIndent = indentLevel * 24;
  const serializedBorders = serializeDocumentParagraphBorders(
    source.paragraphBorders,
  );
  const serializedShading = serializeDocumentParagraphShading(
    source.paragraphShading,
  );
  const tabStops = normalizeDocumentTabStops(source.tabStops);
  return {
    textAlign: oneOfOrNull(source.textAlign, [
      'left',
      'center',
      'right',
      'justify',
    ]),
    paragraphDirection: oneOfOrNull(source.paragraphDirection, ['ltr', 'rtl']),
    indentLevel,
    rightIndent: integerNumber(
      source.rightIndent,
      0,
      MAX_DOCUMENT_INDENT_PX,
      0,
    ),
    firstLineIndent: integerNumber(
      source.firstLineIndent,
      -leftIndent,
      MAX_DOCUMENT_INDENT_PX,
      0,
    ),
    spaceBefore: quarterNumberOrNull(source.spaceBefore, 0, 720),
    spaceAfter: quarterNumberOrNull(source.spaceAfter, 0, 720),
    lineHeight: normalizedLineHeight(source.lineHeight),
    lineRule: oneOfOrNull(source.lineRule, ['auto', 'exact', 'atLeast']),
    autoLineHeight: decimalNumberOrNull(source.autoLineHeight, 0, 20, 4),
    keepLines: booleanOrNull(source.keepLines),
    keepWithNext: booleanOrNull(source.keepWithNext),
    pageBreakBefore: booleanOrNull(source.pageBreakBefore),
    widowControl: booleanOrNull(source.widowControl),
    contextualSpacing: booleanOrNull(source.contextualSpacing),
    outlineLevel: integerNumberOrNull(source.outlineLevel, 0, 9),
    tabStops: tabStops.length
      ? (JSON.parse(serializeDocumentTabStops(tabStops)) as unknown)
      : null,
    paragraphBorders: serializedBorders
      ? (JSON.parse(serializedBorders) as unknown)
      : null,
    paragraphShading: serializedShading
      ? (JSON.parse(serializedShading) as unknown)
      : null,
    defaultCollapsed: booleanOrNull(source.defaultCollapsed),
  };
}

function normalizedLineHeight(value: unknown): string | null {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value !== 'string' || value.trim() !== value) return null;
  return /^(?:\d+(?:\.\d+)?|\d+(?:\.\d+)?(?:px|pt|%))$/i.test(value)
    ? value
    : null;
}

function oneOfOrNull(
  value: unknown,
  allowed: readonly string[],
): string | null {
  return typeof value === 'string' && allowed.includes(value) ? value : null;
}

function booleanOrNull(value: unknown): boolean | null {
  return typeof value === 'boolean' ? value : null;
}

function quarterNumber(
  value: unknown,
  minimum: number,
  maximum: number,
  fallback: number,
): number {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(maximum, Math.max(minimum, Math.round(number * 4) / 4));
}

function quarterNumberOrNull(
  value: unknown,
  minimum: number,
  maximum: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  return quarterNumber(value, minimum, maximum, minimum);
}

function integerNumber(
  value: unknown,
  minimum: number,
  maximum: number,
  fallback: number,
): number {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(maximum, Math.max(minimum, Math.round(number)));
}

function integerNumberOrNull(
  value: unknown,
  minimum: number,
  maximum: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < minimum || number > maximum)
    return null;
  return number;
}

function decimalNumberOrNull(
  value: unknown,
  minimumExclusive: number,
  maximum: number,
  precision: number,
): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  if (
    !Number.isFinite(number) ||
    number <= minimumExclusive ||
    number > maximum
  ) {
    return null;
  }
  return Number(number.toFixed(precision));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
