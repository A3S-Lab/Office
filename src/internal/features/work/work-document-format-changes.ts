import type { Mark, MarkType, Schema } from '@tiptap/pm/model';
import type { Transaction } from '@tiptap/pm/state';
import { normalizeDocumentCharacterScalePercent } from './work-document-character-scale';
import { normalizeDocumentCharacterPositionHalfPoints } from './work-document-character-position';
import { normalizeDocumentCharacterSpacingTwips } from './work-document-character-spacing';
import {
  normalizeDocumentTextCase,
  type WorkDocumentTextCase,
} from './work-document-text-case';
import {
  normalizeDocumentUnderlineColor,
  normalizeDocumentUnderlineStyle,
  type WorkDocumentUnderlineFormatting,
} from './work-document-underline';
import {
  normalizeDocumentStrikeStyle,
  type WorkDocumentStrikeFormatting,
} from './work-document-strike';
import {
  parseDocxThemeReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';

export const DOCUMENT_CHARACTER_FORMAT_MARKS = [
  'bold',
  'italic',
  'underline',
  'strike',
  'subscript',
  'superscript',
  'textStyle',
  'highlight',
] as const;

export type DocumentCharacterFormatMarkName =
  (typeof DOCUMENT_CHARACTER_FORMAT_MARKS)[number];

export interface DocumentCharacterFormatMark {
  type: DocumentCharacterFormatMarkName;
  attrs?: Record<string, boolean | number | string>;
}

const CHARACTER_FORMAT_MARK_NAMES = new Set<string>(
  DOCUMENT_CHARACTER_FORMAT_MARKS,
);
const CHARACTER_FORMAT_MARK_ORDER = new Map<string, number>(
  DOCUMENT_CHARACTER_FORMAT_MARKS.map((name, index) => [name, index]),
);
const MAX_CHARACTER_FORMAT_SNAPSHOT_BYTES = 4_096;
const MAX_CHARACTER_FORMAT_ATTRIBUTE_LENGTH = 512;

const ALLOWED_ATTRIBUTES: Readonly<
  Record<DocumentCharacterFormatMarkName, ReadonlySet<string>>
> = {
  bold: new Set(),
  italic: new Set(),
  underline: new Set([
    'underlineColor',
    'underlineStyle',
    'underlineThemeColor',
  ]),
  strike: new Set(['strikeStyle']),
  subscript: new Set(),
  superscript: new Set(),
  textStyle: new Set([
    'characterScalePercent',
    'characterPositionHalfPoints',
    'characterSpacingTwips',
    'color',
    'fontFamily',
    'fontSize',
    'themeColor',
    'textCase',
    'wordLineHeightFactor',
    'wordSnapToGrid',
  ]),
  highlight: new Set(['color', 'themeFill']),
};

export function isDocumentCharacterFormatMark(value: string): boolean {
  return CHARACTER_FORMAT_MARK_NAMES.has(value);
}

export function serializeDocumentCharacterFormatting(
  marks: readonly Mark[],
): string {
  return JSON.stringify(
    marks
      .flatMap((mark) => {
        const normalized = normalizeCharacterFormatMark({
          type: mark.type.name,
          attrs: mark.attrs,
        });
        return normalized ? [normalized] : [];
      })
      .sort(compareCharacterFormatMarks),
  );
}

export function parseDocumentCharacterFormatting(
  value: unknown,
): DocumentCharacterFormatMark[] | null {
  if (
    typeof value !== 'string' ||
    value.length > MAX_CHARACTER_FORMAT_SNAPSHOT_BYTES
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (
    !Array.isArray(parsed) ||
    parsed.length > DOCUMENT_CHARACTER_FORMAT_MARKS.length
  ) {
    return null;
  }
  const result: DocumentCharacterFormatMark[] = [];
  const names = new Set<string>();
  for (const candidate of parsed) {
    const mark = normalizeCharacterFormatMark(candidate);
    if (!mark || names.has(mark.type)) return null;
    names.add(mark.type);
    result.push(mark);
  }
  result.sort(compareCharacterFormatMarks);
  return JSON.stringify(result) === value ? result : null;
}

export function restoreDocumentCharacterFormatting(
  transaction: Transaction,
  schema: Schema,
  from: number,
  to: number,
  serialized: unknown,
): boolean {
  const formatting = parseDocumentCharacterFormatting(serialized);
  if (
    !formatting ||
    from < 0 ||
    to <= from ||
    to > transaction.doc.content.size
  ) {
    return false;
  }
  for (const name of DOCUMENT_CHARACTER_FORMAT_MARKS) {
    const type = schema.marks[name];
    if (type) transaction.removeMark(from, to, type);
  }
  for (const mark of formatting) {
    const type = schema.marks[mark.type];
    if (!type) return false;
    transaction.addMark(from, to, createCharacterFormatMark(type, mark));
  }
  return true;
}

export function importedDocumentCharacterFormatting(formatting: {
  bold?: boolean;
  italic?: boolean;
  underline?: WorkDocumentUnderlineFormatting;
  strike?: WorkDocumentStrikeFormatting;
  subscript?: boolean;
  superscript?: boolean;
  fontFamily?: string;
  characterScalePercent?: number;
  characterPositionHalfPoints?: number;
  characterSpacingTwips?: number;
  wordLineHeightFactor?: number;
  wordSnapToGrid?: boolean;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  themeColor?: string;
  themeFill?: string;
  textCase?: WorkDocumentTextCase;
}): string {
  const marks: DocumentCharacterFormatMark[] = [];
  for (const name of ['bold', 'italic', 'subscript', 'superscript'] as const) {
    if (formatting[name]) marks.push({ type: name });
  }
  if (formatting.strike) {
    marks.push({
      type: 'strike',
      attrs: { strikeStyle: formatting.strike.style },
    });
  }
  if (formatting.underline) {
    marks.push({
      type: 'underline',
      attrs: compactAttributes({
        underlineColor: formatting.underline.color,
        underlineStyle: formatting.underline.style,
        underlineThemeColor: serializeDocxThemeReference(
          formatting.underline.themeColor ?? null,
        ),
      }),
    });
  }
  const textStyle = compactAttributes({
    characterScalePercent: formatting.characterScalePercent,
    characterPositionHalfPoints: formatting.characterPositionHalfPoints,
    characterSpacingTwips: formatting.characterSpacingTwips,
    color: formatting.color,
    fontFamily: formatting.fontFamily,
    fontSize:
      formatting.fontSize === undefined
        ? undefined
        : `${formatting.fontSize}pt`,
    themeColor: formatting.themeColor,
    textCase: formatting.textCase,
    wordLineHeightFactor: formatting.wordLineHeightFactor,
    wordSnapToGrid: formatting.wordSnapToGrid,
  });
  if (textStyle) marks.push({ type: 'textStyle', attrs: textStyle });
  const highlight = compactAttributes({
    color: formatting.backgroundColor,
    themeFill: formatting.themeFill,
  });
  if (highlight) marks.push({ type: 'highlight', attrs: highlight });
  marks.sort(compareCharacterFormatMarks);
  return JSON.stringify(marks);
}

function normalizeCharacterFormatMark(
  value: unknown,
): DocumentCharacterFormatMark | null {
  if (!isRecord(value) || typeof value.type !== 'string') return null;
  if (!isDocumentCharacterFormatMark(value.type)) return null;
  const type = value.type as DocumentCharacterFormatMarkName;
  const allowed = ALLOWED_ATTRIBUTES[type];
  const source = value.attrs;
  if (source !== undefined && !isRecord(source)) return null;
  const attrs: Record<string, boolean | number | string> = {};
  for (const [key, candidate] of Object.entries(source ?? {})) {
    if (!allowed.has(key) || candidate === null || candidate === undefined) {
      continue;
    }
    if (type === 'textStyle' && key === 'characterSpacingTwips') {
      const spacing = normalizeDocumentCharacterSpacingTwips(candidate);
      if (spacing === null) return null;
      attrs[key] = spacing;
      continue;
    }
    if (type === 'textStyle' && key === 'characterPositionHalfPoints') {
      const position = normalizeDocumentCharacterPositionHalfPoints(candidate);
      if (position === null) return null;
      attrs[key] = position;
      continue;
    }
    if (type === 'textStyle' && key === 'characterScalePercent') {
      const scale = normalizeDocumentCharacterScalePercent(candidate);
      if (scale === null) return null;
      attrs[key] = scale;
      continue;
    }
    if (typeof candidate === 'string') {
      if (!candidate.length) continue;
      if (
        type === 'textStyle' &&
        key === 'textCase' &&
        !normalizeDocumentTextCase(candidate)
      ) {
        return null;
      }
      if (candidate.length > MAX_CHARACTER_FORMAT_ATTRIBUTE_LENGTH) {
        return null;
      }
      const normalized = normalizeCharacterFormatStringAttribute(
        type,
        key,
        candidate,
      );
      if (!normalized) return null;
      attrs[key] = normalized;
      continue;
    }
    if (typeof candidate === 'boolean') {
      attrs[key] = candidate;
      continue;
    }
    if (typeof candidate === 'number' && Number.isFinite(candidate)) {
      attrs[key] = candidate;
      continue;
    }
    return null;
  }
  const keys = Object.keys(attrs).sort();
  const normalizedAttributes = keys.length
    ? Object.fromEntries(keys.map((key) => [key, attrs[key]]))
    : undefined;
  return {
    type,
    ...(normalizedAttributes ? { attrs: normalizedAttributes } : {}),
  };
}

function normalizeCharacterFormatStringAttribute(
  type: DocumentCharacterFormatMarkName,
  key: string,
  value: string,
): string | null {
  if (type === 'strike') {
    return key === 'strikeStyle' ? normalizeDocumentStrikeStyle(value) : null;
  }
  if (type !== 'underline') return value;
  if (key === 'underlineStyle') {
    return normalizeDocumentUnderlineStyle(value);
  }
  if (key === 'underlineColor') {
    return normalizeDocumentUnderlineColor(value);
  }
  if (key === 'underlineThemeColor') {
    return serializeDocxThemeReference(parseDocxThemeReference(value)) ?? null;
  }
  return null;
}

function compareCharacterFormatMarks(
  left: DocumentCharacterFormatMark,
  right: DocumentCharacterFormatMark,
): number {
  return (
    (CHARACTER_FORMAT_MARK_ORDER.get(left.type) ?? Number.MAX_SAFE_INTEGER) -
    (CHARACTER_FORMAT_MARK_ORDER.get(right.type) ?? Number.MAX_SAFE_INTEGER)
  );
}

function createCharacterFormatMark(
  type: MarkType,
  value: DocumentCharacterFormatMark,
): Mark {
  return type.create(value.attrs ?? undefined);
}

function compactAttributes(
  value: Record<string, boolean | number | string | undefined>,
): Record<string, boolean | number | string> | undefined {
  const entries = Object.entries(value).filter(
    (entry): entry is [string, boolean | number | string] =>
      entry[1] !== undefined,
  );
  return entries.length ? Object.fromEntries(entries) : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
