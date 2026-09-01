import { MAX_DOCUMENT_NUMBERING_START } from './work-document-lists';

export type DocumentNumberingChangeType = 'a' | 'A' | 'i' | 'I' | null;

export interface DocumentNumberingChangeSnapshot {
  start: number;
  type: DocumentNumberingChangeType;
  officeNumberingId: string | null;
  officeAbstractNumberingId: string | null;
  officeNumberingLevel: string | null;
  officeNumberingFormat: string | null;
  officeNumberingText: string | null;
  officeNumberingSuffix: string | null;
  officeNumberingAlignment: string | null;
  officeNumberingIndentLeft: string | null;
  officeNumberingIndentRight: string | null;
  officeNumberingIndentStart: string | null;
  officeNumberingIndentEnd: string | null;
  officeNumberingIndentHanging: string | null;
  officeNumberingIndentFirstLine: string | null;
  officeNumberingRestartAfterLevel: string | null;
  level: number;
  originalFormat: number;
  originalSuffix: string;
}

export const DOCUMENT_NUMBERING_CHANGE_ATTRIBUTES = [
  'numberingChangeKind',
  'numberingChangeId',
  'numberingChangeActorId',
  'numberingChangeAuthor',
  'numberingChangeDate',
  'numberingChangeBefore',
] as const;

const NUMBERING_SNAPSHOT_KEYS = [
  'start',
  'type',
  'officeNumberingId',
  'officeAbstractNumberingId',
  'officeNumberingLevel',
  'officeNumberingFormat',
  'officeNumberingText',
  'officeNumberingSuffix',
  'officeNumberingAlignment',
  'officeNumberingIndentLeft',
  'officeNumberingIndentRight',
  'officeNumberingIndentStart',
  'officeNumberingIndentEnd',
  'officeNumberingIndentHanging',
  'officeNumberingIndentFirstLine',
  'officeNumberingRestartAfterLevel',
  'level',
  'originalFormat',
  'originalSuffix',
] as const satisfies readonly (keyof DocumentNumberingChangeSnapshot)[];

const MAX_NUMBERING_SNAPSHOT_BYTES = 65_536;
const MAX_IDENTITY_LENGTH = 1_024;
const MAX_ORIGINAL_SUFFIX_LENGTH = 32;

export function serializeDocumentNumberingChange(
  attributes: Record<string, unknown>,
): string {
  return JSON.stringify(normalizeDocumentNumberingChange(attributes));
}

export function parseDocumentNumberingChange(
  value: unknown,
): DocumentNumberingChangeSnapshot | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_NUMBERING_SNAPSHOT_BYTES
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
    keys.length !== NUMBERING_SNAPSHOT_KEYS.length ||
    keys.some((key, index) => key !== NUMBERING_SNAPSHOT_KEYS[index])
  ) {
    return null;
  }
  const normalized = normalizeDocumentNumberingChange(parsed);
  return JSON.stringify(normalized) === value ? normalized : null;
}

export function restoredDocumentNumberingAttributes(
  attributes: Record<string, unknown>,
  serialized: unknown,
): Record<string, unknown> | null {
  const snapshot = parseDocumentNumberingChange(serialized);
  if (!snapshot) return null;
  return clearDocumentNumberingChangeAttributes({
    ...attributes,
    start: snapshot.start,
    type: snapshot.type,
    officeNumberingId: snapshot.officeNumberingId,
    officeAbstractNumberingId: snapshot.officeAbstractNumberingId,
    officeNumberingLevel: snapshot.officeNumberingLevel,
    officeNumberingFormat: snapshot.officeNumberingFormat,
    officeNumberingText: snapshot.officeNumberingText,
    officeNumberingSuffix: snapshot.officeNumberingSuffix,
    officeNumberingAlignment: snapshot.officeNumberingAlignment,
    officeNumberingIndentLeft: snapshot.officeNumberingIndentLeft,
    officeNumberingIndentRight: snapshot.officeNumberingIndentRight,
    officeNumberingIndentStart: snapshot.officeNumberingIndentStart,
    officeNumberingIndentEnd: snapshot.officeNumberingIndentEnd,
    officeNumberingIndentHanging: snapshot.officeNumberingIndentHanging,
    officeNumberingIndentFirstLine: snapshot.officeNumberingIndentFirstLine,
    officeNumberingRestartAfterLevel: snapshot.officeNumberingRestartAfterLevel,
  });
}

export function clearDocumentNumberingChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  return {
    ...attributes,
    numberingChangeKind: null,
    numberingChangeId: '',
    numberingChangeActorId: '',
    numberingChangeAuthor: '',
    numberingChangeDate: '',
    numberingChangeBefore: '',
  };
}

export function numberingFormatFromType(
  type: DocumentNumberingChangeType,
): number {
  if (type === 'I') return 1;
  if (type === 'i') return 2;
  if (type === 'A') return 3;
  if (type === 'a') return 4;
  return 0;
}

export function numberingTypeFromFormat(
  format: number,
): DocumentNumberingChangeType | undefined {
  if (format === 0) return null;
  if (format === 1) return 'I';
  if (format === 2) return 'i';
  if (format === 3) return 'A';
  if (format === 4) return 'a';
  return undefined;
}

function normalizeDocumentNumberingChange(
  source: Record<string, unknown>,
): DocumentNumberingChangeSnapshot {
  const type = numberingType(source.type);
  const level = boundedInteger(
    source.level ?? source.officeNumberingLevel,
    0,
    8,
    0,
  );
  const originalFormat = boundedInteger(
    source.originalFormat,
    0,
    4,
    numberingFormatFromType(type),
  );
  return {
    start: boundedInteger(source.start, 1, MAX_DOCUMENT_NUMBERING_START, 1),
    type,
    officeNumberingId: boundedString(source.officeNumberingId),
    officeAbstractNumberingId: boundedString(source.officeAbstractNumberingId),
    officeNumberingLevel: boundedString(source.officeNumberingLevel),
    officeNumberingFormat: boundedString(source.officeNumberingFormat),
    officeNumberingText: boundedString(source.officeNumberingText),
    officeNumberingSuffix: boundedString(source.officeNumberingSuffix),
    officeNumberingAlignment: boundedString(source.officeNumberingAlignment),
    officeNumberingIndentLeft: boundedString(source.officeNumberingIndentLeft),
    officeNumberingIndentRight: boundedString(
      source.officeNumberingIndentRight,
    ),
    officeNumberingIndentStart: boundedString(
      source.officeNumberingIndentStart,
    ),
    officeNumberingIndentEnd: boundedString(source.officeNumberingIndentEnd),
    officeNumberingIndentHanging: boundedString(
      source.officeNumberingIndentHanging,
    ),
    officeNumberingIndentFirstLine: boundedString(
      source.officeNumberingIndentFirstLine,
    ),
    officeNumberingRestartAfterLevel: boundedString(
      source.officeNumberingRestartAfterLevel,
    ),
    level,
    originalFormat,
    originalSuffix: originalSuffix(source.originalSuffix),
  };
}

function numberingType(value: unknown): DocumentNumberingChangeType {
  return value === 'a' || value === 'A' || value === 'i' || value === 'I'
    ? value
    : null;
}

function originalSuffix(value: unknown): string {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_ORIGINAL_SUFFIX_LENGTH ||
    value.includes('%') ||
    /[\u0000-\u001f\u007f]/.test(value)
  ) {
    return '.';
  }
  return value;
}

function boundedString(value: unknown): string | null {
  return typeof value === 'string' && value.length <= MAX_IDENTITY_LENGTH
    ? value || null
    : null;
}

function boundedInteger(
  value: unknown,
  minimum: number,
  maximum: number,
  fallback: number,
): number {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= minimum && number <= maximum
    ? number
    : fallback;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
