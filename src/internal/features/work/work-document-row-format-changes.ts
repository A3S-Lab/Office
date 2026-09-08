import {
  normalizeDocumentTableAlignment,
  type DocumentTableAlignment,
} from './work-document-table-geometry';
import {
  normalizeDocumentTableRowHeight,
  normalizeDocumentTableRowHeightRule,
  type DocumentTableRowHeightRule,
} from './work-document-table-row';

export const DOCUMENT_ROW_CHANGE_ATTRIBUTES = [
  'rowChangeKind',
  'rowChangeId',
  'rowChangeAuthor',
  'rowChangeDate',
  'rowChangeBefore',
] as const;

/**
 * Prior snapshot for reviewable row-property revisions.
 * At least one of cantSplit, repeatHeader, height, hidden, or alignment must
 * be present.
 */
export interface DocumentRowFormattingSnapshot {
  cantSplit?: boolean;
  repeatHeader?: boolean;
  height?: DocumentRowFormattingHeight;
  hidden?: boolean;
  alignment?: DocumentTableAlignment;
}

export interface DocumentRowFormattingHeight {
  value: number;
  rule: DocumentTableRowHeightRule;
}

const MAX_ROW_FORMAT_SNAPSHOT_BYTES = 4_096;

export function serializeDocumentRowFormatting(
  attributes: {
    cantSplit?: unknown;
    repeatHeader?: unknown;
    height?: unknown;
    hidden?: unknown;
    alignment?: unknown;
  },
): string {
  const snapshot = normalizeDocumentRowFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Row-formatting snapshot requires cantSplit, repeatHeader, height, hidden, or alignment.',
    );
  }
  return JSON.stringify(orderedSnapshot(snapshot));
}

export function parseDocumentRowFormatting(
  value: unknown,
): DocumentRowFormattingSnapshot | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_ROW_FORMAT_SNAPSHOT_BYTES
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return null;
  }
  const record = parsed as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some(
      (key) =>
        key !== 'cantSplit' &&
        key !== 'repeatHeader' &&
        key !== 'height' &&
        key !== 'hidden' &&
        key !== 'alignment',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentRowFormattingSnapshot(record);
  if (!snapshot) return null;
  return JSON.stringify(orderedSnapshot(snapshot)) === value ? snapshot : null;
}

export function normalizeDocumentRowFormattingSnapshot(
  attributes: {
    cantSplit?: unknown;
    repeatHeader?: unknown;
    height?: unknown;
    hidden?: unknown;
    alignment?: unknown;
  },
): DocumentRowFormattingSnapshot | null {
  const snapshot: DocumentRowFormattingSnapshot = {};
  if ('cantSplit' in attributes && attributes.cantSplit !== undefined) {
    if (typeof attributes.cantSplit !== 'boolean') return null;
    snapshot.cantSplit = attributes.cantSplit;
  }
  if ('repeatHeader' in attributes && attributes.repeatHeader !== undefined) {
    if (typeof attributes.repeatHeader !== 'boolean') return null;
    snapshot.repeatHeader = attributes.repeatHeader;
  }
  if ('height' in attributes && attributes.height !== undefined) {
    const height = normalizeDocumentRowFormattingHeight(attributes.height);
    if (!height) return null;
    snapshot.height = height;
  }
  if ('hidden' in attributes && attributes.hidden !== undefined) {
    if (typeof attributes.hidden !== 'boolean') return null;
    snapshot.hidden = attributes.hidden;
  }
  if ('alignment' in attributes && attributes.alignment !== undefined) {
    const alignment = normalizeDocumentTableAlignment(attributes.alignment);
    if (!alignment) return null;
    snapshot.alignment = alignment;
  }
  return snapshot.cantSplit !== undefined ||
    snapshot.repeatHeader !== undefined ||
    snapshot.height !== undefined ||
    snapshot.hidden !== undefined ||
    snapshot.alignment !== undefined
    ? snapshot
    : null;
}

export function clearDocumentRowChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  const next = { ...attributes };
  for (const key of DOCUMENT_ROW_CHANGE_ATTRIBUTES) {
    next[key] = null;
  }
  return next;
}

export function restoredDocumentRowAttributes(
  attributes: Record<string, unknown>,
  serialized: unknown,
): Record<string, unknown> | null {
  const formatting = parseDocumentRowFormatting(serialized);
  if (!formatting) return null;
  return clearDocumentRowChangeAttributes({
    ...attributes,
    ...(formatting.cantSplit !== undefined
      ? { cantSplit: formatting.cantSplit }
      : {}),
    ...(formatting.repeatHeader !== undefined
      ? { repeatHeader: formatting.repeatHeader }
      : {}),
    ...(formatting.height
      ? {
          rowHeight: formatting.height.value,
          rowHeightRule: formatting.height.rule,
        }
      : {}),
    ...(formatting.hidden !== undefined ? { hidden: formatting.hidden } : {}),
    ...(formatting.alignment !== undefined
      ? { alignment: formatting.alignment }
      : {}),
  });
}

export function normalizeDocumentRowFormattingHeight(
  value: unknown,
): DocumentRowFormattingHeight | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    keys.length !== 2 ||
    !keys.includes('value') ||
    !keys.includes('rule')
  ) {
    return null;
  }
  const height = normalizeDocumentTableRowHeight(record.value);
  const rule = normalizeDocumentTableRowHeightRule(record.rule);
  if (height === null || !rule) return null;
  return { value: height, rule };
}

function orderedSnapshot(
  snapshot: DocumentRowFormattingSnapshot,
): DocumentRowFormattingSnapshot {
  const ordered: DocumentRowFormattingSnapshot = {};
  if (snapshot.cantSplit !== undefined) ordered.cantSplit = snapshot.cantSplit;
  if (snapshot.repeatHeader !== undefined) {
    ordered.repeatHeader = snapshot.repeatHeader;
  }
  if (snapshot.height) {
    ordered.height = {
      value: snapshot.height.value,
      rule: snapshot.height.rule,
    };
  }
  if (snapshot.hidden !== undefined) ordered.hidden = snapshot.hidden;
  if (snapshot.alignment !== undefined) ordered.alignment = snapshot.alignment;
  return ordered;
}
