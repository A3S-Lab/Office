import {
  normalizeDocumentTableAlignment,
  normalizeDocumentTablePreferredWidth,
  type DocumentTableAlignment,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';
import { normalizeDocumentRowCnfStyle } from './work-document-row-cnf-style';
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
 * At least one of cantSplit, repeatHeader, height, hidden, alignment,
 * gridBefore, gridAfter, widthBefore, widthAfter, cnfStyle, divId, or tblCellSpacing must be present.
 */
export interface DocumentRowFormattingSnapshot {
  cantSplit?: boolean;
  repeatHeader?: boolean;
  height?: DocumentRowFormattingHeight;
  hidden?: boolean;
  alignment?: DocumentTableAlignment;
  gridBefore?: number;
  gridAfter?: number;
  widthBefore?: DocumentTablePreferredWidth;
  widthAfter?: DocumentTablePreferredWidth;
  cnfStyle?: string;
  divId?: number;
  tblCellSpacing?: DocumentTablePreferredWidth;
}

export interface DocumentRowFormattingHeight {
  value: number;
  rule: DocumentTableRowHeightRule;
}

const MAX_ROW_FORMAT_SNAPSHOT_BYTES = 4_096;

export function serializeDocumentRowFormatting(attributes: {
  cantSplit?: unknown;
  repeatHeader?: unknown;
  height?: unknown;
  hidden?: unknown;
  alignment?: unknown;
  gridBefore?: unknown;
  gridAfter?: unknown;
  widthBefore?: unknown;
  widthAfter?: unknown;
  cnfStyle?: unknown;
  divId?: unknown;
  tblCellSpacing?: unknown;
}): string {
  const snapshot = normalizeDocumentRowFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Row-formatting snapshot requires cantSplit, repeatHeader, height, hidden, alignment, gridBefore, gridAfter, widthBefore, widthAfter, cnfStyle, divId, or tblCellSpacing.',
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
        key !== 'alignment' &&
        key !== 'gridBefore' &&
        key !== 'gridAfter' &&
        key !== 'widthBefore' &&
        key !== 'widthAfter' &&
        key !== 'cnfStyle' &&
        key !== 'divId' &&
        key !== 'tblCellSpacing',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentRowFormattingSnapshot(record);
  if (!snapshot) return null;
  return JSON.stringify(orderedSnapshot(snapshot)) === value ? snapshot : null;
}

export function normalizeDocumentRowFormattingSnapshot(attributes: {
  cantSplit?: unknown;
  repeatHeader?: unknown;
  height?: unknown;
  hidden?: unknown;
  alignment?: unknown;
  gridBefore?: unknown;
  gridAfter?: unknown;
  widthBefore?: unknown;
  widthAfter?: unknown;
  cnfStyle?: unknown;
  divId?: unknown;
  tblCellSpacing?: unknown;
}): DocumentRowFormattingSnapshot | null {
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
  if ('gridBefore' in attributes && attributes.gridBefore !== undefined) {
    const gridBefore = normalizeDocumentRowGridSpan(attributes.gridBefore);
    if (gridBefore === null) return null;
    snapshot.gridBefore = gridBefore;
  }
  if ('gridAfter' in attributes && attributes.gridAfter !== undefined) {
    const gridAfter = normalizeDocumentRowGridSpan(attributes.gridAfter);
    if (gridAfter === null) return null;
    snapshot.gridAfter = gridAfter;
  }
  if ('widthBefore' in attributes && attributes.widthBefore !== undefined) {
    const widthBefore = normalizeDocumentTablePreferredWidth(
      attributes.widthBefore,
    );
    if (!widthBefore) return null;
    snapshot.widthBefore = widthBefore;
  }
  if ('widthAfter' in attributes && attributes.widthAfter !== undefined) {
    const widthAfter = normalizeDocumentTablePreferredWidth(
      attributes.widthAfter,
    );
    if (!widthAfter) return null;
    snapshot.widthAfter = widthAfter;
  }
  if ('cnfStyle' in attributes && attributes.cnfStyle !== undefined) {
    const cnfStyle = normalizeDocumentRowCnfStyle(attributes.cnfStyle);
    if (!cnfStyle) return null;
    snapshot.cnfStyle = cnfStyle;
  }
  if ('divId' in attributes && attributes.divId !== undefined) {
    const divId = normalizeDocumentRowDivId(attributes.divId);
    if (divId === null) return null;
    snapshot.divId = divId;
  }
  if (
    'tblCellSpacing' in attributes &&
    attributes.tblCellSpacing !== undefined
  ) {
    const tblCellSpacing = normalizeDocumentTablePreferredWidth(
      attributes.tblCellSpacing,
    );
    if (!tblCellSpacing) return null;
    snapshot.tblCellSpacing = tblCellSpacing;
  }
  return snapshot.cantSplit !== undefined ||
    snapshot.repeatHeader !== undefined ||
    snapshot.height !== undefined ||
    snapshot.hidden !== undefined ||
    snapshot.alignment !== undefined ||
    snapshot.gridBefore !== undefined ||
    snapshot.gridAfter !== undefined ||
    snapshot.widthBefore !== undefined ||
    snapshot.widthAfter !== undefined ||
    snapshot.cnfStyle !== undefined ||
    snapshot.divId !== undefined ||
    snapshot.tblCellSpacing !== undefined
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
    ...(formatting.gridBefore !== undefined
      ? { gridBefore: formatting.gridBefore }
      : {}),
    ...(formatting.gridAfter !== undefined
      ? { gridAfter: formatting.gridAfter }
      : {}),
    ...(formatting.widthBefore !== undefined
      ? { widthBefore: formatting.widthBefore }
      : {}),
    ...(formatting.widthAfter !== undefined
      ? { widthAfter: formatting.widthAfter }
      : {}),
    ...(formatting.cnfStyle !== undefined
      ? { cnfStyle: formatting.cnfStyle }
      : {}),
    ...(formatting.divId !== undefined ? { divId: formatting.divId } : {}),
    ...(formatting.tblCellSpacing !== undefined
      ? { tblCellSpacing: formatting.tblCellSpacing }
      : {}),
  });
}

export function normalizeDocumentRowFormattingHeight(
  value: unknown,
): DocumentRowFormattingHeight | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length !== 2 || !keys.includes('value') || !keys.includes('rule')) {
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
  if (snapshot.gridBefore !== undefined)
    ordered.gridBefore = snapshot.gridBefore;
  if (snapshot.gridAfter !== undefined) ordered.gridAfter = snapshot.gridAfter;
  if (snapshot.widthBefore) {
    ordered.widthBefore = {
      type: snapshot.widthBefore.type,
      value: snapshot.widthBefore.value,
    };
  }
  if (snapshot.widthAfter) {
    ordered.widthAfter = {
      type: snapshot.widthAfter.type,
      value: snapshot.widthAfter.value,
    };
  }
  if (snapshot.cnfStyle !== undefined) ordered.cnfStyle = snapshot.cnfStyle;
  if (snapshot.divId !== undefined) ordered.divId = snapshot.divId;
  if (snapshot.tblCellSpacing) {
    ordered.tblCellSpacing = {
      type: snapshot.tblCellSpacing.type,
      value: snapshot.tblCellSpacing.value,
    };
  }
  return ordered;
}

export function normalizeDocumentRowGridSpan(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
    return null;
  }
  return value;
}

/** Non-negative safe-integer HTML div association for w:divId / data-office-row-div-id. */
export function normalizeDocumentRowDivId(value: unknown): number | null {
  if (typeof value === 'number') {
    return Number.isInteger(value) && Number.isSafeInteger(value) && value >= 0
      ? value
      : null;
  }
  if (typeof value === 'string') {
    const normalized = value.trim();
    if (!normalized || !/^[0-9]+$/.test(normalized)) return null;
    const numeric = Number(normalized);
    return Number.isInteger(numeric) &&
      Number.isSafeInteger(numeric) &&
      numeric >= 0
      ? numeric
      : null;
  }
  return null;
}

/** Parse w:divId: missing/negative/non-integer/malformed fail closed. */
export function parseDocxRowDivIdValue(
  value: string | null | undefined,
): number | null {
  if (value === null || value === undefined) return null;
  return normalizeDocumentRowDivId(value);
}

/** @deprecated Prefer normalizeDocumentRowGridSpan */
export function normalizeDocumentRowGridBefore(value: unknown): number | null {
  return normalizeDocumentRowGridSpan(value);
}

export function normalizeDocumentRowGridAfter(value: unknown): number | null {
  return normalizeDocumentRowGridSpan(value);
}
