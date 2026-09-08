import { normalizeTableColor } from './work-document-table-borders';
import {
  normalizeDocumentTableVerticalAlign,
  type DocumentTableVerticalAlign,
} from './work-document-table-cell-formatting';
import {
  normalizeDocumentTableCellMarginOverrides,
  type DocumentTableCellMarginOverrides,
  type DocumentTableCellMarginSide,
} from './work-document-table-geometry';

export const DOCUMENT_CELL_CHANGE_ATTRIBUTES = [
  'cellChangeKind',
  'cellChangeId',
  'cellChangeAuthor',
  'cellChangeDate',
  'cellChangeBefore',
] as const;

/**
 * Prior snapshot for reviewable cell-property revisions.
 * At least one of verticalAlign, fill, or margins must be present.
 */
export interface DocumentCellFormattingSnapshot {
  verticalAlign?: DocumentTableVerticalAlign;
  fill?: string;
  margins?: DocumentTableCellMarginOverrides;
}

const MAX_CELL_FORMAT_SNAPSHOT_BYTES = 4_096;
const MARGIN_SIDES: readonly DocumentTableCellMarginSide[] = [
  'top',
  'right',
  'bottom',
  'left',
];

export function serializeDocumentCellFormatting(
  attributes: {
    verticalAlign?: unknown;
    fill?: unknown;
    margins?: unknown;
  },
): string {
  const snapshot = normalizeDocumentCellFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Cell-formatting snapshot requires verticalAlign, fill, or margins.',
    );
  }
  return JSON.stringify(orderedSnapshot(snapshot));
}

export function parseDocumentCellFormatting(
  value: unknown,
): DocumentCellFormattingSnapshot | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_CELL_FORMAT_SNAPSHOT_BYTES
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
      (key) => key !== 'verticalAlign' && key !== 'fill' && key !== 'margins',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentCellFormattingSnapshot(record);
  if (!snapshot) return null;
  return JSON.stringify(orderedSnapshot(snapshot)) === value ? snapshot : null;
}

export function normalizeDocumentCellFormattingSnapshot(
  attributes: {
    verticalAlign?: unknown;
    fill?: unknown;
    margins?: unknown;
  },
): DocumentCellFormattingSnapshot | null {
  const snapshot: DocumentCellFormattingSnapshot = {};
  if ('verticalAlign' in attributes && attributes.verticalAlign !== undefined) {
    const verticalAlign = normalizeDocumentTableVerticalAlign(
      typeof attributes.verticalAlign === 'string'
        ? attributes.verticalAlign
        : null,
    );
    if (!verticalAlign) return null;
    snapshot.verticalAlign = verticalAlign;
  }
  if ('fill' in attributes && attributes.fill !== undefined) {
    const fill = normalizeTableColor(
      typeof attributes.fill === 'string' ? attributes.fill : null,
    );
    if (!fill) return null;
    snapshot.fill = fill;
  }
  if ('margins' in attributes && attributes.margins !== undefined) {
    const margins = normalizeDocumentTableCellMarginOverrides(
      attributes.margins,
    );
    if (!margins) return null;
    snapshot.margins = orderedMargins(margins);
  }
  return snapshot.verticalAlign || snapshot.fill || snapshot.margins
    ? snapshot
    : null;
}

export function clearDocumentCellChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  const next = { ...attributes };
  for (const key of DOCUMENT_CELL_CHANGE_ATTRIBUTES) {
    next[key] = null;
  }
  return next;
}

export function restoredDocumentCellAttributes(
  attributes: Record<string, unknown>,
  serialized: unknown,
): Record<string, unknown> | null {
  const formatting = parseDocumentCellFormatting(serialized);
  if (!formatting) return null;
  return clearDocumentCellChangeAttributes({
    ...attributes,
    ...(formatting.verticalAlign !== undefined
      ? { verticalAlign: formatting.verticalAlign }
      : {}),
    ...(formatting.fill !== undefined
      ? { backgroundColor: formatting.fill }
      : {}),
    ...(formatting.margins !== undefined
      ? { margins: formatting.margins }
      : {}),
  });
}

function orderedSnapshot(
  snapshot: DocumentCellFormattingSnapshot,
): DocumentCellFormattingSnapshot {
  const ordered: DocumentCellFormattingSnapshot = {};
  if (snapshot.verticalAlign !== undefined) {
    ordered.verticalAlign = snapshot.verticalAlign;
  }
  if (snapshot.fill !== undefined) ordered.fill = snapshot.fill;
  if (snapshot.margins !== undefined) {
    ordered.margins = orderedMargins(snapshot.margins);
  }
  return ordered;
}

function orderedMargins(
  margins: DocumentTableCellMarginOverrides,
): DocumentTableCellMarginOverrides {
  const ordered: DocumentTableCellMarginOverrides = {};
  for (const side of MARGIN_SIDES) {
    if (margins[side] !== undefined) ordered[side] = margins[side];
  }
  return ordered;
}
