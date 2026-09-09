import {
  normalizeDocumentTableAlignment,
  normalizeDocumentTableCellMarginOverrides,
  normalizeDocumentTableCellMargins,
  normalizeDocumentTableLayoutAlgorithm,
  normalizeDocumentTablePreferredWidth,
  type DocumentTableAlignment,
  type DocumentTableCellMarginOverrides,
  type DocumentTableCellMarginSide,
  type DocumentTableLayoutAlgorithm,
  type DocumentTablePreferredWidth,
  DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
} from './work-document-table-geometry';
import { normalizeTableColor } from './work-document-table-borders';
import {
  normalizeDocumentTableLook,
  orderedDocumentTableLook,
  type DocumentTableLook,
} from './work-document-table-look';

export const DOCUMENT_TABLE_CHANGE_ATTRIBUTES = [
  'tableChangeKind',
  'tableChangeId',
  'tableChangeAuthor',
  'tableChangeDate',
  'tableChangeBefore',
] as const;

/**
 * Prior snapshot for reviewable table-property revisions.
 * At least one of layout, alignment, preferred width, indent, cellMargins,
 * bidiVisual, fill, or look must be present.
 */
export interface DocumentTableFormattingSnapshot {
  layout?: DocumentTableLayoutAlgorithm;
  alignment?: DocumentTableAlignment;
  width?: DocumentTablePreferredWidth;
  indent?: number;
  cellMargins?: DocumentTableCellMarginOverrides;
  bidiVisual?: boolean;
  fill?: string;
  look?: DocumentTableLook;
}

const MAX_TABLE_FORMAT_SNAPSHOT_BYTES = 4_096;
const MAX_TABLE_INDENT = 100_000;
const MARGIN_SIDES: readonly DocumentTableCellMarginSide[] = [
  'top',
  'right',
  'bottom',
  'left',
];

export function serializeDocumentTableFormatting(
  attributes: {
    layout?: unknown;
    alignment?: unknown;
    width?: unknown;
    indent?: unknown;
    cellMargins?: unknown;
    bidiVisual?: unknown;
    fill?: unknown;
    look?: unknown;
  },
): string {
  const snapshot = normalizeDocumentTableFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Table-formatting snapshot requires layout, alignment, width, indent, cellMargins, bidiVisual, fill, or look.',
    );
  }
  return JSON.stringify(orderedSnapshot(snapshot));
}

export function parseDocumentTableFormatting(
  value: unknown,
): DocumentTableFormattingSnapshot | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_TABLE_FORMAT_SNAPSHOT_BYTES
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
        key !== 'layout' &&
        key !== 'alignment' &&
        key !== 'width' &&
        key !== 'indent' &&
        key !== 'cellMargins' &&
        key !== 'bidiVisual' &&
        key !== 'fill' &&
        key !== 'look',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentTableFormattingSnapshot(record);
  if (!snapshot) return null;
  const normalized = JSON.stringify(orderedSnapshot(snapshot));
  return normalized === value ? snapshot : null;
}

export function normalizeDocumentTableFormattingSnapshot(
  attributes: {
    layout?: unknown;
    alignment?: unknown;
    width?: unknown;
    indent?: unknown;
    cellMargins?: unknown;
    bidiVisual?: unknown;
    fill?: unknown;
    look?: unknown;
  },
): DocumentTableFormattingSnapshot | null {
  const snapshot: DocumentTableFormattingSnapshot = {};
  if ('layout' in attributes && attributes.layout !== undefined) {
    const layout = normalizeDocumentTableLayoutAlgorithm(attributes.layout);
    if (!layout) return null;
    snapshot.layout = layout;
  }
  if ('alignment' in attributes && attributes.alignment !== undefined) {
    const alignment = normalizeDocumentTableAlignment(attributes.alignment);
    if (!alignment) return null;
    snapshot.alignment = alignment;
  }
  if ('width' in attributes && attributes.width !== undefined) {
    const width = normalizeDocumentTablePreferredWidth(attributes.width);
    if (!width) return null;
    snapshot.width = width;
  }
  if ('indent' in attributes && attributes.indent !== undefined) {
    const indent = normalizeTableIndent(attributes.indent);
    if (indent === null) return null;
    snapshot.indent = indent;
  }
  if ('cellMargins' in attributes && attributes.cellMargins !== undefined) {
    const cellMargins = normalizeDocumentTableCellMarginOverrides(
      attributes.cellMargins,
    );
    if (!cellMargins) return null;
    snapshot.cellMargins = orderedMargins(cellMargins);
  }
  if ('bidiVisual' in attributes && attributes.bidiVisual !== undefined) {
    if (typeof attributes.bidiVisual !== 'boolean') return null;
    snapshot.bidiVisual = attributes.bidiVisual;
  }
  if ('fill' in attributes && attributes.fill !== undefined) {
    const fill = normalizeTableFill(attributes.fill);
    if (!fill) return null;
    snapshot.fill = fill;
  }
  if ('look' in attributes && attributes.look !== undefined) {
    const look = normalizeDocumentTableLook(attributes.look);
    if (!look) return null;
    snapshot.look = orderedDocumentTableLook(look);
  }
  return snapshot.layout ||
    snapshot.alignment ||
    snapshot.width ||
    snapshot.indent !== undefined ||
    snapshot.cellMargins ||
    snapshot.bidiVisual !== undefined ||
    snapshot.fill ||
    snapshot.look
    ? snapshot
    : null;
}

export function clearDocumentTableChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  const next = { ...attributes };
  for (const key of DOCUMENT_TABLE_CHANGE_ATTRIBUTES) {
    next[key] = null;
  }
  return next;
}

export function restoredDocumentTableAttributes(
  attributes: Record<string, unknown>,
  serialized: unknown,
): Record<string, unknown> | null {
  const formatting = parseDocumentTableFormatting(serialized);
  if (!formatting) return null;
  const geometry =
    attributes.geometry &&
    typeof attributes.geometry === 'object' &&
    !Array.isArray(attributes.geometry)
      ? { ...(attributes.geometry as Record<string, unknown>) }
      : {};
  if (formatting.layout) geometry.layout = formatting.layout;
  if (formatting.alignment) geometry.alignment = formatting.alignment;
  if (formatting.width) geometry.width = formatting.width;
  if (formatting.indent !== undefined) geometry.indent = formatting.indent;
  if (formatting.cellMargins) {
    const existing =
      normalizeDocumentTableCellMargins(geometry.cellMargins) ?? {
        ...DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
      };
    const merged = normalizeDocumentTableCellMargins({
      ...existing,
      ...formatting.cellMargins,
    });
    if (!merged) return null;
    geometry.cellMargins = merged;
  }
  const next = clearDocumentTableChangeAttributes({
    ...attributes,
    geometry,
  });
  if (formatting.bidiVisual !== undefined) {
    next.bidiVisual = formatting.bidiVisual;
  }
  if (formatting.fill !== undefined) {
    next.fill = formatting.fill;
  }
  if (formatting.look !== undefined) {
    next.look = orderedDocumentTableLook(formatting.look);
  }
  return next;
}

function orderedSnapshot(
  snapshot: DocumentTableFormattingSnapshot,
): Record<string, unknown> {
  const ordered: Record<string, unknown> = {};
  if (snapshot.layout) ordered.layout = snapshot.layout;
  if (snapshot.alignment) ordered.alignment = snapshot.alignment;
  if (snapshot.width) {
    ordered.width = {
      type: snapshot.width.type,
      value: snapshot.width.value,
    };
  }
  if (snapshot.indent !== undefined) ordered.indent = snapshot.indent;
  if (snapshot.cellMargins) {
    ordered.cellMargins = orderedMargins(snapshot.cellMargins);
  }
  if (snapshot.bidiVisual !== undefined) ordered.bidiVisual = snapshot.bidiVisual;
  if (snapshot.fill !== undefined) ordered.fill = snapshot.fill;
  if (snapshot.look) ordered.look = orderedDocumentTableLook(snapshot.look);
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

function normalizeTableIndent(value: unknown): number | null {
  const numeric =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value)
        : Number.NaN;
  if (!Number.isFinite(numeric) || numeric < 0 || numeric > MAX_TABLE_INDENT) {
    return null;
  }
  return Math.round(numeric * 100) / 100;
}

function normalizeTableFill(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  return normalizeTableColor(value);
}
