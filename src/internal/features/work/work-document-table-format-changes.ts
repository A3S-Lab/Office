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
  normalizeDocumentTableFormattingBorders,
  orderedDocumentTableFormattingBorders,
  type DocumentTableFormattingBorders,
} from './work-document-table-formatting-borders';
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
 * bidiVisual, fill, look, overlap, styleId, cellSpacing, colBandSize,
 * rowBandSize, borders, caption, or description must be present.
 */
export type DocumentTableOverlap = 'never' | 'overlap';

export interface DocumentTableFormattingSnapshot {
  layout?: DocumentTableLayoutAlgorithm;
  alignment?: DocumentTableAlignment;
  width?: DocumentTablePreferredWidth;
  indent?: number;
  cellMargins?: DocumentTableCellMarginOverrides;
  bidiVisual?: boolean;
  fill?: string;
  look?: DocumentTableLook;
  overlap?: DocumentTableOverlap;
  styleId?: string;
  cellSpacing?: number;
  colBandSize?: number;
  rowBandSize?: number;
  borders?: DocumentTableFormattingBorders;
  caption?: string;
  description?: string;
}

const MAX_TABLE_FORMAT_SNAPSHOT_BYTES = 4_096;
const MAX_TABLE_INDENT = 100_000;
const MARGIN_SIDES: readonly DocumentTableCellMarginSide[] = [
  'top',
  'right',
  'bottom',
  'left',
];

export function serializeDocumentTableFormatting(attributes: {
  layout?: unknown;
  alignment?: unknown;
  width?: unknown;
  indent?: unknown;
  cellMargins?: unknown;
  bidiVisual?: unknown;
  fill?: unknown;
  look?: unknown;
  overlap?: unknown;
  styleId?: unknown;
  cellSpacing?: unknown;
  colBandSize?: unknown;
  rowBandSize?: unknown;
  borders?: unknown;
  caption?: unknown;
  description?: unknown;
}): string {
  const snapshot = normalizeDocumentTableFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Table-formatting snapshot requires layout, alignment, width, indent, cellMargins, bidiVisual, fill, look, overlap, styleId, cellSpacing, colBandSize, rowBandSize, borders, caption, or description.',
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
        key !== 'look' &&
        key !== 'overlap' &&
        key !== 'styleId' &&
        key !== 'cellSpacing' &&
        key !== 'colBandSize' &&
        key !== 'rowBandSize' &&
        key !== 'borders' &&
        key !== 'caption' &&
        key !== 'description',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentTableFormattingSnapshot(record);
  if (!snapshot) return null;
  const normalized = JSON.stringify(orderedSnapshot(snapshot));
  return normalized === value ? snapshot : null;
}

export function normalizeDocumentTableFormattingSnapshot(attributes: {
  layout?: unknown;
  alignment?: unknown;
  width?: unknown;
  indent?: unknown;
  cellMargins?: unknown;
  bidiVisual?: unknown;
  fill?: unknown;
  look?: unknown;
  overlap?: unknown;
  styleId?: unknown;
  cellSpacing?: unknown;
  colBandSize?: unknown;
  rowBandSize?: unknown;
  borders?: unknown;
  caption?: unknown;
  description?: unknown;
}): DocumentTableFormattingSnapshot | null {
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
  if ('overlap' in attributes && attributes.overlap !== undefined) {
    const overlap = normalizeDocumentTableOverlap(attributes.overlap);
    if (!overlap) return null;
    snapshot.overlap = overlap;
  }
  if ('styleId' in attributes && attributes.styleId !== undefined) {
    const styleId = normalizeDocumentTableStyleId(attributes.styleId);
    if (!styleId) return null;
    snapshot.styleId = styleId;
  }
  if ('cellSpacing' in attributes && attributes.cellSpacing !== undefined) {
    const cellSpacing = normalizeDocumentTableCellSpacing(
      attributes.cellSpacing,
    );
    if (cellSpacing === null) return null;
    snapshot.cellSpacing = cellSpacing;
  }
  if ('colBandSize' in attributes && attributes.colBandSize !== undefined) {
    const colBandSize = normalizeDocumentTableColBandSize(
      attributes.colBandSize,
    );
    if (colBandSize === null) return null;
    snapshot.colBandSize = colBandSize;
  }
  if ('rowBandSize' in attributes && attributes.rowBandSize !== undefined) {
    const rowBandSize = normalizeDocumentTableRowBandSize(
      attributes.rowBandSize,
    );
    if (rowBandSize === null) return null;
    snapshot.rowBandSize = rowBandSize;
  }
  if ('borders' in attributes && attributes.borders !== undefined) {
    const borders = normalizeDocumentTableFormattingBorders(attributes.borders);
    if (!borders) return null;
    snapshot.borders = orderedDocumentTableFormattingBorders(borders);
  }
  if ('caption' in attributes && attributes.caption !== undefined) {
    const caption = normalizeDocumentTableCaption(attributes.caption);
    if (!caption) return null;
    snapshot.caption = caption;
  }
  if ('description' in attributes && attributes.description !== undefined) {
    const description = normalizeDocumentTableDescription(
      attributes.description,
    );
    if (!description) return null;
    snapshot.description = description;
  }
  return snapshot.layout ||
    snapshot.alignment ||
    snapshot.width ||
    snapshot.indent !== undefined ||
    snapshot.cellMargins ||
    snapshot.bidiVisual !== undefined ||
    snapshot.fill ||
    snapshot.look ||
    snapshot.overlap ||
    snapshot.styleId ||
    snapshot.cellSpacing !== undefined ||
    snapshot.colBandSize !== undefined ||
    snapshot.rowBandSize !== undefined ||
    snapshot.borders ||
    snapshot.caption ||
    snapshot.description
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
    const existing = normalizeDocumentTableCellMargins(
      geometry.cellMargins,
    ) ?? {
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
  if (formatting.overlap !== undefined) {
    next.overlap = formatting.overlap;
  }
  if (formatting.styleId !== undefined) {
    next.styleId = formatting.styleId;
  }
  if (formatting.cellSpacing !== undefined) {
    next.cellSpacing = formatting.cellSpacing;
  }
  if (formatting.colBandSize !== undefined) {
    next.colBandSize = formatting.colBandSize;
  }
  if (formatting.rowBandSize !== undefined) {
    next.rowBandSize = formatting.rowBandSize;
  }
  if (formatting.borders !== undefined) {
    next.borders = orderedDocumentTableFormattingBorders(formatting.borders);
  }
  if (formatting.caption !== undefined) {
    next.caption = formatting.caption;
  }
  if (formatting.description !== undefined) {
    next.description = formatting.description;
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
  if (snapshot.bidiVisual !== undefined)
    ordered.bidiVisual = snapshot.bidiVisual;
  if (snapshot.fill !== undefined) ordered.fill = snapshot.fill;
  if (snapshot.look) ordered.look = orderedDocumentTableLook(snapshot.look);
  if (snapshot.overlap) ordered.overlap = snapshot.overlap;
  if (snapshot.styleId) ordered.styleId = snapshot.styleId;
  if (snapshot.cellSpacing !== undefined)
    ordered.cellSpacing = snapshot.cellSpacing;
  if (snapshot.colBandSize !== undefined)
    ordered.colBandSize = snapshot.colBandSize;
  if (snapshot.rowBandSize !== undefined)
    ordered.rowBandSize = snapshot.rowBandSize;
  if (snapshot.borders) {
    ordered.borders = orderedDocumentTableFormattingBorders(snapshot.borders);
  }
  if (snapshot.caption) ordered.caption = snapshot.caption;
  if (snapshot.description) ordered.description = snapshot.description;
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

export function normalizeDocumentTableOverlap(
  value: unknown,
): DocumentTableOverlap | null {
  return value === 'never' || value === 'overlap' ? value : null;
}

const MAX_TABLE_STYLE_ID_LENGTH = 255;
const MAX_TABLE_CAPTION_LENGTH = 255;

export function normalizeDocumentTableCellSpacing(
  value: unknown,
): number | null {
  return normalizeTableIndent(value);
}

const MAX_TABLE_COL_BAND_SIZE = 64;

export function normalizeDocumentTableColBandSize(
  value: unknown,
): number | null {
  return normalizeDocumentTableBandSize(value);
}

export function normalizeDocumentTableRowBandSize(
  value: unknown,
): number | null {
  return normalizeDocumentTableBandSize(value);
}

function normalizeDocumentTableBandSize(value: unknown): number | null {
  const numeric =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value)
        : Number.NaN;
  if (
    !Number.isInteger(numeric) ||
    numeric <= 0 ||
    numeric > MAX_TABLE_COL_BAND_SIZE
  ) {
    return null;
  }
  return numeric;
}

export function normalizeDocumentTableStyleId(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const styleId = value.trim();
  if (
    !styleId ||
    styleId.length > MAX_TABLE_STYLE_ID_LENGTH ||
    /[\u0000-\u001f\u007f]/.test(styleId)
  ) {
    return null;
  }
  return styleId;
}

export function normalizeDocumentTableCaption(value: unknown): string | null {
  return normalizeDocumentTableStringProperty(value);
}

export function normalizeDocumentTableDescription(
  value: unknown,
): string | null {
  return normalizeDocumentTableStringProperty(value);
}

function normalizeDocumentTableStringProperty(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim();
  if (
    !normalized ||
    normalized.length > MAX_TABLE_CAPTION_LENGTH ||
    /[\u0000-\u001f\u007f]/.test(normalized)
  ) {
    return null;
  }
  return normalized;
}
