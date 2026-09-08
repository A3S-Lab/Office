import { normalizeTableColor } from './work-document-table-borders';
import {
  normalizeDocumentTableVerticalAlign,
  normalizeDocumentTableCellTextDirection,
  type DocumentTableCellTextDirection,
  type DocumentTableVerticalAlign,
} from './work-document-table-cell-formatting';
import {
  normalizeDocumentTableCellMarginOverrides,
  normalizeDocumentTablePreferredWidth,
  type DocumentTableCellMarginOverrides,
  type DocumentTableCellMarginSide,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';

export const DOCUMENT_CELL_CHANGE_ATTRIBUTES = [
  'cellChangeKind',
  'cellChangeId',
  'cellChangeAuthor',
  'cellChangeDate',
  'cellChangeBefore',
] as const;

export type { DocumentTableCellTextDirection } from './work-document-table-cell-formatting';
export { normalizeDocumentTableCellTextDirection } from './work-document-table-cell-formatting';

/**
 * Prior snapshot for reviewable cell-property revisions.
 * At least one of verticalAlign, fill, margins, width, noWrap, textDirection,
 * or fitText must be present.
 */
export interface DocumentCellFormattingSnapshot {
  verticalAlign?: DocumentTableVerticalAlign;
  fill?: string;
  margins?: DocumentTableCellMarginOverrides;
  width?: DocumentTablePreferredWidth;
  noWrap?: boolean;
  textDirection?: DocumentTableCellTextDirection;
  fitText?: boolean;
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
    width?: unknown;
    noWrap?: unknown;
    textDirection?: unknown;
    fitText?: unknown;
  },
): string {
  const snapshot = normalizeDocumentCellFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Cell-formatting snapshot requires verticalAlign, fill, margins, width, noWrap, textDirection, or fitText.',
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
      (key) =>
        key !== 'verticalAlign' &&
        key !== 'fill' &&
        key !== 'margins' &&
        key !== 'width' &&
        key !== 'noWrap' &&
        key !== 'textDirection' &&
        key !== 'fitText',
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
    width?: unknown;
    noWrap?: unknown;
    textDirection?: unknown;
    fitText?: unknown;
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
  if ('width' in attributes && attributes.width !== undefined) {
    const width = normalizeDocumentTablePreferredWidth(attributes.width);
    if (!width) return null;
    snapshot.width = orderedWidth(width);
  }
  if ('noWrap' in attributes && attributes.noWrap !== undefined) {
    if (typeof attributes.noWrap !== 'boolean') return null;
    snapshot.noWrap = attributes.noWrap;
  }
  if ('textDirection' in attributes && attributes.textDirection !== undefined) {
    const textDirection = normalizeDocumentTableCellTextDirection(
      attributes.textDirection,
    );
    if (!textDirection) return null;
    snapshot.textDirection = textDirection;
  }
  if ('fitText' in attributes && attributes.fitText !== undefined) {
    if (typeof attributes.fitText !== 'boolean') return null;
    snapshot.fitText = attributes.fitText;
  }
  return snapshot.verticalAlign ||
    snapshot.fill ||
    snapshot.margins ||
    snapshot.width ||
    snapshot.noWrap !== undefined ||
    snapshot.textDirection !== undefined ||
    snapshot.fitText !== undefined
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
    ...(formatting.width !== undefined
      ? restoredCellWidthAttributes(attributes, formatting.width)
      : {}),
    ...(formatting.noWrap !== undefined ? { noWrap: formatting.noWrap } : {}),
    ...(formatting.textDirection !== undefined
      ? { textDirection: formatting.textDirection }
      : {}),
    ...(formatting.fitText !== undefined
      ? { fitText: formatting.fitText }
      : {}),
  });
}

export function preferredWidthFromCellAttributes(
  attributes: Record<string, unknown>,
): DocumentTablePreferredWidth | null {
  const percentages = normalizedPositiveNumberList(
    attributes.columnWidthPercentages,
  );
  if (percentages) {
    const total = percentages.reduce((sum, value) => sum + value, 0);
    return normalizeDocumentTablePreferredWidth({
      type: 'percent',
      value: Math.round(total * 100) / 100,
    });
  }
  const pixels = normalizedPositiveNumberList(attributes.colwidth);
  if (pixels) {
    const total = pixels.reduce((sum, value) => sum + value, 0);
    return normalizeDocumentTablePreferredWidth({
      type: 'pixels',
      value: Math.round(total * 100) / 100,
    });
  }
  return normalizeDocumentTablePreferredWidth({ type: 'auto', value: null });
}

function restoredCellWidthAttributes(
  attributes: Record<string, unknown>,
  width: DocumentTablePreferredWidth,
): Record<string, unknown> {
  const span = Math.max(
    1,
    Number(attributes.colspan) ||
      (Array.isArray(attributes.colwidth) ? attributes.colwidth.length : 1) ||
      1,
  );
  if (width.type === 'auto') {
    return { colwidth: null, columnWidthPercentages: null };
  }
  if (width.type === 'percent') {
    const share = Math.round(((width.value ?? 0) / span) * 100) / 100;
    return {
      colwidth: null,
      columnWidthPercentages: Array.from({ length: span }, () => share),
    };
  }
  const share = Math.round(((width.value ?? 0) / span) * 100) / 100;
  return {
    colwidth: Array.from({ length: span }, () => share),
    columnWidthPercentages: null,
  };
}

function normalizedPositiveNumberList(value: unknown): number[] | null {
  if (!Array.isArray(value) || !value.length) return null;
  const numbers: number[] = [];
  for (const entry of value) {
    const numeric = typeof entry === 'number' ? entry : Number(entry);
    if (!Number.isFinite(numeric) || numeric <= 0) return null;
    numbers.push(Math.round(numeric * 100) / 100);
  }
  return numbers;
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
  if (snapshot.width !== undefined) ordered.width = orderedWidth(snapshot.width);
  if (snapshot.noWrap !== undefined) ordered.noWrap = snapshot.noWrap;
  if (snapshot.textDirection !== undefined) {
    ordered.textDirection = snapshot.textDirection;
  }
  if (snapshot.fitText !== undefined) ordered.fitText = snapshot.fitText;
  return ordered;
}

function orderedWidth(
  width: DocumentTablePreferredWidth,
): DocumentTablePreferredWidth {
  return width.type === 'auto'
    ? { type: 'auto', value: null }
    : { type: width.type, value: width.value };
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
