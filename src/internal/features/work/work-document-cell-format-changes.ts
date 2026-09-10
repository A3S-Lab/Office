import { normalizeDocumentCnfStyle } from './work-document-cnf-style';
import {
  normalizeTableColor,
  type DocumentTableBorder,
  type DocumentTableCellBorders,
} from './work-document-table-borders';
import {
  normalizeDocumentCellFormattingBorders,
  orderedDocumentCellFormattingBorders,
  type DocumentCellFormattingBorders,
} from './work-document-table-formatting-borders';
import {
  normalizeDocumentTableVerticalAlign,
  normalizeDocumentTableCellTextDirection,
  normalizeDocumentTableCellHMerge,
  normalizeDocumentTableCellVMerge,
  normalizeDocumentTableCellGridSpan,
  type DocumentTableCellHMerge,
  type DocumentTableCellVMerge,
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

export type {
  DocumentTableCellHMerge,
  DocumentTableCellVMerge,
  DocumentTableCellTextDirection,
} from './work-document-table-cell-formatting';
export {
  normalizeDocumentTableCellHMerge,
  normalizeDocumentTableCellVMerge,
  normalizeDocumentTableCellGridSpan,
  normalizeDocumentTableCellTextDirection,
} from './work-document-table-cell-formatting';

/**
 * Prior snapshot for reviewable cell-property revisions.
 * At least one of verticalAlign, fill, margins, width, noWrap, textDirection,
 * fitText, hideMark, cnfStyle, hMerge, vMerge, gridSpan, or borders must be present.
 */
export interface DocumentCellFormattingSnapshot {
  verticalAlign?: DocumentTableVerticalAlign;
  fill?: string;
  margins?: DocumentTableCellMarginOverrides;
  width?: DocumentTablePreferredWidth;
  noWrap?: boolean;
  textDirection?: DocumentTableCellTextDirection;
  fitText?: boolean;
  hideMark?: boolean;
  cnfStyle?: string;
  hMerge?: DocumentTableCellHMerge;
  vMerge?: DocumentTableCellVMerge;
  gridSpan?: number;
  borders?: DocumentCellFormattingBorders;
}

export type { DocumentCellFormattingBorders } from './work-document-table-formatting-borders';

const MAX_CELL_FORMAT_SNAPSHOT_BYTES = 4_096;
const MARGIN_SIDES: readonly DocumentTableCellMarginSide[] = [
  'top',
  'right',
  'bottom',
  'left',
];

export function serializeDocumentCellFormatting(attributes: {
  verticalAlign?: unknown;
  fill?: unknown;
  margins?: unknown;
  width?: unknown;
  noWrap?: unknown;
  textDirection?: unknown;
  fitText?: unknown;
  hideMark?: unknown;
  cnfStyle?: unknown;
  hMerge?: unknown;
  vMerge?: unknown;
  gridSpan?: unknown;
  borders?: unknown;
}): string {
  const snapshot = normalizeDocumentCellFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Cell-formatting snapshot requires verticalAlign, fill, margins, width, noWrap, textDirection, fitText, hideMark, cnfStyle, hMerge, vMerge, gridSpan, or borders.',
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
        key !== 'fitText' &&
        key !== 'hideMark' &&
        key !== 'cnfStyle' &&
        key !== 'hMerge' &&
        key !== 'vMerge' &&
        key !== 'gridSpan' &&
        key !== 'borders',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentCellFormattingSnapshot(record);
  if (!snapshot) return null;
  return JSON.stringify(orderedSnapshot(snapshot)) === value ? snapshot : null;
}

export function normalizeDocumentCellFormattingSnapshot(attributes: {
  verticalAlign?: unknown;
  fill?: unknown;
  margins?: unknown;
  width?: unknown;
  noWrap?: unknown;
  textDirection?: unknown;
  fitText?: unknown;
  hideMark?: unknown;
  cnfStyle?: unknown;
  hMerge?: unknown;
  vMerge?: unknown;
  gridSpan?: unknown;
  borders?: unknown;
}): DocumentCellFormattingSnapshot | null {
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
  if ('hideMark' in attributes && attributes.hideMark !== undefined) {
    if (typeof attributes.hideMark !== 'boolean') return null;
    snapshot.hideMark = attributes.hideMark;
  }
  if ('cnfStyle' in attributes && attributes.cnfStyle !== undefined) {
    const cnfStyle = normalizeDocumentCnfStyle(attributes.cnfStyle);
    if (!cnfStyle) return null;
    snapshot.cnfStyle = cnfStyle;
  }
  if ('hMerge' in attributes && attributes.hMerge !== undefined) {
    const hMerge = normalizeDocumentTableCellHMerge(attributes.hMerge);
    if (!hMerge) return null;
    snapshot.hMerge = hMerge;
  }
  if ('vMerge' in attributes && attributes.vMerge !== undefined) {
    const vMerge = normalizeDocumentTableCellVMerge(attributes.vMerge);
    if (!vMerge) return null;
    snapshot.vMerge = vMerge;
  }
  if ('gridSpan' in attributes && attributes.gridSpan !== undefined) {
    const gridSpan = normalizeDocumentTableCellGridSpan(attributes.gridSpan);
    if (gridSpan === null) return null;
    snapshot.gridSpan = gridSpan;
  }
  if ('borders' in attributes && attributes.borders !== undefined) {
    const borders = normalizeDocumentCellFormattingBorders(attributes.borders);
    if (!borders) return null;
    snapshot.borders = orderedDocumentCellFormattingBorders(borders);
  }
  return snapshot.verticalAlign ||
    snapshot.fill ||
    snapshot.margins ||
    snapshot.width ||
    snapshot.noWrap !== undefined ||
    snapshot.textDirection !== undefined ||
    snapshot.fitText !== undefined ||
    snapshot.hideMark !== undefined ||
    snapshot.cnfStyle !== undefined ||
    snapshot.hMerge !== undefined ||
    snapshot.vMerge !== undefined ||
    snapshot.gridSpan !== undefined ||
    snapshot.borders !== undefined
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
    ...(formatting.hideMark !== undefined
      ? { hideMark: formatting.hideMark }
      : {}),
    ...(formatting.cnfStyle !== undefined
      ? { cnfStyle: formatting.cnfStyle }
      : {}),
    ...(formatting.hMerge !== undefined
      ? { hMerge: formatting.hMerge }
      : {}),
    ...(formatting.vMerge !== undefined
      ? { vMerge: formatting.vMerge }
      : {}),
    ...(formatting.gridSpan !== undefined
      ? { gridSpan: formatting.gridSpan }
      : {}),
    ...(formatting.borders !== undefined
      ? restoredCellBorderAttributes(formatting.borders)
      : {}),
  });
}

function restoredCellBorderAttributes(
  borders: DocumentCellFormattingBorders,
): Record<string, unknown> {
  const ordered = orderedDocumentCellFormattingBorders(borders);
  const full = fullCellBordersFromPartial(ordered);
  const representative =
    uniformCellBorder(full) ?? full.top ?? defaultCellBorder();
  return {
    borders: full,
    borderColor: representative.color,
    borderStyle: representative.style,
    borderWidth: representative.width,
  };
}

function fullCellBordersFromPartial(
  borders: DocumentCellFormattingBorders,
): DocumentTableCellBorders {
  const fallback = defaultCellBorder();
  return {
    top: borders.top ?? fallback,
    right: borders.right ?? fallback,
    bottom: borders.bottom ?? fallback,
    left: borders.left ?? fallback,
  };
}

function defaultCellBorder(): DocumentTableBorder {
  return { color: '#cfd5df', style: 'solid', width: 1 };
}

function uniformCellBorder(
  borders: DocumentTableCellBorders,
): DocumentTableBorder | null {
  const first = borders.top;
  if (
    sameCellBorder(borders.top, first) &&
    sameCellBorder(borders.right, first) &&
    sameCellBorder(borders.bottom, first) &&
    sameCellBorder(borders.left, first)
  ) {
    return first;
  }
  return null;
}

function sameCellBorder(
  left: DocumentTableBorder,
  right: DocumentTableBorder,
): boolean {
  return (
    left.color === right.color &&
    left.style === right.style &&
    left.width === right.width
  );
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
  if (snapshot.width !== undefined)
    ordered.width = orderedWidth(snapshot.width);
  if (snapshot.noWrap !== undefined) ordered.noWrap = snapshot.noWrap;
  if (snapshot.textDirection !== undefined) {
    ordered.textDirection = snapshot.textDirection;
  }
  if (snapshot.fitText !== undefined) ordered.fitText = snapshot.fitText;
  if (snapshot.hideMark !== undefined) ordered.hideMark = snapshot.hideMark;
  if (snapshot.cnfStyle !== undefined) ordered.cnfStyle = snapshot.cnfStyle;
  if (snapshot.hMerge !== undefined) ordered.hMerge = snapshot.hMerge;
  if (snapshot.vMerge !== undefined) ordered.vMerge = snapshot.vMerge;
  if (snapshot.gridSpan !== undefined) ordered.gridSpan = snapshot.gridSpan;
  if (snapshot.borders) {
    ordered.borders = orderedDocumentCellFormattingBorders(snapshot.borders);
  }
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
