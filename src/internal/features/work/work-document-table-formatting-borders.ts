/**
 * Relationship-free table-level `w:tblBorders` snapshots for reviewable
 * `table-formatting` revisions (direct-color edges only).
 */
import {
  normalizeDocumentTableBorder,
  normalizeDocumentTableBorderStyle,
  normalizeDocumentTableBorderWidth,
  normalizeTableColor,
  type DocumentTableBorder,
  type DocumentTableBorderStyle,
} from './work-document-table-borders';

export type DocumentTableFormattingBorderEdge =
  | 'top'
  | 'left'
  | 'bottom'
  | 'right'
  | 'insideH'
  | 'insideV';

export type DocumentTableFormattingBorders = Partial<
  Record<DocumentTableFormattingBorderEdge, DocumentTableBorder>
>;

export const DOCUMENT_TABLE_FORMATTING_BORDER_EDGES: readonly DocumentTableFormattingBorderEdge[] =
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];

const EDGE_ALIASES: Record<string, DocumentTableFormattingBorderEdge> = {
  top: 'top',
  left: 'left',
  start: 'left',
  bottom: 'bottom',
  right: 'right',
  end: 'right',
  insideH: 'insideH',
  insideV: 'insideV',
};

export function normalizeDocumentTableFormattingBorders(
  value: unknown,
): DocumentTableFormattingBorders | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (!keys.length) return null;
  if (
    keys.some(
      (key) =>
        !DOCUMENT_TABLE_FORMATTING_BORDER_EDGES.includes(
          key as DocumentTableFormattingBorderEdge,
        ),
    )
  ) {
    return null;
  }
  const borders: DocumentTableFormattingBorders = {};
  for (const edge of DOCUMENT_TABLE_FORMATTING_BORDER_EDGES) {
    if (!(edge in record)) continue;
    const border = normalizeDocumentTableBorder(record[edge]);
    if (!border) return null;
    borders[edge] = border;
  }
  return Object.keys(borders).length ? orderedDocumentTableFormattingBorders(borders) : null;
}

export function orderedDocumentTableFormattingBorders(
  borders: DocumentTableFormattingBorders,
): DocumentTableFormattingBorders {
  const ordered: DocumentTableFormattingBorders = {};
  for (const edge of DOCUMENT_TABLE_FORMATTING_BORDER_EDGES) {
    if (borders[edge]) ordered[edge] = { ...borders[edge]! };
  }
  return ordered;
}

export function serializeDocumentTableFormattingBordersDataset(
  borders: DocumentTableFormattingBorders | null | undefined,
): string | null {
  const normalized = normalizeDocumentTableFormattingBorders(borders ?? null);
  return normalized ? JSON.stringify(normalized) : null;
}

export function parseDocumentTableFormattingBordersDataset(
  value: unknown,
): DocumentTableFormattingBorders | null {
  if (typeof value !== 'string' || !value.length || value.length > 2_048) {
    return null;
  }
  try {
    return normalizeDocumentTableFormattingBorders(JSON.parse(value));
  } catch {
    return null;
  }
}

export function mapDocxBorderStyleToSnapshot(
  value: string | null,
): DocumentTableBorderStyle | null {
  const normalized = value?.trim();
  if (!normalized) return null;
  if (normalized === 'nil' || normalized === 'none') return 'none';
  if (normalized === 'double' || normalized === 'triple') return 'double';
  if (normalized === 'dotted') return 'dotted';
  if (/dash|dotDash|dotDotDash/i.test(normalized)) return 'dashed';
  if (
    normalized === 'single' ||
    normalized === 'thick' ||
    normalized === 'thinThickSmallGap' ||
    normalized === 'thickThinSmallGap' ||
    normalized === 'thinThickThinSmallGap' ||
    normalized === 'thinThickMediumGap' ||
    normalized === 'thickThinMediumGap' ||
    normalized === 'thinThickThinMediumGap' ||
    normalized === 'thinThickLargeGap' ||
    normalized === 'thickThinLargeGap' ||
    normalized === 'thinThickThinLargeGap' ||
    normalized === 'wave' ||
    normalized === 'doubleWave' ||
    normalized === 'dashSmallGap' ||
    normalized === 'dashDotStroked' ||
    normalized === 'threeDEmboss' ||
    normalized === 'threeDEngrave' ||
    normalized === 'outset' ||
    normalized === 'inset'
  ) {
    return 'solid';
  }
  // Fail closed on art / unknown border vals for reviewable revisions.
  return null;
}

export function mapSnapshotBorderStyleToDocx(
  style: DocumentTableBorderStyle,
): string {
  if (style === 'none') return 'nil';
  if (style === 'double') return 'double';
  if (style === 'dotted') return 'dotted';
  if (style === 'dashed') return 'dashed';
  return 'single';
}

export function resolveFormattingBorderEdge(
  localName: string,
): DocumentTableFormattingBorderEdge | null {
  return EDGE_ALIASES[localName] ?? null;
}

export function normalizeRevisionBorderWidthFromSz(
  sz: string | null,
  style: DocumentTableBorderStyle,
): number | null {
  if (style === 'none') return 0;
  const size = Number(sz);
  if (!Number.isFinite(size) || size < 0) return null;
  const width = size > 0 ? size / 6 : 1;
  return normalizeDocumentTableBorderWidth(
    Math.max(0.5, Math.min(6, Math.round(width * 2) / 2)),
  );
}

export function revisionBorderFromDocxEdge(
  edge: Element,
  attribute: (element: Element, localName: string) => string | null,
): DocumentTableBorder | null {
  // Theme-bound edges stay opaque / fail-closed for this reviewable subset.
  if (
    attribute(edge, 'themeColor') ||
    attribute(edge, 'themeTint') ||
    attribute(edge, 'themeShade')
  ) {
    return null;
  }
  const style = mapDocxBorderStyleToSnapshot(attribute(edge, 'val'));
  if (!style) return null;
  const space = attribute(edge, 'space');
  if (space !== null && space !== '' && Number(space) !== 0) return null;
  const rawColor = attribute(edge, 'color');
  const color =
    style === 'none'
      ? '#000000'
      : rawColor === null || rawColor.trim().toLowerCase() === 'auto'
        ? '#000000'
        : normalizeTableColor(
            rawColor.startsWith('#') ? rawColor : `#${rawColor}`,
          );
  if (!color) return null;
  const width = normalizeRevisionBorderWidthFromSz(attribute(edge, 'sz'), style);
  if (width === null) return null;
  return normalizeDocumentTableBorder({ color, style, width });
}

export function docxSzFromSnapshotBorderWidth(width: number): string {
  return String(Math.max(2, Math.round(width * 6)));
}
