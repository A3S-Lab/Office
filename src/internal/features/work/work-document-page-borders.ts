import {
  type DocumentBorderPresentation,
  type DocumentParagraphBorder,
  documentBorderPresentation,
  normalizeDocumentParagraphBorder,
} from './work-document-paragraph-borders';
import { millimetersToPixels } from './work-document-layout';
import type { WorkDocumentMargins } from './work-types';

export const DOCUMENT_PAGE_BORDER_EDGES = [
  'top',
  'left',
  'bottom',
  'right',
] as const;

export type WorkDocumentPageBorderEdge =
  (typeof DOCUMENT_PAGE_BORDER_EDGES)[number];
export type WorkDocumentPageBorderDisplay =
  | 'allPages'
  | 'firstPage'
  | 'notFirstPage';
export type WorkDocumentPageBorderOffsetFrom = 'page' | 'text';
export type WorkDocumentPageBorderZOrder = 'front' | 'back';

export interface WorkDocumentPageBorders {
  display?: WorkDocumentPageBorderDisplay;
  offsetFrom?: WorkDocumentPageBorderOffsetFrom;
  zOrder?: WorkDocumentPageBorderZOrder;
  edges: Partial<Record<WorkDocumentPageBorderEdge, DocumentParagraphBorder>>;
}

export interface ResolvedDocumentPageBorders {
  display: WorkDocumentPageBorderDisplay;
  offsetFrom: WorkDocumentPageBorderOffsetFrom;
  zOrder: WorkDocumentPageBorderZOrder;
  edges: Partial<
    Record<WorkDocumentPageBorderEdge, DocumentBorderPresentation>
  >;
  insets: Record<WorkDocumentPageBorderEdge, number>;
}

const PAGE_BORDER_PROPERTY_SET = new Set([
  'display',
  'offsetFrom',
  'zOrder',
  'edges',
]);
const PAGE_BORDER_EDGE_SET = new Set<string>(DOCUMENT_PAGE_BORDER_EDGES);
const PAGE_BORDER_DISPLAY_SET = new Set<WorkDocumentPageBorderDisplay>([
  'allPages',
  'firstPage',
  'notFirstPage',
]);
const PAGE_BORDER_OFFSET_SET = new Set<WorkDocumentPageBorderOffsetFrom>([
  'page',
  'text',
]);
const PAGE_BORDER_Z_ORDER_SET = new Set<WorkDocumentPageBorderZOrder>([
  'front',
  'back',
]);
const MAX_SERIALIZED_PAGE_BORDERS = 32_768;
const POINTS_TO_PIXELS = 96 / 72;

export function normalizeDocumentPageBorders(
  source: unknown,
): WorkDocumentPageBorders | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null;
  }
  const record = source as Record<string, unknown>;
  if (
    Object.keys(record).some((key) => !PAGE_BORDER_PROPERTY_SET.has(key)) ||
    !record.edges ||
    typeof record.edges !== 'object' ||
    Array.isArray(record.edges)
  ) {
    return null;
  }
  const display = optionalEnum(record.display, PAGE_BORDER_DISPLAY_SET);
  const offsetFrom = optionalEnum(record.offsetFrom, PAGE_BORDER_OFFSET_SET);
  const zOrder = optionalEnum(record.zOrder, PAGE_BORDER_Z_ORDER_SET);
  if (display === null || offsetFrom === null || zOrder === null) return null;
  const edgeRecord = record.edges as Record<string, unknown>;
  if (Object.keys(edgeRecord).some((key) => !PAGE_BORDER_EDGE_SET.has(key))) {
    return null;
  }
  const edges: WorkDocumentPageBorders['edges'] = {};
  for (const edge of DOCUMENT_PAGE_BORDER_EDGES) {
    if (edgeRecord[edge] === undefined) continue;
    const border = normalizeDocumentParagraphBorder(edgeRecord[edge]);
    if (!border) return null;
    edges[edge] = border;
  }
  return {
    ...(display ? { display } : {}),
    ...(offsetFrom ? { offsetFrom } : {}),
    ...(zOrder ? { zOrder } : {}),
    edges,
  };
}

export function parseDocumentPageBorders(
  source: unknown,
): WorkDocumentPageBorders | null {
  if (typeof source !== 'string') return normalizeDocumentPageBorders(source);
  if (!source.trim() || source.length > MAX_SERIALIZED_PAGE_BORDERS) {
    return null;
  }
  try {
    return normalizeDocumentPageBorders(JSON.parse(source));
  } catch {
    return null;
  }
}

export function serializeDocumentPageBorders(
  source: unknown,
): string | undefined {
  const borders = normalizeDocumentPageBorders(source);
  if (!borders) return undefined;
  return JSON.stringify({
    ...(borders.display ? { display: borders.display } : {}),
    ...(borders.offsetFrom ? { offsetFrom: borders.offsetFrom } : {}),
    ...(borders.zOrder ? { zOrder: borders.zOrder } : {}),
    edges: Object.fromEntries(
      DOCUMENT_PAGE_BORDER_EDGES.flatMap((edge) => {
        const border = borders.edges[edge];
        return border ? [[edge, border]] : [];
      }),
    ),
  });
}

export function documentPageBordersVisible(
  source: unknown,
  sectionPage: number,
): boolean {
  const borders = normalizeDocumentPageBorders(source);
  if (!borders) return false;
  const display = borders.display ?? 'allPages';
  if (display === 'firstPage') return sectionPage === 1;
  if (display === 'notFirstPage') return sectionPage > 1;
  return true;
}

export function resolveDocumentPageBorders(
  source: unknown,
  margins: WorkDocumentMargins,
): ResolvedDocumentPageBorders | null {
  const borders = normalizeDocumentPageBorders(source);
  if (!borders) return null;
  const offsetFrom = borders.offsetFrom ?? 'text';
  const presentations: ResolvedDocumentPageBorders['edges'] = {};
  const insets = {} as Record<WorkDocumentPageBorderEdge, number>;
  for (const edge of DOCUMENT_PAGE_BORDER_EDGES) {
    const border = borders.edges[edge];
    const presentation = border
      ? documentBorderPresentation(border)
      : undefined;
    if (presentation) presentations[edge] = presentation;
    const space = (border?.space ?? 0) * POINTS_TO_PIXELS;
    insets[edge] =
      offsetFrom === 'page'
        ? space
        : Math.max(
            0,
            millimetersToPixels(margins[edge]) -
              space -
              (presentation?.width ?? 0),
          );
  }
  return {
    display: borders.display ?? 'allPages',
    offsetFrom,
    zOrder: borders.zOrder ?? 'front',
    edges: presentations,
    insets,
  };
}

function optionalEnum<T extends string>(
  value: unknown,
  values: ReadonlySet<T>,
): T | null | undefined {
  if (value === undefined) return undefined;
  return typeof value === 'string' && values.has(value as T)
    ? (value as T)
    : null;
}
