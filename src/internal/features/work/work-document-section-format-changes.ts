import {
  normalizeDocumentColumns,
} from './work-document-columns';
import {
  documentPageMarginBody,
  documentPageMarginsForLayout,
  normalizeDocumentPageMargins,
  serializeDocumentPageMargins,
  type WorkDocumentPageMargins,
} from './work-document-page-margins';
import {
  documentPageGeometryForLayout,
  normalizeDocumentPageGeometry,
  normalizeDocumentPaperSource,
  parseDocumentPageGeometry,
  serializeDocumentPageGeometry,
  serializeDocumentPaperSource,
  type WorkDocumentPageGeometry,
  type WorkDocumentPaperSource,
} from './work-document-page-size';
import {
  normalizeDocumentPageChrome,
  parseDocumentPageChrome,
  serializeDocumentPageChrome,
} from './work-document-page-chrome';
import type {
  WorkDocumentColumns,
  WorkDocumentPaperSize,
} from './work-types';

export const DOCUMENT_SECTION_CHANGE_ATTRIBUTES = [
  'sectionChangeKind',
  'sectionChangeId',
  'sectionChangeAuthor',
  'sectionChangeDate',
  'sectionChangeBefore',
] as const;

/**
 * Equal-width column layout for reviewable section-property revisions.
 * Unequal `custom` widths stay outside this snapshot.
 */
export interface DocumentSectionEqualColumnsSnapshot {
  count: number;
  spacing: number;
  separator: boolean;
}

/**
 * Prior snapshot for reviewable section-property revisions.
 * At least one of orientation, pageGeometry, pageMargins, paperSource,
 * columns, differentFirstPage, or rtlGutter must be present.
 */
export interface DocumentSectionFormattingSnapshot {
  orientation?: 'portrait' | 'landscape';
  pageGeometry?: WorkDocumentPageGeometry;
  pageMargins?: WorkDocumentPageMargins;
  paperSource?: WorkDocumentPaperSource;
  columns?: DocumentSectionEqualColumnsSnapshot;
  differentFirstPage?: boolean;
  rtlGutter?: boolean;
}

const MAX_SECTION_FORMAT_SNAPSHOT_BYTES = 4_096;
const PAGE_MARGIN_TWIP_KEYS = [
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
] as const;

export function serializeDocumentSectionFormatting(
  attributes: {
    orientation?: unknown;
    pageGeometry?: unknown;
    pageMargins?: unknown;
    paperSource?: unknown;
    columns?: unknown;
    differentFirstPage?: unknown;
    rtlGutter?: unknown;
  },
): string {
  const snapshot = normalizeDocumentSectionFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Section-formatting snapshot requires orientation, pageGeometry, pageMargins, paperSource, columns, differentFirstPage, or rtlGutter.',
    );
  }
  return JSON.stringify(orderedSnapshot(snapshot));
}

export function parseDocumentSectionFormatting(
  value: unknown,
): DocumentSectionFormattingSnapshot | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_SECTION_FORMAT_SNAPSHOT_BYTES
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
        key !== 'orientation' &&
        key !== 'pageGeometry' &&
        key !== 'pageMargins' &&
        key !== 'paperSource' &&
        key !== 'columns' &&
        key !== 'differentFirstPage' &&
        key !== 'rtlGutter',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentSectionFormattingSnapshot(record);
  if (!snapshot) return null;
  return JSON.stringify(orderedSnapshot(snapshot)) === value ? snapshot : null;
}

export function normalizeDocumentSectionFormattingSnapshot(
  attributes: {
    orientation?: unknown;
    pageGeometry?: unknown;
    pageMargins?: unknown;
    paperSource?: unknown;
    columns?: unknown;
    differentFirstPage?: unknown;
    rtlGutter?: unknown;
  },
): DocumentSectionFormattingSnapshot | null {
  const snapshot: DocumentSectionFormattingSnapshot = {};
  if ('orientation' in attributes && attributes.orientation !== undefined) {
    if (
      attributes.orientation !== 'portrait' &&
      attributes.orientation !== 'landscape'
    ) {
      return null;
    }
    snapshot.orientation = attributes.orientation;
  }
  if ('pageGeometry' in attributes && attributes.pageGeometry !== undefined) {
    const pageGeometry = normalizeDocumentPageGeometry(attributes.pageGeometry);
    if (!pageGeometry) return null;
    snapshot.pageGeometry = orderedPageGeometry(pageGeometry);
  }
  if ('pageMargins' in attributes && attributes.pageMargins !== undefined) {
    const pageMargins = normalizeRevisionPageMargins(attributes.pageMargins);
    if (!pageMargins) return null;
    snapshot.pageMargins = pageMargins;
  }
  if ('paperSource' in attributes && attributes.paperSource !== undefined) {
    const paperSource = normalizeRevisionPaperSource(attributes.paperSource);
    if (!paperSource) return null;
    snapshot.paperSource = paperSource;
  }
  if ('columns' in attributes && attributes.columns !== undefined) {
    const columns = normalizeRevisionEqualColumns(attributes.columns);
    if (!columns) return null;
    snapshot.columns = columns;
  }
  if (
    'differentFirstPage' in attributes &&
    attributes.differentFirstPage !== undefined
  ) {
    if (typeof attributes.differentFirstPage !== 'boolean') return null;
    snapshot.differentFirstPage = attributes.differentFirstPage;
  }
  if ('rtlGutter' in attributes && attributes.rtlGutter !== undefined) {
    if (typeof attributes.rtlGutter !== 'boolean') return null;
    snapshot.rtlGutter = attributes.rtlGutter;
  }
  return snapshot.orientation ||
    snapshot.pageGeometry ||
    snapshot.pageMargins ||
    snapshot.paperSource ||
    snapshot.columns ||
    snapshot.differentFirstPage !== undefined ||
    snapshot.rtlGutter !== undefined
    ? snapshot
    : null;
}

export function clearDocumentSectionChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  const next = { ...attributes };
  for (const key of DOCUMENT_SECTION_CHANGE_ATTRIBUTES) {
    next[key] = null;
  }
  return next;
}

export function restoredDocumentSectionAttributes(
  attributes: Record<string, unknown>,
  serialized: unknown,
): Record<string, unknown> | null {
  const formatting = parseDocumentSectionFormatting(serialized);
  if (!formatting) return null;
  let orientation =
    attributes.orientation === 'landscape' ? 'landscape' : 'portrait';
  let pageGeometry = attributes.pageGeometry;
  if (formatting.pageGeometry) {
    pageGeometry =
      serializeDocumentPageGeometry(formatting.pageGeometry) ?? '';
    orientation =
      formatting.pageGeometry.orientation ??
      (formatting.pageGeometry.width <= formatting.pageGeometry.height
        ? 'portrait'
        : 'landscape');
  } else if (formatting.orientation) {
    const currentOrientation = orientation;
    orientation = formatting.orientation;
    if (
      formatting.orientation !== currentOrientation &&
      typeof pageGeometry === 'string' &&
      pageGeometry.trim()
    ) {
      const parsed = parseDocumentPageGeometry(pageGeometry);
      if (parsed) {
        pageGeometry =
          serializeDocumentPageGeometry({
            ...parsed,
            width: parsed.height,
            height: parsed.width,
            orientation: formatting.orientation,
          }) ?? '';
      }
    }
  }
  let pageMargins = attributes.pageMargins;
  let marginTop = attributes.marginTop;
  let marginRight = attributes.marginRight;
  let marginBottom = attributes.marginBottom;
  let marginLeft = attributes.marginLeft;
  if (formatting.pageMargins) {
    const existing = normalizeDocumentPageMargins(attributes.pageMargins);
    const restored = normalizeDocumentPageMargins({
      ...formatting.pageMargins,
      ...(existing?.mirrorMargins !== undefined
        ? { mirrorMargins: existing.mirrorMargins }
        : {}),
      ...(existing?.gutterAtTop !== undefined
        ? { gutterAtTop: existing.gutterAtTop }
        : {}),
      ...(existing?.gutterOnRight !== undefined
        ? { gutterOnRight: existing.gutterOnRight }
        : {}),
    });
    if (!restored) return null;
    pageMargins = serializeDocumentPageMargins(restored) ?? '';
    const body = documentPageMarginBody(restored, {
      top: Number(attributes.marginTop) || 25,
      right: Number(attributes.marginRight) || 23,
      bottom: Number(attributes.marginBottom) || 25,
      left: Number(attributes.marginLeft) || 23,
    });
    marginTop = body.top;
    marginRight = body.right;
    marginBottom = body.bottom;
    marginLeft = body.left;
  }
  let paperSource = attributes.paperSource;
  if (formatting.paperSource) {
    paperSource = serializeDocumentPaperSource(formatting.paperSource) ?? '';
  }
  let columnCount = attributes.columnCount;
  let columnSpacing = attributes.columnSpacing;
  let columnSeparator = attributes.columnSeparator;
  let columnLayout = attributes.columnLayout;
  if (formatting.columns) {
    const restored = normalizeDocumentColumns(formatting.columns);
    columnCount = restored.count;
    columnSpacing = restored.spacing;
    columnSeparator = restored.separator;
    columnLayout = JSON.stringify(restored);
  }
  let pageChrome = attributes.pageChrome;
  if (formatting.differentFirstPage !== undefined) {
    const chrome = parseDocumentPageChrome(
      typeof attributes.pageChrome === 'string' ? attributes.pageChrome : '',
      {
        headerText:
          typeof attributes.headerText === 'string'
            ? attributes.headerText
            : undefined,
        footerText:
          typeof attributes.footerText === 'string'
            ? attributes.footerText
            : undefined,
        showPageNumbers:
          typeof attributes.showPageNumbers === 'boolean'
            ? attributes.showPageNumbers
            : undefined,
      },
    );
    pageChrome = serializeDocumentPageChrome({
      ...chrome,
      differentFirstPage: formatting.differentFirstPage,
    });
  }
  if (formatting.rtlGutter !== undefined) {
    const existing = normalizeDocumentPageMargins(pageMargins) ?? {
      top: 25,
      right: 23,
      bottom: 25,
      left: 23,
      header: 12.5,
      footer: 12.5,
      gutter: 0,
    };
    const restored = normalizeDocumentPageMargins({
      ...existing,
      gutterOnRight: formatting.rtlGutter,
    });
    if (!restored) return null;
    pageMargins = serializeDocumentPageMargins(restored) ?? '';
  }
  return clearDocumentSectionChangeAttributes({
    ...attributes,
    orientation,
    pageGeometry,
    pageMargins,
    paperSource,
    pageChrome,
    marginTop,
    marginRight,
    marginBottom,
    marginLeft,
    columnCount,
    columnSpacing,
    columnSeparator,
    columnLayout,
  });
}

export function sectionFormattingSnapshotFromLayout(layout: {
  orientation: 'portrait' | 'landscape';
  margins: { top: number; right: number; bottom: number; left: number };
  pageMargins?: WorkDocumentPageMargins;
  pageGeometry?: WorkDocumentPageGeometry;
  pageSize?: WorkDocumentPaperSize;
  paperSource?: WorkDocumentPaperSource;
  columns?: WorkDocumentColumns;
  pageChrome?: unknown;
  headerText?: string;
  footerText?: string;
  showPageNumbers?: boolean;
}): DocumentSectionFormattingSnapshot | null {
  const pageMargins = documentPageMarginsForLayout(layout);
  const pageGeometry = documentPageGeometryForLayout({
    orientation: layout.orientation,
    pageGeometry: layout.pageGeometry,
    pageSize: layout.pageSize ?? 'a4',
  });
  const paperSource = normalizeRevisionPaperSource(layout.paperSource);
  const columns = normalizeRevisionEqualColumns(layout.columns);
  const chrome = normalizeDocumentPageChrome(
    typeof layout.pageChrome === 'string'
      ? parseDocumentPageChrome(layout.pageChrome, {
          headerText: layout.headerText,
          footerText: layout.footerText,
          showPageNumbers: layout.showPageNumbers,
        })
      : (layout.pageChrome as Parameters<typeof normalizeDocumentPageChrome>[0]),
    {
      headerText: layout.headerText,
      footerText: layout.footerText,
      showPageNumbers: layout.showPageNumbers,
    },
  );
  return normalizeDocumentSectionFormattingSnapshot({
    orientation: layout.orientation,
    pageGeometry,
    pageMargins: twipOnlyPageMargins(pageMargins),
    ...(paperSource ? { paperSource } : {}),
    ...(columns ? { columns } : {}),
    differentFirstPage: chrome.differentFirstPage,
    rtlGutter: pageMargins.gutterOnRight === true,
  });
}

function normalizeRevisionPageMargins(
  value: unknown,
): WorkDocumentPageMargins | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const keys = Object.keys(value);
  if (
    !keys.length ||
    keys.some(
      (key) =>
        !(PAGE_MARGIN_TWIP_KEYS as readonly string[]).includes(key),
    )
  ) {
    return null;
  }
  const normalized = normalizeDocumentPageMargins(value);
  return normalized ? twipOnlyPageMargins(normalized) : null;
}

function normalizeRevisionPaperSource(
  value: unknown,
): WorkDocumentPaperSource | null {
  const normalized = normalizeDocumentPaperSource(value);
  if (!normalized) return null;
  if (normalized.first === undefined && normalized.other === undefined) {
    return null;
  }
  return orderedPaperSource(normalized);
}

function twipOnlyPageMargins(
  pageMargins: WorkDocumentPageMargins,
): WorkDocumentPageMargins {
  return {
    top: pageMargins.top,
    right: pageMargins.right,
    bottom: pageMargins.bottom,
    left: pageMargins.left,
    header: pageMargins.header,
    footer: pageMargins.footer,
    gutter: pageMargins.gutter,
  };
}

function orderedPageGeometry(
  geometry: WorkDocumentPageGeometry,
): WorkDocumentPageGeometry {
  return {
    width: geometry.width,
    height: geometry.height,
    ...(geometry.orientation !== undefined
      ? { orientation: geometry.orientation }
      : {}),
    ...(geometry.code !== undefined ? { code: geometry.code } : {}),
  };
}

function orderedPaperSource(
  source: WorkDocumentPaperSource,
): WorkDocumentPaperSource {
  return {
    ...(source.first !== undefined ? { first: source.first } : {}),
    ...(source.other !== undefined ? { other: source.other } : {}),
  };
}

function normalizeRevisionEqualColumns(
  value: unknown,
): DocumentSectionEqualColumnsSnapshot | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if ('custom' in record && record.custom !== undefined) return null;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some(
      (key) => key !== 'count' && key !== 'spacing' && key !== 'separator',
    )
  ) {
    return null;
  }
  if (
    !('count' in record) ||
    !('spacing' in record) ||
    !('separator' in record)
  ) {
    return null;
  }
  const normalized = normalizeDocumentColumns({
    count: record.count as number,
    spacing: record.spacing as number,
    separator: Boolean(record.separator),
  });
  if (normalized.custom) return null;
  return {
    count: normalized.count,
    spacing: normalized.spacing,
    separator: normalized.separator,
  };
}

function orderedSnapshot(
  snapshot: DocumentSectionFormattingSnapshot,
): Record<string, unknown> {
  const ordered: Record<string, unknown> = {};
  if (snapshot.orientation) ordered.orientation = snapshot.orientation;
  if (snapshot.pageGeometry) {
    ordered.pageGeometry = orderedPageGeometry(snapshot.pageGeometry);
  }
  if (snapshot.pageMargins) {
    const margins: Record<string, number> = {};
    for (const key of PAGE_MARGIN_TWIP_KEYS) {
      margins[key] = snapshot.pageMargins[key];
    }
    ordered.pageMargins = margins;
  }
  if (snapshot.paperSource) {
    ordered.paperSource = orderedPaperSource(snapshot.paperSource);
  }
  if (snapshot.differentFirstPage !== undefined) {
    ordered.differentFirstPage = snapshot.differentFirstPage;
  }
  if (snapshot.rtlGutter !== undefined) {
    ordered.rtlGutter = snapshot.rtlGutter;
  }
  if (snapshot.columns) {
    ordered.columns = {
      count: snapshot.columns.count,
      spacing: snapshot.columns.spacing,
      separator: snapshot.columns.separator,
    };
  }
  return ordered;
}
