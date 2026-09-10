import { normalizeDocumentColumns } from './work-document-columns';
import {
  documentPageMarginBody,
  documentPageMarginsForLayout,
  normalizeDocumentPageMargins,
  serializeDocumentPageMargins,
  type WorkDocumentPageMargins,
} from './work-document-page-margins';
import {
  DOCUMENT_PAGE_BORDER_EDGES,
  normalizeDocumentPageBorders,
  serializeDocumentPageBorders,
  type WorkDocumentPageBorders,
} from './work-document-page-borders';
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
  WorkDocumentGrid,
  WorkDocumentGridType,
  WorkDocumentLnNumType,
  WorkDocumentPaperSize,
  WorkDocumentFootnoteNumRestart,
  WorkDocumentFootnotePos,
  WorkDocumentFootnotePr,
  WorkDocumentEndnotePos,
  WorkDocumentEndnotePr,
  WorkDocumentPgNumFmt,
  WorkDocumentPgNumType,
  WorkDocumentSectionBreakType,
  WorkDocumentSectionTextDirection,
  WorkDocumentSectionVerticalAlign,
} from './work-types';

export const DOCUMENT_SECTION_CHANGE_ATTRIBUTES = [
  'sectionChangeKind',
  'sectionChangeId',
  'sectionChangeAuthor',
  'sectionChangeDate',
  'sectionChangeBefore',
] as const;

/**
 * Column layout for reviewable section-property revisions.
 * Equal-width snapshots omit `custom`; unequal-width snapshots include it.
 */
export interface DocumentSectionColumnsSnapshot {
  count: number;
  spacing: number;
  separator: boolean;
  custom?: Array<{ widthPercent: number; spacing: number }>;
}

/** @deprecated Use {@link DocumentSectionColumnsSnapshot}. */
export type DocumentSectionEqualColumnsSnapshot =
  DocumentSectionColumnsSnapshot;

/**
 * Prior snapshot for reviewable section-property revisions.
 * At least one of orientation, pageGeometry, pageMargins, paperSource,
 * columns, differentFirstPage, rtlGutter, documentGrid, lnNumType, pgNumType,
 * formProt, noEndnote, verticalAlign, textDirection, bidi, footnotePr,
 * endnotePr, breakAfter, or pageBorders must be present.
 */
export interface DocumentSectionFormattingSnapshot {
  orientation?: 'portrait' | 'landscape';
  /** Section break type (`w:type` / ST_SectionMark); current-section field. */
  breakAfter?: WorkDocumentSectionBreakType;
  pageGeometry?: WorkDocumentPageGeometry;
  pageMargins?: WorkDocumentPageMargins;
  paperSource?: WorkDocumentPaperSource;
  /** Bounded section page borders (`w:pgBorders`). */
  pageBorders?: WorkDocumentPageBorders;
  columns?: DocumentSectionColumnsSnapshot;
  differentFirstPage?: boolean;
  rtlGutter?: boolean;
  documentGrid?: WorkDocumentGrid;
  lnNumType?: WorkDocumentLnNumType;
  pgNumType?: WorkDocumentPgNumType;
  formProt?: boolean;
  noEndnote?: boolean;
  verticalAlign?: WorkDocumentSectionVerticalAlign;
  textDirection?: WorkDocumentSectionTextDirection;
  bidi?: boolean;
  footnotePr?: WorkDocumentFootnotePr;
  endnotePr?: WorkDocumentEndnotePr;
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

export function serializeDocumentSectionFormatting(attributes: {
  orientation?: unknown;
  breakAfter?: unknown;
  pageGeometry?: unknown;
  pageMargins?: unknown;
  paperSource?: unknown;
  pageBorders?: unknown;
  columns?: unknown;
  differentFirstPage?: unknown;
  rtlGutter?: unknown;
  documentGrid?: unknown;
  lnNumType?: unknown;
  pgNumType?: unknown;
  formProt?: unknown;
  noEndnote?: unknown;
  verticalAlign?: unknown;
  textDirection?: unknown;
  bidi?: unknown;
  footnotePr?: unknown;
  endnotePr?: unknown;
}): string {
  const snapshot = normalizeDocumentSectionFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Section-formatting snapshot requires orientation, breakAfter, pageGeometry, pageMargins, paperSource, pageBorders, columns, differentFirstPage, rtlGutter, documentGrid, lnNumType, pgNumType, formProt, noEndnote, verticalAlign, textDirection, bidi, footnotePr, or endnotePr.',
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
        key !== 'breakAfter' &&
        key !== 'pageGeometry' &&
        key !== 'pageMargins' &&
        key !== 'paperSource' &&
        key !== 'pageBorders' &&
        key !== 'columns' &&
        key !== 'differentFirstPage' &&
        key !== 'rtlGutter' &&
        key !== 'documentGrid' &&
        key !== 'lnNumType' &&
        key !== 'pgNumType' &&
        key !== 'formProt' &&
        key !== 'noEndnote' &&
        key !== 'verticalAlign' &&
        key !== 'textDirection' &&
        key !== 'bidi' &&
        key !== 'footnotePr' &&
        key !== 'endnotePr',
    )
  ) {
    return null;
  }
  const snapshot = normalizeDocumentSectionFormattingSnapshot(record);
  if (!snapshot) return null;
  return JSON.stringify(orderedSnapshot(snapshot)) === value ? snapshot : null;
}

export function normalizeDocumentSectionFormattingSnapshot(attributes: {
  orientation?: unknown;
  breakAfter?: unknown;
  pageGeometry?: unknown;
  pageMargins?: unknown;
  paperSource?: unknown;
  pageBorders?: unknown;
  columns?: unknown;
  differentFirstPage?: unknown;
  rtlGutter?: unknown;
  documentGrid?: unknown;
  lnNumType?: unknown;
  pgNumType?: unknown;
  formProt?: unknown;
  noEndnote?: unknown;
  verticalAlign?: unknown;
  textDirection?: unknown;
  bidi?: unknown;
  footnotePr?: unknown;
  endnotePr?: unknown;
}): DocumentSectionFormattingSnapshot | null {
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
  if ('breakAfter' in attributes && attributes.breakAfter !== undefined) {
    if (!isSectionBreakAfter(attributes.breakAfter)) return null;
    snapshot.breakAfter = attributes.breakAfter;
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
  if ('pageBorders' in attributes && attributes.pageBorders !== undefined) {
    const pageBorders = normalizeRevisionPageBorders(attributes.pageBorders);
    if (pageBorders === null) return null;
    snapshot.pageBorders = pageBorders;
  }
  if ('columns' in attributes && attributes.columns !== undefined) {
    const columns = normalizeRevisionColumns(attributes.columns);
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
  if ('documentGrid' in attributes && attributes.documentGrid !== undefined) {
    const documentGrid = normalizeRevisionDocumentGrid(attributes.documentGrid);
    if (!documentGrid) return null;
    snapshot.documentGrid = documentGrid;
  }
  if ('lnNumType' in attributes && attributes.lnNumType !== undefined) {
    const lnNumType = normalizeRevisionLnNumType(attributes.lnNumType);
    if (!lnNumType) return null;
    snapshot.lnNumType = lnNumType;
  }
  if ('pgNumType' in attributes && attributes.pgNumType !== undefined) {
    const pgNumType = normalizeRevisionPgNumType(attributes.pgNumType);
    if (!pgNumType) return null;
    snapshot.pgNumType = pgNumType;
  }
  if ('formProt' in attributes && attributes.formProt !== undefined) {
    if (typeof attributes.formProt !== 'boolean') return null;
    snapshot.formProt = attributes.formProt;
  }
  if ('noEndnote' in attributes && attributes.noEndnote !== undefined) {
    if (typeof attributes.noEndnote !== 'boolean') return null;
    snapshot.noEndnote = attributes.noEndnote;
  }
  if ('verticalAlign' in attributes && attributes.verticalAlign !== undefined) {
    if (!isSectionVerticalAlign(attributes.verticalAlign)) return null;
    snapshot.verticalAlign = attributes.verticalAlign;
  }
  if ('textDirection' in attributes && attributes.textDirection !== undefined) {
    if (!isSectionTextDirection(attributes.textDirection)) return null;
    snapshot.textDirection = attributes.textDirection;
  }
  if ('bidi' in attributes && attributes.bidi !== undefined) {
    if (typeof attributes.bidi !== 'boolean') return null;
    snapshot.bidi = attributes.bidi;
  }
  if ('footnotePr' in attributes && attributes.footnotePr !== undefined) {
    const footnotePr = normalizeRevisionFootnotePr(attributes.footnotePr);
    if (footnotePr === null) return null;
    snapshot.footnotePr = footnotePr;
  }
  if ('endnotePr' in attributes && attributes.endnotePr !== undefined) {
    const endnotePr = normalizeRevisionEndnotePr(attributes.endnotePr);
    if (endnotePr === null) return null;
    snapshot.endnotePr = endnotePr;
  }
  return snapshot.orientation ||
    snapshot.breakAfter !== undefined ||
    snapshot.pageGeometry ||
    snapshot.pageMargins ||
    snapshot.paperSource ||
    snapshot.pageBorders !== undefined ||
    snapshot.columns ||
    snapshot.differentFirstPage !== undefined ||
    snapshot.rtlGutter !== undefined ||
    snapshot.documentGrid ||
    snapshot.lnNumType ||
    snapshot.pgNumType ||
    snapshot.formProt !== undefined ||
    snapshot.noEndnote !== undefined ||
    snapshot.verticalAlign !== undefined ||
    snapshot.textDirection !== undefined ||
    snapshot.bidi !== undefined ||
    snapshot.footnotePr !== undefined ||
    snapshot.endnotePr !== undefined
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
  let breakAfter = attributes.breakAfter;
  if (formatting.breakAfter !== undefined) {
    breakAfter = formatting.breakAfter;
  }
  let pageGeometry = attributes.pageGeometry;
  if (formatting.pageGeometry) {
    pageGeometry = serializeDocumentPageGeometry(formatting.pageGeometry) ?? '';
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
  let documentGridType = attributes.documentGridType;
  let documentGridLinePitch = attributes.documentGridLinePitch;
  let documentGridCharSpace = attributes.documentGridCharSpace;
  if (formatting.documentGrid) {
    documentGridType = formatting.documentGrid.type;
    documentGridLinePitch = formatting.documentGrid.linePitch;
    documentGridCharSpace = formatting.documentGrid.charSpace ?? null;
  }
  let lnNumCountBy = attributes.lnNumCountBy;
  let lnNumStart = attributes.lnNumStart;
  let lnNumDistance = attributes.lnNumDistance;
  let lnNumRestart = attributes.lnNumRestart;
  if (formatting.lnNumType) {
    lnNumCountBy = formatting.lnNumType.countBy ?? null;
    lnNumStart = formatting.lnNumType.start ?? null;
    lnNumDistance = formatting.lnNumType.distance ?? null;
    lnNumRestart = formatting.lnNumType.restart ?? '';
  }
  let pgNumFmt = attributes.pgNumFmt;
  let pgNumStart = attributes.pgNumStart;
  let pageNumberStart = attributes.pageNumberStart;
  let pgNumChapStyle = attributes.pgNumChapStyle;
  let pgNumChapSep = attributes.pgNumChapSep;
  if (formatting.pgNumType) {
    pgNumFmt = formatting.pgNumType.fmt ?? '';
    pgNumStart = formatting.pgNumType.start ?? null;
    pgNumChapStyle = formatting.pgNumType.chapStyle ?? null;
    pgNumChapSep = formatting.pgNumType.chapSep ?? '';
    pageNumberStart =
      formatting.pgNumType.start !== undefined && formatting.pgNumType.start > 0
        ? Math.min(9999, formatting.pgNumType.start)
        : null;
  }
  let formProt = attributes.formProt;
  if (formatting.formProt !== undefined) {
    formProt = formatting.formProt;
  }
  let noEndnote = attributes.noEndnote;
  if (formatting.noEndnote !== undefined) {
    noEndnote = formatting.noEndnote;
  }
  let verticalAlign = attributes.verticalAlign;
  if (formatting.verticalAlign !== undefined) {
    verticalAlign = formatting.verticalAlign;
  }
  let textDirection = attributes.textDirection;
  if (formatting.textDirection !== undefined) {
    textDirection = formatting.textDirection;
  }
  let bidi = attributes.bidi;
  if (formatting.bidi !== undefined) {
    bidi = formatting.bidi;
  }
  let pageBorders = attributes.pageBorders;
  if (formatting.pageBorders !== undefined) {
    pageBorders = serializeDocumentPageBorders(formatting.pageBorders) ?? '';
  }
  let footnotePr = attributes.footnotePr;
  if (formatting.footnotePr !== undefined) {
    footnotePr = serializeDocumentFootnotePr(formatting.footnotePr);
  }
  let endnotePr = attributes.endnotePr;
  if (formatting.endnotePr !== undefined) {
    endnotePr = serializeDocumentEndnotePr(formatting.endnotePr);
  }
  return clearDocumentSectionChangeAttributes({
    ...attributes,
    orientation,
    breakAfter,
    pageGeometry,
    pageMargins,
    paperSource,
    pageBorders,
    pageChrome,
    marginTop,
    marginRight,
    marginBottom,
    marginLeft,
    columnCount,
    columnSpacing,
    columnSeparator,
    columnLayout,
    documentGridType,
    documentGridLinePitch,
    documentGridCharSpace,
    lnNumCountBy,
    lnNumStart,
    lnNumDistance,
    lnNumRestart,
    pgNumFmt,
    pgNumStart,
    pgNumChapStyle,
    pgNumChapSep,
    pageNumberStart,
    formProt,
    noEndnote,
    verticalAlign,
    textDirection,
    bidi,
    footnotePr,
    endnotePr,
  });
}

export function sectionFormattingSnapshotFromLayout(layout: {
  orientation: 'portrait' | 'landscape';
  breakAfter?: WorkDocumentSectionBreakType;
  margins: { top: number; right: number; bottom: number; left: number };
  pageMargins?: WorkDocumentPageMargins;
  pageGeometry?: WorkDocumentPageGeometry;
  pageSize?: WorkDocumentPaperSize;
  paperSource?: WorkDocumentPaperSource;
  columns?: WorkDocumentColumns;
  pageChrome?: unknown;
  documentGrid?: WorkDocumentGrid;
  lnNumType?: WorkDocumentLnNumType;
  pgNumType?: WorkDocumentPgNumType;
  formProt?: boolean;
  noEndnote?: boolean;
  verticalAlign?: WorkDocumentSectionVerticalAlign;
  textDirection?: WorkDocumentSectionTextDirection;
  bidi?: boolean;
  footnotePr?: WorkDocumentFootnotePr;
  endnotePr?: WorkDocumentEndnotePr;
  pageBorders?: WorkDocumentPageBorders;
  pageNumberStart?: number;
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
  const columns = normalizeRevisionColumns(layout.columns);
  const chrome = normalizeDocumentPageChrome(
    typeof layout.pageChrome === 'string'
      ? parseDocumentPageChrome(layout.pageChrome, {
          headerText: layout.headerText,
          footerText: layout.footerText,
          showPageNumbers: layout.showPageNumbers,
        })
      : (layout.pageChrome as Parameters<
          typeof normalizeDocumentPageChrome
        >[0]),
    {
      headerText: layout.headerText,
      footerText: layout.footerText,
      showPageNumbers: layout.showPageNumbers,
    },
  );
  const documentGrid = normalizeRevisionDocumentGrid(layout.documentGrid);
  const lnNumType = normalizeRevisionLnNumType(layout.lnNumType);
  const pgNumType = normalizeRevisionPgNumType(
    layout.pgNumType ??
      (layout.pageNumberStart !== undefined
        ? { start: layout.pageNumberStart }
        : undefined),
  );
  return normalizeDocumentSectionFormattingSnapshot({
    orientation: layout.orientation,
    ...(layout.breakAfter !== undefined
      ? { breakAfter: layout.breakAfter }
      : {}),
    pageGeometry,
    pageMargins: twipOnlyPageMargins(pageMargins),
    ...(paperSource ? { paperSource } : {}),
    ...(columns ? { columns } : {}),
    differentFirstPage: chrome.differentFirstPage,
    rtlGutter: pageMargins.gutterOnRight === true,
    ...(documentGrid ? { documentGrid } : {}),
    ...(lnNumType ? { lnNumType } : {}),
    ...(pgNumType ? { pgNumType } : {}),
    formProt: layout.formProt === true,
    noEndnote: layout.noEndnote === true,
    ...(layout.verticalAlign !== undefined
      ? { verticalAlign: layout.verticalAlign }
      : {}),
    ...(layout.textDirection !== undefined
      ? { textDirection: layout.textDirection }
      : {}),
    bidi: layout.bidi === true,
    ...(layout.footnotePr !== undefined
      ? { footnotePr: layout.footnotePr }
      : {}),
    ...(layout.endnotePr !== undefined ? { endnotePr: layout.endnotePr } : {}),
    ...(layout.pageBorders !== undefined
      ? { pageBorders: layout.pageBorders }
      : {}),
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
      (key) => !(PAGE_MARGIN_TWIP_KEYS as readonly string[]).includes(key),
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

function normalizeRevisionPageBorders(
  value: unknown,
): WorkDocumentPageBorders | null {
  const normalized = normalizeDocumentPageBorders(value);
  if (!normalized) return null;
  return orderedPageBorders(normalized);
}

function orderedPageBorders(
  borders: WorkDocumentPageBorders,
): WorkDocumentPageBorders {
  const edges: WorkDocumentPageBorders['edges'] = {};
  for (const edge of DOCUMENT_PAGE_BORDER_EDGES) {
    const border = borders.edges[edge];
    if (!border) continue;
    edges[edge] = border;
  }
  return {
    ...(borders.display ? { display: borders.display } : {}),
    ...(borders.offsetFrom ? { offsetFrom: borders.offsetFrom } : {}),
    ...(borders.zOrder ? { zOrder: borders.zOrder } : {}),
    edges,
  };
}

function normalizeRevisionColumns(
  value: unknown,
): DocumentSectionColumnsSnapshot | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some(
      (key) =>
        key !== 'count' &&
        key !== 'spacing' &&
        key !== 'separator' &&
        key !== 'custom',
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
  const customSource = record.custom;
  const custom =
    customSource === undefined
      ? undefined
      : Array.isArray(customSource)
        ? customSource.map((entry) => {
            if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
              return null;
            }
            const item = entry as Record<string, unknown>;
            if (
              !('widthPercent' in item) ||
              !('spacing' in item) ||
              Object.keys(item).some(
                (key) => key !== 'widthPercent' && key !== 'spacing',
              )
            ) {
              return null;
            }
            return {
              widthPercent: item.widthPercent as number,
              spacing: item.spacing as number,
            };
          })
        : null;
  if (
    customSource !== undefined &&
    (!custom || custom.some((entry) => !entry))
  ) {
    return null;
  }
  const customColumns = custom?.filter(
    (entry): entry is { widthPercent: number; spacing: number } =>
      entry !== null,
  );
  const normalized = normalizeDocumentColumns({
    count: record.count as number,
    spacing: record.spacing as number,
    separator: Boolean(record.separator),
    ...(customColumns?.length ? { custom: customColumns } : {}),
  });
  const snapshot: DocumentSectionColumnsSnapshot = {
    count: normalized.count,
    spacing: normalized.spacing,
    separator: normalized.separator,
  };
  if (normalized.custom) {
    snapshot.custom = normalized.custom.map((column) => ({
      widthPercent: column.widthPercent,
      spacing: column.spacing,
    }));
  } else if (customSource !== undefined) {
    return null;
  }
  return snapshot;
}

function normalizeRevisionDocumentGrid(
  value: unknown,
): WorkDocumentGrid | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some(
      (key) => key !== 'type' && key !== 'linePitch' && key !== 'charSpace',
    )
  ) {
    return null;
  }
  if (!('type' in record) || !('linePitch' in record)) return null;
  if (!validDocumentGridType(record.type)) return null;
  const linePitch = Number(record.linePitch);
  if (!Number.isFinite(linePitch) || linePitch <= 0) return null;
  const next: WorkDocumentGrid = {
    type: record.type,
    linePitch: Math.min(720, Number(linePitch.toFixed(2))),
  };
  if ('charSpace' in record) {
    const charSpace = Number(record.charSpace);
    if (
      !Number.isInteger(charSpace) ||
      charSpace < -MAX_REVISION_DOC_GRID_CHAR_SPACE ||
      charSpace > MAX_REVISION_DOC_GRID_CHAR_SPACE
    ) {
      return null;
    }
    next.charSpace = charSpace;
  }
  return next;
}

function validDocumentGridType(value: unknown): value is WorkDocumentGridType {
  return (
    value === 'default' ||
    value === 'lines' ||
    value === 'linesAndChars' ||
    value === 'snapToChars'
  );
}

const MAX_REVISION_DOC_GRID_CHAR_SPACE = 720 * 4096;

function orderedDocumentGrid(grid: WorkDocumentGrid): WorkDocumentGrid {
  return {
    type: grid.type,
    linePitch: grid.linePitch,
    ...(grid.charSpace !== undefined ? { charSpace: grid.charSpace } : {}),
  };
}

function normalizeRevisionLnNumType(
  value: unknown,
): WorkDocumentLnNumType | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some(
      (key) =>
        key !== 'countBy' &&
        key !== 'start' &&
        key !== 'distance' &&
        key !== 'restart',
    )
  ) {
    return null;
  }
  const next: WorkDocumentLnNumType = {};
  if ('countBy' in record) {
    const countBy = Number(record.countBy);
    if (!Number.isInteger(countBy) || countBy < 1 || countBy > 32_767) {
      return null;
    }
    next.countBy = countBy;
  }
  if ('start' in record) {
    const start = Number(record.start);
    if (!Number.isInteger(start) || start < 0 || start > 32_767) {
      return null;
    }
    next.start = start;
  }
  if ('distance' in record) {
    const distance = Number(record.distance);
    if (!Number.isInteger(distance) || distance < 1 || distance > 31_680) {
      return null;
    }
    next.distance = distance;
  }
  if ('restart' in record) {
    if (
      record.restart !== 'newPage' &&
      record.restart !== 'newSection' &&
      record.restart !== 'continuous'
    ) {
      return null;
    }
    next.restart = record.restart;
  }
  return next.countBy !== undefined ||
    next.start !== undefined ||
    next.distance !== undefined ||
    next.restart !== undefined
    ? orderedLnNumType(next)
    : null;
}

function orderedLnNumType(value: WorkDocumentLnNumType): WorkDocumentLnNumType {
  return {
    ...(value.countBy !== undefined ? { countBy: value.countBy } : {}),
    ...(value.start !== undefined ? { start: value.start } : {}),
    ...(value.distance !== undefined ? { distance: value.distance } : {}),
    ...(value.restart !== undefined ? { restart: value.restart } : {}),
  };
}

function normalizeRevisionPgNumType(
  value: unknown,
): WorkDocumentPgNumType | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some(
      (key) =>
        key !== 'fmt' &&
        key !== 'start' &&
        key !== 'chapStyle' &&
        key !== 'chapSep',
    )
  ) {
    return null;
  }
  const next: WorkDocumentPgNumType = {};
  if ('fmt' in record) {
    if (!validPgNumFmt(record.fmt)) return null;
    next.fmt = record.fmt;
  }
  if ('start' in record) {
    const start = Number(record.start);
    if (!Number.isInteger(start) || start < 0 || start > 32_767) {
      return null;
    }
    next.start = start;
  }
  if ('chapStyle' in record) {
    const chapStyle = Number(record.chapStyle);
    if (!Number.isInteger(chapStyle) || chapStyle < 1 || chapStyle > 9) {
      return null;
    }
    next.chapStyle = chapStyle;
  }
  if ('chapSep' in record) {
    if (!validPgNumChapSep(record.chapSep)) return null;
    next.chapSep = record.chapSep;
  }
  return next.fmt !== undefined ||
    next.start !== undefined ||
    next.chapStyle !== undefined ||
    next.chapSep !== undefined
    ? orderedPgNumType(next)
    : null;
}

function validPgNumFmt(value: unknown): value is WorkDocumentPgNumFmt {
  return (
    value === 'decimal' ||
    value === 'upperRoman' ||
    value === 'lowerRoman' ||
    value === 'upperLetter' ||
    value === 'lowerLetter'
  );
}

function validPgNumChapSep(
  value: unknown,
): value is NonNullable<WorkDocumentPgNumType['chapSep']> {
  return (
    value === 'hyphen' ||
    value === 'period' ||
    value === 'colon' ||
    value === 'emDash' ||
    value === 'enDash'
  );
}

function orderedPgNumType(value: WorkDocumentPgNumType): WorkDocumentPgNumType {
  return {
    ...(value.fmt !== undefined ? { fmt: value.fmt } : {}),
    ...(value.start !== undefined ? { start: value.start } : {}),
    ...(value.chapStyle !== undefined ? { chapStyle: value.chapStyle } : {}),
    ...(value.chapSep !== undefined ? { chapSep: value.chapSep } : {}),
  };
}

const FOOTNOTE_POS = new Set([
  'pageBottom',
  'beneathText',
  'sectEnd',
  'docEnd',
]);
const FOOTNOTE_NUM_RESTARTS = new Set(['continuous', 'eachSect', 'eachPage']);
const FOOTNOTE_NUM_FMTS = new Set([
  'decimal',
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
]);

export function serializeDocumentFootnotePr(
  value: WorkDocumentFootnotePr,
): string {
  return JSON.stringify(orderedFootnotePr(value));
}

export function parseDocumentFootnotePr(
  value: unknown,
): WorkDocumentFootnotePr | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    return normalizeRevisionFootnotePr(value) ?? undefined;
  }
  if (!value.length) return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return undefined;
  }
  return normalizeRevisionFootnotePr(parsed) ?? undefined;
}

function normalizeRevisionFootnotePr(
  value: unknown,
): WorkDocumentFootnotePr | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return null;
  }
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    keys.some(
      (key) =>
        key !== 'pos' &&
        key !== 'numFmt' &&
        key !== 'numStart' &&
        key !== 'numRestart',
    )
  ) {
    return null;
  }
  const next: WorkDocumentFootnotePr = {};
  if ('pos' in record && record.pos !== undefined) {
    if (typeof record.pos !== 'string' || !FOOTNOTE_POS.has(record.pos)) {
      return null;
    }
    next.pos = record.pos as WorkDocumentFootnotePos;
  }
  if ('numFmt' in record && record.numFmt !== undefined) {
    if (
      typeof record.numFmt !== 'string' ||
      !FOOTNOTE_NUM_FMTS.has(record.numFmt)
    ) {
      return null;
    }
    next.numFmt = record.numFmt as WorkDocumentPgNumFmt;
  }
  if ('numStart' in record && record.numStart !== undefined) {
    const start = Number(record.numStart);
    if (!Number.isInteger(start) || start < 0 || start > 32_767) return null;
    next.numStart = start;
  }
  if ('numRestart' in record && record.numRestart !== undefined) {
    if (
      typeof record.numRestart !== 'string' ||
      !FOOTNOTE_NUM_RESTARTS.has(record.numRestart)
    ) {
      return null;
    }
    next.numRestart = record.numRestart as WorkDocumentFootnoteNumRestart;
  }
  return orderedFootnotePr(next);
}

function orderedFootnotePr(
  value: WorkDocumentFootnotePr,
): WorkDocumentFootnotePr {
  return {
    ...(value.pos !== undefined ? { pos: value.pos } : {}),
    ...(value.numFmt !== undefined ? { numFmt: value.numFmt } : {}),
    ...(value.numStart !== undefined ? { numStart: value.numStart } : {}),
    ...(value.numRestart !== undefined ? { numRestart: value.numRestart } : {}),
  };
}

const ENDNOTE_POS = new Set(['sectEnd', 'docEnd']);

export function serializeDocumentEndnotePr(
  value: WorkDocumentEndnotePr,
): string {
  return JSON.stringify(orderedEndnotePr(value));
}

export function parseDocumentEndnotePr(
  value: unknown,
): WorkDocumentEndnotePr | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    return normalizeRevisionEndnotePr(value) ?? undefined;
  }
  if (!value.length) return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return undefined;
  }
  return normalizeRevisionEndnotePr(parsed) ?? undefined;
}

function normalizeRevisionEndnotePr(
  value: unknown,
): WorkDocumentEndnotePr | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return null;
  }
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    keys.some(
      (key) =>
        key !== 'pos' &&
        key !== 'numFmt' &&
        key !== 'numStart' &&
        key !== 'numRestart',
    )
  ) {
    return null;
  }
  const next: WorkDocumentEndnotePr = {};
  if ('pos' in record && record.pos !== undefined) {
    if (typeof record.pos !== 'string' || !ENDNOTE_POS.has(record.pos)) {
      return null;
    }
    next.pos = record.pos as WorkDocumentEndnotePos;
  }
  if ('numFmt' in record && record.numFmt !== undefined) {
    if (
      typeof record.numFmt !== 'string' ||
      !FOOTNOTE_NUM_FMTS.has(record.numFmt)
    ) {
      return null;
    }
    next.numFmt = record.numFmt as WorkDocumentPgNumFmt;
  }
  if ('numStart' in record && record.numStart !== undefined) {
    const start = Number(record.numStart);
    if (!Number.isInteger(start) || start < 0 || start > 32_767) return null;
    next.numStart = start;
  }
  if ('numRestart' in record && record.numRestart !== undefined) {
    if (
      typeof record.numRestart !== 'string' ||
      !FOOTNOTE_NUM_RESTARTS.has(record.numRestart)
    ) {
      return null;
    }
    next.numRestart = record.numRestart as WorkDocumentFootnoteNumRestart;
  }
  return orderedEndnotePr(next);
}

function orderedEndnotePr(value: WorkDocumentEndnotePr): WorkDocumentEndnotePr {
  return {
    ...(value.pos !== undefined ? { pos: value.pos } : {}),
    ...(value.numFmt !== undefined ? { numFmt: value.numFmt } : {}),
    ...(value.numStart !== undefined ? { numStart: value.numStart } : {}),
    ...(value.numRestart !== undefined ? { numRestart: value.numRestart } : {}),
  };
}

function isSectionVerticalAlign(
  value: unknown,
): value is WorkDocumentSectionVerticalAlign {
  return (
    value === 'top' ||
    value === 'center' ||
    value === 'both' ||
    value === 'bottom'
  );
}

function isSectionBreakAfter(
  value: unknown,
): value is WorkDocumentSectionBreakType {
  return (
    value === 'nextPage' ||
    value === 'nextColumn' ||
    value === 'continuous' ||
    value === 'evenPage' ||
    value === 'oddPage'
  );
}

function isSectionTextDirection(
  value: unknown,
): value is WorkDocumentSectionTextDirection {
  return (
    value === 'lrTb' ||
    value === 'tbRl' ||
    value === 'btLr' ||
    value === 'lrTbV' ||
    value === 'tbRlV' ||
    value === 'tbLrV'
  );
}

function orderedSnapshot(
  snapshot: DocumentSectionFormattingSnapshot,
): Record<string, unknown> {
  const ordered: Record<string, unknown> = {};
  if (snapshot.orientation) ordered.orientation = snapshot.orientation;
  if (snapshot.breakAfter !== undefined) {
    ordered.breakAfter = snapshot.breakAfter;
  }
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
  if (snapshot.pageBorders !== undefined) {
    ordered.pageBorders = orderedPageBorders(snapshot.pageBorders);
  }
  if (snapshot.differentFirstPage !== undefined) {
    ordered.differentFirstPage = snapshot.differentFirstPage;
  }
  if (snapshot.rtlGutter !== undefined) {
    ordered.rtlGutter = snapshot.rtlGutter;
  }
  if (snapshot.documentGrid) {
    ordered.documentGrid = orderedDocumentGrid(snapshot.documentGrid);
  }
  if (snapshot.lnNumType) {
    ordered.lnNumType = orderedLnNumType(snapshot.lnNumType);
  }
  if (snapshot.pgNumType) {
    ordered.pgNumType = orderedPgNumType(snapshot.pgNumType);
  }
  if (snapshot.formProt !== undefined) {
    ordered.formProt = snapshot.formProt;
  }
  if (snapshot.verticalAlign !== undefined) {
    ordered.verticalAlign = snapshot.verticalAlign;
  }
  if (snapshot.noEndnote !== undefined) {
    ordered.noEndnote = snapshot.noEndnote;
  }
  if (snapshot.textDirection !== undefined) {
    ordered.textDirection = snapshot.textDirection;
  }
  if (snapshot.bidi !== undefined) {
    ordered.bidi = snapshot.bidi;
  }
  if (snapshot.footnotePr !== undefined) {
    ordered.footnotePr = orderedFootnotePr(snapshot.footnotePr);
  }
  if (snapshot.endnotePr !== undefined) {
    ordered.endnotePr = orderedEndnotePr(snapshot.endnotePr);
  }
  if (snapshot.columns) {
    ordered.columns = snapshot.columns.custom
      ? {
          count: snapshot.columns.count,
          spacing: snapshot.columns.spacing,
          separator: snapshot.columns.separator,
          custom: snapshot.columns.custom.map((column) => ({
            widthPercent: column.widthPercent,
            spacing: column.spacing,
          })),
        }
      : {
          count: snapshot.columns.count,
          spacing: snapshot.columns.spacing,
          separator: snapshot.columns.separator,
        };
  }
  return ordered;
}
