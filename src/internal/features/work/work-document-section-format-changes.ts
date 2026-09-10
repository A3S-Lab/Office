import { normalizeDocumentColumns } from './work-document-columns';
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
  WorkDocumentGrid,
  WorkDocumentGridType,
  WorkDocumentLnNumType,
  WorkDocumentPaperSize,
  WorkDocumentPgNumFmt,
  WorkDocumentPgNumType,
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
export type DocumentSectionEqualColumnsSnapshot = DocumentSectionColumnsSnapshot;

/**
 * Prior snapshot for reviewable section-property revisions.
 * At least one of orientation, pageGeometry, pageMargins, paperSource,
 * columns, differentFirstPage, rtlGutter, documentGrid, lnNumType, pgNumType,
 * formProt, noEndnote, or verticalAlign must be present.
 */
export interface DocumentSectionFormattingSnapshot {
  orientation?: 'portrait' | 'landscape';
  pageGeometry?: WorkDocumentPageGeometry;
  pageMargins?: WorkDocumentPageMargins;
  paperSource?: WorkDocumentPaperSource;
  columns?: DocumentSectionColumnsSnapshot;
  differentFirstPage?: boolean;
  rtlGutter?: boolean;
  documentGrid?: WorkDocumentGrid;
  lnNumType?: WorkDocumentLnNumType;
  pgNumType?: WorkDocumentPgNumType;
  formProt?: boolean;
  noEndnote?: boolean;
  verticalAlign?: WorkDocumentSectionVerticalAlign;
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
  pageGeometry?: unknown;
  pageMargins?: unknown;
  paperSource?: unknown;
  columns?: unknown;
  differentFirstPage?: unknown;
  rtlGutter?: unknown;
  documentGrid?: unknown;
  lnNumType?: unknown;
  pgNumType?: unknown;
  formProt?: unknown;
  noEndnote?: unknown;
  verticalAlign?: unknown;
}): string {
  const snapshot = normalizeDocumentSectionFormattingSnapshot(attributes);
  if (!snapshot) {
    throw new Error(
      'Section-formatting snapshot requires orientation, pageGeometry, pageMargins, paperSource, columns, differentFirstPage, rtlGutter, documentGrid, lnNumType, pgNumType, formProt, noEndnote, or verticalAlign.',
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
        key !== 'rtlGutter' &&
        key !== 'documentGrid' &&
        key !== 'lnNumType' &&
        key !== 'pgNumType' &&
        key !== 'formProt' &&
        key !== 'noEndnote' &&
        key !== 'verticalAlign',
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
  pageGeometry?: unknown;
  pageMargins?: unknown;
  paperSource?: unknown;
  columns?: unknown;
  differentFirstPage?: unknown;
  rtlGutter?: unknown;
  documentGrid?: unknown;
  lnNumType?: unknown;
  pgNumType?: unknown;
  formProt?: unknown;
  noEndnote?: unknown;
  verticalAlign?: unknown;
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
  return snapshot.orientation ||
    snapshot.pageGeometry ||
    snapshot.pageMargins ||
    snapshot.paperSource ||
    snapshot.columns ||
    snapshot.differentFirstPage !== undefined ||
    snapshot.rtlGutter !== undefined ||
    snapshot.documentGrid ||
    snapshot.lnNumType ||
    snapshot.pgNumType ||
    snapshot.formProt !== undefined ||
    snapshot.noEndnote !== undefined ||
    snapshot.verticalAlign !== undefined
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
  if (formatting.documentGrid) {
    documentGridType = formatting.documentGrid.type;
    documentGridLinePitch = formatting.documentGrid.linePitch;
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
  if (formatting.pgNumType) {
    pgNumFmt = formatting.pgNumType.fmt ?? '';
    pgNumStart = formatting.pgNumType.start ?? null;
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
    documentGridType,
    documentGridLinePitch,
    lnNumCountBy,
    lnNumStart,
    lnNumDistance,
    lnNumRestart,
    pgNumFmt,
    pgNumStart,
    pageNumberStart,
    formProt,
    noEndnote,
    verticalAlign,
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
  documentGrid?: WorkDocumentGrid;
  lnNumType?: WorkDocumentLnNumType;
  pgNumType?: WorkDocumentPgNumType;
  formProt?: boolean;
  noEndnote?: boolean;
  verticalAlign?: WorkDocumentSectionVerticalAlign;
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
  if (customSource !== undefined && (!custom || custom.some((entry) => !entry))) {
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
  if (!keys.length || keys.some((key) => key !== 'type' && key !== 'linePitch')) {
    return null;
  }
  if (!('type' in record) || !('linePitch' in record)) return null;
  if (!validDocumentGridType(record.type)) return null;
  const linePitch = Number(record.linePitch);
  if (!Number.isFinite(linePitch) || linePitch <= 0) return null;
  return {
    type: record.type,
    linePitch: Math.min(720, Number(linePitch.toFixed(2))),
  };
}

function validDocumentGridType(value: unknown): value is WorkDocumentGridType {
  return (
    value === 'default' ||
    value === 'lines' ||
    value === 'linesAndChars' ||
    value === 'snapToChars'
  );
}

function orderedDocumentGrid(grid: WorkDocumentGrid): WorkDocumentGrid {
  return {
    type: grid.type,
    linePitch: grid.linePitch,
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
    keys.some((key) => key !== 'fmt' && key !== 'start')
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
  return next.fmt !== undefined || next.start !== undefined
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

function orderedPgNumType(value: WorkDocumentPgNumType): WorkDocumentPgNumType {
  return {
    ...(value.fmt !== undefined ? { fmt: value.fmt } : {}),
    ...(value.start !== undefined ? { start: value.start } : {}),
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
