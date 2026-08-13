import type {
  WorkDocumentMargins,
  WorkDocumentSectionLayout,
} from './work-types';

export const DOCUMENT_PAGE_MARGIN_KEYS = [
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
] as const;

export type WorkDocumentPageMarginKey =
  (typeof DOCUMENT_PAGE_MARGIN_KEYS)[number];
export type WorkDocumentPageMarginEdge = 'top' | 'right' | 'bottom' | 'left';
export type WorkDocumentPageMarginVerticalEdge = 'top' | 'bottom';
export type WorkDocumentPageMarginMode = 'clearChrome' | 'fromPageEdge';
export type WorkDocumentGutterPosition = 'left' | 'right' | 'top';

/**
 * Canonical Word page-margin geometry in twentieths of a point.
 *
 * Keeping the native integer unit avoids changing valid OOXML values merely
 * because the browser and the editing UI project the same geometry in pixels
 * and millimetres. Negative top or bottom values retain Word's explicit
 * header/footer-overlap semantics.
 */
export interface WorkDocumentPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  gutter: number;
  mirrorMargins?: boolean;
  gutterAtTop?: boolean;
  gutterOnRight?: boolean;
}

export interface ResolvedDocumentPageMargins {
  body: WorkDocumentMargins;
  unboundedBody: WorkDocumentMargins;
  headerDistance: number;
  footerDistance: number;
  gutter: number;
  gutterPosition: WorkDocumentGutterPosition;
  mirrorMargins: boolean;
  topMode: WorkDocumentPageMarginMode;
  bottomMode: WorkDocumentPageMarginMode;
  bounded: boolean;
}

const PAGE_MARGIN_PROPERTY_SET = new Set([
  ...DOCUMENT_PAGE_MARGIN_KEYS,
  'mirrorMargins',
  'gutterAtTop',
  'gutterOnRight',
]);
const MAX_SERIALIZED_PAGE_MARGINS = 4_096;
const MAX_WORD_TWIPS = 31_680;
const MIN_SIGNED_WORD_TWIPS = -2_147_483_648;
const MAX_SIGNED_WORD_TWIPS = 2_147_483_647;
const TWIPS_PER_MILLIMETER = 1_440 / 25.4;
const DEFAULT_HEADER_FOOTER_TWIPS = 708;
const MINIMUM_RENDERED_BODY_MILLIMETERS = 1;

export function normalizeDocumentPageMargins(
  source: unknown,
): WorkDocumentPageMargins | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null;
  }
  const record = source as Record<string, unknown>;
  if (Object.keys(record).some((key) => !PAGE_MARGIN_PROPERTY_SET.has(key))) {
    return null;
  }
  const result = {} as WorkDocumentPageMargins;
  for (const key of DOCUMENT_PAGE_MARGIN_KEYS) {
    const value = record[key];
    const signed = key === 'top' || key === 'bottom';
    if (
      !Number.isSafeInteger(value) ||
      (signed
        ? Number(value) < MIN_SIGNED_WORD_TWIPS ||
          Number(value) > MAX_SIGNED_WORD_TWIPS
        : Number(value) < 0 || Number(value) > MAX_WORD_TWIPS)
    ) {
      return null;
    }
    result[key] = normalizeNegativeZero(Number(value));
  }
  for (const key of [
    'mirrorMargins',
    'gutterAtTop',
    'gutterOnRight',
  ] as const) {
    const value = record[key];
    if (value === undefined) continue;
    if (typeof value !== 'boolean') return null;
    result[key] = value;
  }
  return result;
}

export function parseDocumentPageMargins(
  source: unknown,
): WorkDocumentPageMargins | null {
  if (typeof source !== 'string') return normalizeDocumentPageMargins(source);
  if (!source.trim() || source.length > MAX_SERIALIZED_PAGE_MARGINS) {
    return null;
  }
  try {
    return normalizeDocumentPageMargins(JSON.parse(source));
  } catch {
    return null;
  }
}

export function serializeDocumentPageMargins(
  source: unknown,
): string | undefined {
  const pageMargins = normalizeDocumentPageMargins(source);
  if (!pageMargins) return undefined;
  return JSON.stringify({
    top: pageMargins.top,
    right: pageMargins.right,
    bottom: pageMargins.bottom,
    left: pageMargins.left,
    header: pageMargins.header,
    footer: pageMargins.footer,
    gutter: pageMargins.gutter,
    ...(pageMargins.mirrorMargins !== undefined
      ? { mirrorMargins: pageMargins.mirrorMargins }
      : {}),
    ...(pageMargins.gutterAtTop !== undefined
      ? { gutterAtTop: pageMargins.gutterAtTop }
      : {}),
    ...(pageMargins.gutterOnRight !== undefined
      ? { gutterOnRight: pageMargins.gutterOnRight }
      : {}),
  });
}

export function documentPageMarginsForLayout(
  layout: Pick<WorkDocumentSectionLayout, 'margins' | 'pageMargins'>,
): WorkDocumentPageMargins {
  return (
    normalizeDocumentPageMargins(layout.pageMargins) ??
    documentPageMarginsFromBody(layout.margins)
  );
}

export function documentPageMarginsFromBody(
  margins: WorkDocumentMargins,
): WorkDocumentPageMargins {
  return {
    top: millimetersToTwips(margins.top),
    right: unsignedMillimetersToTwips(margins.right),
    bottom: millimetersToTwips(margins.bottom),
    left: unsignedMillimetersToTwips(margins.left),
    header: DEFAULT_HEADER_FOOTER_TWIPS,
    footer: DEFAULT_HEADER_FOOTER_TWIPS,
    gutter: 0,
  };
}

export function documentPageMarginBody(
  source: unknown,
  fallback: WorkDocumentMargins,
): WorkDocumentMargins {
  const pageMargins = normalizeDocumentPageMargins(source);
  if (!pageMargins) return { ...fallback };
  return {
    top: twipsToDisplayMillimeters(Math.abs(pageMargins.top)),
    right: twipsToDisplayMillimeters(pageMargins.right),
    bottom: twipsToDisplayMillimeters(Math.abs(pageMargins.bottom)),
    left: twipsToDisplayMillimeters(pageMargins.left),
  };
}

export function documentPageHorizontalMarginTwips(
  layout: Pick<WorkDocumentSectionLayout, 'margins' | 'pageMargins'>,
): number {
  const pageMargins = documentPageMarginsForLayout(layout);
  const topGutter =
    pageMargins.gutterAtTop === true && pageMargins.mirrorMargins !== true;
  return (
    pageMargins.left + pageMargins.right + (topGutter ? 0 : pageMargins.gutter)
  );
}

export function resolveDocumentPageMargins(
  layout: Pick<
    WorkDocumentSectionLayout,
    'margins' | 'orientation' | 'pageMargins' | 'pageSize'
  >,
  physicalPage = 1,
): ResolvedDocumentPageMargins {
  const pageMargins = documentPageMarginsForLayout(layout);
  const oddPage = Math.max(1, Math.trunc(physicalPage)) % 2 === 1;
  const mirrored = pageMargins.mirrorMargins === true;
  let left = twipsToMillimeters(pageMargins.left);
  let right = twipsToMillimeters(pageMargins.right);
  if (mirrored && !oddPage) [left, right] = [right, left];

  let top = twipsToMillimeters(Math.abs(pageMargins.top));
  const bottom = twipsToMillimeters(Math.abs(pageMargins.bottom));
  const gutter = twipsToMillimeters(pageMargins.gutter);
  const topGutter = pageMargins.gutterAtTop === true && !mirrored;
  let gutterPosition: WorkDocumentGutterPosition;
  if (topGutter) {
    gutterPosition = 'top';
    top += gutter;
  } else {
    const sourceSide = pageMargins.gutterOnRight === true ? 'right' : 'left';
    gutterPosition =
      mirrored && !oddPage ? oppositeSide(sourceSide) : sourceSide;
    if (gutterPosition === 'left') left += gutter;
    else right += gutter;
  }

  const unboundedBody = { top, right, bottom, left };
  const page = documentPageSizeMillimeters(layout);
  const horizontal = boundMarginPair(left, right, page.width);
  const vertical = boundMarginPair(top, bottom, page.height);
  const body = {
    top: vertical.start,
    right: horizontal.end,
    bottom: vertical.end,
    left: horizontal.start,
  };
  return {
    body,
    unboundedBody,
    headerDistance: Math.min(
      page.height,
      twipsToMillimeters(pageMargins.header),
    ),
    footerDistance: Math.min(
      page.height,
      twipsToMillimeters(pageMargins.footer),
    ),
    gutter,
    gutterPosition,
    mirrorMargins: mirrored,
    topMode: pageMargins.top < 0 ? 'fromPageEdge' : 'clearChrome',
    bottomMode: pageMargins.bottom < 0 ? 'fromPageEdge' : 'clearChrome',
    bounded: horizontal.bounded || vertical.bounded,
  };
}

export function updateDocumentPageMarginMillimeters(
  layout: WorkDocumentSectionLayout,
  key: WorkDocumentPageMarginKey,
  millimeters: number,
): WorkDocumentSectionLayout {
  const pageMargins = documentPageMarginsForLayout(layout);
  const signed = key === 'top' || key === 'bottom';
  const nextTwips = signed
    ? millimetersToTwips(Math.abs(millimeters))
    : unsignedMillimetersToTwips(millimeters);
  pageMargins[key] =
    signed && pageMargins[key] < 0 && nextTwips !== 0 ? -nextTwips : nextTwips;
  return layoutWithPageMargins(layout, pageMargins);
}

export function updateDocumentPageMarginMode(
  layout: WorkDocumentSectionLayout,
  edge: WorkDocumentPageMarginVerticalEdge,
  mode: WorkDocumentPageMarginMode,
): WorkDocumentSectionLayout {
  const pageMargins = documentPageMarginsForLayout(layout);
  const magnitude = Math.abs(pageMargins[edge]);
  pageMargins[edge] =
    mode === 'fromPageEdge' && magnitude > 0 ? -magnitude : magnitude;
  return layoutWithPageMargins(layout, pageMargins);
}

export function updateDocumentGutterPosition(
  layout: WorkDocumentSectionLayout,
  position: WorkDocumentGutterPosition,
): WorkDocumentSectionLayout {
  const pageMargins = documentPageMarginsForLayout(layout);
  pageMargins.gutterAtTop = position === 'top';
  pageMargins.gutterOnRight = position === 'right';
  if (position === 'top') pageMargins.mirrorMargins = false;
  return layoutWithPageMargins(layout, pageMargins);
}

export function updateDocumentMirrorMargins(
  layout: WorkDocumentSectionLayout,
  mirrored: boolean,
): WorkDocumentSectionLayout {
  const pageMargins = documentPageMarginsForLayout(layout);
  pageMargins.mirrorMargins = mirrored;
  if (mirrored) pageMargins.gutterAtTop = false;
  return layoutWithPageMargins(layout, pageMargins);
}

export function reconcileDocumentPageMarginUpdate(
  previous: WorkDocumentSectionLayout,
  next: WorkDocumentSectionLayout,
): WorkDocumentSectionLayout {
  const previousPageMargins = normalizeDocumentPageMargins(
    previous.pageMargins,
  );
  const nextPageMargins = normalizeDocumentPageMargins(next.pageMargins);
  if (!previousPageMargins && !nextPageMargins) return next;
  if (!nextPageMargins) {
    return {
      ...next,
      pageMargins: documentPageMarginsFromBody(next.margins),
    };
  }
  const pageMargins = { ...nextPageMargins };
  if (samePageMargins(previousPageMargins, nextPageMargins)) {
    for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
      if (next.margins[edge] === previous.margins[edge]) continue;
      const twips =
        edge === 'top' || edge === 'bottom'
          ? millimetersToTwips(Math.abs(next.margins[edge]))
          : unsignedMillimetersToTwips(next.margins[edge]);
      pageMargins[edge] =
        (edge === 'top' || edge === 'bottom') &&
        (previousPageMargins?.[edge] ?? 0) < 0 &&
        twips !== 0
          ? -twips
          : twips;
    }
  }
  return layoutWithPageMargins(next, pageMargins);
}

export function documentPageMarginGlobalsChanged(
  previous: WorkDocumentSectionLayout,
  next: WorkDocumentSectionLayout,
): boolean {
  const before = normalizeDocumentPageMargins(previous.pageMargins);
  const after = normalizeDocumentPageMargins(next.pageMargins);
  return (
    before?.mirrorMargins !== after?.mirrorMargins ||
    before?.gutterAtTop !== after?.gutterAtTop
  );
}

export function synchronizeDocumentPageMarginGlobals(
  layout: WorkDocumentSectionLayout,
  source: WorkDocumentPageMargins,
): WorkDocumentSectionLayout {
  const pageMargins = documentPageMarginsForLayout(layout);
  copyOptionalBoolean(pageMargins, source, 'mirrorMargins');
  copyOptionalBoolean(pageMargins, source, 'gutterAtTop');
  return layoutWithPageMargins(layout, pageMargins);
}

export function twipsToMillimeters(value: number): number {
  return roundFour(value / TWIPS_PER_MILLIMETER);
}

function twipsToDisplayMillimeters(value: number): number {
  return Math.round(twipsToMillimeters(value) * 10) / 10;
}

export function millimetersToTwips(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(
    MIN_SIGNED_WORD_TWIPS,
    Math.min(MAX_SIGNED_WORD_TWIPS, Math.round(value * TWIPS_PER_MILLIMETER)),
  );
}

function unsignedMillimetersToTwips(value: number): number {
  return Math.max(0, Math.min(MAX_WORD_TWIPS, millimetersToTwips(value)));
}

function layoutWithPageMargins(
  layout: WorkDocumentSectionLayout,
  source: WorkDocumentPageMargins,
): WorkDocumentSectionLayout {
  const pageMargins = normalizeDocumentPageMargins(source);
  if (!pageMargins) return layout;
  return {
    ...layout,
    margins: documentPageMarginBody(pageMargins, layout.margins),
    pageMargins,
  };
}

function samePageMargins(
  left: WorkDocumentPageMargins | null,
  right: WorkDocumentPageMargins | null,
): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function copyOptionalBoolean(
  target: WorkDocumentPageMargins,
  source: WorkDocumentPageMargins,
  key: 'gutterAtTop' | 'mirrorMargins',
): void {
  if (source[key] === undefined) delete target[key];
  else target[key] = source[key];
}

function oppositeSide(side: 'left' | 'right'): 'left' | 'right' {
  return side === 'left' ? 'right' : 'left';
}

function documentPageSizeMillimeters(
  layout: Pick<WorkDocumentSectionLayout, 'orientation' | 'pageSize'>,
): { width: number; height: number } {
  const portrait =
    layout.pageSize === 'letter'
      ? { width: 215.9, height: 279.4 }
      : { width: 210, height: 297 };
  return layout.orientation === 'landscape'
    ? { width: portrait.height, height: portrait.width }
    : portrait;
}

function boundMarginPair(
  start: number,
  end: number,
  extent: number,
): { start: number; end: number; bounded: boolean } {
  const maximum = Math.max(0, extent - MINIMUM_RENDERED_BODY_MILLIMETERS);
  const sum = start + end;
  if (sum <= maximum || sum <= 0) {
    return { start, end, bounded: false };
  }
  const scale = maximum / sum;
  return {
    start: roundFour(start * scale),
    end: roundFour(end * scale),
    bounded: true,
  };
}

function roundFour(value: number): number {
  return Number(value.toFixed(4));
}

function normalizeNegativeZero(value: number): number {
  return Object.is(value, -0) ? 0 : value;
}
