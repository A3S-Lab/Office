import type {
  WorkDocumentPaperSize,
  WorkDocumentSectionLayout,
} from './work-types';

export interface WorkDocumentPageGeometry {
  width: number;
  height: number;
  orientation?: 'portrait' | 'landscape';
  code?: number;
}

export interface WorkDocumentPaperSource {
  first?: number;
  other?: number;
}

export interface ResolvedDocumentPageSize {
  width: number;
  height: number;
  nativeWidth: number;
  nativeHeight: number;
  widthPixels: number;
  heightPixels: number;
  widthPoints: number;
  heightPoints: number;
  orientation: 'portrait' | 'landscape';
  preset: WorkDocumentPaperSize;
  bounded: boolean;
}

export const DOCUMENT_PAPER_SIZE_PRESETS = [
  'a3',
  'a4',
  'a5',
  'letter',
  'legal',
  'tabloid',
] as const satisfies readonly Exclude<WorkDocumentPaperSize, 'custom'>[];

const PAGE_GEOMETRY_PROPERTY_SET = new Set([
  'width',
  'height',
  'orientation',
  'code',
]);
const PAPER_SOURCE_PROPERTY_SET = new Set(['first', 'other']);
const MAX_SERIALIZED_PAGE_SETUP = 4_096;
export const MAX_WORD_PAGE_TWIPS = 31_680;
export const MAX_WORD_PAPER_CODE = 118;
export const MAX_WORD_PAPER_SOURCE_CODE = 65_535;
const TWIPS_PER_MILLIMETER = 1_440 / 25.4;
const PIXELS_PER_MILLIMETER = 96 / 25.4;
const MINIMUM_RENDERED_PAGE_TWIPS = 1_440;
const PRESET_TOLERANCE_TWIPS = 2;

const PORTRAIT_PAGE_PRESETS: Record<
  Exclude<WorkDocumentPaperSize, 'custom'>,
  { width: number; height: number }
> = {
  a3: { width: 16_838, height: 23_811 },
  a4: { width: 11_906, height: 16_838 },
  a5: { width: 8_391, height: 11_906 },
  letter: { width: 12_240, height: 15_840 },
  legal: { width: 12_240, height: 20_160 },
  tabloid: { width: 15_840, height: 24_480 },
};
const PORTRAIT_PAGE_MILLIMETERS: Record<
  Exclude<WorkDocumentPaperSize, 'custom'>,
  { width: number; height: number }
> = {
  a3: { width: 297, height: 420 },
  a4: { width: 210, height: 297 },
  a5: { width: 148, height: 210 },
  letter: { width: 215.9, height: 279.4 },
  legal: { width: 215.9, height: 355.6 },
  tabloid: { width: 279.4, height: 431.8 },
};

export function normalizeDocumentPageGeometry(
  source: unknown,
): WorkDocumentPageGeometry | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null;
  }
  const record = source as Record<string, unknown>;
  if (Object.keys(record).some((key) => !PAGE_GEOMETRY_PROPERTY_SET.has(key))) {
    return null;
  }
  const width = pageTwips(record.width);
  const height = pageTwips(record.height);
  if (width === null || height === null) return null;
  const orientation = record.orientation;
  if (
    orientation !== undefined &&
    orientation !== 'portrait' &&
    orientation !== 'landscape'
  ) {
    return null;
  }
  const code = record.code;
  if (
    code !== undefined &&
    (!Number.isSafeInteger(code) ||
      Number(code) < 0 ||
      Number(code) > MAX_WORD_PAPER_CODE)
  ) {
    return null;
  }
  return {
    width,
    height,
    ...(orientation !== undefined ? { orientation } : {}),
    ...(code !== undefined ? { code: Number(code) } : {}),
  };
}

export function normalizeDocumentPaperSource(
  source: unknown,
): WorkDocumentPaperSource | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null;
  }
  const record = source as Record<string, unknown>;
  if (Object.keys(record).some((key) => !PAPER_SOURCE_PROPERTY_SET.has(key))) {
    return null;
  }
  const result: WorkDocumentPaperSource = {};
  for (const key of ['first', 'other'] as const) {
    const value = record[key];
    if (value === undefined) continue;
    if (
      !Number.isSafeInteger(value) ||
      Number(value) < 0 ||
      Number(value) > MAX_WORD_PAPER_SOURCE_CODE
    ) {
      return null;
    }
    result[key] = Number(value);
  }
  return result;
}

export function parseDocumentPageGeometry(
  source: unknown,
): WorkDocumentPageGeometry | null {
  return parseSerializedPageSetup(source, normalizeDocumentPageGeometry);
}

export function parseDocumentPaperSource(
  source: unknown,
): WorkDocumentPaperSource | null {
  return parseSerializedPageSetup(source, normalizeDocumentPaperSource);
}

export function serializeDocumentPageGeometry(
  source: unknown,
): string | undefined {
  const geometry = normalizeDocumentPageGeometry(source);
  if (!geometry) return undefined;
  return JSON.stringify({
    width: geometry.width,
    height: geometry.height,
    ...(geometry.orientation !== undefined
      ? { orientation: geometry.orientation }
      : {}),
    ...(geometry.code !== undefined ? { code: geometry.code } : {}),
  });
}

export function serializeDocumentPaperSource(
  source: unknown,
): string | undefined {
  const paperSource = normalizeDocumentPaperSource(source);
  if (!paperSource) return undefined;
  return JSON.stringify({
    ...(paperSource.first !== undefined ? { first: paperSource.first } : {}),
    ...(paperSource.other !== undefined ? { other: paperSource.other } : {}),
  });
}

export function normalizeDocumentPaperSize(
  source: unknown,
  fallback: WorkDocumentPaperSize = 'a4',
): WorkDocumentPaperSize {
  return source === 'a3' ||
    source === 'a5' ||
    source === 'letter' ||
    source === 'legal' ||
    source === 'tabloid' ||
    source === 'custom'
    ? source
    : fallback;
}

export function documentPageGeometryForLayout(
  layout: Pick<
    WorkDocumentSectionLayout,
    'orientation' | 'pageGeometry' | 'pageSize'
  >,
): WorkDocumentPageGeometry {
  return (
    normalizeDocumentPageGeometry(layout.pageGeometry) ??
    documentPageGeometryFromPreset(
      layout.pageSize === 'custom' ? 'a4' : layout.pageSize,
      layout.orientation,
    )
  );
}

export function documentPageGeometryFromPreset(
  preset: Exclude<WorkDocumentPaperSize, 'custom'>,
  orientation: 'portrait' | 'landscape',
): WorkDocumentPageGeometry {
  const portrait = PORTRAIT_PAGE_PRESETS[preset] ?? PORTRAIT_PAGE_PRESETS.a4;
  return orientation === 'landscape'
    ? {
        width: portrait.height,
        height: portrait.width,
        orientation: 'landscape',
      }
    : { width: portrait.width, height: portrait.height };
}

export function documentPaperSizeForGeometry(
  source: unknown,
): WorkDocumentPaperSize {
  const geometry = normalizeDocumentPageGeometry(source);
  if (!geometry) return 'custom';
  for (const preset of DOCUMENT_PAPER_SIZE_PRESETS) {
    const portrait = PORTRAIT_PAGE_PRESETS[preset];
    if (
      dimensionsMatch(geometry, portrait.width, portrait.height) ||
      dimensionsMatch(geometry, portrait.height, portrait.width)
    ) {
      return preset;
    }
  }
  return 'custom';
}

export function documentPageOrientationForGeometry(
  source: unknown,
  fallback: 'portrait' | 'landscape' = 'portrait',
): 'portrait' | 'landscape' {
  const geometry = normalizeDocumentPageGeometry(source);
  if (!geometry) return fallback;
  return (
    geometry.orientation ??
    (geometry.width > geometry.height ? 'landscape' : 'portrait')
  );
}

export function resolveDocumentPageSize(
  layout: Pick<
    WorkDocumentSectionLayout,
    'orientation' | 'pageGeometry' | 'pageSize'
  >,
): ResolvedDocumentPageSize {
  const nativeGeometry = normalizeDocumentPageGeometry(layout.pageGeometry);
  const geometry = nativeGeometry ?? documentPageGeometryForLayout(layout);
  const renderedWidth = Math.max(MINIMUM_RENDERED_PAGE_TWIPS, geometry.width);
  const renderedHeight = Math.max(MINIMUM_RENDERED_PAGE_TWIPS, geometry.height);
  const preset = documentPaperSizeForGeometry(geometry);
  const presetMillimeters =
    !nativeGeometry && preset !== 'custom'
      ? PORTRAIT_PAGE_MILLIMETERS[preset]
      : null;
  const width = presetMillimeters
    ? layout.orientation === 'landscape'
      ? presetMillimeters.height
      : presetMillimeters.width
    : twipsToMillimeters(renderedWidth);
  const height = presetMillimeters
    ? layout.orientation === 'landscape'
      ? presetMillimeters.width
      : presetMillimeters.height
    : twipsToMillimeters(renderedHeight);
  return {
    width,
    height,
    nativeWidth: twipsToMillimeters(geometry.width),
    nativeHeight: twipsToMillimeters(geometry.height),
    widthPixels: roundFour(width * PIXELS_PER_MILLIMETER),
    heightPixels: roundFour(height * PIXELS_PER_MILLIMETER),
    widthPoints: nativeGeometry
      ? roundFour(renderedWidth / 20)
      : roundFour((width * 72) / 25.4),
    heightPoints: nativeGeometry
      ? roundFour(renderedHeight / 20)
      : roundFour((height * 72) / 25.4),
    orientation: documentPageOrientationForGeometry(
      geometry,
      layout.orientation,
    ),
    preset,
    bounded:
      Boolean(nativeGeometry) &&
      (renderedWidth !== geometry.width || renderedHeight !== geometry.height),
  };
}

export function applyDocumentPageGeometry(
  layout: WorkDocumentSectionLayout,
  source: WorkDocumentPageGeometry,
): WorkDocumentSectionLayout {
  const pageGeometry = normalizeDocumentPageGeometry(source);
  if (!pageGeometry) return layout;
  return {
    ...layout,
    pageSize: documentPaperSizeForGeometry(pageGeometry),
    orientation: documentPageOrientationForGeometry(
      pageGeometry,
      layout.orientation,
    ),
    pageGeometry,
  };
}

export function updateDocumentPaperSizePreset(
  layout: WorkDocumentSectionLayout,
  pageSize: Exclude<WorkDocumentPaperSize, 'custom'>,
): WorkDocumentSectionLayout {
  const next = { ...layout, pageSize };
  delete next.pageGeometry;
  return next;
}

export function updateDocumentCustomPageMillimeters(
  layout: WorkDocumentSectionLayout,
  dimension: 'height' | 'width',
  millimeters: number,
): WorkDocumentSectionLayout {
  const current = documentPageGeometryForLayout(layout);
  const pageGeometry = {
    ...current,
    [dimension]: pageMillimetersToTwips(millimeters),
  };
  delete pageGeometry.code;
  return applyDocumentPageGeometry(layout, pageGeometry);
}

export function updateDocumentPageOrientation(
  layout: WorkDocumentSectionLayout,
  orientation: 'portrait' | 'landscape',
): WorkDocumentSectionLayout {
  const pageGeometry = normalizeDocumentPageGeometry(layout.pageGeometry);
  if (!pageGeometry) return { ...layout, orientation };
  if (orientation === layout.orientation) {
    return applyDocumentPageGeometry(layout, {
      ...pageGeometry,
      orientation,
    });
  }
  return applyDocumentPageGeometry(layout, {
    ...pageGeometry,
    width: pageGeometry.height,
    height: pageGeometry.width,
    orientation,
  });
}

export function reconcileDocumentPageSizeUpdate(
  previous: WorkDocumentSectionLayout,
  next: WorkDocumentSectionLayout,
): WorkDocumentSectionLayout {
  const before = normalizeDocumentPageGeometry(previous.pageGeometry);
  const after = normalizeDocumentPageGeometry(next.pageGeometry);
  if (JSON.stringify(before) !== JSON.stringify(after)) {
    return after
      ? applyDocumentPageGeometry(next, after)
      : normalizePreset(next);
  }
  if (previous.pageSize !== next.pageSize && next.pageSize !== 'custom') {
    return updateDocumentPaperSizePreset(next, next.pageSize);
  }
  if (previous.orientation !== next.orientation) {
    const basis = before
      ? { ...next, orientation: previous.orientation, pageGeometry: before }
      : { ...next, orientation: previous.orientation };
    return updateDocumentPageOrientation(basis, next.orientation);
  }
  return after ? applyDocumentPageGeometry(next, after) : normalizePreset(next);
}

export function pageMillimetersToTwips(value: number): number {
  if (!Number.isFinite(value)) return MINIMUM_RENDERED_PAGE_TWIPS;
  return Math.max(
    1,
    Math.min(MAX_WORD_PAGE_TWIPS, Math.round(value * TWIPS_PER_MILLIMETER)),
  );
}

export function pageTwipsToMillimeters(value: number): number {
  return twipsToMillimeters(value);
}

function normalizePreset(
  layout: WorkDocumentSectionLayout,
): WorkDocumentSectionLayout {
  const pageSize = normalizeDocumentPaperSize(layout.pageSize);
  const normalized = {
    ...layout,
    pageSize: pageSize === 'custom' ? 'a4' : pageSize,
  };
  delete normalized.pageGeometry;
  return normalized;
}

function pageTwips(value: unknown): number | null {
  return Number.isSafeInteger(value) &&
    Number(value) >= 1 &&
    Number(value) <= MAX_WORD_PAGE_TWIPS
    ? Number(value)
    : null;
}

function parseSerializedPageSetup<T>(
  source: unknown,
  normalize: (value: unknown) => T | null,
): T | null {
  if (typeof source !== 'string') return normalize(source);
  if (!source.trim() || source.length > MAX_SERIALIZED_PAGE_SETUP) return null;
  try {
    return normalize(JSON.parse(source));
  } catch {
    return null;
  }
}

function dimensionsMatch(
  geometry: WorkDocumentPageGeometry,
  width: number,
  height: number,
): boolean {
  return (
    Math.abs(geometry.width - width) <= PRESET_TOLERANCE_TWIPS &&
    Math.abs(geometry.height - height) <= PRESET_TOLERANCE_TWIPS
  );
}

function twipsToMillimeters(value: number): number {
  return roundFour(value / TWIPS_PER_MILLIMETER);
}

function roundFour(value: number): number {
  return Number(value.toFixed(4));
}
