export const DOCUMENT_OPEN_TYPE_ATTRIBUTE = 'data-office-opentype-features';

export const DOCUMENT_OPEN_TYPE_LIGATURES = [
  'none',
  'standard',
  'contextual',
  'historical',
  'discretional',
  'standardContextual',
  'standardHistorical',
  'contextualHistorical',
  'standardDiscretional',
  'contextualDiscretional',
  'historicalDiscretional',
  'standardContextualHistorical',
  'standardContextualDiscretional',
  'standardHistoricalDiscretional',
  'contextualHistoricalDiscretional',
  'all',
] as const;

export type WorkDocumentOpenTypeLigatures =
  (typeof DOCUMENT_OPEN_TYPE_LIGATURES)[number];

export const DOCUMENT_OPEN_TYPE_NUMBER_FORMS = [
  'default',
  'lining',
  'oldStyle',
] as const;

export type WorkDocumentOpenTypeNumberForm =
  (typeof DOCUMENT_OPEN_TYPE_NUMBER_FORMS)[number];

export const DOCUMENT_OPEN_TYPE_NUMBER_SPACINGS = [
  'default',
  'proportional',
  'tabular',
] as const;

export type WorkDocumentOpenTypeNumberSpacing =
  (typeof DOCUMENT_OPEN_TYPE_NUMBER_SPACINGS)[number];

export interface WorkDocumentOpenTypeFeatures {
  ligatures?: WorkDocumentOpenTypeLigatures;
  numberForm?: WorkDocumentOpenTypeNumberForm;
  numberSpacing?: WorkDocumentOpenTypeNumberSpacing;
  stylisticSets?: number[];
  contextualAlternates?: boolean;
}

export interface WorkDocumentOpenTypeFeaturePatch {
  ligatures?: WorkDocumentOpenTypeLigatures | null;
  numberForm?: WorkDocumentOpenTypeNumberForm | null;
  numberSpacing?: WorkDocumentOpenTypeNumberSpacing | null;
  stylisticSets?: readonly number[] | null;
  contextualAlternates?: boolean | null;
}

export interface WorkDocumentOpenTypeCssProperties {
  fontFeatureSettings?: string;
  fontVariantLigatures?: string;
  fontVariantNumeric?: string;
}

const OPEN_TYPE_FEATURE_KEYS = [
  'ligatures',
  'numberForm',
  'numberSpacing',
  'stylisticSets',
  'contextualAlternates',
] as const;
const OPEN_TYPE_FEATURE_KEY_SET = new Set<string>(OPEN_TYPE_FEATURE_KEYS);
const LIGATURE_SET = new Set<string>(DOCUMENT_OPEN_TYPE_LIGATURES);
const NUMBER_FORM_SET = new Set<string>(DOCUMENT_OPEN_TYPE_NUMBER_FORMS);
const NUMBER_SPACING_SET = new Set<string>(DOCUMENT_OPEN_TYPE_NUMBER_SPACINGS);
const MAX_STYLISTIC_SET_ENTRIES = 4_096;
const MIN_STYLISTIC_SET_ID = 1;
const MAX_STYLISTIC_SET_ID = 20;
const MAX_SERIALIZED_OPEN_TYPE_BYTES = 1_024;

const LIGATURE_STANDARD = 1;
const LIGATURE_CONTEXTUAL = 2;
const LIGATURE_HISTORICAL = 4;
const LIGATURE_DISCRETIONAL = 8;
const LIGATURE_FLAGS = new Map<WorkDocumentOpenTypeLigatures, number>([
  ['none', 0],
  ['standard', LIGATURE_STANDARD],
  ['contextual', LIGATURE_CONTEXTUAL],
  ['historical', LIGATURE_HISTORICAL],
  ['discretional', LIGATURE_DISCRETIONAL],
  ['standardContextual', LIGATURE_STANDARD | LIGATURE_CONTEXTUAL],
  ['standardHistorical', LIGATURE_STANDARD | LIGATURE_HISTORICAL],
  ['contextualHistorical', LIGATURE_CONTEXTUAL | LIGATURE_HISTORICAL],
  ['standardDiscretional', LIGATURE_STANDARD | LIGATURE_DISCRETIONAL],
  ['contextualDiscretional', LIGATURE_CONTEXTUAL | LIGATURE_DISCRETIONAL],
  ['historicalDiscretional', LIGATURE_HISTORICAL | LIGATURE_DISCRETIONAL],
  [
    'standardContextualHistorical',
    LIGATURE_STANDARD | LIGATURE_CONTEXTUAL | LIGATURE_HISTORICAL,
  ],
  [
    'standardContextualDiscretional',
    LIGATURE_STANDARD | LIGATURE_CONTEXTUAL | LIGATURE_DISCRETIONAL,
  ],
  [
    'standardHistoricalDiscretional',
    LIGATURE_STANDARD | LIGATURE_HISTORICAL | LIGATURE_DISCRETIONAL,
  ],
  [
    'contextualHistoricalDiscretional',
    LIGATURE_CONTEXTUAL | LIGATURE_HISTORICAL | LIGATURE_DISCRETIONAL,
  ],
  [
    'all',
    LIGATURE_STANDARD |
      LIGATURE_CONTEXTUAL |
      LIGATURE_HISTORICAL |
      LIGATURE_DISCRETIONAL,
  ],
]);

export function normalizeDocumentOpenTypeLigatures(
  value: unknown,
): WorkDocumentOpenTypeLigatures | null {
  return typeof value === 'string' && LIGATURE_SET.has(value)
    ? (value as WorkDocumentOpenTypeLigatures)
    : null;
}

export function normalizeDocumentOpenTypeNumberForm(
  value: unknown,
): WorkDocumentOpenTypeNumberForm | null {
  return typeof value === 'string' && NUMBER_FORM_SET.has(value)
    ? (value as WorkDocumentOpenTypeNumberForm)
    : null;
}

export function normalizeDocumentOpenTypeNumberSpacing(
  value: unknown,
): WorkDocumentOpenTypeNumberSpacing | null {
  return typeof value === 'string' && NUMBER_SPACING_SET.has(value)
    ? (value as WorkDocumentOpenTypeNumberSpacing)
    : null;
}

export function normalizeDocumentOpenTypeStylisticSets(
  value: unknown,
): number[] | null {
  if (!Array.isArray(value) || value.length > MAX_STYLISTIC_SET_ENTRIES) {
    return null;
  }
  const result: number[] = [];
  const seen = new Set<number>();
  for (const candidate of value) {
    if (
      !Number.isSafeInteger(candidate) ||
      candidate < MIN_STYLISTIC_SET_ID ||
      candidate > MAX_STYLISTIC_SET_ID
    ) {
      return null;
    }
    if (!seen.has(candidate)) {
      seen.add(candidate);
      result.push(candidate);
    }
  }
  return result;
}

export function normalizeDocumentOpenTypeFeatures(
  value: unknown,
): WorkDocumentOpenTypeFeatures | null {
  if (!isRecord(value)) return null;
  if (Object.keys(value).some((key) => !OPEN_TYPE_FEATURE_KEY_SET.has(key))) {
    return null;
  }
  const result: WorkDocumentOpenTypeFeatures = {};
  if (value.ligatures !== undefined) {
    const ligatures = normalizeDocumentOpenTypeLigatures(value.ligatures);
    if (ligatures === null) return null;
    result.ligatures = ligatures;
  }
  if (value.numberForm !== undefined) {
    const numberForm = normalizeDocumentOpenTypeNumberForm(value.numberForm);
    if (numberForm === null) return null;
    result.numberForm = numberForm;
  }
  if (value.numberSpacing !== undefined) {
    const numberSpacing = normalizeDocumentOpenTypeNumberSpacing(
      value.numberSpacing,
    );
    if (numberSpacing === null) return null;
    result.numberSpacing = numberSpacing;
  }
  if (value.stylisticSets !== undefined) {
    const stylisticSets = normalizeDocumentOpenTypeStylisticSets(
      value.stylisticSets,
    );
    if (stylisticSets === null) return null;
    result.stylisticSets = stylisticSets;
  }
  if (value.contextualAlternates !== undefined) {
    if (typeof value.contextualAlternates !== 'boolean') return null;
    result.contextualAlternates = value.contextualAlternates;
  }
  return Object.keys(result).length ? result : null;
}

export function serializeDocumentOpenTypeFeatures(
  value: unknown,
): string | null {
  const features = normalizeDocumentOpenTypeFeatures(value);
  return features ? JSON.stringify(features) : null;
}

export function parseDocumentOpenTypeFeatures(
  value: unknown,
): WorkDocumentOpenTypeFeatures | null {
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > MAX_SERIALIZED_OPEN_TYPE_BYTES
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return null;
  }
  const features = normalizeDocumentOpenTypeFeatures(parsed);
  return features && JSON.stringify(features) === value ? features : null;
}

export function documentOpenTypeFeaturesFromElement(
  element: HTMLElement,
): WorkDocumentOpenTypeFeatures | null {
  return parseDocumentOpenTypeFeatures(
    element.getAttribute(DOCUMENT_OPEN_TYPE_ATTRIBUTE),
  );
}

export function documentOpenTypeDomAttributes(
  value: unknown,
): Record<string, string> {
  const features = normalizeDocumentOpenTypeFeatures(value);
  const serialized = features
    ? serializeDocumentOpenTypeFeatures(features)
    : null;
  if (!features || !serialized) return {};
  const style = documentOpenTypeCss(features);
  return {
    [DOCUMENT_OPEN_TYPE_ATTRIBUTE]: serialized,
    ...(style ? { style } : {}),
  };
}

export function documentOpenTypeCss(value: unknown): string {
  const properties = documentOpenTypeCssProperties(value);
  return [
    properties.fontFeatureSettings
      ? `font-feature-settings: ${properties.fontFeatureSettings}`
      : '',
    properties.fontVariantLigatures
      ? `font-variant-ligatures: ${properties.fontVariantLigatures}`
      : '',
    properties.fontVariantNumeric
      ? `font-variant-numeric: ${properties.fontVariantNumeric}`
      : '',
  ]
    .filter(Boolean)
    .join('; ');
}

export function documentOpenTypeCssProperties(
  value: unknown,
): WorkDocumentOpenTypeCssProperties {
  const features = normalizeDocumentOpenTypeFeatures(value);
  if (!features) return {};
  const featureSettings = openTypeFeatureSettingsValue(
    features.ligatures,
    features.stylisticSets,
  );
  const contextualAlternates = contextualAlternateValue(
    features.contextualAlternates,
  );
  const numberVariants = numberStylesValue(
    features.numberForm,
    features.numberSpacing,
  );
  return {
    ...(featureSettings ? { fontFeatureSettings: featureSettings } : {}),
    ...(contextualAlternates
      ? { fontVariantLigatures: contextualAlternates }
      : {}),
    ...(numberVariants ? { fontVariantNumeric: numberVariants } : {}),
  };
}

export function isDocumentOpenTypeFeaturePatch(
  value: unknown,
): value is WorkDocumentOpenTypeFeaturePatch {
  if (!isRecord(value) || !Object.keys(value).length) return false;
  if (Object.keys(value).some((key) => !OPEN_TYPE_FEATURE_KEY_SET.has(key))) {
    return false;
  }
  if (
    value.ligatures !== undefined &&
    value.ligatures !== null &&
    normalizeDocumentOpenTypeLigatures(value.ligatures) === null
  ) {
    return false;
  }
  if (
    value.numberForm !== undefined &&
    value.numberForm !== null &&
    normalizeDocumentOpenTypeNumberForm(value.numberForm) === null
  ) {
    return false;
  }
  if (
    value.numberSpacing !== undefined &&
    value.numberSpacing !== null &&
    normalizeDocumentOpenTypeNumberSpacing(value.numberSpacing) === null
  ) {
    return false;
  }
  if (
    value.stylisticSets !== undefined &&
    value.stylisticSets !== null &&
    normalizeDocumentOpenTypeStylisticSets(value.stylisticSets) === null
  ) {
    return false;
  }
  return (
    value.contextualAlternates === undefined ||
    value.contextualAlternates === null ||
    typeof value.contextualAlternates === 'boolean'
  );
}

export function patchDocumentOpenTypeFeatures(
  source: unknown,
  patch: WorkDocumentOpenTypeFeaturePatch,
): WorkDocumentOpenTypeFeatures | null {
  if (!isDocumentOpenTypeFeaturePatch(patch)) return null;
  const result: WorkDocumentOpenTypeFeatures = {
    ...(normalizeDocumentOpenTypeFeatures(source) ?? {}),
  };
  for (const key of OPEN_TYPE_FEATURE_KEYS) {
    if (!(key in patch)) continue;
    const value = patch[key];
    if (value === null) delete result[key];
    else if (value !== undefined) {
      if (key === 'stylisticSets') {
        result.stylisticSets =
          normalizeDocumentOpenTypeStylisticSets(value) ?? [];
      } else {
        Object.assign(result, { [key]: value });
      }
    }
  }
  return normalizeDocumentOpenTypeFeatures(result);
}

function openTypeFeatureSettingsValue(
  ligatures: WorkDocumentOpenTypeLigatures | undefined,
  stylisticSets: number[] | undefined,
): string {
  if (ligatures === undefined && stylisticSets === undefined) return '';
  const values: string[] = [];
  const flags =
    ligatures === undefined ? undefined : LIGATURE_FLAGS.get(ligatures);
  if (flags !== undefined) {
    values.push(
      `"liga" ${flags & LIGATURE_STANDARD ? 1 : 0}`,
      `"clig" ${flags & LIGATURE_CONTEXTUAL ? 1 : 0}`,
      `"hlig" ${flags & LIGATURE_HISTORICAL ? 1 : 0}`,
      `"dlig" ${flags & LIGATURE_DISCRETIONAL ? 1 : 0}`,
    );
  }
  if (stylisticSets !== undefined) {
    values.push(
      ...stylisticSets.map((id) => `"ss${String(id).padStart(2, '0')}" 1`),
    );
  }
  return values.length ? values.join(', ') : 'normal';
}

function contextualAlternateValue(value: boolean | undefined): string {
  if (value === undefined) return '';
  return value ? 'contextual' : 'no-contextual';
}

function numberStylesValue(
  form: WorkDocumentOpenTypeNumberForm | undefined,
  spacing: WorkDocumentOpenTypeNumberSpacing | undefined,
): string {
  if (form === undefined && spacing === undefined) return '';
  const values = [
    form === 'lining'
      ? 'lining-nums'
      : form === 'oldStyle'
        ? 'oldstyle-nums'
        : '',
    spacing === 'proportional'
      ? 'proportional-nums'
      : spacing === 'tabular'
        ? 'tabular-nums'
        : '',
  ].filter(Boolean);
  return values.length ? values.join(' ') : 'normal';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
