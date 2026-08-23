export const DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE =
  'data-office-legacy-text-outline';
export const DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE =
  'data-office-legacy-text-shadow';
export const DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE =
  'data-office-legacy-text-emboss';
export const DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE =
  'data-office-legacy-text-imprint';

export const DOCUMENT_LEGACY_TEXT_EFFECT_NAMES = [
  'outline',
  'shadow',
  'emboss',
  'imprint',
] as const;

export type WorkDocumentLegacyTextEffectName =
  (typeof DOCUMENT_LEGACY_TEXT_EFFECT_NAMES)[number];

export interface WorkDocumentLegacyTextEffects {
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
}

export interface WorkDocumentLegacyTextEffectCss {
  WebkitTextFillColor?: string;
  WebkitTextStroke?: string;
  paintOrder?: string;
  textShadow?: string;
}

const ATTRIBUTE_BY_EFFECT: Readonly<
  Record<WorkDocumentLegacyTextEffectName, string>
> = {
  outline: DOCUMENT_LEGACY_TEXT_OUTLINE_ATTRIBUTE,
  shadow: DOCUMENT_LEGACY_TEXT_SHADOW_ATTRIBUTE,
  emboss: DOCUMENT_LEGACY_TEXT_EMBOSS_ATTRIBUTE,
  imprint: DOCUMENT_LEGACY_TEXT_IMPRINT_ATTRIBUTE,
};

const TEXT_STYLE_ATTRIBUTE_BY_EFFECT: Readonly<
  Record<WorkDocumentLegacyTextEffectName, string>
> = {
  outline: 'legacyTextOutline',
  shadow: 'legacyTextShadow',
  emboss: 'legacyTextEmboss',
  imprint: 'legacyTextImprint',
};

export function normalizeDocumentLegacyTextEffect(
  value: unknown,
): boolean | null {
  if (value === true || value === 'true' || value === '1') return true;
  if (value === false || value === 'false' || value === '0') return false;
  return null;
}

export function normalizeDocumentLegacyTextEffects(
  value: unknown,
): WorkDocumentLegacyTextEffects | null {
  if (!isRecord(value)) return null;
  const allowed = new Set<string>(DOCUMENT_LEGACY_TEXT_EFFECT_NAMES);
  if (Object.keys(value).some((key) => !allowed.has(key))) return null;
  const effects: WorkDocumentLegacyTextEffects = {};
  for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
    if (!Object.hasOwn(value, name)) continue;
    if (typeof value[name] !== 'boolean') return null;
    effects[name] = value[name];
  }
  return documentLegacyTextEffectsConflict(effects) ? null : effects;
}

export function documentLegacyTextEffectsConflict(
  effects: WorkDocumentLegacyTextEffects,
): boolean {
  if (
    effects.emboss === true &&
    (effects.outline === true ||
      effects.shadow === true ||
      effects.imprint === true)
  ) {
    return true;
  }
  return (
    effects.imprint === true &&
    (effects.outline === true ||
      effects.shadow === true ||
      effects.emboss === true)
  );
}

export function documentLegacyTextEffectFromElement(
  element: HTMLElement,
  effect: WorkDocumentLegacyTextEffectName,
): boolean | null {
  return normalizeDocumentLegacyTextEffect(
    element.getAttribute(ATTRIBUTE_BY_EFFECT[effect]),
  );
}

export function documentLegacyTextEffectsFromElement(
  element: HTMLElement,
): WorkDocumentLegacyTextEffects | null {
  const effects: WorkDocumentLegacyTextEffects = {};
  for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
    const attribute = ATTRIBUTE_BY_EFFECT[name];
    const value = documentLegacyTextEffectFromElement(element, name);
    if (element.hasAttribute(attribute) && value === null) return null;
    if (value !== null) effects[name] = value;
  }
  return documentLegacyTextEffectsConflict(effects) ? null : effects;
}

export function documentLegacyTextEffectsFromTextStyleAttributes(
  attributes: Record<string, unknown>,
): WorkDocumentLegacyTextEffects | null {
  const effects: WorkDocumentLegacyTextEffects = {};
  for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
    const rawValue = attributes[TEXT_STYLE_ATTRIBUTE_BY_EFFECT[name]];
    if (rawValue === null || rawValue === undefined) continue;
    const value = normalizeDocumentLegacyTextEffect(rawValue);
    if (value === null) return null;
    if (value !== null) effects[name] = value;
  }
  return documentLegacyTextEffectsConflict(effects) ? null : effects;
}

export function documentLegacyTextEffectsDomAttributes(
  value: WorkDocumentLegacyTextEffects,
): Record<string, string> {
  const effects = normalizeDocumentLegacyTextEffects(value);
  if (!effects) return {};
  return Object.fromEntries(
    DOCUMENT_LEGACY_TEXT_EFFECT_NAMES.flatMap((name) =>
      effects[name] === undefined
        ? []
        : [[ATTRIBUTE_BY_EFFECT[name], String(effects[name])]],
    ),
  );
}

export function documentLegacyTextEffectsCss(
  value: WorkDocumentLegacyTextEffects,
): WorkDocumentLegacyTextEffectCss {
  const effects = normalizeDocumentLegacyTextEffects(value);
  if (!effects) return {};
  if (effects.emboss) {
    return {
      textShadow:
        '-0.045em -0.045em 0 rgba(255, 255, 255, 0.9), 0.045em 0.045em 0 rgba(0, 0, 0, 0.52)',
    };
  }
  if (effects.imprint) {
    return {
      textShadow:
        '0.045em 0.045em 0 rgba(255, 255, 255, 0.9), -0.045em -0.045em 0 rgba(0, 0, 0, 0.52)',
    };
  }
  return {
    ...(effects.outline
      ? {
          WebkitTextFillColor: 'transparent',
          WebkitTextStroke: '0.045em currentColor',
          paintOrder: 'stroke fill',
        }
      : {}),
    ...(effects.shadow ? { textShadow: '0.08em 0.08em 0 currentColor' } : {}),
  };
}

export function documentLegacyTextStyleAttributeName(
  effect: WorkDocumentLegacyTextEffectName,
): string {
  return TEXT_STYLE_ATTRIBUTE_BY_EFFECT[effect];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
