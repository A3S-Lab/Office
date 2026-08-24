import {
  DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE,
  type WorkDocumentScriptFontSlot,
} from './work-document-script-fonts';

export const DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE =
  'data-office-proofing-languages';
export const DOCUMENT_NO_PROOF_ATTRIBUTE = 'data-office-no-proof';

export interface WorkDocumentProofingLanguages {
  latin?: string;
  eastAsia?: string;
  bidi?: string;
}

export type WorkDocumentProofingLanguageSlot =
  keyof WorkDocumentProofingLanguages;

const PROOFING_LANGUAGE_KEYS = new Set<WorkDocumentProofingLanguageSlot>([
  'latin',
  'eastAsia',
  'bidi',
]);
const PROOFING_LANGUAGE_ORDER: readonly WorkDocumentProofingLanguageSlot[] = [
  'latin',
  'eastAsia',
  'bidi',
];
const MAX_LANGUAGE_TAG_LENGTH = 85;
const MAX_SERIALIZED_PROOFING_LANGUAGES_BYTES = 384;

export function normalizeDocumentLanguageTag(source: unknown): string | null {
  if (typeof source !== 'string') return null;
  if (
    !source ||
    source !== source.trim() ||
    source.length > MAX_LANGUAGE_TAG_LENGTH ||
    /[\p{Cc}\p{Cs}]/u.test(source)
  ) {
    return null;
  }
  return /^(?:x-none|[a-z0-9]{1,8}(?:-[a-z0-9]{1,8})*)$/iu.test(source)
    ? source
    : null;
}

export function normalizeDocumentProofingLanguages(
  source: unknown,
): WorkDocumentProofingLanguages | null {
  if (!isRecord(source)) return null;
  const keys = Object.keys(source);
  if (
    !keys.length ||
    keys.some(
      (key) =>
        !PROOFING_LANGUAGE_KEYS.has(key as WorkDocumentProofingLanguageSlot),
    )
  ) {
    return null;
  }
  const normalized: WorkDocumentProofingLanguages = {};
  for (const key of PROOFING_LANGUAGE_ORDER) {
    if (source[key] === undefined) continue;
    const language = normalizeDocumentLanguageTag(source[key]);
    if (!language) return null;
    normalized[key] = language;
  }
  return Object.keys(normalized).length ? normalized : null;
}

export function serializeDocumentProofingLanguages(
  source: unknown,
): string | undefined {
  const normalized = normalizeDocumentProofingLanguages(source);
  if (!normalized) return undefined;
  const serialized = JSON.stringify(normalized);
  return serialized.length <= MAX_SERIALIZED_PROOFING_LANGUAGES_BYTES
    ? serialized
    : undefined;
}

export function parseDocumentProofingLanguages(
  source: unknown,
): WorkDocumentProofingLanguages | null {
  if (typeof source !== 'string') {
    return normalizeDocumentProofingLanguages(source);
  }
  if (!source || source.length > MAX_SERIALIZED_PROOFING_LANGUAGES_BYTES) {
    return null;
  }
  try {
    const normalized = normalizeDocumentProofingLanguages(JSON.parse(source));
    return normalized && JSON.stringify(normalized) === source
      ? normalized
      : null;
  } catch {
    return null;
  }
}

export function normalizeDocumentNoProof(source: unknown): boolean | null {
  if (source === true || source === 'true' || source === '1') return true;
  if (source === false || source === 'false' || source === '0') return false;
  return null;
}

export function documentProofingLanguagesFromElement(
  element: HTMLElement,
): WorkDocumentProofingLanguages | null {
  return parseDocumentProofingLanguages(
    element.getAttribute(DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE),
  );
}

export function documentNoProofFromElement(
  element: HTMLElement,
): boolean | null {
  return normalizeDocumentNoProof(
    element.getAttribute(DOCUMENT_NO_PROOF_ATTRIBUTE),
  );
}

export function documentProofingLanguageForScript(
  source: unknown,
  slot?: WorkDocumentScriptFontSlot | null,
): string | undefined {
  const languages = normalizeDocumentProofingLanguages(source);
  if (!languages) return undefined;
  if (slot === 'eastAsia') return languages.eastAsia ?? languages.latin;
  if (slot === 'complexScript') return languages.bidi ?? languages.latin;
  return languages.latin ?? languages.eastAsia ?? languages.bidi;
}

export function documentProofingDomAttributes(
  languagesSource: unknown,
  noProofSource?: unknown,
  slot?: WorkDocumentScriptFontSlot | null,
): Record<string, string> {
  const languages = normalizeDocumentProofingLanguages(languagesSource);
  const serialized = serializeDocumentProofingLanguages(languages);
  const noProof = normalizeDocumentNoProof(noProofSource);
  const language = documentProofingLanguageForScript(languages, slot);
  return {
    ...(serialized
      ? { [DOCUMENT_PROOFING_LANGUAGES_ATTRIBUTE]: serialized }
      : {}),
    ...(noProof === null
      ? {}
      : { [DOCUMENT_NO_PROOF_ATTRIBUTE]: String(noProof) }),
    ...(slot ? { [DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE]: slot } : {}),
    ...(language && language !== 'x-none' ? { lang: language } : {}),
    ...(noProof === true ? { spellcheck: 'false' } : {}),
  };
}

export function patchDocumentProofingLanguages(
  source: unknown,
  patch: Partial<Record<WorkDocumentProofingLanguageSlot, string | null>>,
): WorkDocumentProofingLanguages | null {
  const current = normalizeDocumentProofingLanguages(source) ?? {};
  const next: WorkDocumentProofingLanguages = { ...current };
  for (const slot of PROOFING_LANGUAGE_ORDER) {
    const value = patch[slot];
    if (value === undefined) continue;
    if (value === null) {
      delete next[slot];
      continue;
    }
    const language = normalizeDocumentLanguageTag(value);
    if (!language) return null;
    next[slot] = language;
  }
  return Object.keys(next).length ? next : null;
}

function isRecord(source: unknown): source is Record<string, unknown> {
  return (
    typeof source === 'object' && source !== null && !Array.isArray(source)
  );
}
