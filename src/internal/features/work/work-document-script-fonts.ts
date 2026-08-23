export type WorkDocumentScriptFontSlot =
  | 'ascii'
  | 'highAnsi'
  | 'eastAsia'
  | 'complexScript';

export type WorkDocumentScriptFontHint = 'default' | 'eastAsia' | 'cs';

export type WorkDocumentThemeFont =
  | 'majorEastAsia'
  | 'majorBidi'
  | 'majorAscii'
  | 'majorHAnsi'
  | 'minorEastAsia'
  | 'minorBidi'
  | 'minorAscii'
  | 'minorHAnsi';

export interface WorkDocumentScriptFontFace {
  /** Exact direct w:rFonts value, when present. */
  name?: string;
  /** Exact theme reference, when present. */
  theme?: WorkDocumentThemeFont;
  /** Browser-resolved family retained independently from source identity. */
  resolved?: string;
}

export interface WorkDocumentScriptFonts {
  ascii?: WorkDocumentScriptFontFace;
  highAnsi?: WorkDocumentScriptFontFace;
  eastAsia?: WorkDocumentScriptFontFace;
  complexScript?: WorkDocumentScriptFontFace;
  hint?: WorkDocumentScriptFontHint;
}

export interface WorkDocumentScriptFontSegment {
  from: number;
  to: number;
  slot: WorkDocumentScriptFontSlot;
}

export interface WorkDocumentScriptFontPatch {
  latin?: string | null;
  eastAsia?: string | null;
  complexScript?: string | null;
}

export const DOCUMENT_SCRIPT_FONTS_ATTRIBUTE = 'data-office-script-fonts';
export const DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE =
  'data-office-script-font-slot';

const MAX_FONT_NAME_LENGTH = 127;
const MAX_SERIALIZED_SCRIPT_FONTS_LENGTH = 4_096;
const SCRIPT_FONT_KEYS = new Set([
  'ascii',
  'highAnsi',
  'eastAsia',
  'complexScript',
  'hint',
]);
const SCRIPT_FONT_FACE_KEYS = new Set(['name', 'theme', 'resolved']);
const SCRIPT_FONT_HINTS = new Set<WorkDocumentScriptFontHint>([
  'default',
  'eastAsia',
  'cs',
]);
const THEME_FONTS = new Set<WorkDocumentThemeFont>([
  'majorEastAsia',
  'majorBidi',
  'majorAscii',
  'majorHAnsi',
  'minorEastAsia',
  'minorBidi',
  'minorAscii',
  'minorHAnsi',
]);
const SLOT_FALLBACK_ORDER: Readonly<
  Record<WorkDocumentScriptFontSlot, readonly WorkDocumentScriptFontSlot[]>
> = {
  ascii: ['ascii', 'highAnsi', 'eastAsia', 'complexScript'],
  highAnsi: ['highAnsi', 'ascii', 'eastAsia', 'complexScript'],
  eastAsia: ['eastAsia', 'highAnsi', 'ascii', 'complexScript'],
  complexScript: ['complexScript', 'highAnsi', 'ascii', 'eastAsia'],
};
const NEUTRAL_SCRIPT_CHARACTER = /^[\p{Cc}\p{Cf}\p{M}\p{N}\p{P}\p{S}\p{Z}]$/u;

export function normalizeDocumentScriptFonts(
  source: unknown,
): WorkDocumentScriptFonts | null {
  if (!isRecordWithKeys(source, SCRIPT_FONT_KEYS)) return null;
  const normalized: WorkDocumentScriptFonts = {};
  for (const slot of scriptFontSlots) {
    if (source[slot] === undefined) continue;
    const face = normalizeDocumentScriptFontFace(source[slot]);
    if (!face) return null;
    normalized[slot] = face;
  }
  if (source.hint !== undefined) {
    const hint = normalizeDocumentScriptFontHint(source.hint);
    if (!hint) return null;
    normalized.hint = hint;
  }
  return Object.keys(normalized).length ? normalized : null;
}

export function serializeDocumentScriptFonts(source: unknown): string | null {
  const fonts = normalizeDocumentScriptFonts(source);
  return fonts ? JSON.stringify(fonts) : null;
}

export function parseDocumentScriptFonts(
  source: string | null | undefined,
): WorkDocumentScriptFonts | null {
  if (!source || source.length > MAX_SERIALIZED_SCRIPT_FONTS_LENGTH) {
    return null;
  }
  try {
    return normalizeDocumentScriptFonts(JSON.parse(source));
  } catch {
    return null;
  }
}

export function documentScriptFontsFromElement(
  element: HTMLElement,
): WorkDocumentScriptFonts | null {
  return parseDocumentScriptFonts(
    element.getAttribute(DOCUMENT_SCRIPT_FONTS_ATTRIBUTE),
  );
}

export function documentScriptFontSlotFromElement(
  element: HTMLElement,
): WorkDocumentScriptFontSlot | null {
  return normalizeDocumentScriptFontSlot(
    element.getAttribute(DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE),
  );
}

export function documentScriptFontsDomAttributes(
  source: unknown,
  slot: unknown,
): Record<string, string> {
  const fonts = normalizeDocumentScriptFonts(source);
  const normalizedSlot = normalizeDocumentScriptFontSlot(slot);
  if (!fonts) return {};
  const serialized = serializeDocumentScriptFonts(fonts);
  if (!serialized) return {};
  const family = documentScriptFontFamily(
    fonts,
    normalizedSlot ?? documentScriptFontSlotFromHint(fonts.hint),
  );
  return {
    [DOCUMENT_SCRIPT_FONTS_ATTRIBUTE]: serialized,
    ...(normalizedSlot
      ? { [DOCUMENT_SCRIPT_FONT_SLOT_ATTRIBUTE]: normalizedSlot }
      : {}),
    ...(family ? { style: `font-family: ${family}` } : {}),
  };
}

export function documentScriptFontFamily(
  source: unknown,
  slot: WorkDocumentScriptFontSlot,
): string | undefined {
  const fonts = normalizeDocumentScriptFonts(source);
  if (!fonts) return undefined;
  const families: string[] = [];
  const seen = new Set<string>();
  for (const candidate of documentScriptFontFallbackSlots(slot)) {
    const family = documentScriptFontFaceFamily(fonts[candidate]);
    const key = family?.toLocaleLowerCase();
    if (!family || !key || seen.has(key)) continue;
    seen.add(key);
    families.push(cssFontFamily(family));
  }
  return families.length ? families.join(', ') : undefined;
}

export function documentScriptFontFallbackSlots(
  slot: WorkDocumentScriptFontSlot,
): readonly WorkDocumentScriptFontSlot[] {
  return SLOT_FALLBACK_ORDER[slot];
}

export function documentScriptFontFamilyForRendering(
  source: unknown,
  slot: WorkDocumentScriptFontSlot,
  currentFontFamily?: unknown,
): string | undefined {
  const projected = documentScriptFontFamily(source, slot);
  const safeCurrent = safeCssFontFamilyList(currentFontFamily);
  if (!safeCurrent || !projected) return projected;
  const currentPrimary = documentFontNameFromCssFamily(safeCurrent);
  const projectedPrimary = documentFontNameFromCssFamily(projected);
  return currentPrimary &&
    projectedPrimary &&
    currentPrimary.toLocaleLowerCase() === projectedPrimary.toLocaleLowerCase()
    ? safeCurrent
    : projected;
}

export function documentScriptFontDirectFamily(
  source: unknown,
  slot: WorkDocumentScriptFontSlot,
): string | null {
  const fonts = normalizeDocumentScriptFonts(source);
  return fonts ? (documentScriptFontFaceFamily(fonts[slot]) ?? null) : null;
}

export function documentScriptFontsForAllText(
  fontFamily: unknown,
): WorkDocumentScriptFonts | null {
  const name = documentFontNameFromCssFamily(fontFamily);
  if (!name) return null;
  const face = { name, resolved: name };
  return {
    ascii: face,
    highAnsi: face,
    eastAsia: face,
    complexScript: face,
    hint: 'default',
  };
}

export function patchDocumentScriptFonts(
  source: unknown,
  patch: WorkDocumentScriptFontPatch,
  fallbackFontFamily?: unknown,
): WorkDocumentScriptFonts | null {
  const current =
    normalizeDocumentScriptFonts(source) ??
    documentScriptFontsForAllText(fallbackFontFamily) ??
    {};
  const next: WorkDocumentScriptFonts = { ...current };
  if (patch.latin !== undefined) {
    const face = directFontFace(patch.latin);
    if (face) {
      next.ascii = face;
      next.highAnsi = face;
    } else {
      delete next.ascii;
      delete next.highAnsi;
    }
  }
  if (patch.eastAsia !== undefined) {
    const face = directFontFace(patch.eastAsia);
    if (face) next.eastAsia = face;
    else delete next.eastAsia;
  }
  if (patch.complexScript !== undefined) {
    const face = directFontFace(patch.complexScript);
    if (face) next.complexScript = face;
    else delete next.complexScript;
  }
  return normalizeDocumentScriptFonts(next);
}

export function documentScriptFontSegments(
  text: string,
  hint: WorkDocumentScriptFontHint = 'default',
  forceComplexScript = false,
): WorkDocumentScriptFontSegment[] {
  if (!text) return [];
  if (forceComplexScript) {
    return [{ from: 0, to: text.length, slot: 'complexScript' }];
  }
  const characters: Array<{
    from: number;
    to: number;
    slot: WorkDocumentScriptFontSlot | null;
  }> = [];
  let offset = 0;
  for (const character of text) {
    const from = offset;
    offset += character.length;
    characters.push({
      from,
      to: offset,
      slot: strongDocumentScriptFontSlot(character),
    });
  }
  const fallback = documentScriptFontSlotFromHint(hint);
  let previous: WorkDocumentScriptFontSlot | null = null;
  for (let index = 0; index < characters.length; index += 1) {
    const entry = characters[index];
    if (!entry) continue;
    if (entry.slot) {
      previous = entry.slot;
      continue;
    }
    let next = previous;
    if (!next) {
      for (let cursor = index + 1; cursor < characters.length; cursor += 1) {
        const candidate = characters[cursor]?.slot;
        if (candidate) {
          next = candidate;
          break;
        }
      }
    }
    entry.slot = next ?? fallback;
  }
  const segments: WorkDocumentScriptFontSegment[] = [];
  for (const entry of characters) {
    const slot = entry.slot ?? fallback;
    const prior = segments[segments.length - 1];
    if (prior?.slot === slot && prior.to === entry.from) {
      prior.to = entry.to;
    } else {
      segments.push({ from: entry.from, to: entry.to, slot });
    }
  }
  return segments;
}

export function normalizeDocumentScriptFontSlot(
  value: unknown,
): WorkDocumentScriptFontSlot | null {
  return scriptFontSlots.includes(value as WorkDocumentScriptFontSlot)
    ? (value as WorkDocumentScriptFontSlot)
    : null;
}

export function normalizeDocumentScriptFontHint(
  value: unknown,
): WorkDocumentScriptFontHint | null {
  return SCRIPT_FONT_HINTS.has(value as WorkDocumentScriptFontHint)
    ? (value as WorkDocumentScriptFontHint)
    : null;
}

export function normalizeDocumentThemeFont(
  value: unknown,
): WorkDocumentThemeFont | null {
  return THEME_FONTS.has(value as WorkDocumentThemeFont)
    ? (value as WorkDocumentThemeFont)
    : null;
}

export function documentScriptFontSlotFromHint(
  hint: WorkDocumentScriptFontHint | undefined,
): WorkDocumentScriptFontSlot {
  if (hint === 'eastAsia') return 'eastAsia';
  if (hint === 'cs') return 'complexScript';
  return 'ascii';
}

export function documentFontNameFromCssFamily(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const source = value.trim();
  if (!source) return null;
  let family = '';
  const quote = source[0];
  if (quote === '"' || quote === "'") {
    let closed = false;
    for (let index = 1; index < source.length; index += 1) {
      const character = source[index];
      if (character === quote) {
        closed = true;
        break;
      }
      if (character !== '\\') {
        family += character;
        continue;
      }
      const decoded = decodeCssEscape(source, index + 1);
      if (!decoded) return null;
      family += decoded.value;
      index = decoded.end - 1;
    }
    if (!closed) return null;
  } else {
    family = source.split(',')[0] ?? '';
  }
  return normalizeDocumentFontName(family);
}

export function cssDocumentFontFamily(value: unknown): string | null {
  const family = normalizeDocumentFontName(value);
  return family ? cssFontFamily(family) : null;
}

export function normalizeDocumentFontName(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim();
  return normalized &&
    normalized.length <= MAX_FONT_NAME_LENGTH &&
    !/[\p{Cc}\p{Cs}]/u.test(normalized)
    ? normalized
    : null;
}

export const scriptFontSlots = [
  'ascii',
  'highAnsi',
  'eastAsia',
  'complexScript',
] as const satisfies readonly WorkDocumentScriptFontSlot[];

function normalizeDocumentScriptFontFace(
  source: unknown,
): WorkDocumentScriptFontFace | null {
  if (!isRecordWithKeys(source, SCRIPT_FONT_FACE_KEYS)) return null;
  const name =
    source.name === undefined
      ? undefined
      : normalizeDocumentFontName(source.name);
  const resolved =
    source.resolved === undefined
      ? undefined
      : normalizeDocumentFontName(source.resolved);
  const theme =
    source.theme === undefined
      ? undefined
      : normalizeDocumentThemeFont(source.theme);
  if (
    (source.name !== undefined && !name) ||
    (source.resolved !== undefined && !resolved) ||
    theme === null ||
    (!name && !theme && !resolved)
  ) {
    return null;
  }
  return {
    ...(name ? { name } : {}),
    ...(theme ? { theme } : {}),
    ...(resolved ? { resolved } : {}),
  };
}

function directFontFace(value: unknown): WorkDocumentScriptFontFace | null {
  if (value === null) return null;
  const name =
    documentFontNameFromCssFamily(value) ?? normalizeDocumentFontName(value);
  return name ? { name, resolved: name } : null;
}

function documentScriptFontFaceFamily(
  face: WorkDocumentScriptFontFace | undefined,
): string | undefined {
  return face?.resolved ?? face?.name;
}

function strongDocumentScriptFontSlot(
  character: string,
): WorkDocumentScriptFontSlot | null {
  if (NEUTRAL_SCRIPT_CHARACTER.test(character)) return null;
  const codePoint = character.codePointAt(0);
  if (codePoint === undefined) return null;
  if (isComplexScriptCodePoint(codePoint)) return 'complexScript';
  if (isEastAsianCodePoint(codePoint)) return 'eastAsia';
  return codePoint <= 0x7f ? 'ascii' : 'highAnsi';
}

function isComplexScriptCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x0590 && codePoint <= 0x08ff) ||
    (codePoint >= 0x0900 && codePoint <= 0x109f) ||
    (codePoint >= 0x1780 && codePoint <= 0x18af) ||
    (codePoint >= 0x1900 && codePoint <= 0x1cff) ||
    (codePoint >= 0xa800 && codePoint <= 0xa8ff) ||
    (codePoint >= 0xa980 && codePoint <= 0xa9df) ||
    (codePoint >= 0xaa00 && codePoint <= 0xaa7f) ||
    (codePoint >= 0xabc0 && codePoint <= 0xabff) ||
    (codePoint >= 0xfb1d && codePoint <= 0xfdff) ||
    (codePoint >= 0xfe70 && codePoint <= 0xfeff) ||
    (codePoint >= 0x10a00 && codePoint <= 0x10fff) ||
    (codePoint >= 0x11000 && codePoint <= 0x11fff) ||
    (codePoint >= 0x1e900 && codePoint <= 0x1edff) ||
    (codePoint >= 0x1ee00 && codePoint <= 0x1eeff)
  );
}

function isEastAsianCodePoint(codePoint: number): boolean {
  return (
    (codePoint >= 0x1100 && codePoint <= 0x11ff) ||
    (codePoint >= 0x2e80 && codePoint <= 0xa4cf) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7af) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xfe10 && codePoint <= 0xfe6f) ||
    (codePoint >= 0xff00 && codePoint <= 0xffef) ||
    (codePoint >= 0x20000 && codePoint <= 0x323af)
  );
}

function cssFontFamily(value: string): string {
  return /^(?:-?[\p{L}_])[\p{L}\p{N}_-]*$/u.test(value)
    ? value
    : `"${Array.from(value, cssStringCharacter).join('')}"`;
}

function cssStringCharacter(character: string): string {
  return /[\\":;{}<>]/u.test(character)
    ? `\\${character.codePointAt(0)?.toString(16)} `
    : character;
}

function safeCssFontFamilyList(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const source = value.trim();
  if (!source || source.length > 1_024 || /[;{}]/u.test(source)) return null;
  const tokens: string[] = [];
  let start = 0;
  let quote = '';
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index] ?? '';
    if (quote) {
      if (character === '\\') {
        index += 1;
        if (index >= source.length) return null;
      } else if (character === quote) {
        quote = '';
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character !== ',') continue;
    tokens.push(source.slice(start, index).trim());
    start = index + 1;
  }
  if (quote) return null;
  tokens.push(source.slice(start).trim());
  return tokens.length && tokens.every(validCssFontFamilyToken) ? source : null;
}

function validCssFontFamilyToken(value: string): boolean {
  if (!value) return false;
  if (value.startsWith('"')) {
    return (
      /^"(?:[^"\\\r\n\f]|\\(?:[\da-f]{1,6}\s?|[^\r\n\f]))*"$/iu.test(value) &&
      documentFontNameFromCssFamily(value) !== null
    );
  }
  if (value.startsWith("'")) {
    return (
      /^'(?:[^'\\\r\n\f]|\\(?:[\da-f]{1,6}\s?|[^\r\n\f]))*'$/iu.test(value) &&
      documentFontNameFromCssFamily(value) !== null
    );
  }
  return (
    /[\p{L}_]/u.test(value) &&
    /^[\p{L}_-][\p{L}\p{N}_ -]*$/u.test(value) &&
    !/^(?:inherit|initial|revert(?:-layer)?|unset)$/iu.test(value)
  );
}

function decodeCssEscape(
  source: string,
  start: number,
): { value: string; end: number } | null {
  if (start >= source.length) return null;
  const hexadecimal = /^[\da-f]{1,6}/i.exec(source.slice(start))?.[0];
  if (hexadecimal) {
    const codePoint = Number.parseInt(hexadecimal, 16);
    if (
      codePoint === 0 ||
      codePoint > 0x10ffff ||
      (codePoint >= 0xd800 && codePoint <= 0xdfff)
    ) {
      return null;
    }
    let end = start + hexadecimal.length;
    if (/\s/u.test(source[end] ?? '')) end += 1;
    return { value: String.fromCodePoint(codePoint), end };
  }
  const value = source[start];
  return !value || /[\r\n\f]/u.test(value) ? null : { value, end: start + 1 };
}

function isRecordWithKeys(
  value: unknown,
  allowed: ReadonlySet<string>,
): value is Record<string, unknown> {
  return (
    typeof value === 'object' &&
    value !== null &&
    !Array.isArray(value) &&
    Object.keys(value).every((key) => allowed.has(key))
  );
}
