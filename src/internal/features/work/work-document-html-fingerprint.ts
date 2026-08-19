const DOCUMENT_HTML_FINGERPRINT_VERSION = 'p1';
const DOCUMENT_HTML_HASH_BASE = 0x01000193;
const LEGACY_FNV_OFFSET = 0x811c9dc5;

export interface DocumentHtmlFingerprintSegment {
  hash: number;
  length: number;
  power: number;
}

/**
 * Hashes one HTML segment with a polynomial that can be combined without
 * rescanning unchanged prefixes and suffixes.
 */
export function createDocumentHtmlFingerprintSegment(
  source: string,
  from = 0,
  to = source.length,
): DocumentHtmlFingerprintSegment {
  if (
    !Number.isSafeInteger(from) ||
    !Number.isSafeInteger(to) ||
    from < 0 ||
    to < from ||
    to > source.length
  ) {
    throw new RangeError('The document HTML fingerprint range is invalid.');
  }
  let hash = 0;
  for (let index = from; index < to; index += 1) {
    hash =
      (Math.imul(hash, DOCUMENT_HTML_HASH_BASE) + source.charCodeAt(index)) >>>
      0;
  }
  const length = to - from;
  return {
    hash,
    length,
    power: documentHtmlHashPower(length),
  };
}

export function combineDocumentHtmlFingerprintSegments(
  segments: Iterable<DocumentHtmlFingerprintSegment>,
): DocumentHtmlFingerprintSegment {
  let combined: DocumentHtmlFingerprintSegment = {
    hash: 0,
    length: 0,
    power: 1,
  };
  for (const segment of segments) {
    combined = {
      hash: (Math.imul(combined.hash, segment.power) + segment.hash) >>> 0,
      length: combined.length + segment.length,
      power: Math.imul(combined.power, segment.power) >>> 0,
    };
  }
  return combined;
}

export function documentHtmlFingerprintForSegment(
  segment: DocumentHtmlFingerprintSegment,
): string {
  return `${DOCUMENT_HTML_FINGERPRINT_VERSION}:${segment.length.toString(36)}:${segment.hash.toString(36)}`;
}

export function documentHtmlFingerprint(html: string): string {
  return documentHtmlFingerprintForSegment(
    createDocumentHtmlFingerprintSegment(html),
  );
}

/** Accepts fingerprints written before composable document hashing shipped. */
export function documentHtmlFingerprintMatches(
  html: string,
  candidate: string,
): boolean {
  return candidate.startsWith(`${DOCUMENT_HTML_FINGERPRINT_VERSION}:`)
    ? candidate === documentHtmlFingerprint(html)
    : candidate === legacyDocumentHtmlFingerprint(html);
}

function legacyDocumentHtmlFingerprint(html: string): string {
  let hash = LEGACY_FNV_OFFSET;
  for (let index = 0; index < html.length; index += 1) {
    hash ^= html.charCodeAt(index);
    hash = Math.imul(hash, DOCUMENT_HTML_HASH_BASE);
  }
  return `${html.length.toString(36)}:${(hash >>> 0).toString(36)}`;
}

function documentHtmlHashPower(length: number): number {
  let exponent = length;
  let factor = DOCUMENT_HTML_HASH_BASE;
  let power = 1;
  while (exponent > 0) {
    if (exponent % 2 === 1) power = Math.imul(power, factor) >>> 0;
    factor = Math.imul(factor, factor) >>> 0;
    exponent = Math.floor(exponent / 2);
  }
  return power;
}
