const WILDCARD_ESCAPES = new Set(['*', '?', '~']);

type WildcardToken =
  | { kind: 'literal'; value: string }
  | { kind: 'many' }
  | { kind: 'single' };

type FixedWidthWildcardToken = Exclude<WildcardToken, { kind: 'many' }>;

interface FixedWidthWildcardSegment {
  anchor: { offset: number; text: string } | null;
  bitset: FixedWidthWildcardBitset | null;
  tokens: readonly FixedWidthWildcardToken[];
}

interface FixedWidthWildcardBitset {
  literalMasks: ReadonlyMap<string, bigint>;
  matchBit: bigint;
  singleMask: bigint;
}

const BITSET_MINIMUM_SEARCH_WORK = 1_000_000;
const BITSET_MINIMUM_TOKENS = 256;
const BITSET_MAXIMUM_LITERAL_VALUES = 256;

/**
 * Compiles the WPS/OOXML wildcard language into a controlled token matcher. `*`
 * matches zero or more characters, `?` matches one normalized Unicode
 * character, and `~` escapes `*`, `?`, or another `~`.
 */
export function workSpreadsheetWildcardMatcher(
  source: string,
): (value: string) => boolean {
  const tokens = wildcardTokens(source);
  const literalMatcher = literalWildcardMatcher(tokens);
  if (literalMatcher) {
    return (value) => literalMatcher(comparableWildcardText(value));
  }
  const fixedWidthMatcher = fixedWidthWildcardMatcher(tokens);
  return (value) => fixedWidthMatcher(comparableWildcardText(value));
}

export function workSpreadsheetHasUnescapedWildcard(source: string): boolean {
  const characters = [...source];
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];
    if (
      character === '~' &&
      WILDCARD_ESCAPES.has(characters[index + 1] ?? '')
    ) {
      index += 1;
      continue;
    }
    if (character === '*' || character === '?') return true;
  }
  return false;
}

function wildcardTokens(source: string): WildcardToken[] {
  const characters = [...source];
  const tokens: WildcardToken[] = [];
  let literal = '';
  const flushLiteral = () => {
    if (!literal) return;
    for (const value of comparableWildcardText(literal)) {
      tokens.push({ kind: 'literal', value });
    }
    literal = '';
  };

  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index] ?? '';
    const escaped = characters[index + 1] ?? '';
    if (character === '~' && WILDCARD_ESCAPES.has(escaped)) {
      literal += escaped;
      index += 1;
      continue;
    }
    if (character === '*') {
      flushLiteral();
      if (tokens.at(-1)?.kind !== 'many') tokens.push({ kind: 'many' });
      continue;
    }
    if (character === '?') {
      flushLiteral();
      tokens.push({ kind: 'single' });
      continue;
    }
    literal += character;
  }
  flushLiteral();
  return tokens;
}

function literalWildcardMatcher(
  tokens: readonly WildcardToken[],
): ((value: string) => boolean) | null {
  if (tokens.some((token) => token.kind === 'single')) return null;
  const startsWithMany = tokens[0]?.kind === 'many';
  const endsWithMany = tokens.at(-1)?.kind === 'many';
  const segments: string[] = [];
  let segment = '';
  for (const token of tokens) {
    if (token.kind === 'many') {
      if (segment) segments.push(segment);
      segment = '';
      continue;
    }
    if (token.kind !== 'literal') return null;
    segment += token.value;
  }
  if (segment) segments.push(segment);
  if (!segments.length) {
    return tokens.length ? () => true : (value) => value === '';
  }
  if (!startsWithMany && !endsWithMany && segments.length === 1) {
    return (value) => value === segments[0];
  }
  return (value) => {
    let minimumIndex = 0;
    let segmentIndex = 0;
    if (!startsWithMany) {
      const first = segments[0] ?? '';
      if (!value.startsWith(first)) return false;
      minimumIndex = first.length;
      segmentIndex = 1;
    }
    const searchableSegments = endsWithMany
      ? segments.length
      : segments.length - 1;
    for (; segmentIndex < searchableSegments; segmentIndex += 1) {
      const current = segments[segmentIndex] ?? '';
      const found = value.indexOf(current, minimumIndex);
      if (found < 0) return false;
      minimumIndex = found + current.length;
    }
    if (endsWithMany) return true;
    const last = segments.at(-1) ?? '';
    return value.endsWith(last) && value.length - last.length >= minimumIndex;
  };
}

function fixedWidthWildcardMatcher(
  tokens: readonly WildcardToken[],
): (value: string) => boolean {
  const startsWithMany = tokens[0]?.kind === 'many';
  const endsWithMany = tokens.at(-1)?.kind === 'many';
  const segments: FixedWidthWildcardSegment[] = [];
  let current: FixedWidthWildcardToken[] = [];
  for (const token of tokens) {
    if (token.kind === 'many') {
      if (current.length) segments.push(fixedWidthWildcardSegment(current));
      current = [];
      continue;
    }
    current.push(token);
  }
  if (current.length) segments.push(fixedWidthWildcardSegment(current));
  if (!segments.length) return () => true;
  if (!startsWithMany && !endsWithMany && segments.length === 1) {
    const only = segments[0];
    return (value) => {
      const characters = [...value];
      return (
        characters.length === only.tokens.length &&
        fixedWidthWildcardSegmentMatches(only, characters, 0)
      );
    };
  }
  return (value) => {
    const characters = [...value];
    let minimumIndex = 0;
    let segmentIndex = 0;
    let maximumEnd = characters.length;
    let finalSegment = segments.length;
    if (!startsWithMany) {
      const first = segments[0];
      if (!fixedWidthWildcardSegmentMatches(first, characters, 0)) return false;
      minimumIndex = first.tokens.length;
      segmentIndex = 1;
    }
    if (!endsWithMany) {
      finalSegment -= 1;
      const last = segments[finalSegment];
      const start = characters.length - last.tokens.length;
      if (
        start < minimumIndex ||
        !fixedWidthWildcardSegmentMatches(last, characters, start)
      ) {
        return false;
      }
      maximumEnd = start;
    }
    if (segmentIndex >= finalSegment) return true;
    const offsets = wildcardCodePointOffsets(characters);
    for (; segmentIndex < finalSegment; segmentIndex += 1) {
      const segment = segments[segmentIndex];
      const found = findFixedWidthWildcardSegment(
        segment,
        value,
        characters,
        offsets,
        minimumIndex,
        maximumEnd,
      );
      if (found < 0) return false;
      minimumIndex = found + segment.tokens.length;
    }
    return true;
  };
}

function fixedWidthWildcardSegment(
  tokens: readonly FixedWidthWildcardToken[],
): FixedWidthWildcardSegment {
  let anchor: FixedWidthWildcardSegment['anchor'] = null;
  let literalStart = -1;
  const considerLiteralRun = (end: number) => {
    if (literalStart < 0) return;
    const text = tokens
      .slice(literalStart, end)
      .map((token) => (token.kind === 'literal' ? token.value : ''))
      .join('');
    if (!anchor || end - literalStart >= [...anchor.text].length) {
      anchor = { offset: literalStart, text };
    }
    literalStart = -1;
  };
  for (let index = 0; index <= tokens.length; index += 1) {
    if (tokens[index]?.kind === 'literal') {
      if (literalStart < 0) literalStart = index;
    } else {
      considerLiteralRun(index);
    }
  }
  return { anchor, bitset: fixedWidthWildcardBitset(tokens), tokens };
}

function findFixedWidthWildcardSegment(
  segment: FixedWidthWildcardSegment,
  value: string,
  characters: readonly string[],
  offsets: readonly number[],
  minimumIndex: number,
  maximumEnd: number,
): number {
  if (minimumIndex + segment.tokens.length > maximumEnd) return -1;
  if (!segment.anchor) return minimumIndex;
  if (
    segment.bitset &&
    segment.tokens.length * (maximumEnd - minimumIndex) >=
      BITSET_MINIMUM_SEARCH_WORK
  ) {
    return findFixedWidthWildcardSegmentWithBitset(
      segment,
      characters,
      minimumIndex,
      maximumEnd,
    );
  }
  const minimumAnchorIndex = minimumIndex + segment.anchor.offset;
  let anchorOffset = value.indexOf(
    segment.anchor.text,
    offsets[minimumAnchorIndex] ?? value.length,
  );
  while (anchorOffset >= 0) {
    const anchorIndex = codePointIndexAtOffset(offsets, anchorOffset);
    if (anchorIndex >= 0) {
      const start = anchorIndex - segment.anchor.offset;
      if (start + segment.tokens.length > maximumEnd) return -1;
      if (
        start >= minimumIndex &&
        fixedWidthWildcardSegmentMatches(segment, characters, start)
      ) {
        return start;
      }
    }
    anchorOffset = value.indexOf(segment.anchor.text, anchorOffset + 1);
  }
  return -1;
}

function fixedWidthWildcardBitset(
  tokens: readonly FixedWidthWildcardToken[],
): FixedWidthWildcardBitset | null {
  if (tokens.length < BITSET_MINIMUM_TOKENS) return null;
  const literalValues = new Set(
    tokens.flatMap((token) => (token.kind === 'literal' ? [token.value] : [])),
  );
  if (literalValues.size > BITSET_MAXIMUM_LITERAL_VALUES) return null;
  const literalMasks = new Map<string, bigint>();
  let singleMask = 0n;
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    const bit = 1n << BigInt(index);
    if (token.kind === 'single') {
      singleMask |= bit;
      continue;
    }
    literalMasks.set(token.value, (literalMasks.get(token.value) ?? 0n) | bit);
  }
  return {
    literalMasks,
    matchBit: 1n << BigInt(tokens.length - 1),
    singleMask,
  };
}

function findFixedWidthWildcardSegmentWithBitset(
  segment: FixedWidthWildcardSegment,
  characters: readonly string[],
  minimumIndex: number,
  maximumEnd: number,
): number {
  const bitset = segment.bitset;
  if (!bitset) return -1;
  let state = 0n;
  for (let index = minimumIndex; index < maximumEnd; index += 1) {
    const mask =
      bitset.singleMask | (bitset.literalMasks.get(characters[index]) ?? 0n);
    state = ((state << 1n) | 1n) & mask;
    if ((state & bitset.matchBit) !== 0n) {
      return index - segment.tokens.length + 1;
    }
  }
  return -1;
}

function fixedWidthWildcardSegmentMatches(
  segment: FixedWidthWildcardSegment,
  characters: readonly string[],
  start: number,
): boolean {
  if (start < 0 || start + segment.tokens.length > characters.length) {
    return false;
  }
  return segment.tokens.every(
    (token, index) =>
      token.kind === 'single' || token.value === characters[start + index],
  );
}

function wildcardCodePointOffsets(characters: readonly string[]): number[] {
  const offsets = [0];
  for (const character of characters) {
    offsets.push((offsets.at(-1) ?? 0) + character.length);
  }
  return offsets;
}

function codePointIndexAtOffset(
  offsets: readonly number[],
  target: number,
): number {
  let low = 0;
  let high = offsets.length - 1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const value = offsets[middle] ?? -1;
    if (value === target) return middle;
    if (value < target) low = middle + 1;
    else high = middle - 1;
  }
  return -1;
}

function comparableWildcardText(value: string): string {
  return value.normalize('NFKC').toLocaleLowerCase('zh-CN');
}
