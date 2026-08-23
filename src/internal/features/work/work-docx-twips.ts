export const STRICT_WORDPROCESSING_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

const UNIVERSAL_TWIP_RATIOS: Readonly<
  Record<string, readonly [bigint, bigint]>
> = {
  mm: [7_200n, 127n],
  cm: [72_000n, 127n],
  in: [1_440n, 1n],
  pt: [20n, 1n],
  pc: [240n, 1n],
  pi: [240n, 1n],
};

const UNIVERSAL_HALF_POINT_RATIOS: Readonly<
  Record<string, readonly [bigint, bigint]>
> = {
  mm: [720n, 127n],
  cm: [7_200n, 127n],
  in: [144n, 1n],
  pt: [2n, 1n],
  pc: [24n, 1n],
  pi: [24n, 1n],
};

export function parseDocxTwipsMeasure(
  source: string,
  options: {
    minimum: number;
    maximum: number;
    signed: boolean;
    strict: boolean;
  },
): number | null {
  const integer = parseBoundedDocxInteger(source, {
    minimum: options.minimum,
    maximum: options.maximum,
    signed: options.signed,
  });
  if (integer !== null) return integer;
  return options.strict
    ? parseStrictUniversalMeasure(source, UNIVERSAL_TWIP_RATIOS, options)
    : null;
}

export function parseDocxHalfPointsMeasure(
  source: string,
  options: {
    minimum: number;
    maximum: number;
    signed: boolean;
    strict: boolean;
  },
): number | null {
  const integer = parseBoundedDocxInteger(source, {
    minimum: options.minimum,
    maximum: options.maximum,
    signed: options.signed,
  });
  if (integer !== null) return integer;
  return options.strict
    ? parseStrictUniversalMeasure(source, UNIVERSAL_HALF_POINT_RATIOS, options)
    : null;
}

export function parseBoundedDocxInteger(
  source: string,
  options: {
    minimum: number;
    maximum: number;
    signed?: boolean;
  },
): number | null {
  if (source.length > 32) return null;
  const pattern = options.signed ? /^[+-]?\d+$/u : /^\+?\d+$/u;
  if (!pattern.test(source)) return null;
  const value = Number(source);
  return Number.isSafeInteger(value) &&
    value >= options.minimum &&
    value <= options.maximum
    ? Object.is(value, -0)
      ? 0
      : value
    : null;
}

function parseStrictUniversalMeasure(
  source: string,
  conversions: Readonly<Record<string, readonly [bigint, bigint]>>,
  options: { minimum: number; maximum: number; signed: boolean },
): number | null {
  if (source.length > 64) return null;
  const match = /^(-?)(\d+)(?:\.(\d+))?(mm|cm|in|pt|pc|pi)$/u.exec(source);
  if (!match) return null;
  const [, sign, whole = '', fraction = '', unit = ''] = match;
  if (!options.signed && sign) return null;
  const conversion = conversions[unit];
  if (!conversion) return null;
  const denominator = 10n ** BigInt(fraction.length) * conversion[1];
  const numerator =
    BigInt(`${whole}${fraction}`) * conversion[0] * (sign ? -1n : 1n);
  if (numerator % denominator !== 0n) return null;
  const units = numerator / denominator;
  if (units < BigInt(options.minimum) || units > BigInt(options.maximum)) {
    return null;
  }
  return Number(units);
}
