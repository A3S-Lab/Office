export interface OfficeNumberNormalizationOptions {
  decimalPlaces?: number;
  integer?: boolean;
  isValid?: (value: number) => boolean;
  maximum?: number;
  minimum?: number;
}

export function normalizeRequiredOfficeNumber(
  value: string,
  options: OfficeNumberNormalizationOptions = {},
): number | null {
  if (!value.trim()) return null;
  return normalizeOfficeNumber(value, options);
}

export function normalizeOptionalOfficeNumber(
  value: string,
  options: OfficeNumberNormalizationOptions = {},
): number | undefined | null {
  if (!value.trim()) return undefined;
  return normalizeOfficeNumber(value, options);
}

function normalizeOfficeNumber(
  value: string,
  options: OfficeNumberNormalizationOptions,
): number | null {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return null;
  let normalized = options.integer
    ? Math.round(parsed)
    : roundDecimalPlaces(parsed, options.decimalPlaces);
  normalized = Math.min(
    options.maximum ?? Number.POSITIVE_INFINITY,
    Math.max(options.minimum ?? Number.NEGATIVE_INFINITY, normalized),
  );
  return !options.isValid || options.isValid(normalized) ? normalized : null;
}

function roundDecimalPlaces(value: number, decimalPlaces?: number): number {
  if (decimalPlaces === undefined) return value;
  const factor = 10 ** Math.max(0, Math.trunc(decimalPlaces));
  return Math.round(value * factor) / factor;
}
