export const DEFAULT_SPREADSHEET_FONT_SIZE = 10;

export const spreadsheetFontSizes = [
  9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72,
] as const;

export type SpreadsheetFontSizeDirection = 'grow' | 'shrink';

export function nextSpreadsheetFontSize(
  current: number,
  direction: SpreadsheetFontSizeDirection,
): number | null {
  if (!Number.isFinite(current) || current < 1 || current > 409) return null;
  return direction === 'grow'
    ? (spreadsheetFontSizes.find((size) => size > current) ?? null)
    : ([...spreadsheetFontSizes].reverse().find((size) => size < current) ??
        null);
}
