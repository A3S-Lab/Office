import type { WorkSpreadsheetDataValidationErrorStyle } from './work-types';

const MILLISECONDS_PER_DAY = 86_400_000;

export const SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT = 32;
export const SPREADSHEET_DATA_VALIDATION_HINT_LIMIT = 255;
export const SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT = 225;
export const SPREADSHEET_DATA_VALIDATION_FORMULA_LIMIT = 255;

export const SPREADSHEET_DATA_VALIDATION_ERROR_STYLES = [
  'stop',
  'warning',
  'information',
] as const satisfies readonly WorkSpreadsheetDataValidationErrorStyle[];

export function normalizeSpreadsheetDataValidationErrorStyle(
  value: unknown,
): WorkSpreadsheetDataValidationErrorStyle {
  return SPREADSHEET_DATA_VALIDATION_ERROR_STYLES.includes(
    value as WorkSpreadsheetDataValidationErrorStyle,
  )
    ? (value as WorkSpreadsheetDataValidationErrorStyle)
    : 'stop';
}

export function boundedSpreadsheetDataValidationText(
  value: unknown,
  limit: number,
): string {
  if (typeof value !== 'string' || !Number.isSafeInteger(limit) || limit < 0)
    return '';
  return Array.from(value.trim()).slice(0, limit).join('');
}

export function normalizeSpreadsheetDateValidationBoundary(
  value: string,
  uses1904DateSystem = false,
): string | null {
  const trimmed = value.trim().replace(/^=/, '');
  const iso = normalizeIsoDate(trimmed);
  if (iso) return iso;

  const formula = /^DATE\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})\)$/i.exec(trimmed);
  if (formula) {
    return normalizeIsoDate(
      `${formula[1]}-${String(formula[2]).padStart(2, '0')}-${String(
        formula[3],
      ).padStart(2, '0')}`,
    );
  }

  const serial = Number(trimmed);
  if (!Number.isInteger(serial) || serial < 0) return null;
  if (!uses1904DateSystem && (serial === 0 || serial === 60)) return null;
  const adjustedSerial = uses1904DateSystem
    ? serial
    : serial > 60
      ? serial - 1
      : serial;
  const epoch = uses1904DateSystem
    ? Date.UTC(1904, 0, 1)
    : Date.UTC(1899, 11, 31);
  return isoDateFromTimestamp(epoch + adjustedSerial * MILLISECONDS_PER_DAY);
}

export function spreadsheetDateValidationFormula(value: string): string {
  const normalized = normalizeIsoDate(value.trim());
  if (!normalized) return value.trim().replace(/^=/, '');
  const [year, month, day] = normalized.split('-');
  return `DATE(${year},${Number(month)},${Number(day)})`;
}

function normalizeIsoDate(value: string): string | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return null;
  const timestamp = Date.UTC(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
  );
  const normalized = isoDateFromTimestamp(timestamp);
  return normalized === value ? normalized : null;
}

function isoDateFromTimestamp(timestamp: number): string | null {
  const date = new Date(timestamp);
  const year = date.getUTCFullYear();
  if (!Number.isFinite(timestamp) || year < 1900 || year > 9999) return null;
  return `${String(year).padStart(4, '0')}-${String(
    date.getUTCMonth() + 1,
  ).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}
