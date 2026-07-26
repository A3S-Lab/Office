import type { Cell } from '@fortune-sheet/core';

export type SpreadsheetNumberFormatPreset =
  | 'general'
  | 'number'
  | 'percent'
  | 'custom';

type SpreadsheetNumberFormatChoice = Exclude<
  SpreadsheetNumberFormatPreset,
  'custom'
>;

const spreadsheetNumberFormatCodes = {
  general: 'General',
  number: '#,##0.00',
  percent: '0.00%',
} as const satisfies Record<SpreadsheetNumberFormatChoice, string>;

export function spreadsheetNumberFormatPreset(
  formatCode?: string,
): SpreadsheetNumberFormatPreset {
  const code = normalizedSpreadsheetNumberFormat(formatCode);
  if (code.toLocaleLowerCase() === 'general') return 'general';

  const semanticCode = code
    .replace(/"(?:[^"]|"")*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]*\]/g, '');
  if (semanticCode.includes('%')) return 'percent';
  if (isSpreadsheetDateTimeFormat(semanticCode)) return 'custom';
  if (/[0#?]/.test(semanticCode)) return 'number';
  return 'custom';
}

export function spreadsheetNumberFormatCode(
  preset: SpreadsheetNumberFormatChoice,
): string {
  return spreadsheetNumberFormatCodes[preset];
}

export function spreadsheetNumberFormatValue(
  formatCode: string,
  cell?: Cell | null,
): NonNullable<Cell['ct']> & { fa: string; t: string } {
  const fa = normalizedSpreadsheetNumberFormat(formatCode);
  const preset = spreadsheetNumberFormatPreset(fa);
  if (preset === 'number' || preset === 'percent') return { fa, t: 'n' };
  if (preset === 'custom' && cell?.ct?.t) return { fa, t: cell.ct.t };
  if (typeof cell?.v === 'number' || cell?.ct?.t === 'n') return { fa, t: 'n' };
  if (typeof cell?.v === 'boolean') return { fa, t: 'b' };
  return { fa, t: 'g' };
}

export function adjustSpreadsheetNumberFormat(
  formatCode: string | undefined,
  direction: -1 | 1,
): string {
  const code = normalizedSpreadsheetNumberFormat(formatCode);
  const preset = spreadsheetNumberFormatPreset(code);
  if (preset === 'custom') return code;
  if (preset === 'general') return direction > 0 ? '0.0' : '0';
  return code
    .split(';')
    .map((section) => adjustSpreadsheetNumberFormatSection(section, direction))
    .join(';');
}

function normalizedSpreadsheetNumberFormat(formatCode?: string): string {
  return formatCode?.trim() || 'General';
}

function isSpreadsheetDateTimeFormat(formatCode: string): boolean {
  if (/(^|[^a-z])(y+|d+|h+|s+)(?=$|[^a-z])/i.test(formatCode)) return true;
  if (/(^|[^a-z])m{2,}(?=$|[^a-z])/i.test(formatCode)) return true;
  return (
    /[/:]/.test(formatCode) && /(^|[^a-z])m+(?=$|[^a-z])/i.test(formatCode)
  );
}

function adjustSpreadsheetNumberFormatSection(
  section: string,
  direction: -1 | 1,
): string {
  const match = /([0#?][0#?,]*)(?:\.([0#?]+))?/.exec(section);
  if (!match || match.index === undefined) return section;

  const decimals = match[2] ?? '';
  const nextLength = Math.max(0, Math.min(10, decimals.length + direction));
  const nextDecimals =
    nextLength <= decimals.length
      ? decimals.slice(0, nextLength)
      : `${decimals}${'0'.repeat(nextLength - decimals.length)}`;
  const replacement = `${match[1]}${nextDecimals ? `.${nextDecimals}` : ''}`;
  return `${section.slice(0, match.index)}${replacement}${section.slice(
    match.index + match[0].length,
  )}`;
}
