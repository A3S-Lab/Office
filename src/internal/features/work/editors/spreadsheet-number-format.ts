import type { Cell } from '@fortune-sheet/core';

export type SpreadsheetNumberFormatPreset =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'percent'
  | 'date'
  | 'time'
  | 'scientific'
  | 'fraction'
  | 'text'
  | 'custom';

export type SpreadsheetNumberFormatChoice = Exclude<
  SpreadsheetNumberFormatPreset,
  'custom'
>;

export const spreadsheetNumberFormatPresetLabels = {
  general: '常规',
  number: '数字',
  currency: '货币',
  accounting: '会计专用',
  percent: '百分比',
  date: '短日期',
  time: '时间',
  scientific: '科学计数',
  fraction: '分数',
  text: '文本',
  custom: '自定义',
} as const satisfies Record<SpreadsheetNumberFormatPreset, string>;

const spreadsheetNumberFormatCodes = {
  general: 'General',
  number: '#,##0.00',
  currency: '[$¥-804]#,##0.00',
  accounting: '_([$¥-804]* #,##0.00_);_([$¥-804]* (#,##0.00)',
  percent: '0.00%',
  date: 'yyyy-MM-dd',
  time: 'hh:mm',
  scientific: '0.00E+00',
  fraction: '# ?/?',
  text: '@',
} as const satisfies Record<SpreadsheetNumberFormatChoice, string>;

export function spreadsheetNumberFormatPreset(
  formatCode?: string,
): SpreadsheetNumberFormatPreset {
  const code = normalizedSpreadsheetNumberFormat(formatCode);
  if (code.toLocaleLowerCase() === 'general') return 'general';
  if (code === '@') return 'text';

  const semanticCode = spreadsheetNumberFormatSemanticCode(code);
  const hasDate = isSpreadsheetDateFormat(semanticCode);
  const hasTime = isSpreadsheetTimeFormat(semanticCode);
  if (hasDate && hasTime) return 'custom';
  if (hasDate) return 'date';
  if (hasTime) return 'time';
  if (semanticCode.includes('%')) return 'percent';
  if (isSpreadsheetScientificFormat(semanticCode)) return 'scientific';
  if (isSpreadsheetFractionFormat(semanticCode)) return 'fraction';
  if (isSpreadsheetAccountingFormat(code)) return 'accounting';
  if (hasSpreadsheetCurrencyMarker(code)) return 'currency';
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
  if (preset === 'date' || preset === 'time') return { fa, t: 'd' };
  if (preset === 'text') return { fa, t: 's' };
  if (
    preset === 'number' ||
    preset === 'currency' ||
    preset === 'accounting' ||
    preset === 'percent' ||
    preset === 'scientific' ||
    preset === 'fraction'
  ) {
    return { fa, t: 'n' };
  }
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
  if (
    preset === 'custom' ||
    preset === 'date' ||
    preset === 'time' ||
    preset === 'fraction' ||
    preset === 'text'
  ) {
    return code;
  }
  if (preset === 'general') return direction > 0 ? '0.0' : '0';
  return code
    .split(';')
    .map((section) => adjustSpreadsheetNumberFormatSection(section, direction))
    .join(';');
}

function normalizedSpreadsheetNumberFormat(formatCode?: string): string {
  return formatCode?.trim() || 'General';
}

function spreadsheetNumberFormatSemanticCode(formatCode: string): string {
  return formatCode
    .replace(/"(?:[^"]|"")*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[([hms]+)\]/gi, '$1')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/[_*]./g, '');
}

function isSpreadsheetDateFormat(formatCode: string): boolean {
  return (
    /(^|[^a-z])(y+|d+)(?=$|[^a-z])/i.test(formatCode) ||
    /(^|[^a-z])m{3,4}(?=$|[^a-z])/i.test(formatCode)
  );
}

function isSpreadsheetTimeFormat(formatCode: string): boolean {
  return (
    /(^|[^a-z])(h+|s+)(?=$|[^a-z])/i.test(formatCode) ||
    /am\/pm/i.test(formatCode) ||
    /[0#?]+:[0#?]+/.test(formatCode)
  );
}

function isSpreadsheetScientificFormat(formatCode: string): boolean {
  return /[0#?](?:[0#?,]*)(?:\.[0#?]+)?e[+-][0#?]+/i.test(formatCode);
}

function isSpreadsheetFractionFormat(formatCode: string): boolean {
  return /(?:[0#?]+\s+)?[0#?]+\s*\/\s*[0#?]+/.test(formatCode);
}

function hasSpreadsheetCurrencyMarker(formatCode: string): boolean {
  return /\[\$[^\]]*\]|[$€£¥￥₹₽₩]/u.test(formatCode);
}

function isSpreadsheetAccountingFormat(formatCode: string): boolean {
  return Boolean(
    hasSpreadsheetCurrencyMarker(formatCode) &&
      (formatCode.includes('*') ||
        /(?:\[\$[^\]]*\]|[$€£¥￥₹₽₩])\s*\([0#?,. ]+\)/u.test(formatCode)),
  );
}

function adjustSpreadsheetNumberFormatSection(
  section: string,
  direction: -1 | 1,
): string {
  const token = /"(?:[^"]|"")*"|\\.|\[[^\]]*\]|([0#?][0#?,]*)(?:\.([0#?]+))?/g;
  let match: RegExpExecArray | null = null;
  for (const candidate of section.matchAll(token)) {
    if (candidate[1]) {
      match = candidate;
      break;
    }
  }
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
