export const browserScalarFunctionArities = new Map<
  string,
  readonly [minimum: number, maximum: number]
>([
  ['ABS', [1, 1]],
  ['AND', [1, 255]],
  ['AVERAGE', [1, 255]],
  ['AVERAGEIF', [2, 3]],
  ['AVERAGEIFS', [3, 255]],
  ['COLUMN', [0, 0]],
  ['CONCAT', [1, 255]],
  ['CONCATENATE', [1, 255]],
  ['COUNT', [1, 255]],
  ['COUNTA', [1, 255]],
  ['COUNTIF', [2, 2]],
  ['COUNTIFS', [2, 255]],
  ['DATE', [3, 3]],
  ['DAY', [1, 1]],
  ['FALSE', [0, 0]],
  ['HLOOKUP', [3, 4]],
  ['IF', [2, 3]],
  ['IFERROR', [2, 2]],
  ['INDEX', [2, 3]],
  ['LEFT', [1, 2]],
  ['LEN', [1, 1]],
  ['MATCH', [2, 3]],
  ['MAX', [1, 255]],
  ['MID', [3, 3]],
  ['MIN', [1, 255]],
  ['MOD', [2, 2]],
  ['MONTH', [1, 1]],
  ['NA', [0, 0]],
  ['NOT', [1, 1]],
  ['NUMBERVALUE', [1, 3]],
  ['OR', [1, 255]],
  ['PI', [0, 0]],
  ['POWER', [2, 2]],
  ['PRODUCT', [1, 255]],
  ['RIGHT', [1, 2]],
  ['ROUND', [2, 2]],
  ['ROW', [0, 0]],
  ['SQRT', [1, 1]],
  ['SUBTOTAL', [2, 255]],
  ['SUM', [1, 255]],
  ['SUMIF', [2, 3]],
  ['SUMIFS', [3, 255]],
  ['SUMPRODUCT', [1, 255]],
  ['SEQUENCE', [1, 4]],
  ['FILTER', [2, 3]],
  ['SORT', [1, 4]],
  ['TRANSPOSE', [1, 1]],
  ['UNIQUE', [1, 3]],
  ['TRUE', [0, 0]],
  ['VALUE', [1, 1]],
  ['VLOOKUP', [3, 4]],
  ['YEAR', [1, 1]],
]);

export function normalizeSpreadsheetFunctionName(name: string): string {
  let normalized = name.toUpperCase();
  while (normalized.startsWith('_XLFN.') || normalized.startsWith('_XLWS.')) {
    normalized = normalized.slice(6);
  }
  return normalized;
}

/**
 * Evaluate the bounded browser implementation of `SUBTOTAL`.
 *
 * Fortune delegates unknown functions to Formula.js. Formula.js returns an
 * undefined value for an unknown function number, which the parser can then
 * coerce into a boolean. Keeping this small implementation here makes the
 * JavaScript fallback deterministic and aligned with the Rust kernel. The
 * calculation request does not carry hidden-row metadata yet, so the 1–11
 * and 101–111 function-number families intentionally have the same result.
 */
export function evaluateParserSubtotal(
  parameters: readonly unknown[],
): unknown {
  const codeValue = parserNumericValue(parameters[0]);
  if (codeValue === undefined) return 'VALUE!';
  const code = Math.trunc(codeValue);
  const values = parameters
    .slice(1)
    .flatMap((parameter) => flattenParserValues(parameter));
  for (const value of values) {
    const error = parserErrorValue(value);
    if (error) return error;
  }
  switch (code) {
    case 1:
    case 101:
      return parserSubtotalNumeric(values, 'average');
    case 2:
    case 102:
      return values.filter((value) => typeof value === 'number').length;
    case 3:
    case 103:
      return values.filter((value) => value !== null && value !== undefined)
        .length;
    case 4:
    case 104:
      return parserSubtotalNumeric(values, 'max');
    case 5:
    case 105:
      return parserSubtotalNumeric(values, 'min');
    case 6:
    case 106:
      return parserSubtotalNumeric(values, 'product');
    case 7:
    case 107:
      return parserSubtotalNumeric(values, 'stddev');
    case 8:
    case 108:
      return parserSubtotalNumeric(values, 'stddevp');
    case 9:
    case 109:
      return parserSubtotalNumeric(values, 'sum');
    case 10:
    case 110:
      return parserSubtotalNumeric(values, 'var');
    case 11:
    case 111:
      return parserSubtotalNumeric(values, 'varp');
    default:
      return 'VALUE!';
  }
}

/** Excel 1900-date-system epoch (includes the historical leap-day bug). */
const EXCEL_1900_EPOCH_UTC = Date.UTC(1899, 11, 30);
/** Excel 1904-date-system epoch (Mac legacy / workbook date1904). */
const EXCEL_1904_EPOCH_UTC = Date.UTC(1904, 0, 1);

export type SpreadsheetFormulaDateSystem = '1900' | '1904';

function excelEpochUtc(
  dateSystem: SpreadsheetFormulaDateSystem = '1900',
): number {
  return dateSystem === '1904' ? EXCEL_1904_EPOCH_UTC : EXCEL_1900_EPOCH_UTC;
}

/**
 * Bounded Traditional Office `DATE` / `YEAR` / `MONTH` / `DAY` helpers that keep
 * Excel serial numbers instead of JavaScript `Date` objects.
 */
export function evaluateParserDate(
  parameters: readonly unknown[],
  dateSystem: SpreadsheetFormulaDateSystem = '1900',
): unknown {
  const year = parserNumericValue(parameters[0]);
  const month = parserNumericValue(parameters[1]);
  const day = parserNumericValue(parameters[2]);
  if (year === undefined || month === undefined || day === undefined) {
    return 'VALUE!';
  }
  const utc = Date.UTC(
    Math.trunc(year),
    Math.trunc(month) - 1,
    Math.trunc(day),
  );
  if (!Number.isFinite(utc)) return 'NUM!';
  return Math.round((utc - excelEpochUtc(dateSystem)) / 86_400_000);
}

export function evaluateParserYear(
  parameters: readonly unknown[],
  dateSystem: SpreadsheetFormulaDateSystem = '1900',
): unknown {
  const parts = excelPartsFromParserDate(parameters[0], dateSystem);
  return parts ? parts.year : 'VALUE!';
}

export function evaluateParserMonth(
  parameters: readonly unknown[],
  dateSystem: SpreadsheetFormulaDateSystem = '1900',
): unknown {
  const parts = excelPartsFromParserDate(parameters[0], dateSystem);
  return parts ? parts.month : 'VALUE!';
}

export function evaluateParserDay(
  parameters: readonly unknown[],
  dateSystem: SpreadsheetFormulaDateSystem = '1900',
): unknown {
  const parts = excelPartsFromParserDate(parameters[0], dateSystem);
  return parts ? parts.day : 'VALUE!';
}

/**
 * Bounded Traditional Office `NUMBERVALUE` with explicit decimal/group
 * separators. One-argument calls keep en-US defaults (`.` decimal, `,`
 * thousands) and refuse ambiguous comma-decimal text so locale guessing stays
 * fail-closed — use `NUMBERVALUE(text, ",", ".")` for EU-style decimals.
 */
export function evaluateParserNumberValue(
  parameters: readonly unknown[],
): unknown {
  if (!parameters.length || parameters.length > 3) return 'VALUE!';
  const raw = parameters[0];
  if (raw instanceof Error) {
    return parserErrorValue(raw) ?? 'VALUE!';
  }
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? raw : 'VALUE!';
  }
  if (typeof raw === 'boolean') return raw ? 1 : 0;
  if (raw === null || raw === undefined) return 'VALUE!';
  const text = String(raw).trim();
  if (!text) return 'VALUE!';

  const decimalSeparator =
    parameters.length >= 2 &&
    parameters[1] !== null &&
    parameters[1] !== undefined &&
    String(parameters[1]) !== ''
      ? String(parameters[1]).charAt(0)
      : '.';
  const groupSeparator =
    parameters.length >= 3 &&
    parameters[2] !== null &&
    parameters[2] !== undefined &&
    String(parameters[2]) !== ''
      ? String(parameters[2]).charAt(0)
      : parameters.length >= 2
        ? ''
        : ',';

  if (
    parameters.length === 1 &&
    decimalSeparator === '.' &&
    groupSeparator === ',' &&
    /^\d+,\d+$/.test(text.replace(/\s/g, ''))
  ) {
    return 'VALUE!';
  }

  let working = text.replace(/\s/g, '');
  if (groupSeparator) {
    working = working.split(groupSeparator).join('');
  }
  if (decimalSeparator !== '.') {
    const parts = working.split(decimalSeparator);
    if (parts.length > 2) return 'VALUE!';
    working = parts.join('.');
  } else if ((working.match(/\./g) ?? []).length > 1) {
    return 'VALUE!';
  }

  let percent = false;
  if (working.endsWith('%')) {
    percent = true;
    working = working.slice(0, -1);
  }
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(working)) {
    return 'VALUE!';
  }
  const parsed = Number(working);
  if (!Number.isFinite(parsed)) return 'VALUE!';
  return percent ? parsed / 100 : parsed;
}

/**
 * Bounded Traditional Office `VALUE` with locale guessing for unambiguous
 * decimal/group separators. Prefer the last separator as decimal when both
 * `,` and `.` appear; treat one-separator text as EU decimal when the
 * fractional part has 1–2 digits, otherwise as US thousands groups of three.
 * Remaining ambiguous shapes stay `#VALUE!` (use `NUMBERVALUE` with explicit
 * separators).
 */
export function evaluateParserValue(parameters: readonly unknown[]): unknown {
  if (parameters.length !== 1) return 'VALUE!';
  const raw = parameters[0];
  if (raw instanceof Error) {
    return parserErrorValue(raw) ?? 'VALUE!';
  }
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? raw : 'VALUE!';
  }
  if (typeof raw === 'boolean') return raw ? 1 : 0;
  if (raw === null || raw === undefined) return 'VALUE!';
  const text = String(raw).trim();
  if (!text) return 'VALUE!';
  return localeGuessParseNumberText(text);
}

function localeGuessParseNumberText(text: string): number | string {
  let working = text.replace(/\s/g, '');
  let percent = false;
  if (working.endsWith('%')) {
    percent = true;
    working = working.slice(0, -1);
  }
  if (!working) return 'VALUE!';

  const hasComma = working.includes(',');
  const hasDot = working.includes('.');

  if (hasComma && hasDot) {
    const lastComma = working.lastIndexOf(',');
    const lastDot = working.lastIndexOf('.');
    if (lastComma > lastDot) {
      return finalizeLocaleGuessNumber(
        stripLocaleGroupSeparators(working, '.').replace(',', '.'),
        percent,
      );
    }
    return finalizeLocaleGuessNumber(
      stripLocaleGroupSeparators(working, ','),
      percent,
    );
  }

  if (hasComma) {
    const parts = working.split(',');
    if (parts.length === 2) {
      const [left, right] = parts;
      if (
        left !== undefined &&
        right !== undefined &&
        /^[+-]?\d+$/.test(left) &&
        /^\d{1,2}$/.test(right)
      ) {
        return finalizeLocaleGuessNumber(`${left}.${right}`, percent);
      }
      if (
        left !== undefined &&
        right !== undefined &&
        /^[+-]?\d{1,3}$/.test(left) &&
        /^\d{3}$/.test(right)
      ) {
        return finalizeLocaleGuessNumber(`${left}${right}`, percent);
      }
      return 'VALUE!';
    }
    if (
      parts.length > 2 &&
      parts.every((part, index) =>
        index === 0 ? /^[+-]?\d{1,3}$/.test(part) : /^\d{3}$/.test(part),
      )
    ) {
      return finalizeLocaleGuessNumber(parts.join(''), percent);
    }
    return 'VALUE!';
  }

  if (hasDot) {
    const parts = working.split('.');
    if (
      parts.length > 2 &&
      parts.every((part, index) =>
        index === 0 ? /^[+-]?\d{1,3}$/.test(part) : /^\d{3}$/.test(part),
      )
    ) {
      return finalizeLocaleGuessNumber(parts.join(''), percent);
    }
  }

  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(working)) {
    return 'VALUE!';
  }
  return finalizeLocaleGuessNumber(working, percent);
}

function stripLocaleGroupSeparators(text: string, group: string): string {
  return text.split(group).join('');
}

function finalizeLocaleGuessNumber(
  working: string,
  percent: boolean,
): number | string {
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(working)) {
    return 'VALUE!';
  }
  const parsed = Number(working);
  if (!Number.isFinite(parsed)) return 'VALUE!';
  return percent ? parsed / 100 : parsed;
}

/**
 * Bounded Traditional Office `SEQUENCE` returning a row-major numeric grid for
 * dynamic-array spill materialization in the JavaScript fallback.
 */
export function evaluateParserSequence(
  parameters: readonly unknown[],
): unknown {
  if (!parameters.length || parameters.length > 4) return 'VALUE!';
  const rows = positiveSequenceDimension(parameters[0]);
  const columns =
    parameters.length >= 2 ? positiveSequenceDimension(parameters[1]) : 1;
  const start = parameters.length >= 3 ? parserNumericValue(parameters[2]) : 1;
  const step = parameters.length >= 4 ? parserNumericValue(parameters[3]) : 1;
  if (
    rows === undefined ||
    columns === undefined ||
    start === undefined ||
    step === undefined
  ) {
    return 'VALUE!';
  }
  const cells = rows * columns;
  if (cells > 10_000) return 'VALUE!';
  const grid: number[][] = [];
  let next = start;
  for (let row = 0; row < rows; row += 1) {
    const line: number[] = [];
    for (let column = 0; column < columns; column += 1) {
      line.push(next);
      next += step;
    }
    grid.push(line);
  }
  return grid;
}

/**
 * Bounded Traditional Office `TRANSPOSE` returning a transposed numeric/text
 * grid for dynamic-array spill materialization.
 */
export function evaluateParserTranspose(
  parameters: readonly unknown[],
): unknown {
  if (parameters.length !== 1) return 'VALUE!';
  const grid = normalizeParserGrid(parameters[0]);
  if (!grid) return 'VALUE!';
  const rowCount = grid.length;
  const columnCount = Math.max(0, ...grid.map((row) => row.length));
  if (rowCount * columnCount > 10_000) return 'VALUE!';
  const transposed: unknown[][] = [];
  for (let column = 0; column < columnCount; column += 1) {
    const line: unknown[] = [];
    for (let row = 0; row < rowCount; row += 1) {
      line.push(grid[row]?.[column] ?? null);
    }
    transposed.push(line);
  }
  return transposed;
}

/**
 * Bounded Traditional Office `UNIQUE` (row-wise, keep first occurrence).
 * Optional `by_col` / `exactly_once` arguments stay fail-closed as `#VALUE!`
 * when non-default so the spill surface remains intentional.
 */
export function evaluateParserUnique(parameters: readonly unknown[]): unknown {
  if (!parameters.length || parameters.length > 3) return 'VALUE!';
  const byCol = parameters.length >= 2 ? parserNumericValue(parameters[1]) : 0;
  const exactlyOnce =
    parameters.length >= 3 ? parserNumericValue(parameters[2]) : 0;
  if (byCol === undefined || exactlyOnce === undefined) return 'VALUE!';
  if (byCol !== 0 || exactlyOnce !== 0) return 'VALUE!';
  const grid = normalizeParserGrid(parameters[0]);
  if (!grid) return 'VALUE!';
  const seen = new Set<string>();
  const uniqueRows: unknown[][] = [];
  for (const row of grid) {
    const key = JSON.stringify(row.map((cell) => normalizeUniqueKey(cell)));
    if (seen.has(key)) continue;
    seen.add(key);
    uniqueRows.push(row);
    if (uniqueRows.length > 10_000) return 'VALUE!';
  }
  return uniqueRows.length ? uniqueRows : [[null]];
}

/**
 * Bounded Traditional Office `FILTER` (row or column include masks).
 * Empty results without `if_empty` fail closed as `#CALC!`.
 */
export function evaluateParserFilter(parameters: readonly unknown[]): unknown {
  if (parameters.length < 2 || parameters.length > 3) return 'VALUE!';
  const grid = normalizeParserGrid(parameters[0]);
  const include = normalizeParserGrid(parameters[1]);
  if (!grid || !include) return 'VALUE!';
  const rowCount = grid.length;
  const columnCount = Math.max(0, ...grid.map((row) => row.length));
  if (rowCount * columnCount > 10_000) return 'VALUE!';
  const includeRows = include.length;
  const includeColumns = Math.max(0, ...include.map((row) => row.length));
  const filterRows = includeRows === rowCount && includeColumns === 1;
  const filterColumns = includeRows === 1 && includeColumns === columnCount;
  if (!filterRows && !filterColumns) return 'VALUE!';

  let filtered: unknown[][];
  if (filterRows) {
    filtered = [];
    for (let row = 0; row < rowCount; row += 1) {
      if (!parserIncludeTruthy(include[row]?.[0])) continue;
      filtered.push(padParserRow(grid[row] ?? [], columnCount));
      if (filtered.length > 10_000) return 'VALUE!';
    }
  } else {
    const keptColumns: number[] = [];
    for (let column = 0; column < columnCount; column += 1) {
      if (!parserIncludeTruthy(include[0]?.[column])) continue;
      keptColumns.push(column);
    }
    filtered = grid.map((row) =>
      keptColumns.map((column) => row[column] ?? null),
    );
  }

  if (filtered.length === 0 || (filtered[0]?.length ?? 0) === 0) {
    if (parameters.length < 3) return 'CALC!';
    const fallback = normalizeParserGrid(parameters[2]);
    return fallback ?? parameters[2] ?? null;
  }
  return filtered;
}

/**
 * Bounded Traditional Office `SORT` (row-wise by one column index).
 * Non-default `by_col` stays fail-closed as `#VALUE!`.
 */
export function evaluateParserSort(parameters: readonly unknown[]): unknown {
  if (!parameters.length || parameters.length > 4) return 'VALUE!';
  const grid = normalizeParserGrid(parameters[0]);
  if (!grid) return 'VALUE!';
  const sortIndex =
    parameters.length >= 2 ? parserNumericValue(parameters[1]) : 1;
  const sortOrder =
    parameters.length >= 3 ? parserNumericValue(parameters[2]) : 1;
  const byCol = parameters.length >= 4 ? parserNumericValue(parameters[3]) : 0;
  if (
    sortIndex === undefined ||
    sortOrder === undefined ||
    byCol === undefined
  ) {
    return 'VALUE!';
  }
  if (byCol !== 0) return 'VALUE!';
  if (sortOrder !== 1 && sortOrder !== -1) return 'VALUE!';
  const columnCount = Math.max(0, ...grid.map((row) => row.length));
  const column = Math.trunc(sortIndex) - 1;
  if (column < 0 || column >= columnCount) return 'VALUE!';
  if (grid.length * columnCount > 10_000) return 'VALUE!';
  const sorted = grid
    .map((row) => padParserRow(row, columnCount))
    .map((row, index) => ({ row, index }))
    .sort((left, right) => {
      const compared = compareParserSortValues(
        left.row[column],
        right.row[column],
      );
      if (compared !== 0) return sortOrder === 1 ? compared : -compared;
      return left.index - right.index;
    })
    .map((entry) => entry.row);
  return sorted.length ? sorted : [[null]];
}

function parserIncludeTruthy(value: unknown): boolean {
  if (value instanceof Error) return false;
  if (value === null || value === undefined || value === '') return false;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0 && Number.isFinite(value);
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return false;
    const numeric = Number(trimmed);
    if (Number.isFinite(numeric)) return numeric !== 0;
    return trimmed.toUpperCase() === 'TRUE';
  }
  return Boolean(value);
}

function padParserRow(row: readonly unknown[], columnCount: number): unknown[] {
  const padded: unknown[] = [];
  for (let column = 0; column < columnCount; column += 1) {
    padded.push(row[column] ?? null);
  }
  return padded;
}

function compareParserSortValues(left: unknown, right: unknown): number {
  const leftBlank = left === null || left === undefined || left === '';
  const rightBlank = right === null || right === undefined || right === '';
  if (leftBlank && rightBlank) return 0;
  if (leftBlank) return 1;
  if (rightBlank) return -1;
  const leftNumber = parserNumericValue(left);
  const rightNumber = parserNumericValue(right);
  if (leftNumber !== undefined && rightNumber !== undefined) {
    return leftNumber === rightNumber ? 0 : leftNumber < rightNumber ? -1 : 1;
  }
  const leftText = String(left).toLocaleLowerCase();
  const rightText = String(right).toLocaleLowerCase();
  return leftText.localeCompare(rightText);
}

function normalizeUniqueKey(value: unknown): unknown {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && Object.is(value, -0)) return 0;
  if (value instanceof Error) return value.message;
  return value;
}

function normalizeParserGrid(value: unknown): unknown[][] | null {
  if (value instanceof Error) return null;
  if (!Array.isArray(value)) return [[value]];
  if (value.length === 0) return null;
  if (value.every((row) => Array.isArray(row))) {
    return value as unknown[][];
  }
  return [value];
}

function positiveSequenceDimension(value: unknown): number | undefined {
  const number = parserNumericValue(value);
  if (number === undefined || !Number.isFinite(number)) return undefined;
  const truncated = Math.trunc(number);
  return truncated >= 1 ? truncated : undefined;
}

function excelPartsFromParserDate(
  value: unknown,
  dateSystem: SpreadsheetFormulaDateSystem = '1900',
): { year: number; month: number; day: number } | null {
  const epoch = excelEpochUtc(dateSystem);
  let serial: number | undefined;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    serial = Math.round(
      (Date.UTC(
        value.getUTCFullYear(),
        value.getUTCMonth(),
        value.getUTCDate(),
      ) -
        epoch) /
        86_400_000,
    );
  } else {
    serial = parserNumericValue(value);
  }
  if (serial === undefined || !Number.isFinite(serial)) return null;
  const utc = epoch + Math.floor(serial) * 86_400_000;
  const date = new Date(utc);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
}

type ParserSubtotalOperation =
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'stddev'
  | 'stddevp'
  | 'sum'
  | 'var'
  | 'varp';

function parserSubtotalNumeric(
  values: readonly unknown[],
  operation: ParserSubtotalOperation,
): number | string {
  const numbers = values.filter(
    (value): value is number =>
      typeof value === 'number' && Number.isFinite(value),
  );
  const count = numbers.length;
  const sum = numbers.reduce((total, value) => total + value, 0);
  if (operation === 'sum') return finiteParserNumber(sum);
  if (operation === 'average') {
    return count ? finiteParserNumber(sum / count) : 'DIV/0!';
  }
  if (operation === 'max') return finiteParserNumber(Math.max(...numbers, 0));
  if (operation === 'min') return finiteParserNumber(Math.min(...numbers, 0));
  if (operation === 'product') {
    return finiteParserNumber(
      count ? numbers.reduce((total, value) => total * value, 1) : 0,
    );
  }
  if (operation === 'stddev' || operation === 'var') {
    if (count < 2) return 'DIV/0!';
  } else if (count === 0) {
    return 'DIV/0!';
  }
  const mean = sum / count;
  const divisor =
    operation === 'stddev' || operation === 'var' ? count - 1 : count;
  const variance =
    numbers.reduce((total, value) => total + (value - mean) ** 2, 0) / divisor;
  return finiteParserNumber(
    operation === 'stddev' || operation === 'stddevp'
      ? Math.sqrt(variance)
      : variance,
  );
}

function flattenParserValues(value: unknown): unknown[] {
  if (!Array.isArray(value)) return [value];
  return value.flatMap((entry) => flattenParserValues(entry));
}

function parserNumericValue(value: unknown): number | undefined {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : undefined;
  }
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value !== 'string' || !value.trim()) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function parserErrorValue(value: unknown): string | undefined {
  if (value instanceof Error) {
    return value.message.startsWith('#') ? value.message.slice(1) : 'VALUE!';
  }
  return typeof value === 'string' && value.startsWith('#')
    ? value.slice(1)
    : undefined;
}

function finiteParserNumber(value: number): number | string {
  return Number.isFinite(value) ? value : 'NUM!';
}

export function normalizeFormulaForFortuneParser(formula: string): string {
  const source = formula.replace(/^=/, '');
  let output = '';
  let cursor = 0;
  let bracketDepth = 0;
  while (cursor < source.length) {
    const character = source[cursor] ?? '';
    if (character === '"' || character === "'") {
      const quote = character;
      output += quote;
      cursor += 1;
      while (cursor < source.length) {
        const quoted = source[cursor] ?? '';
        output += quoted;
        cursor += 1;
        if (quoted !== quote) continue;
        if (source[cursor] === quote) {
          output += quote;
          cursor += 1;
          continue;
        }
        break;
      }
      continue;
    }
    if (character === '[') {
      bracketDepth += 1;
      output += character;
      cursor += 1;
      continue;
    }
    if (character === ']' && bracketDepth > 0) {
      bracketDepth -= 1;
      output += character;
      cursor += 1;
      continue;
    }
    if (bracketDepth === 0 && isAsciiFormulaNameStart(character)) {
      const start = cursor;
      cursor += 1;
      while (
        cursor < source.length &&
        isAsciiFormulaNameContinue(source[cursor] ?? '')
      ) {
        cursor += 1;
      }
      const token = source.slice(start, cursor);
      const normalized = token.toUpperCase();
      const normalizedFunction = normalizeSpreadsheetFunctionName(token);
      const previous = adjacentNonWhitespace(source, start - 1, -1);
      const next = adjacentNonWhitespace(source, cursor, 1);
      if (next === '(' && normalizedFunction !== normalized) {
        output += normalizedFunction;
      } else if (
        (normalized === 'TRUE' || normalized === 'FALSE') &&
        previous !== '!' &&
        next !== '!' &&
        next !== '('
      ) {
        output += `${token}()`;
      } else {
        output += token;
      }
      continue;
    }
    output += character;
    cursor += 1;
  }
  return output;
}

function adjacentNonWhitespace(
  value: string,
  start: number,
  direction: -1 | 1,
): string | undefined {
  let cursor = start;
  while (cursor >= 0 && cursor < value.length) {
    const character = value[cursor];
    if (character && !/\s/u.test(character)) return character;
    cursor += direction;
  }
  return undefined;
}

function isAsciiFormulaNameStart(value: string): boolean {
  return /[A-Za-z_]/u.test(value);
}

function isAsciiFormulaNameContinue(value: string): boolean {
  return /[A-Za-z0-9_.]/u.test(value);
}
