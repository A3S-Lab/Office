export const browserScalarFunctionArities = new Map<
  string,
  readonly [minimum: number, maximum: number]
>([
  ['ABS', [1, 1]],
  ['AND', [1, 255]],
  ['AVERAGE', [1, 255]],
  ['COLUMN', [0, 0]],
  ['CONCAT', [1, 255]],
  ['CONCATENATE', [1, 255]],
  ['COUNT', [1, 255]],
  ['COUNTA', [1, 255]],
  ['FALSE', [0, 0]],
  ['IF', [2, 3]],
  ['IFERROR', [2, 2]],
  ['MAX', [1, 255]],
  ['MIN', [1, 255]],
  ['MOD', [2, 2]],
  ['NA', [0, 0]],
  ['NOT', [1, 1]],
  ['OR', [1, 255]],
  ['PI', [0, 0]],
  ['POWER', [2, 2]],
  ['ROUND', [2, 2]],
  ['ROW', [0, 0]],
  ['SQRT', [1, 1]],
  ['SUBTOTAL', [2, 255]],
  ['SUM', [1, 255]],
  ['TRUE', [0, 0]],
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
