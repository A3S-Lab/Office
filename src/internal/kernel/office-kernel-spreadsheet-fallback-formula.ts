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
