import { locale } from '@fortune-sheet/core';
import { spreadsheetFormulaFunctions } from './work-spreadsheet-formulas';

const FORTUNE_SUPPORTED_FUNCTIONS = new Set(
  locale({ lang: 'en' } as Parameters<typeof locale>[0]).functionlist.map(
    (item) => item.n.toUpperCase(),
  ),
);

export function unsupportedSpreadsheetFormulaFunctions(
  formula: string,
): string[] {
  return spreadsheetFormulaFunctions(formula).filter(
    (name) => !FORTUNE_SUPPORTED_FUNCTIONS.has(name),
  );
}
