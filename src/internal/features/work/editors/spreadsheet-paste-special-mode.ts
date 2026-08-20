import type { WorkSpreadsheetSheet } from '../work-types';
import type { SpreadsheetPasteContent } from './spreadsheet-paste-special-types';

export function pasteContentCopiesFormulas(
  content: SpreadsheetPasteContent,
): boolean {
  return [
    'all',
    'all-except-borders',
    'formulas',
    'formulas-and-number-formats',
  ].includes(content);
}

export function pasteContentCopiesBorders(
  content: SpreadsheetPasteContent,
): boolean {
  return content === 'all' || content === 'formats';
}

export function pasteContentCopiesMerges(
  content: SpreadsheetPasteContent,
): boolean {
  return content === 'all' || content === 'all-except-borders';
}

export function pasteContentCopiesValidation(
  content: SpreadsheetPasteContent,
): boolean {
  return (
    content === 'all' ||
    content === 'all-except-borders' ||
    content === 'validation'
  );
}

export function pasteContentCopiesHyperlinks(
  content: SpreadsheetPasteContent,
): boolean {
  return content === 'all' || content === 'all-except-borders';
}

export function pasteContentCopiesProtection(
  content: SpreadsheetPasteContent,
): boolean {
  return (
    content === 'all' ||
    content === 'all-except-borders' ||
    content === 'formats'
  );
}

export function columnsContainReadOnlyState(
  sheet: WorkSpreadsheetSheet,
  range: { column: [number, number] },
): boolean {
  for (let column = range.column[0]; column <= range.column[1]; column += 1) {
    if (sheet.config?.colReadOnly?.[column]) return true;
  }
  return false;
}
