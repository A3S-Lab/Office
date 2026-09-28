import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';
import {
  evaluateSpreadsheetLocalFormula,
  MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH,
  MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS,
  spreadsheetLocalFormulaBoolean,
} from '../work-spreadsheet-local-formula';
import { parseSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationItem,
} from '../work-types';
import {
  spreadsheetCellRangeContains,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';

/** OOXML limits a data-validation formula to 255 characters. */
export const MAX_SPREADSHEET_CUSTOM_VALIDATION_FORMULA_LENGTH =
  MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH;

/** Keep synchronous browser validation bounded even for a hostile workbook. */
export const MAX_SPREADSHEET_CUSTOM_VALIDATION_REFERENCED_CELLS =
  MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS;

export interface SpreadsheetCustomValidationResult {
  message?: string;
  supported: boolean;
  valid: boolean;
}

interface ValidationAnchor {
  column: number;
  row: number;
}

/**
 * Evaluates a local Excel custom-validation formula synchronously.
 *
 * The editor deliberately evaluates against cached cell values. Formula
 * recalculation remains a separate kernel concern; a formula cell without a
 * cached value therefore fails closed instead of silently accepting input.
 */
export function evaluateSpreadsheetCustomValidation(
  content: WorkSpreadsheetContent,
  sheetId: string,
  row: number,
  column: number,
  item: WorkSpreadsheetDataValidationItem,
  value: unknown,
): SpreadsheetCustomValidationResult {
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  if (!sheet) {
    return unsupportedCustomValidation(
      officeMessage(resolveOfficeMessages(), 'spreadsheet.dv.custom.sheetMissing'),
    );
  }
  const proposed = spreadsheetCustomValidationValue(value);
  const normalizedProposed = spreadsheetCustomValidationBlank(proposed)
    ? ''
    : proposed;
  if (item.allowBlank !== false && normalizedProposed === '') {
    return { supported: true, valid: true };
  }
  const formula = item.value1.trim().replace(/^=/, '').trim();
  if (!formula) {
    return unsupportedCustomValidation(
      officeMessage(resolveOfficeMessages(), 'spreadsheet.dv.custom.empty'),
    );
  }
  if (
    Array.from(formula).length >
    MAX_SPREADSHEET_CUSTOM_VALIDATION_FORMULA_LENGTH
  ) {
    return unsupportedCustomValidation(
      officeMessage(resolveOfficeMessages(), 'spreadsheet.dv.custom.tooLong'),
    );
  }

  const ranges = parseSpreadsheetCellRanges(item.rangeTxt);
  const anchor = spreadsheetCustomValidationAnchor(ranges, row, column);
  const result = evaluateSpreadsheetLocalFormula(formula, {
    sheets: content.sheets.flatMap((candidate) =>
      candidate.id ? [{ ...candidate, id: candidate.id }] : [],
    ),
    sheetId,
    row,
    column,
    anchorRow: anchor?.row,
    anchorColumn: anchor?.column,
    proposedValue: normalizedProposed,
    useProposedValue: true,
  });
  if (!result.supported) {
    return {
      supported: false,
      valid: false,
      message: result.message
        ? officeMessage(resolveOfficeMessages(), 'spreadsheet.dv.custom.prefixedDetail', {
            detail: result.message,
          })
        : result.message,
    };
  }
  return {
    supported: true,
    valid: spreadsheetLocalFormulaBoolean(result.value),
  };
}

function spreadsheetCustomValidationAnchor(
  ranges: readonly SpreadsheetCellRange[] | null,
  row: number,
  column: number,
): ValidationAnchor {
  const range = ranges?.find((candidate) =>
    spreadsheetCellRangeContains(candidate, row, column),
  );
  return {
    row: range?.row[0] ?? row,
    column: range?.column[0] ?? column,
  };
}

function spreadsheetCustomValidationValue(value: unknown): unknown {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return value;
  if ('v' in value) return (value as { v?: unknown }).v;
  return value;
}

function spreadsheetCustomValidationBlank(value: unknown): boolean {
  return value == null || (typeof value === 'string' && value.trim() === '');
}

function unsupportedCustomValidation(
  message: string,
): SpreadsheetCustomValidationResult {
  return { supported: false, valid: false, message };
}
