import {
  Parser,
  SUPPORTED_FORMULAS,
  type FormulaParserCoordinate,
} from '@fortune-sheet/formula-parser';
import type { Cell } from '@fortune-sheet/core';
import { normalizeFormulaForFortuneParser } from '../../../kernel/office-kernel-spreadsheet-fallback-formula';
import { parseSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  spreadsheetCellRangeContains,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';

/** OOXML limits a data-validation formula to 255 characters. */
export const MAX_SPREADSHEET_CUSTOM_VALIDATION_FORMULA_LENGTH = 255;

/** Keep synchronous browser validation bounded even for a hostile workbook. */
export const MAX_SPREADSHEET_CUSTOM_VALIDATION_REFERENCED_CELLS = 1_024;

export interface SpreadsheetCustomValidationResult {
  message?: string;
  supported: boolean;
  valid: boolean;
}

interface Coordinate {
  column: number;
  row: number;
  sheet: WorkSpreadsheetSheet;
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
    return unsupportedCustomValidation('找不到自定义公式所在的工作表。');
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
    return unsupportedCustomValidation('自定义公式为空。');
  }
  if (
    Array.from(formula).length >
    MAX_SPREADSHEET_CUSTOM_VALIDATION_FORMULA_LENGTH
  ) {
    return unsupportedCustomValidation('自定义公式超过 255 个字符。');
  }

  const ranges = parseSpreadsheetCellRanges(item.rangeTxt);
  const anchor = spreadsheetCustomValidationAnchor(ranges, row, column);
  const evaluator = new SpreadsheetCustomFormulaEvaluator(
    content,
    sheet,
    row,
    column,
    anchor,
    normalizedProposed,
  );
  return evaluator.evaluate(formula);
}

class SpreadsheetCustomFormulaEvaluator {
  private readonly sheetsByName = new Map<string, WorkSpreadsheetSheet>();
  private referencedCells = 0;
  private unsupportedMessage: string | undefined;

  constructor(
    private readonly content: WorkSpreadsheetContent,
    private readonly currentSheet: WorkSpreadsheetSheet,
    private readonly targetRow: number,
    private readonly targetColumn: number,
    private readonly anchor: ValidationAnchor,
    private readonly proposedValue: unknown,
  ) {
    for (const sheet of content.sheets) {
      this.sheetsByName.set(sheet.name.trim().toLocaleLowerCase(), sheet);
    }
  }

  evaluate(formula: string): SpreadsheetCustomValidationResult {
    const allowedFunctions = new Set(
      SUPPORTED_FORMULAS.map((name) => name.toUpperCase()),
    );
    const parser = new Parser();
    parser
      .setFunction('ROW', (parameters) =>
        parameters.length ? null : this.targetRow + 1,
      )
      .setFunction('COLUMN', (parameters) =>
        parameters.length ? null : this.targetColumn + 1,
      );
    let unsupportedFunction: string | undefined;
    parser.on('callFunction', (name) => {
      const normalized = name.toUpperCase();
      if (!allowedFunctions.has(normalized)) unsupportedFunction ??= normalized;
    });
    parser.on('callCellValue', (coordinate, _options, done) => {
      const resolved = this.resolveCoordinate(this.currentSheet, coordinate);
      if (!resolved) {
        done(null);
        return;
      }
      done(this.readCell(resolved));
    });
    parser.on('callRangeValue', (start, end, _options, done) => {
      done(this.readRange(start, end));
    });

    const parsed = parser.parse(normalizeFormulaForFortuneParser(formula), {
      sheetId: this.currentSheet.id,
    });
    if (unsupportedFunction) {
      return unsupportedCustomValidation(
        `自定义公式函数“${unsupportedFunction}”无法在浏览器本地求值。`,
      );
    }
    if (this.unsupportedMessage) {
      return unsupportedCustomValidation(this.unsupportedMessage);
    }
    if (parsed.error) {
      return unsupportedCustomValidation(
        '自定义公式无法求值，请检查函数、引用和括号。',
      );
    }
    return {
      supported: true,
      valid: spreadsheetCustomValidationBoolean(parsed.result),
    };
  }

  private resolveCoordinate(
    fallbackSheet: WorkSpreadsheetSheet,
    coordinate: FormulaParserCoordinate,
  ): Coordinate | null {
    const sheet = coordinate.sheetName
      ? this.sheetsByName.get(coordinate.sheetName.trim().toLocaleLowerCase())
      : fallbackSheet;
    if (!sheet) {
      this.unsupportedMessage = '自定义公式引用了不存在的工作表。';
      return null;
    }
    if (coordinate.row.index < 0 || coordinate.column.index < 0) {
      this.unsupportedMessage =
        '自定义公式暂不支持整行或整列引用，请改用有限区域。';
      return null;
    }
    const resolvedRow = coordinate.row.isAbsolute
      ? coordinate.row.index
      : this.targetRow + (coordinate.row.index - this.anchor.row);
    const resolvedColumn = coordinate.column.isAbsolute
      ? coordinate.column.index
      : this.targetColumn + (coordinate.column.index - this.anchor.column);
    if (
      resolvedRow < 0 ||
      resolvedColumn < 0 ||
      resolvedRow >= 1_048_576 ||
      resolvedColumn >= 16_384
    ) {
      this.unsupportedMessage = '自定义公式引用超出了工作表边界。';
      return null;
    }
    return { sheet, row: resolvedRow, column: resolvedColumn };
  }

  private readRange(
    start: FormulaParserCoordinate,
    end: FormulaParserCoordinate,
  ): unknown[][] {
    const startCoordinate = this.resolveCoordinate(this.currentSheet, start);
    const rangeSheet = start.sheetName
      ? this.sheetsByName.get(start.sheetName.trim().toLocaleLowerCase())
      : this.currentSheet;
    const endCoordinate = rangeSheet
      ? this.resolveCoordinate(rangeSheet, end)
      : null;
    if (
      !startCoordinate ||
      !endCoordinate ||
      startCoordinate.sheet.id !== endCoordinate.sheet.id
    ) {
      this.unsupportedMessage ??= '自定义公式区域引用必须位于同一工作表。';
      return [];
    }
    const rowStart = Math.min(startCoordinate.row, endCoordinate.row);
    const rowEnd = Math.max(startCoordinate.row, endCoordinate.row);
    const columnStart = Math.min(startCoordinate.column, endCoordinate.column);
    const columnEnd = Math.max(startCoordinate.column, endCoordinate.column);
    const area = (rowEnd - rowStart + 1) * (columnEnd - columnStart + 1);
    if (
      area <= 0 ||
      area >
        MAX_SPREADSHEET_CUSTOM_VALIDATION_REFERENCED_CELLS -
          this.referencedCells
    ) {
      this.unsupportedMessage = '自定义公式一次最多读取 1,024 个单元格。';
      return [];
    }
    const values: unknown[][] = [];
    for (let currentRow = rowStart; currentRow <= rowEnd; currentRow += 1) {
      const row: unknown[] = [];
      for (
        let currentColumn = columnStart;
        currentColumn <= columnEnd;
        currentColumn += 1
      ) {
        row.push(
          this.readCell({
            sheet: startCoordinate.sheet,
            row: currentRow,
            column: currentColumn,
          }),
        );
      }
      values.push(row);
    }
    return values;
  }

  private readCell(coordinate: Coordinate): unknown {
    this.referencedCells += 1;
    if (
      this.referencedCells > MAX_SPREADSHEET_CUSTOM_VALIDATION_REFERENCED_CELLS
    ) {
      this.unsupportedMessage = '自定义公式一次最多读取 1,024 个单元格。';
      return null;
    }
    if (
      coordinate.sheet.id === this.currentSheet.id &&
      coordinate.row === this.targetRow &&
      coordinate.column === this.targetColumn
    ) {
      return this.proposedValue;
    }
    const cell = spreadsheetCustomValidationCellAt(
      coordinate.sheet,
      coordinate.row,
      coordinate.column,
    );
    if (!cell) return null;
    if (cell.f && cell.v === undefined) {
      this.unsupportedMessage ??= '自定义公式引用了尚未计算的公式单元格。';
      return null;
    }
    return cell.v ?? null;
  }
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

function spreadsheetCustomValidationCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  if (sheet.data !== undefined) return sheet.data[row]?.[column] ?? null;
  return (
    sheet.celldata?.find((entry) => entry.r === row && entry.c === column)?.v ??
    null
  );
}

function spreadsheetCustomValidationValue(value: unknown): unknown {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return value;
  if ('v' in value) return (value as { v?: unknown }).v;
  return value;
}

function spreadsheetCustomValidationBlank(value: unknown): boolean {
  return value == null || (typeof value === 'string' && value.trim() === '');
}

function spreadsheetCustomValidationBoolean(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return Number.isFinite(value) && value !== 0;
  return typeof value === 'string' && value.trim().toUpperCase() === 'TRUE';
}

function unsupportedCustomValidation(
  message: string,
): SpreadsheetCustomValidationResult {
  return { supported: false, valid: false, message };
}
