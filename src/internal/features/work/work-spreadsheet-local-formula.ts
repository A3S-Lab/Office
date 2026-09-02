import {
  Parser,
  SUPPORTED_FORMULAS,
  type FormulaParserCoordinate,
} from '@fortune-sheet/formula-parser';
import type { Cell, CellMatrix } from '@fortune-sheet/core';
import { normalizeFormulaForFortuneParser } from '../../kernel/office-kernel-spreadsheet-fallback-formula';

/** OOXML and the browser editor keep local rule formulas deliberately small. */
export const MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH = 255;

/** Keep synchronous browser formula evaluation bounded for hostile workbooks. */
export const MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS = 1_024;

export interface SpreadsheetFormulaSheet {
  id: string;
  name: string;
  data?: CellMatrix;
  celldata?: readonly SpreadsheetFormulaCell[];
}

export interface SpreadsheetFormulaCell {
  r: number;
  c: number;
  v: unknown;
}

export interface SpreadsheetLocalFormulaOptions {
  sheets: readonly SpreadsheetFormulaSheet[];
  sheetId: string;
  row: number;
  column: number;
  /** The first row of the rule/validation range. Defaults to the target row. */
  anchorRow?: number;
  /** The first column of the rule/validation range. Defaults to the target column. */
  anchorColumn?: number;
  /** Substitute the target cell with this value, as data validation does. */
  proposedValue?: unknown;
  useProposedValue?: boolean;
}

export interface SpreadsheetLocalFormulaResult {
  supported: boolean;
  value?: unknown;
  message?: string;
}

interface Coordinate {
  column: number;
  row: number;
  sheet: SpreadsheetFormulaSheet;
}

interface FormulaAnchor {
  column: number;
  row: number;
}

const MAX_SHEET_ROW = 1_048_576;
const MAX_SHEET_COLUMN = 16_384;
const LOCAL_FORMULA_FUNCTIONS = new Set(
  SUPPORTED_FORMULAS.map((name) => name.toUpperCase()),
);

/**
 * Evaluate a bounded Excel formula against the workbook's cached local data.
 *
 * This is intentionally shared by custom data validation and conditional
 * formatting. It never recalculates formula cells or reaches a remote source;
 * an unavailable dependency therefore fails closed.
 */
export function evaluateSpreadsheetLocalFormula(
  formula: string,
  options: SpreadsheetLocalFormulaOptions,
): SpreadsheetLocalFormulaResult {
  const source = formula.trim().replace(/^=/, '').trim();
  if (!source) return unsupportedLocalFormula('公式为空。');
  if (Array.from(source).length > MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH) {
    return unsupportedLocalFormula('公式超过 255 个字符。');
  }
  const currentSheet = options.sheets.find(
    (candidate) => candidate.id === options.sheetId,
  );
  if (!currentSheet) return unsupportedLocalFormula('找不到公式所在的工作表。');

  const evaluator = new SpreadsheetLocalFormulaEvaluator(options, currentSheet);
  return evaluator.evaluate(source);
}

export function spreadsheetLocalFormulaBoolean(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return Number.isFinite(value) && value !== 0;
  return typeof value === 'string' && value.trim().toUpperCase() === 'TRUE';
}

class SpreadsheetLocalFormulaEvaluator {
  private readonly sheetsByName = new Map<string, SpreadsheetFormulaSheet>();
  private readonly anchor: FormulaAnchor;
  private referencedCells = 0;
  private unsupportedMessage: string | undefined;
  private readonly celldataBySheet = new Map<string, Map<string, Cell>>();

  constructor(
    private readonly options: SpreadsheetLocalFormulaOptions,
    private readonly currentSheet: SpreadsheetFormulaSheet,
  ) {
    this.anchor = {
      row: options.anchorRow ?? options.row,
      column: options.anchorColumn ?? options.column,
    };
    for (const sheet of options.sheets) {
      this.sheetsByName.set(sheet.name.trim().toLocaleLowerCase(), sheet);
    }
  }

  evaluate(formula: string): SpreadsheetLocalFormulaResult {
    const parser = new Parser();
    parser
      .setFunction('ROW', (parameters) =>
        parameters.length ? null : this.options.row + 1,
      )
      .setFunction('COLUMN', (parameters) =>
        parameters.length ? null : this.options.column + 1,
      );
    let unsupportedFunction: string | undefined;
    parser.on('callFunction', (name) => {
      const normalized = name.toUpperCase();
      if (!LOCAL_FORMULA_FUNCTIONS.has(normalized))
        unsupportedFunction ??= normalized;
    });
    parser.on('callCellValue', (coordinate, _parserOptions, done) => {
      const resolved = this.resolveCoordinate(this.currentSheet, coordinate);
      if (!resolved) {
        done(null);
        return;
      }
      done(this.readCell(resolved));
    });
    parser.on('callRangeValue', (start, end, _parserOptions, done) => {
      done(this.readRange(start, end));
    });

    const parsed = parser.parse(normalizeFormulaForFortuneParser(formula), {
      sheetId: this.currentSheet.id,
    });
    if (unsupportedFunction) {
      return unsupportedLocalFormula(
        `公式函数“${unsupportedFunction}”无法在浏览器本地求值。`,
      );
    }
    if (this.unsupportedMessage) {
      return unsupportedLocalFormula(this.unsupportedMessage);
    }
    if (parsed.error || parsed.result === undefined) {
      return unsupportedLocalFormula('公式无法求值，请检查函数、引用和括号。');
    }
    return { supported: true, value: parsed.result };
  }

  private resolveCoordinate(
    fallbackSheet: SpreadsheetFormulaSheet,
    coordinate: FormulaParserCoordinate,
  ): Coordinate | null {
    const sheet = coordinate.sheetName
      ? this.sheetsByName.get(coordinate.sheetName.trim().toLocaleLowerCase())
      : fallbackSheet;
    if (!sheet) {
      this.unsupportedMessage = '公式引用了不存在的工作表。';
      return null;
    }
    if (coordinate.row.index < 0 || coordinate.column.index < 0) {
      this.unsupportedMessage = '公式暂不支持整行或整列引用，请改用有限区域。';
      return null;
    }
    const row = coordinate.row.isAbsolute
      ? coordinate.row.index
      : this.options.row + (coordinate.row.index - this.anchor.row);
    const column = coordinate.column.isAbsolute
      ? coordinate.column.index
      : this.options.column + (coordinate.column.index - this.anchor.column);
    if (
      row < 0 ||
      column < 0 ||
      row >= MAX_SHEET_ROW ||
      column >= MAX_SHEET_COLUMN
    ) {
      this.unsupportedMessage = '公式引用超出了工作表边界。';
      return null;
    }
    return { sheet, row, column };
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
      this.unsupportedMessage ??= '公式区域引用必须位于同一工作表。';
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
        MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS - this.referencedCells
    ) {
      this.unsupportedMessage = '公式一次最多读取 1,024 个单元格。';
      return [];
    }
    const values: unknown[][] = [];
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const valuesRow: unknown[] = [];
      for (let column = columnStart; column <= columnEnd; column += 1) {
        valuesRow.push(
          this.readCell({ sheet: startCoordinate.sheet, row, column }),
        );
      }
      values.push(valuesRow);
    }
    return values;
  }

  private readCell(coordinate: Coordinate): unknown {
    this.referencedCells += 1;
    if (this.referencedCells > MAX_SPREADSHEET_LOCAL_FORMULA_REFERENCED_CELLS) {
      this.unsupportedMessage = '公式一次最多读取 1,024 个单元格。';
      return null;
    }
    if (
      this.options.useProposedValue &&
      coordinate.sheet.id === this.currentSheet.id &&
      coordinate.row === this.options.row &&
      coordinate.column === this.options.column
    ) {
      return this.options.proposedValue;
    }
    const cell = this.cellAt(
      coordinate.sheet,
      coordinate.row,
      coordinate.column,
    );
    if (!cell) return null;
    if (
      typeof cell.f === 'string' &&
      cell.f.length > 0 &&
      cell.v === undefined
    ) {
      this.unsupportedMessage ??= '公式引用了尚未计算的公式单元格。';
      return null;
    }
    return cell.v ?? null;
  }

  private cellAt(
    sheet: SpreadsheetFormulaSheet,
    row: number,
    column: number,
  ): Cell | null {
    if (sheet.data !== undefined) return sheet.data[row]?.[column] ?? null;
    let cells = this.celldataBySheet.get(sheet.id);
    if (!cells) {
      cells = new Map();
      for (const entry of sheet.celldata ?? []) {
        const value =
          entry.v && typeof entry.v === 'object' && 'v' in entry.v
            ? (entry.v as Cell)
            : ({ v: entry.v } as Cell);
        cells.set(`${entry.r}_${entry.c}`, value);
      }
      this.celldataBySheet.set(sheet.id, cells);
    }
    return cells.get(`${row}_${column}`) ?? null;
  }
}

function unsupportedLocalFormula(
  message: string,
): SpreadsheetLocalFormulaResult {
  return { supported: false, value: false, message };
}
