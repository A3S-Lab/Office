import type { FormulaParserCoordinate } from '@fortune-sheet/formula-parser';
import {
  browserScalarFunctionArities,
  evaluateParserSubtotal,
  normalizeFormulaForFortuneParser,
  normalizeSpreadsheetFunctionName,
} from './office-kernel-spreadsheet-fallback-formula';
import {
  boundedSpreadsheetIndex,
  MAX_SPREADSHEET_DEPENDENCY_DEPTH,
  MAX_SPREADSHEET_RANGE_CELLS,
  validateSpreadsheetCalculationRequest,
} from './office-kernel-spreadsheet-fallback-validation';
import {
  evaluateParserIfError,
  recognizedSpreadsheetError,
  spreadsheetValueForParser,
  spreadsheetValueFromParser,
} from './office-kernel-spreadsheet-fallback-value';
import type {
  OfficeKernelSpreadsheetCalculationIssue,
  OfficeKernelSpreadsheetCalculationRequest,
  OfficeKernelSpreadsheetCalculationResult,
  OfficeKernelSpreadsheetCoordinate,
  OfficeKernelSpreadsheetInputCell,
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetValue,
} from './office-kernel-spreadsheet-protocol';
import {
  OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
  OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
} from './office-kernel-spreadsheet-protocol';
import {
  expandSpreadsheetStructuredReferences,
  SpreadsheetStructuredReferenceCatalog,
  SpreadsheetStructuredReferenceError,
} from './office-kernel-spreadsheet-structured-reference';
import { OFFICE_KERNEL_PROTOCOL_VERSION } from './office-kernel-version';

type FormulaParserModule = typeof import('@fortune-sheet/formula-parser');

interface IndexedCell {
  cell: OfficeKernelSpreadsheetInputCell;
  sheet: OfficeKernelSpreadsheetInputSheet;
}

interface EvaluationState {
  issue?: OfficeKernelSpreadsheetCalculationIssue;
  successful: boolean;
  value: OfficeKernelSpreadsheetValue;
}

export async function calculateSpreadsheetInJavaScript(
  request: OfficeKernelSpreadsheetCalculationRequest,
): Promise<OfficeKernelSpreadsheetCalculationResult> {
  validateSpreadsheetCalculationRequest(request);
  const formulaParser = await import('@fortune-sheet/formula-parser');
  const evaluator = new JavaScriptSpreadsheetEvaluator(request, formulaParser);
  return evaluator.calculate();
}

class JavaScriptSpreadsheetEvaluator {
  private readonly cells = new Map<string, IndexedCell>();
  private readonly sheetsByName = new Map<
    string,
    OfficeKernelSpreadsheetInputSheet
  >();
  private readonly states = new Map<string, EvaluationState>();
  private readonly stack: string[] = [];
  private readonly calculationOrder: OfficeKernelSpreadsheetCoordinate[] = [];
  private readonly issues: OfficeKernelSpreadsheetCalculationIssue[] = [];
  private readonly tableCatalog: SpreadsheetStructuredReferenceCatalog;

  constructor(
    private readonly request: OfficeKernelSpreadsheetCalculationRequest,
    private readonly formulaParser: FormulaParserModule,
  ) {
    this.tableCatalog = new SpreadsheetStructuredReferenceCatalog(
      request.sheets,
    );
    for (const sheet of request.sheets) {
      this.sheetsByName.set(sheet.name.toLowerCase(), sheet);
      for (const cell of sheet.cells) {
        this.cells.set(cellKey(sheet.id, cell.row, cell.column), {
          cell,
          sheet,
        });
      }
    }
  }

  calculate(): OfficeKernelSpreadsheetCalculationResult {
    const targets = this.request.targets?.length
      ? this.request.targets
      : this.request.sheets.flatMap((sheet) =>
          sheet.cells
            .flatMap((cell) =>
              cell.formula
                ? [{ sheetId: sheet.id, row: cell.row, column: cell.column }]
                : [],
            )
            .sort(
              (left, right) =>
                left.row - right.row || left.column - right.column,
            ),
        );
    for (const target of targets) this.evaluateCoordinate(target);

    const cells = this.calculationOrder.flatMap((coordinate) => {
      const state = this.states.get(coordinateKey(coordinate));
      return state?.successful ? [{ ...coordinate, value: state.value }] : [];
    });
    return {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'spreadsheetCalculationResult',
      requestId: this.request.requestId,
      revision: this.request.revision,
      documentRevision: this.request.documentRevision,
      engine: 'javascript',
      cells,
      calculationOrder: this.calculationOrder.filter(
        (coordinate) => this.states.get(coordinateKey(coordinate))?.successful,
      ),
      issues: this.issues,
    };
  }

  private evaluateCoordinate(
    coordinate: OfficeKernelSpreadsheetCoordinate,
  ): OfficeKernelSpreadsheetValue {
    const key = coordinateKey(coordinate);
    const cached = this.states.get(key);
    if (cached) return cached.value;
    const indexed = this.cells.get(key);
    if (!indexed) return { kind: 'blank' };
    if (!indexed.cell.formula) return indexed.cell.value;

    const cycleStart = this.stack.indexOf(key);
    if (cycleStart >= 0) {
      this.markCycle(cycleStart);
      return this.states.get(key)?.value ?? indexed.cell.value;
    }
    if (this.stack.length >= MAX_SPREADSHEET_DEPENDENCY_DEPTH) {
      this.recordState(
        key,
        indexed.cell.value,
        false,
        calculationIssue(
          coordinate,
          'office.kernel.spreadsheet.dependency_depth_exceeded',
          `Spreadsheet formula dependency depth may not exceed ${MAX_SPREADSHEET_DEPENDENCY_DEPTH}.`,
        ),
      );
      return indexed.cell.value;
    }

    this.stack.push(key);
    let state: EvaluationState;
    try {
      state = this.evaluateFormula(coordinate, indexed);
    } finally {
      this.stack.pop();
    }
    const cycleState = this.states.get(key);
    if (cycleState) return cycleState.value;
    this.recordState(key, state.value, state.successful, state.issue);
    return state.value;
  }

  private evaluateFormula(
    coordinate: OfficeKernelSpreadsheetCoordinate,
    indexed: IndexedCell,
  ): EvaluationState {
    const formula = indexed.cell.formula ?? '';
    let expandedFormula: string;
    try {
      expandedFormula = expandSpreadsheetStructuredReferences(
        formula,
        this.tableCatalog,
        indexed.sheet,
        coordinate.row,
        coordinate.column,
      );
    } catch (error) {
      const message =
        error instanceof SpreadsheetStructuredReferenceError
          ? error.message
          : 'Structured reference expansion failed.';
      return failedEvaluation(
        indexed.cell.value,
        calculationIssue(
          coordinate,
          'office.kernel.spreadsheet.formula_unsupported',
          message,
        ),
      );
    }
    const parser = new this.formulaParser.Parser();
    let unresolvedDependency = false;
    let unsupportedReference = false;
    let unsupportedFunction: string | undefined;
    let materializedRangeCells = 0;

    parser
      .setFunction('IFERROR', evaluateParserIfError)
      .setFunction('ROW', (parameters) =>
        parameters.length ? null : coordinate.row + 1,
      )
      .setFunction('COLUMN', (parameters) =>
        parameters.length ? null : coordinate.column + 1,
      );
    parser.on('callFunction', (name, parameters, done) => {
      const normalized = normalizeSpreadsheetFunctionName(name);
      const arity = browserScalarFunctionArities.get(normalized);
      if (
        !arity ||
        parameters.length < arity[0] ||
        parameters.length > arity[1]
      ) {
        unsupportedFunction ??= name.toUpperCase();
      }
      if (normalized === 'SUBTOTAL') {
        done(evaluateParserSubtotal(parameters));
      }
    });
    parser.on('callCellValue', (cell, _options, done) => {
      const dependency = this.resolveCoordinate(indexed.sheet, cell);
      if (!dependency) {
        unsupportedReference = true;
        done(null);
        return;
      }
      const value = this.evaluateCoordinate(dependency);
      if (this.dependencyRequiresCompatibility(dependency)) {
        unresolvedDependency = true;
      }
      done(spreadsheetValueForParser(value));
    });
    parser.on('callRangeValue', (start, end, _options, done) => {
      done(
        this.rangeValues(
          indexed.sheet,
          start,
          end,
          () => {
            unresolvedDependency = true;
          },
          () => {
            unsupportedReference = true;
          },
          (cells) => {
            if (cells > MAX_SPREADSHEET_RANGE_CELLS - materializedRangeCells) {
              return false;
            }
            materializedRangeCells += cells;
            return true;
          },
        ),
      );
    });

    const parsed = parser.parse(
      normalizeFormulaForFortuneParser(expandedFormula),
      {
        sheetId: indexed.sheet.id,
      },
    );
    if (unsupportedFunction) {
      return failedEvaluation(
        indexed.cell.value,
        calculationIssue(
          coordinate,
          'office.kernel.spreadsheet.formula_unsupported',
          `Formula function '${unsupportedFunction}' is not supported.`,
        ),
      );
    }
    if (unsupportedReference) {
      return failedEvaluation(
        indexed.cell.value,
        calculationIssue(
          coordinate,
          'office.kernel.spreadsheet.formula_unsupported',
          'A formula reference cannot be materialized by the browser fallback.',
        ),
      );
    }
    if (unresolvedDependency) {
      return failedEvaluation(
        indexed.cell.value,
        calculationIssue(
          coordinate,
          'office.kernel.spreadsheet.dependency_unresolved',
          'A formula dependency requires compatibility calculation.',
        ),
      );
    }
    if (parsed.error) {
      const error = recognizedSpreadsheetError(parsed.error);
      if (error) {
        return {
          successful: true,
          value: { kind: 'error', value: error },
        };
      }
      return failedEvaluation(
        indexed.cell.value,
        calculationIssue(
          coordinate,
          'office.kernel.spreadsheet.formula_invalid',
          'Formula calculation failed with #VALUE!.',
        ),
      );
    }
    return {
      successful: true,
      value: spreadsheetValueFromParser(parsed.result),
    };
  }

  private rangeValues(
    currentSheet: OfficeKernelSpreadsheetInputSheet,
    start: FormulaParserCoordinate,
    end: FormulaParserCoordinate,
    onUnresolvedDependency: () => void,
    onUnsupportedReference: () => void,
    reserveRangeCells: (cells: number) => boolean,
  ): unknown[][] {
    const startCoordinate = this.resolveCoordinate(currentSheet, start);
    // A qualified range such as `Sales!A2:A3` carries the worksheet only on
    // its first endpoint in Fortune's parser. Resolve the second endpoint
    // against that same worksheet instead of accidentally defaulting it to
    // the formula cell's sheet.
    const rangeSheet = start.sheetName
      ? this.sheetsByName.get(start.sheetName.toLowerCase())
      : currentSheet;
    const endCoordinate = rangeSheet
      ? this.resolveCoordinate(rangeSheet, end)
      : null;
    if (
      !startCoordinate ||
      !endCoordinate ||
      startCoordinate.sheetId !== endCoordinate.sheetId
    ) {
      onUnsupportedReference();
      return [];
    }

    const rowStart = Math.min(startCoordinate.row, endCoordinate.row);
    const rowEnd = Math.max(startCoordinate.row, endCoordinate.row);
    const columnStart = Math.min(startCoordinate.column, endCoordinate.column);
    const columnEnd = Math.max(startCoordinate.column, endCoordinate.column);
    const cells = (rowEnd - rowStart + 1) * (columnEnd - columnStart + 1);
    if (cells > MAX_SPREADSHEET_RANGE_CELLS || !reserveRangeCells(cells)) {
      onUnsupportedReference();
      return [];
    }
    const values: unknown[][] = [];
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const rowValues: unknown[] = [];
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const dependency = {
          sheetId: startCoordinate.sheetId,
          row,
          column,
        };
        const value = this.evaluateCoordinate(dependency);
        if (this.dependencyRequiresCompatibility(dependency)) {
          onUnresolvedDependency();
          return [];
        }
        rowValues.push(spreadsheetValueForParser(value));
      }
      values.push(rowValues);
    }
    return values;
  }

  private dependencyRequiresCompatibility(
    coordinate: OfficeKernelSpreadsheetCoordinate,
  ): boolean {
    const state = this.states.get(coordinateKey(coordinate));
    return state?.successful === false;
  }

  private resolveCoordinate(
    currentSheet: OfficeKernelSpreadsheetInputSheet,
    coordinate: FormulaParserCoordinate,
  ): OfficeKernelSpreadsheetCoordinate | null {
    const sheet = coordinate.sheetName
      ? this.sheetsByName.get(coordinate.sheetName.toLowerCase())
      : currentSheet;
    if (
      !sheet ||
      !boundedSpreadsheetIndex(
        coordinate.row.index,
        OFFICE_KERNEL_SPREADSHEET_MAX_ROWS,
      ) ||
      !boundedSpreadsheetIndex(
        coordinate.column.index,
        OFFICE_KERNEL_SPREADSHEET_MAX_COLUMNS,
      )
    ) {
      return null;
    }
    return {
      sheetId: sheet.id,
      row: coordinate.row.index,
      column: coordinate.column.index,
    };
  }

  private recordState(
    key: string,
    value: OfficeKernelSpreadsheetValue,
    successful: boolean,
    issue?: OfficeKernelSpreadsheetCalculationIssue,
  ): void {
    if (this.states.has(key)) return;
    this.states.set(key, { value, successful, issue });
    const indexed = this.cells.get(key);
    if (indexed?.cell.formula) {
      this.calculationOrder.push({
        sheetId: indexed.sheet.id,
        row: indexed.cell.row,
        column: indexed.cell.column,
      });
    }
    if (issue) this.issues.push(issue);
  }

  private markCycle(start: number): void {
    for (const key of this.stack.slice(start)) {
      const indexed = this.cells.get(key);
      if (!indexed) continue;
      this.recordState(
        key,
        indexed.cell.value,
        false,
        calculationIssue(
          {
            sheetId: indexed.sheet.id,
            row: indexed.cell.row,
            column: indexed.cell.column,
          },
          'office.kernel.spreadsheet.circular_reference',
          'Formula dependency cycle is not supported by this calculation pass.',
        ),
      );
    }
  }
}

function failedEvaluation(
  value: OfficeKernelSpreadsheetValue,
  issue: OfficeKernelSpreadsheetCalculationIssue,
): EvaluationState {
  return { successful: false, value, issue };
}

function calculationIssue(
  cell: OfficeKernelSpreadsheetCoordinate,
  code: string,
  message: string,
): OfficeKernelSpreadsheetCalculationIssue {
  return { cell, code, message };
}

function coordinateKey(coordinate: OfficeKernelSpreadsheetCoordinate): string {
  return cellKey(coordinate.sheetId, coordinate.row, coordinate.column);
}

function cellKey(sheetId: string, row: number, column: number): string {
  return `${sheetId}\u0000${row}\u0000${column}`;
}
