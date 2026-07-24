declare module '@fortune-sheet/formula-parser' {
  export interface FormulaParserAxis {
    index: number;
    isAbsolute: boolean;
  }

  export interface FormulaParserCoordinate {
    label: string;
    row: FormulaParserAxis;
    column: FormulaParserAxis;
    sheetName?: string | null;
  }

  export interface FormulaParserOptions {
    sheetId?: string;
  }

  export interface FormulaParserResult {
    error: string | null;
    result: unknown;
  }

  export class Parser {
    setFunction(
      name: string,
      listener: (parameters: unknown[]) => unknown,
    ): this;
    on(
      event: 'callCellValue',
      listener: (
        coordinate: FormulaParserCoordinate,
        options: FormulaParserOptions,
        done: (value: unknown) => void,
      ) => void,
    ): this;
    on(
      event: 'callRangeValue',
      listener: (
        start: FormulaParserCoordinate,
        end: FormulaParserCoordinate,
        options: FormulaParserOptions,
        done: (value: unknown[][]) => void,
      ) => void,
    ): this;
    on(
      event: 'callFunction',
      listener: (
        name: string,
        parameters: unknown[],
        done: (value: unknown) => void,
      ) => void,
    ): this;
    parse(
      expression: string,
      options?: FormulaParserOptions,
    ): FormulaParserResult;
  }

  export const SUPPORTED_FORMULAS: string[];
}
