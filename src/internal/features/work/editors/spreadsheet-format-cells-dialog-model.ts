import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  spreadsheetUnderlineStyle,
  type SpreadsheetUnderlineStyle,
} from '../work-spreadsheet-underline';
import { spreadsheetVisibleTextRotationFromCell } from '../work-spreadsheet-text-orientation';
import type { XlsxGradientFill } from '../work-xlsx-gradient-fill';
import type { XlsxPatternFill } from '../work-xlsx-pattern-fill';
import {
  spreadsheetCellBordersAt,
  type SpreadsheetCellBorderFormat,
  type SpreadsheetCellBorderStyle,
} from './spreadsheet-cell-border';
import {
  MAX_SPREADSHEET_CELL_FORMAT_CELLS,
  spreadsheetCellProtectionAt,
  type SpreadsheetCellFormatPatch,
  type SpreadsheetHorizontalAlignment,
  type SpreadsheetVerticalAlignment,
} from './spreadsheet-cell-format';
import {
  normalizeSpreadsheetCellFillFormat,
  spreadsheetCellFillFormat,
  type SpreadsheetCellFillFormat,
} from './spreadsheet-cell-fill-format';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';

export {
  spreadsheetFormatCellsTabs,
  type SpreadsheetFormatCellsTabId,
} from './spreadsheet-format-cells-intent';

export interface SpreadsheetFormatCellsField<T> {
  value: T;
  mixed: boolean;
}

export interface SpreadsheetFormatCellsFields {
  numberFormat: SpreadsheetFormatCellsField<string>;
  horizontalAlignment: SpreadsheetFormatCellsField<SpreadsheetHorizontalAlignment>;
  verticalAlignment: SpreadsheetFormatCellsField<SpreadsheetVerticalAlignment>;
  wrapText: SpreadsheetFormatCellsField<boolean>;
  rotation: SpreadsheetFormatCellsField<number>;
  fontFamily: SpreadsheetFormatCellsField<string>;
  fontSize: SpreadsheetFormatCellsField<number>;
  fontColor: SpreadsheetFormatCellsField<string>;
  bold: SpreadsheetFormatCellsField<boolean>;
  italic: SpreadsheetFormatCellsField<boolean>;
  underline: SpreadsheetFormatCellsField<SpreadsheetUnderlineStyle>;
  strike: SpreadsheetFormatCellsField<boolean>;
  fill: SpreadsheetFormatCellsField<SpreadsheetCellFillFormat>;
  borders: SpreadsheetFormatCellsField<readonly SpreadsheetCellBorderFormat[]>;
  locked: SpreadsheetFormatCellsField<boolean>;
  hidden: SpreadsheetFormatCellsField<boolean>;
}

export interface SpreadsheetFormatCellsDialogSource {
  sheetId: string;
  range: SpreadsheetCellRange;
  activeCell: Cell | null;
  fields: SpreadsheetFormatCellsFields;
}

export interface SpreadsheetFormatCellsDraft {
  numberFormat: string;
  horizontalAlignment: SpreadsheetHorizontalAlignment;
  verticalAlignment: SpreadsheetVerticalAlignment;
  wrapText: boolean;
  rotation: number;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  bold: boolean;
  italic: boolean;
  underline: SpreadsheetUnderlineStyle;
  strike: boolean;
  fill: SpreadsheetFormatCellsFillDraft;
  borders: SpreadsheetCellBorderFormat[];
  borderColor: string;
  borderStyle: SpreadsheetCellBorderStyle;
  locked: boolean;
  hidden: boolean;
}

export interface SpreadsheetFormatCellsFillDraft {
  gradient: XlsxGradientFill;
  mode: SpreadsheetCellFillFormat['kind'];
  pattern: XlsxPatternFill;
  solidColor: string;
}

export type SpreadsheetFormatCellsTouched = Partial<
  Record<keyof SpreadsheetCellFormatPatch, true>
>;

export interface SpreadsheetFormatCellsDraftErrors {
  fill?: string;
  numberFormat?: string;
  fontSize?: string;
  rotation?: string;
}

export function createSpreadsheetFormatCellsDialogSource(
  content: WorkSpreadsheetContent,
  sheetId: string,
  rangeInput: SpreadsheetCellRange,
  cells: readonly (readonly (Cell | null)[])[],
  requestedActiveCell: { row: number; column: number },
): SpreadsheetFormatCellsDialogSource | null {
  const range = normalizeSpreadsheetCellRange(rangeInput);
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  if (
    !range ||
    !sheet ||
    spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_CELL_FORMAT_CELLS
  ) {
    return null;
  }
  const activeRow = clampIndex(
    requestedActiveCell.row,
    range.row[0],
    range.row[1],
  );
  const activeColumn = clampIndex(
    requestedActiveCell.column,
    range.column[0],
    range.column[1],
  );
  const activeCell =
    cells[activeRow - range.row[0]]?.[activeColumn - range.column[0]] ?? null;
  const values = <T>(
    resolve: (cell: Cell | null, row: number, column: number) => T,
  ): SpreadsheetFormatCellsField<T> => {
    const activeValue = resolve(activeCell, activeRow, activeColumn);
    let mixed = false;
    for (let row = range.row[0]; row <= range.row[1] && !mixed; row += 1) {
      for (
        let column = range.column[0];
        column <= range.column[1];
        column += 1
      ) {
        const cell =
          cells[row - range.row[0]]?.[column - range.column[0]] ?? null;
        if (!sameDialogValue(resolve(cell, row, column), activeValue)) {
          mixed = true;
          break;
        }
      }
    }
    return { value: activeValue, mixed };
  };
  const borders = borderFormatsAt(sheet, activeRow, activeColumn);
  return {
    sheetId,
    range,
    activeCell,
    fields: {
      numberFormat: values((cell) => cell?.ct?.fa?.trim() || 'General'),
      horizontalAlignment: values((cell) => horizontalAlignment(cell?.ht)),
      verticalAlignment: values((cell) => verticalAlignment(cell?.vt)),
      wrapText: values((cell) => String(cell?.tb) === '2'),
      rotation: values((cell) => spreadsheetVisibleTextRotationFromCell(cell)),
      fontFamily: values((cell) =>
        typeof cell?.ff === 'string' && cell.ff.trim()
          ? cell.ff.trim()
          : 'Aptos',
      ),
      fontSize: values((cell) =>
        typeof cell?.fs === 'number' && Number.isFinite(cell.fs) ? cell.fs : 10,
      ),
      fontColor: values((cell) => normalizedColor(cell?.fc, '#172033')),
      bold: values((cell) => Number(cell?.bl) === 1),
      italic: values((cell) => Number(cell?.it) === 1),
      underline: values((cell) => spreadsheetUnderlineStyle(cell?.un)),
      strike: values((cell) => Number(cell?.cl) === 1),
      fill: values((cell) => spreadsheetCellFillFormat(cell)),
      borders: {
        value: borders,
        mixed: spreadsheetCellRangeArea(range) > 1,
      },
      locked: values(
        (cell, row, column) =>
          spreadsheetCellProtectionAt(sheet, row, column, cell).locked,
      ),
      hidden: values(
        (cell, row, column) =>
          spreadsheetCellProtectionAt(sheet, row, column, cell).hidden,
      ),
    },
  };
}

export function createSpreadsheetFormatCellsDraft(
  source: SpreadsheetFormatCellsDialogSource,
): SpreadsheetFormatCellsDraft {
  const borders = source.fields.borders.value.map((format) => ({ ...format }));
  return {
    numberFormat: source.fields.numberFormat.value,
    horizontalAlignment: source.fields.horizontalAlignment.value,
    verticalAlignment: source.fields.verticalAlignment.value,
    wrapText: source.fields.wrapText.value,
    rotation: source.fields.rotation.value,
    fontFamily: source.fields.fontFamily.value,
    fontSize: source.fields.fontSize.value,
    fontColor: source.fields.fontColor.value,
    bold: source.fields.bold.value,
    italic: source.fields.italic.value,
    underline: source.fields.underline.value,
    strike: source.fields.strike.value,
    fill: createSpreadsheetFormatCellsFillDraft(source.fields.fill.value),
    borders,
    borderColor: borders[0]?.color ?? '#172033',
    borderStyle: borders[0]?.style ?? 'thin',
    locked: source.fields.locked.value,
    hidden: source.fields.hidden.value,
  };
}

export function spreadsheetFormatCellsPatch(
  source: SpreadsheetFormatCellsDialogSource,
  draft: SpreadsheetFormatCellsDraft,
  touched: SpreadsheetFormatCellsTouched,
): SpreadsheetCellFormatPatch {
  const patch: SpreadsheetCellFormatPatch = {};
  if (
    shouldEmit(
      source.fields.numberFormat,
      draft.numberFormat,
      touched.numberFormat,
    )
  )
    patch.numberFormat = draft.numberFormat.trim();
  if (
    shouldEmit(
      source.fields.horizontalAlignment,
      draft.horizontalAlignment,
      touched.horizontalAlignment,
    )
  )
    patch.horizontalAlignment = draft.horizontalAlignment;
  if (
    shouldEmit(
      source.fields.verticalAlignment,
      draft.verticalAlignment,
      touched.verticalAlignment,
    )
  )
    patch.verticalAlignment = draft.verticalAlignment;
  if (shouldEmit(source.fields.wrapText, draft.wrapText, touched.wrapText))
    patch.wrapText = draft.wrapText;
  if (shouldEmit(source.fields.rotation, draft.rotation, touched.rotation))
    patch.rotation = draft.rotation;
  if (
    shouldEmit(source.fields.fontFamily, draft.fontFamily, touched.fontFamily)
  )
    patch.fontFamily = draft.fontFamily.trim();
  if (shouldEmit(source.fields.fontSize, draft.fontSize, touched.fontSize))
    patch.fontSize = draft.fontSize;
  if (shouldEmit(source.fields.fontColor, draft.fontColor, touched.fontColor))
    patch.fontColor = draft.fontColor;
  if (shouldEmit(source.fields.bold, draft.bold, touched.bold))
    patch.bold = draft.bold;
  if (shouldEmit(source.fields.italic, draft.italic, touched.italic))
    patch.italic = draft.italic;
  if (shouldEmit(source.fields.underline, draft.underline, touched.underline))
    patch.underline = draft.underline;
  if (shouldEmit(source.fields.strike, draft.strike, touched.strike))
    patch.strike = draft.strike;
  const fill = spreadsheetFormatCellsActiveFill(draft);
  if (shouldEmit(source.fields.fill, fill, touched.fill)) patch.fill = fill;
  if (shouldEmit(source.fields.borders, draft.borders, touched.borders))
    patch.borders = draft.borders.map((format) => ({ ...format }));
  if (shouldEmit(source.fields.locked, draft.locked, touched.locked))
    patch.locked = draft.locked;
  if (shouldEmit(source.fields.hidden, draft.hidden, touched.hidden))
    patch.hidden = draft.hidden;
  return patch;
}

export function spreadsheetFormatCellsDraftErrors(
  draft: SpreadsheetFormatCellsDraft,
): SpreadsheetFormatCellsDraftErrors {
  const errors: SpreadsheetFormatCellsDraftErrors = {};
  if (!draft.numberFormat.trim()) {
    errors.numberFormat = '请输入数字格式代码。';
  } else if (draft.numberFormat.trim().length > 255) {
    errors.numberFormat = '数字格式代码不能超过 255 个字符。';
  }
  if (
    !Number.isFinite(draft.fontSize) ||
    draft.fontSize < 1 ||
    draft.fontSize > 409
  ) {
    errors.fontSize = '字号需为 1–409 之间的数字。';
  }
  if (
    !Number.isInteger(draft.rotation) ||
    draft.rotation < -90 ||
    draft.rotation > 90
  ) {
    errors.rotation = '文字旋转角度需为 -90–90 之间的整数。';
  }
  if (
    !normalizeSpreadsheetCellFillFormat(spreadsheetFormatCellsActiveFill(draft))
  ) {
    errors.fill =
      '请检查填充设置。渐变色标需按位置从小到大排列，位置和路径边界需保持在 0%–100%。';
  }
  return errors;
}

export function spreadsheetFormatCellsActiveFill(
  draft: SpreadsheetFormatCellsDraft,
): SpreadsheetCellFillFormat {
  if (draft.fill.mode === 'none') return { kind: 'none' };
  if (draft.fill.mode === 'solid') {
    return { color: draft.fill.solidColor, kind: 'solid' };
  }
  if (draft.fill.mode === 'pattern') {
    return { kind: 'pattern', value: draft.fill.pattern };
  }
  return { kind: 'gradient', value: draft.fill.gradient };
}

function shouldEmit<T>(
  source: SpreadsheetFormatCellsField<T>,
  value: T,
  touched: true | undefined,
): boolean {
  return Boolean(
    touched && (source.mixed || !sameDialogValue(source.value, value)),
  );
}

function createSpreadsheetFormatCellsFillDraft(
  source: SpreadsheetCellFillFormat,
): SpreadsheetFormatCellsFillDraft {
  const fill = normalizeSpreadsheetCellFillFormat(source) ?? { kind: 'none' };
  const solidColor = spreadsheetFillFallbackColor(fill);
  return {
    gradient:
      fill.kind === 'gradient'
        ? fill.value
        : {
            degree: 0,
            stops: [
              { color: solidColor, position: 0 },
              {
                color: solidColor === '#ffffff' ? '#4472c4' : '#ffffff',
                position: 1,
              },
            ],
            type: 'linear',
          },
    mode: fill.kind,
    pattern:
      fill.kind === 'pattern'
        ? fill.value
        : {
            backgroundColor: solidColor,
            foregroundColor: solidColor === '#ffffff' ? '#4472c4' : '#ffffff',
            patternType: 'lightGrid',
          },
    solidColor,
  };
}

function spreadsheetFillFallbackColor(fill: SpreadsheetCellFillFormat): string {
  if (fill.kind === 'solid') return fill.color;
  if (fill.kind === 'pattern') return fill.value.backgroundColor;
  if (fill.kind === 'gradient') return fill.value.stops[0]?.color ?? '#ffffff';
  return '#ffffff';
}

function horizontalAlignment(
  value: Cell['ht'],
): SpreadsheetHorizontalAlignment {
  if (String(value) === '0') return 'center';
  if (String(value) === '1') return 'left';
  if (String(value) === '2') return 'right';
  return 'general';
}

function verticalAlignment(value: Cell['vt']): SpreadsheetVerticalAlignment {
  if (Number(value) === 0) return 'middle';
  if (Number(value) === 2) return 'bottom';
  return 'top';
}

function borderFormatsAt(
  sheet: WorkSpreadsheetContent['sheets'][number],
  row: number,
  column: number,
): SpreadsheetCellBorderFormat[] {
  const resolved = spreadsheetCellBordersAt(sheet, row, column);
  return (
    ['top', 'bottom', 'left', 'right', 'diagonalDown', 'diagonalUp'] as const
  ).flatMap((target) => {
    const line = resolved[target];
    if (!line) return [];
    return [
      {
        target,
        color: normalizedColor(line.color, '#172033') ?? '#172033',
        style: borderStyle(line.style),
      },
    ];
  });
}

function borderStyle(value: string): SpreadsheetCellBorderStyle {
  return (
    (
      {
        '1': 'thin',
        '3': 'dotted',
        '4': 'dashed',
        '5': 'dash-dot',
        '6': 'dash-dot-dot',
        '8': 'medium',
        '9': 'medium-dashed',
        '10': 'medium-dash-dot',
        '11': 'medium-dash-dot-dot',
        '13': 'thick',
      } as const
    )[value] ?? 'thin'
  );
}

function normalizedColor<T extends string | null>(
  value: unknown,
  fallback: T,
): string | T {
  if (typeof value !== 'string') return fallback;
  const color = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(color)) return color;
  if (/^#[0-9a-f]{3}$/.test(color)) {
    return `#${[...color.slice(1)]
      .map((character) => character.repeat(2))
      .join('')}`;
  }
  return fallback;
}

function sameDialogValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  return JSON.stringify(left) === JSON.stringify(right);
}

function clampIndex(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, Math.trunc(value)));
}
