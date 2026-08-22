import type { Cell } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { deleteXlsxNativeFills } from '../work-xlsx-native-fill';
import {
  canSetSpreadsheetCellBorders,
  type SpreadsheetCellBorderFormat,
  type SpreadsheetResolvedCellBorderLine,
  type SpreadsheetResolvedCellBorders,
  spreadsheetNativeBorderStyle,
} from './spreadsheet-cell-border';
import { setSpreadsheetCellBordersPerCell } from './spreadsheet-cell-border-per-cell';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';

export const spreadsheetCellStylePresetIds = [
  'normal',
  'good',
  'bad',
  'neutral',
  'calculation',
  'checkCell',
  'explanatoryText',
  'input',
  'linkedCell',
  'note',
  'output',
  'warningText',
  'heading1',
  'heading2',
  'heading3',
  'heading4',
  'total',
] as const;

export type SpreadsheetCellStyleChoice =
  (typeof spreadsheetCellStylePresetIds)[number];
export type SpreadsheetCellStylePreset = SpreadsheetCellStyleChoice | 'custom';

export const MAX_SPREADSHEET_CELL_STYLE_CELLS = 10_000;

interface SpreadsheetCellStyleFormat {
  bg: string;
  bl: 0 | 1;
  cl: 0 | 1;
  fc: string;
  ff: string;
  fs: number;
  it: 0 | 1;
  un: 0 | 1;
}

export interface SpreadsheetCellStylePreview {
  backgroundColor: string;
  color: string;
  fontSize?: string;
  fontStyle?: 'italic' | 'normal';
  fontWeight?: number;
  textDecoration?: 'none' | 'underline';
}

export interface SpreadsheetCellStyleDefinition {
  id: SpreadsheetCellStyleChoice;
  label: string;
  group: '常用' | '数据和模型' | '标题和汇总';
  description: string;
  format: SpreadsheetCellStyleFormat;
  borders?: readonly SpreadsheetCellBorderFormat[];
  preview: SpreadsheetCellStylePreview;
}

const normalFormat: SpreadsheetCellStyleFormat = {
  bg: '#ffffff',
  bl: 0,
  cl: 0,
  fc: '#172033',
  ff: 'Aptos',
  fs: 10,
  it: 0,
  un: 0,
};

const style = (
  id: SpreadsheetCellStyleChoice,
  label: string,
  group: SpreadsheetCellStyleDefinition['group'],
  description: string,
  format: Partial<SpreadsheetCellStyleFormat>,
  borders?: readonly SpreadsheetCellBorderFormat[],
): SpreadsheetCellStyleDefinition => {
  const completeFormat = { ...normalFormat, ...format };
  return {
    id,
    label,
    group,
    description,
    format: completeFormat,
    borders,
    preview: {
      backgroundColor: completeFormat.bg,
      color: completeFormat.fc,
      fontSize: `${Math.max(10, completeFormat.fs)}px`,
      fontStyle: completeFormat.it ? 'italic' : 'normal',
      fontWeight: completeFormat.bl ? 700 : 500,
      textDecoration: completeFormat.un ? 'underline' : 'none',
    },
  };
};

const allBorder = (
  color: string,
  styleName: SpreadsheetCellBorderFormat['style'] = 'thin',
): readonly SpreadsheetCellBorderFormat[] => [
  { target: 'all', color, style: styleName },
];

const edgeBorder = (
  target: 'bottom' | 'top',
  color: string,
  styleName: SpreadsheetCellBorderFormat['style'],
): readonly SpreadsheetCellBorderFormat[] => [
  { target, color, style: styleName },
];

export const spreadsheetCellStyleDefinitions = [
  style('normal', '常规', '常用', '恢复默认字体、颜色和无边框外观。', {}),
  style('good', '好', '常用', '突出正向结果或已完成状态。', {
    bg: '#c6efce',
    fc: '#006100',
  }),
  style('bad', '差', '常用', '突出错误、失败或超限状态。', {
    bg: '#ffc7ce',
    fc: '#9c0006',
  }),
  style('neutral', '适中', '常用', '突出需要关注但尚未定性的状态。', {
    bg: '#ffeb9c',
    fc: '#9c5700',
  }),
  style(
    'calculation',
    '计算',
    '数据和模型',
    '标记模型中的计算单元格。',
    {
      bl: 1,
      fc: '#fa7d00',
    },
    edgeBorder('bottom', '#7f7f7f', 'thin'),
  ),
  style(
    'checkCell',
    '检查单元格',
    '数据和模型',
    '标记需要人工复核的值。',
    {
      bg: '#a5a5a5',
      bl: 1,
      fc: '#ffffff',
    },
    allBorder('#7f7f7f'),
  ),
  style(
    'explanatoryText',
    '解释性文本',
    '数据和模型',
    '为模型补充低强调度的说明。',
    { fc: '#7f7f7f', it: 1 },
  ),
  style(
    'input',
    '输入',
    '数据和模型',
    '标记允许用户录入的单元格。',
    {
      bg: '#ffffcc',
      fc: '#3f3f76',
    },
    allBorder('#7f8fa6'),
  ),
  style(
    'linkedCell',
    '链接单元格',
    '数据和模型',
    '标记来自其他位置的链接值。',
    {
      fc: '#0563c1',
      un: 1,
    },
  ),
  style('note', '注释', '数据和模型', '标记工作表中的说明或备注。', {
    bg: '#ffffcc',
    fc: '#3f3f3f',
  }),
  style(
    'output',
    '输出',
    '数据和模型',
    '标记模型输出或最终结果。',
    {
      bg: '#f2f2f2',
      bl: 1,
      fc: '#3f3f3f',
    },
    allBorder('#7f8fa6'),
  ),
  style('warningText', '警告文本', '数据和模型', '用高对比文字提示风险。', {
    bl: 1,
    fc: '#c00000',
  }),
  style(
    'heading1',
    '标题 1',
    '标题和汇总',
    '用于工作表的一级标题。',
    {
      bl: 1,
      fc: '#1f4e78',
      fs: 15,
    },
    edgeBorder('bottom', '#5b9bd5', 'thick'),
  ),
  style(
    'heading2',
    '标题 2',
    '标题和汇总',
    '用于工作表的二级标题。',
    {
      bl: 1,
      fc: '#1f4e78',
      fs: 13,
    },
    edgeBorder('bottom', '#5b9bd5', 'medium'),
  ),
  style('heading3', '标题 3', '标题和汇总', '用于工作表的三级标题。', {
    bl: 1,
    fc: '#1f4e78',
    fs: 11,
  }),
  style('heading4', '标题 4', '标题和汇总', '用于工作表的四级标题。', {
    bl: 1,
    fc: '#1f4e78',
    fs: 11,
    it: 1,
  }),
  style(
    'total',
    '总计',
    '标题和汇总',
    '强调汇总行或最终合计。',
    {
      bl: 1,
    },
    edgeBorder('top', '#172033', 'medium'),
  ),
] as const satisfies readonly SpreadsheetCellStyleDefinition[];

const definitionById = new Map(
  spreadsheetCellStyleDefinitions.map((definition) => [
    definition.id,
    definition,
  ]),
);

export function spreadsheetCellStyleDefinition(
  preset: SpreadsheetCellStyleChoice,
): SpreadsheetCellStyleDefinition {
  const definition = definitionById.get(preset);
  if (!definition) throw new Error(`Unknown spreadsheet cell style: ${preset}`);
  return definition;
}

export function spreadsheetCellStylePreset(
  cell: Cell | null | undefined,
  borders?: SpreadsheetResolvedCellBorders,
): SpreadsheetCellStylePreset {
  const format = normalizedCellStyleFormat(cell);
  return (
    spreadsheetCellStyleDefinitions.find(
      (definition) =>
        sameCellStyleFormat(format, definition.format) &&
        (borders
          ? sameCellStyleBorders(borders, definition.borders ?? [])
          : !definition.borders?.length),
    )?.id ?? 'custom'
  );
}

export function canApplySpreadsheetCellStyle(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  preset: SpreadsheetCellStyleChoice,
): boolean {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  return Boolean(
    definitionById.has(preset) &&
      normalizedRange &&
      spreadsheetCellRangeArea(normalizedRange) <=
        MAX_SPREADSHEET_CELL_STYLE_CELLS &&
      canSetSpreadsheetCellBorders(content, sheetId, normalizedRange, {
        target: 'none',
        color: '#000000',
        style: 'thin',
      }),
  );
}

export function applySpreadsheetCellStyle(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  preset: SpreadsheetCellStyleChoice,
): WorkSpreadsheetContent | null {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  const definition = definitionById.get(preset);
  const sheetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = content.sheets[sheetIndex];
  if (
    !definition ||
    !sheet ||
    !normalizedRange ||
    !canApplySpreadsheetCellStyle(content, sheetId, normalizedRange, preset)
  ) {
    return null;
  }

  const data = cloneSparseMatrix(sheet.data);
  data.length = Math.max(data.length, normalizedRange.row[1] + 1);
  for (
    let row = normalizedRange.row[0];
    row <= normalizedRange.row[1];
    row += 1
  ) {
    const values = data[row] ?? [];
    data[row] = values;
    for (
      let column = normalizedRange.column[0];
      column <= normalizedRange.column[1];
      column += 1
    ) {
      const cell = {
        ...(values[column] ?? {}),
        ...definition.format,
      };
      deleteXlsxNativeFills(cell);
      values[column] = cell;
    }
  }

  const nextSheet: WorkSpreadsheetSheet = {
    ...sheet,
    row: Math.max(sheet.row ?? 0, normalizedRange.row[1] + 1),
    column: Math.max(sheet.column ?? 0, normalizedRange.column[1] + 1),
    data,
  };
  const sheets = [...content.sheets];
  sheets[sheetIndex] = nextSheet;
  return setSpreadsheetCellBordersPerCell(
    { ...content, sheets },
    sheetId,
    normalizedRange,
    definition.borders ?? [],
  );
}

function normalizedCellStyleFormat(
  cell: Cell | null | undefined,
): SpreadsheetCellStyleFormat {
  return {
    bg: normalizedStyleColor(cell?.bg, normalFormat.bg),
    bl: Number(cell?.bl) === 1 ? 1 : 0,
    cl: Number(cell?.cl) === 1 ? 1 : 0,
    fc: normalizedStyleColor(cell?.fc, normalFormat.fc),
    ff:
      typeof cell?.ff === 'string' && cell.ff.trim()
        ? cell.ff.trim()
        : normalFormat.ff,
    fs:
      typeof cell?.fs === 'number' && Number.isFinite(cell.fs)
        ? cell.fs
        : normalFormat.fs,
    it: Number(cell?.it) === 1 ? 1 : 0,
    un: Number(cell?.un) === 1 ? 1 : 0,
  };
}

function normalizedStyleColor(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  const color = value.trim().toLocaleLowerCase();
  if (/^#[0-9a-f]{6}$/.test(color)) return color;
  if (/^#[0-9a-f]{3}$/.test(color)) {
    return `#${[...color.slice(1)]
      .map((character) => character.repeat(2))
      .join('')}`;
  }
  return fallback;
}

function sameCellStyleFormat(
  left: SpreadsheetCellStyleFormat,
  right: SpreadsheetCellStyleFormat,
): boolean {
  return (
    left.bg === right.bg &&
    left.bl === right.bl &&
    left.cl === right.cl &&
    left.fc === right.fc &&
    left.ff === right.ff &&
    left.fs === right.fs &&
    left.it === right.it &&
    left.un === right.un
  );
}

function sameCellStyleBorders(
  actual: SpreadsheetResolvedCellBorders,
  formats: readonly SpreadsheetCellBorderFormat[],
): boolean {
  const expected: SpreadsheetResolvedCellBorders = {};
  for (const format of formats) {
    const line: SpreadsheetResolvedCellBorderLine = {
      color: format.color.toLowerCase(),
      style: spreadsheetNativeBorderStyle(format.style),
    };
    if (format.target === 'all' || format.target === 'left') {
      expected.left = line;
    }
    if (format.target === 'all' || format.target === 'right') {
      expected.right = line;
    }
    if (format.target === 'all' || format.target === 'top') {
      expected.top = line;
    }
    if (format.target === 'all' || format.target === 'bottom') {
      expected.bottom = line;
    }
    if (format.target === 'diagonalDown') expected.diagonalDown = line;
    if (format.target === 'diagonalUp') expected.diagonalUp = line;
  }
  return (
    ['top', 'bottom', 'left', 'right', 'diagonalDown', 'diagonalUp'] as const
  ).every(
    (side) =>
      actual[side]?.color === expected[side]?.color &&
      actual[side]?.style === expected[side]?.style,
  );
}
