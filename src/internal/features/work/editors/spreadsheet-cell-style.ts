import type { Cell } from '@fortune-sheet/core';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
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
export type SpreadsheetCellStyleGroup =
  | 'common'
  | 'dataAndModel'
  | 'titlesAndTotals';

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
  group: SpreadsheetCellStyleGroup;
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
  group: SpreadsheetCellStyleGroup,
  format: Partial<SpreadsheetCellStyleFormat>,
  borders: readonly SpreadsheetCellBorderFormat[] | undefined,
  catalog: OfficeMessageCatalog,
): SpreadsheetCellStyleDefinition => {
  const completeFormat = { ...normalFormat, ...format };
  return {
    id,
    label: officeMessage(catalog, `spreadsheet.cellStyle.${id}`),
    group,
    description: officeMessage(catalog, `spreadsheet.cellStyle.${id}.desc`),
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

const CELL_STYLE_SPECS = [
  { id: 'normal', group: 'common', format: {}, borders: undefined },
  {
    id: 'good',
    group: 'common',
    format: { bg: '#c6efce', fc: '#006100' },
    borders: undefined,
  },
  {
    id: 'bad',
    group: 'common',
    format: { bg: '#ffc7ce', fc: '#9c0006' },
    borders: undefined,
  },
  {
    id: 'neutral',
    group: 'common',
    format: { bg: '#ffeb9c', fc: '#9c5700' },
    borders: undefined,
  },
  {
    id: 'calculation',
    group: 'dataAndModel',
    format: { bl: 1, fc: '#fa7d00' },
    borders: edgeBorder('bottom', '#7f7f7f', 'thin'),
  },
  {
    id: 'checkCell',
    group: 'dataAndModel',
    format: { bg: '#a5a5a5', bl: 1, fc: '#ffffff' },
    borders: allBorder('#7f7f7f'),
  },
  {
    id: 'explanatoryText',
    group: 'dataAndModel',
    format: { fc: '#7f7f7f', it: 1 },
    borders: undefined,
  },
  {
    id: 'input',
    group: 'dataAndModel',
    format: { bg: '#ffffcc', fc: '#3f3f76' },
    borders: allBorder('#7f8fa6'),
  },
  {
    id: 'linkedCell',
    group: 'dataAndModel',
    format: { fc: '#0563c1', un: 1 },
    borders: undefined,
  },
  {
    id: 'note',
    group: 'dataAndModel',
    format: { bg: '#ffffcc', fc: '#3f3f3f' },
    borders: undefined,
  },
  {
    id: 'output',
    group: 'dataAndModel',
    format: { bg: '#f2f2f2', bl: 1, fc: '#3f3f3f' },
    borders: allBorder('#7f8fa6'),
  },
  {
    id: 'warningText',
    group: 'dataAndModel',
    format: { bl: 1, fc: '#c00000' },
    borders: undefined,
  },
  {
    id: 'heading1',
    group: 'titlesAndTotals',
    format: { bl: 1, fc: '#1f4e78', fs: 15 },
    borders: edgeBorder('bottom', '#5b9bd5', 'thick'),
  },
  {
    id: 'heading2',
    group: 'titlesAndTotals',
    format: { bl: 1, fc: '#1f4e78', fs: 13 },
    borders: edgeBorder('bottom', '#5b9bd5', 'medium'),
  },
  {
    id: 'heading3',
    group: 'titlesAndTotals',
    format: { bl: 1, fc: '#1f4e78', fs: 11 },
    borders: undefined,
  },
  {
    id: 'heading4',
    group: 'titlesAndTotals',
    format: { bl: 1, fc: '#1f4e78', fs: 11, it: 1 },
    borders: undefined,
  },
  {
    id: 'total',
    group: 'titlesAndTotals',
    format: { bl: 1 },
    borders: edgeBorder('top', '#172033', 'medium'),
  },
] as const satisfies readonly {
  id: SpreadsheetCellStyleChoice;
  group: SpreadsheetCellStyleGroup;
  format: Partial<SpreadsheetCellStyleFormat>;
  borders: readonly SpreadsheetCellBorderFormat[] | undefined;
}[];

let cachedDefinitions: readonly SpreadsheetCellStyleDefinition[] | null = null;
let cachedDefinitionsLocale: string | null = null;

export function spreadsheetCellStyleDefinitions(
  catalog: OfficeMessageCatalog = resolveOfficeMessages(),
): readonly SpreadsheetCellStyleDefinition[] {
  const localeKey = officeMessage(catalog, 'spreadsheet.cellStyle.normal');
  if (cachedDefinitions && cachedDefinitionsLocale === localeKey) {
    return cachedDefinitions;
  }
  cachedDefinitionsLocale = localeKey;
  cachedDefinitions = Object.freeze(
    CELL_STYLE_SPECS.map((spec) =>
      style(spec.id, spec.group, spec.format, spec.borders, catalog),
    ),
  );
  return cachedDefinitions;
}

export const spreadsheetCellStyleGroups = [
  'common',
  'dataAndModel',
  'titlesAndTotals',
] as const satisfies readonly SpreadsheetCellStyleGroup[];

export function spreadsheetCellStyleGroupLabel(
  group: SpreadsheetCellStyleGroup,
  catalog: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  return officeMessage(catalog, `spreadsheet.cellStyle.group.${group}`);
}

const definitionById = () =>
  new Map(
    spreadsheetCellStyleDefinitions().map((definition) => [
      definition.id,
      definition,
    ]),
  );

export function spreadsheetCellStyleDefinition(
  preset: SpreadsheetCellStyleChoice,
): SpreadsheetCellStyleDefinition {
  const definition = definitionById().get(preset);
  if (!definition) throw new Error(`Unknown spreadsheet cell style: ${preset}`);
  return definition;
}

export function spreadsheetCellStylePreset(
  cell: Cell | null | undefined,
  borders?: SpreadsheetResolvedCellBorders,
): SpreadsheetCellStylePreset {
  const format = normalizedCellStyleFormat(cell);
  return (
    spreadsheetCellStyleDefinitions().find(
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
    definitionById().has(preset) &&
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
  const definition = definitionById().get(preset);
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
