import type {
  WorkSpreadsheetTable,
  WorkSpreadsheetTableStyle,
} from '../work-types';

export interface SpreadsheetTableStylePalette {
  border: string;
  header: string;
  headerText: string;
  primaryRow: string;
  secondaryRow: string;
  stripeColumn: string;
  text: string;
  total: string;
  totalText: string;
}

export interface SpreadsheetTableStyleChoice {
  label: string;
  ooxmlName: string;
  palette: SpreadsheetTableStylePalette;
  style: Exclude<WorkSpreadsheetTableStyle, { family: 'none' }>;
}

export interface SpreadsheetTableCellRenderStyle {
  background: string;
  bold: boolean;
  borderColor: string;
  role: 'body' | 'header' | 'totals';
  tableId: string;
  textColor: string;
}

const ACCENTS = [
  '#64748b',
  '#2563eb',
  '#0f766e',
  '#7c3aed',
  '#c2410c',
  '#be123c',
  '#15803d',
] as const;

const DARK_ACCENTS = [
  '#1e293b',
  '#1e3a8a',
  '#134e4a',
  '#4c1d95',
  '#7c2d12',
  '#881337',
  '#14532d',
  '#0f172a',
  '#312e81',
  '#3f3f46',
  '#422006',
] as const;

let cachedChoices: readonly SpreadsheetTableStyleChoice[] | null = null;

export function spreadsheetTableStyleChoices(): readonly SpreadsheetTableStyleChoice[] {
  cachedChoices ??= Object.freeze([
    ...tableStyleFamilyChoices('light', 21),
    ...tableStyleFamilyChoices('medium', 28),
    ...tableStyleFamilyChoices('dark', 11),
  ]);
  return cachedChoices;
}

export function spreadsheetTableStylePalette(
  style: WorkSpreadsheetTableStyle,
): SpreadsheetTableStylePalette | null {
  if (style.family === 'none') return null;
  const maximum =
    style.family === 'light' ? 21 : style.family === 'medium' ? 28 : 11;
  if (
    !Number.isInteger(style.number) ||
    style.number < 1 ||
    style.number > maximum
  ) {
    return null;
  }
  if (style.family === 'light') return lightPalette(style.number);
  if (style.family === 'medium') return mediumPalette(style.number);
  return darkPalette(style.number);
}

export function createSpreadsheetTableRenderResolver(
  tables: readonly WorkSpreadsheetTable[],
): (row: number, column: number) => SpreadsheetTableCellRenderStyle | null {
  const renderable = tables
    .filter((table) => spreadsheetTableStylePalette(table.style))
    .slice()
    .sort(
      (left, right) =>
        left.range.row[0] - right.range.row[0] ||
        left.range.column[0] - right.range.column[0],
    );
  let cachedRow = -1;
  let rowTables: WorkSpreadsheetTable[] = [];
  return (row, column) => {
    if (row !== cachedRow) {
      cachedRow = row;
      rowTables = renderable.filter(
        (table) => row >= table.range.row[0] && row <= table.range.row[1],
      );
    }
    const table = rowTables.find(
      (candidate) =>
        column >= candidate.range.column[0] &&
        column <= candidate.range.column[1],
    );
    return table ? spreadsheetTableCellRenderStyle(table, row, column) : null;
  };
}

function spreadsheetTableCellRenderStyle(
  table: WorkSpreadsheetTable,
  row: number,
  column: number,
): SpreadsheetTableCellRenderStyle | null {
  const palette = spreadsheetTableStylePalette(table.style);
  if (!palette) return null;
  const header = table.headerRow && row === table.range.row[0];
  const totals = table.totalsRow && row === table.range.row[1];
  const dataStart = table.range.row[0] + Number(table.headerRow);
  const dataIndex = Math.max(0, row - dataStart);
  const secondary = table.showRowStripes && dataIndex % 2 === 1;
  let background = header
    ? palette.header
    : totals
      ? palette.total
      : secondary
        ? palette.secondaryRow
        : palette.primaryRow;
  if (
    !header &&
    !totals &&
    table.showColumnStripes &&
    (column - table.range.column[0]) % 2 === 1
  ) {
    background = mixHex(background, palette.stripeColumn, 0.42);
  }
  const firstColumn = column === table.range.column[0];
  const lastColumn = column === table.range.column[1];
  return {
    background,
    bold:
      header ||
      totals ||
      (table.showFirstColumn && firstColumn) ||
      (table.showLastColumn && lastColumn),
    borderColor: palette.border,
    role: header ? 'header' : totals ? 'totals' : 'body',
    tableId: table.id,
    textColor: header
      ? palette.headerText
      : totals
        ? palette.totalText
        : palette.text,
  };
}

function tableStyleFamilyChoices(
  family: 'dark' | 'light' | 'medium',
  count: number,
): SpreadsheetTableStyleChoice[] {
  return Array.from({ length: count }, (_, index) => {
    const number = index + 1;
    const style = { family, number } as const;
    return {
      label: `${familyLabel(family)} ${number}`,
      ooxmlName: `TableStyle${capitalize(family)}${number}`,
      palette: spreadsheetTableStylePalette(style)!,
      style,
    };
  });
}

function lightPalette(number: number): SpreadsheetTableStylePalette {
  const accent = ACCENTS[(number - 1) % ACCENTS.length] ?? ACCENTS[0];
  const variant = Math.floor((number - 1) / ACCENTS.length);
  const headerWeight = [0.84, 0.73, 0.62][variant] ?? 0.72;
  return {
    border: mixHex(accent, '#ffffff', 0.5),
    header: mixHex(accent, '#ffffff', headerWeight),
    headerText: mixHex(accent, '#0f172a', 0.38),
    primaryRow: '#ffffff',
    secondaryRow: mixHex(accent, '#ffffff', 0.9 - variant * 0.025),
    stripeColumn: mixHex(accent, '#ffffff', 0.82),
    text: '#1f2937',
    total: mixHex(accent, '#ffffff', 0.78),
    totalText: '#172033',
  };
}

function mediumPalette(number: number): SpreadsheetTableStylePalette {
  const accent = ACCENTS[(number - 1) % ACCENTS.length] ?? ACCENTS[0];
  const variant = Math.floor((number - 1) / ACCENTS.length);
  const header = mixHex(accent, '#111827', variant * 0.08);
  return {
    border: mixHex(accent, '#ffffff', 0.34),
    header,
    headerText: '#ffffff',
    primaryRow: '#ffffff',
    secondaryRow: mixHex(accent, '#ffffff', 0.86 - variant * 0.02),
    stripeColumn: mixHex(accent, '#ffffff', 0.76),
    text: '#172033',
    total: mixHex(accent, '#ffffff', 0.72),
    totalText: mixHex(accent, '#111827', 0.48),
  };
}

function darkPalette(number: number): SpreadsheetTableStylePalette {
  const accent = DARK_ACCENTS[number - 1] ?? DARK_ACCENTS[0];
  return {
    border: mixHex(accent, '#ffffff', 0.32),
    header: accent,
    headerText: '#ffffff',
    primaryRow: mixHex(accent, '#ffffff', 0.9),
    secondaryRow: mixHex(accent, '#ffffff', 0.78),
    stripeColumn: mixHex(accent, '#ffffff', 0.68),
    text: '#172033',
    total: mixHex(accent, '#ffffff', 0.62),
    totalText: mixHex(accent, '#111827', 0.34),
  };
}

function mixHex(left: string, right: string, rightWeight: number): string {
  const weight = Math.min(1, Math.max(0, rightWeight));
  const leftRgb = parseHex(left);
  const rightRgb = parseHex(right);
  return `#${leftRgb
    .map((value, index) =>
      Math.round(value * (1 - weight) + (rightRgb[index] ?? value) * weight)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

function parseHex(value: string): [number, number, number] {
  const hex = value.replace(/^#/, '');
  return [0, 2, 4].map((offset) =>
    Number.parseInt(hex.slice(offset, offset + 2), 16),
  ) as [number, number, number];
}

function familyLabel(family: 'dark' | 'light' | 'medium'): string {
  if (family === 'light') return '浅色';
  if (family === 'medium') return '中等';
  return '深色';
}

function capitalize(value: string): string {
  return `${value.charAt(0).toUpperCase()}${value.slice(1)}`;
}
