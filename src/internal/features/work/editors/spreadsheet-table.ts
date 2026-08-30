import type { Cell, Selection } from '@fortune-sheet/core';
import { sparseArrayEntries, sparseArrayIndexes } from '../spreadsheet-sparse';
import {
  formatSpreadsheetCellRanges,
  isValidSpreadsheetDefinedName,
  parseSpreadsheetCellRanges,
} from '../work-spreadsheet-ranges';
import { createWorkId } from '../work-templates';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
  WorkSpreadsheetTableStyle,
} from '../work-types';
import { spreadsheetSelectionOrCurrentRegion } from './spreadsheet-auto-filter';
import { canMutateSpreadsheetCellRanges } from './spreadsheet-cell-mutation-guard';
import {
  normalizeSpreadsheetCellRange,
  spreadsheetCellRangeArea,
  spreadsheetCellRangeContains,
  spreadsheetCellRangesIntersect,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';
import { spreadsheetSheetBounds } from './spreadsheet-keyboard-navigation';
import { materializeSpreadsheetTableAppearance } from './spreadsheet-table-conversion';
import { MAX_SPREADSHEET_TABLE_CELLS } from './spreadsheet-table-limits';
import { reconcileSpreadsheetTableCalculatedColumns } from './spreadsheet-table-calculated-columns';

export { MAX_SPREADSHEET_TABLE_CELLS } from './spreadsheet-table-limits';

export interface SpreadsheetTableTarget {
  sheetId: string;
  selection: Selection;
}

export interface SpreadsheetTableDialogValue {
  headerRow: boolean;
  rangeReference: string;
}

export interface SpreadsheetTableDialogSource {
  name: string;
  range: SpreadsheetCellRange;
  rangeReference: string;
  sheetId: string;
  sheetName: string;
  value: SpreadsheetTableDialogValue;
}

export interface SpreadsheetTableRequest {
  headerRow: boolean;
  name: string;
  range: SpreadsheetCellRange;
  sheetId: string;
  style?: WorkSpreadsheetTableStyle;
}

export type SpreadsheetTableValidation =
  | {
      columns: string[];
      name: string;
      ok: true;
      range: SpreadsheetCellRange;
      sheet: WorkSpreadsheetSheet;
    }
  | {
      code: SpreadsheetTableErrorCode;
      message: string;
      ok: false;
    };

export type SpreadsheetTableErrorCode =
  | 'auto-filter-overlap'
  | 'invalid-column-name'
  | 'invalid-name'
  | 'invalid-range'
  | 'invalid-style'
  | 'merged-range'
  | 'name-conflict'
  | 'out-of-bounds'
  | 'protected-range'
  | 'pivot-table'
  | 'range-too-large'
  | 'sheet-not-found'
  | 'table-overlap';

export interface SpreadsheetTableDesignPatch {
  name?: string;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  style?: WorkSpreadsheetTableStyle;
}

export function createSpreadsheetTableDialogSource(
  content: WorkSpreadsheetContent,
  target: SpreadsheetTableTarget,
): SpreadsheetTableDialogSource | null {
  const sheet = content.sheets.find(
    (candidate) => candidate.id === target.sheetId,
  );
  if (!sheet) return null;
  const range = spreadsheetSelectionOrCurrentRegion(sheet, target.selection);
  if (!range) return null;
  let request: SpreadsheetTableRequest = {
    headerRow: true,
    name: nextSpreadsheetTableName(content),
    range,
    sheetId: target.sheetId,
  };
  if (!validateSpreadsheetTableRequest(content, request).ok) {
    request = { ...request, headerRow: false };
    if (!validateSpreadsheetTableRequest(content, request).ok) return null;
  }
  const rangeReference = formatSpreadsheetCellRanges([range]);
  return {
    name: request.name,
    range,
    rangeReference,
    sheetId: target.sheetId,
    sheetName: sheet.name,
    value: { headerRow: request.headerRow, rangeReference },
  };
}

export function spreadsheetTableRangeFromText(
  value: string,
): SpreadsheetCellRange | null {
  const ranges = parseSpreadsheetCellRanges(value);
  return ranges?.length === 1 ? (ranges[0] ?? null) : null;
}

export function validateSpreadsheetTableRequest(
  content: WorkSpreadsheetContent,
  request: SpreadsheetTableRequest,
): SpreadsheetTableValidation {
  const sheet = content.sheets.find(
    (candidate) => candidate.id === request.sheetId,
  );
  if (!sheet) return tableError('sheet-not-found');
  if (sheet.isPivotTable || sheet.pivotTables?.length) {
    return tableError('pivot-table');
  }
  const range = normalizeSpreadsheetCellRange(request.range);
  if (!range) return tableError('invalid-range');
  const name = request.name.trim();
  if (!validSpreadsheetTableName(name)) return tableError('invalid-name');
  if (spreadsheetTableNameExists(content, name)) {
    return tableError('name-conflict');
  }
  const bounds = spreadsheetSheetBounds(sheet);
  if (range.row[1] > bounds.lastRow || range.column[1] > bounds.lastColumn) {
    return tableError('out-of-bounds');
  }
  const height = range.row[1] - range.row[0] + 1;
  if (height <= Number(request.headerRow)) return tableError('invalid-range');
  if (spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_TABLE_CELLS) {
    return tableError('range-too-large');
  }
  if (spreadsheetTableIntersectsMerge(sheet, range)) {
    return tableError('merged-range');
  }
  if (
    (sheet.tables ?? []).some((table) =>
      spreadsheetCellRangesIntersect(table.range, range),
    )
  ) {
    return tableError('table-overlap');
  }
  const autoFilter = normalizeSpreadsheetCellRange(
    sheet.filter_select ?? { row: [], column: [] },
  );
  if (autoFilter && spreadsheetCellRangesIntersect(autoFilter, range)) {
    return tableError('auto-filter-overlap');
  }
  if (!canMutateSpreadsheetCellRanges(sheet, [range])) {
    return tableError('protected-range');
  }
  const style = request.style ?? { family: 'medium', number: 2 };
  if (!validSpreadsheetTableStyle(style)) return tableError('invalid-style');
  const columns = spreadsheetTableColumnNames(sheet, range, request.headerRow);
  if (
    !columns.length ||
    columns.some((column) => !validTableColumnName(column))
  ) {
    return tableError('invalid-column-name');
  }
  return { columns, name, ok: true, range, sheet };
}

export function applySpreadsheetTable(
  content: WorkSpreadsheetContent,
  request: SpreadsheetTableRequest,
): WorkSpreadsheetContent | null {
  const validation = validateSpreadsheetTableRequest(content, request);
  if (!validation.ok) return null;
  const style = request.style ?? { family: 'medium', number: 2 };
  const tableWithoutCalculatedColumns: WorkSpreadsheetTable = {
    id: createWorkId('spreadsheet-table'),
    name: validation.name,
    range: cloneRange(validation.range),
    columns: validation.columns.map((name) => ({ name })),
    filters: [],
    headerRow: request.headerRow,
    totalsRow: false,
    style,
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: style.family !== 'none',
    showColumnStripes: false,
  };
  const table: WorkSpreadsheetTable = {
    ...tableWithoutCalculatedColumns,
    columns: reconcileSpreadsheetTableCalculatedColumns(
      validation.sheet,
      tableWithoutCalculatedColumns,
    ),
  };
  const nextSheet = request.headerRow
    ? stampSpreadsheetTableHeaders(
        validation.sheet,
        validation.range,
        validation.columns,
      )
    : validation.sheet;
  return replaceSpreadsheetTableSheet(content, {
    ...nextSheet,
    tables: [...(nextSheet.tables ?? []), table],
  });
}

export function updateSpreadsheetTable(
  content: WorkSpreadsheetContent,
  sheetId: string,
  tableId: string,
  patch: SpreadsheetTableDesignPatch,
): WorkSpreadsheetContent | null {
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  const table = sheet?.tables?.find((candidate) => candidate.id === tableId);
  if (!sheet || !table) return null;
  const name = patch.name?.trim() ?? table.name;
  if (
    !validSpreadsheetTableName(name) ||
    spreadsheetTableNameExists(content, name, tableId)
  ) {
    return null;
  }
  const style = patch.style ?? table.style;
  if (!validSpreadsheetTableStyle(style)) return null;
  const showFirstColumn = patch.showFirstColumn ?? table.showFirstColumn;
  const showLastColumn = patch.showLastColumn ?? table.showLastColumn;
  const showRowStripes = patch.showRowStripes ?? table.showRowStripes;
  const showColumnStripes = patch.showColumnStripes ?? table.showColumnStripes;
  if (
    style.family === 'none' &&
    (showFirstColumn || showLastColumn || showRowStripes || showColumnStripes)
  ) {
    return null;
  }
  const tables = sheet.tables?.map((candidate) =>
    candidate.id === tableId
      ? {
          ...candidate,
          name,
          ...(patch.name === undefined ? {} : { displayName: undefined }),
          style,
          showFirstColumn,
          showLastColumn,
          showRowStripes,
          showColumnStripes,
        }
      : candidate,
  );
  return replaceSpreadsheetTableSheet(content, { ...sheet, tables });
}

export function convertSpreadsheetTableToRange(
  content: WorkSpreadsheetContent,
  sheetId: string,
  tableId: string,
): WorkSpreadsheetContent | null {
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  const table = sheet?.tables?.find((candidate) => candidate.id === tableId);
  if (
    !sheet ||
    !table ||
    spreadsheetTableHasStructuredReferences(content, table)
  ) {
    return null;
  }
  const remaining = sheet.tables?.filter(
    (candidate) => candidate.id !== tableId,
  );
  const withoutTable = replaceSpreadsheetTableSheet(content, {
    ...sheet,
    tables: remaining?.length ? remaining : undefined,
  });
  return materializeSpreadsheetTableAppearance(withoutTable, sheetId, table);
}

export function spreadsheetTableAtCell(
  sheet: WorkSpreadsheetSheet | undefined,
  row: number,
  column: number,
): WorkSpreadsheetTable | null {
  if (!sheet) return null;
  return (
    (sheet.tables ?? []).find((table) =>
      spreadsheetCellRangeContains(table.range, row, column),
    ) ?? null
  );
}

export function spreadsheetTableFailureMessage(
  content: WorkSpreadsheetContent,
  request: SpreadsheetTableRequest,
): string | null {
  const validation = validateSpreadsheetTableRequest(content, request);
  return validation.ok ? null : validation.message;
}

function nextSpreadsheetTableName(content: WorkSpreadsheetContent): string {
  for (let index = 1; index <= 65_536; index += 1) {
    const candidate = `Table${index}`;
    if (!spreadsheetTableNameExists(content, candidate)) return candidate;
  }
  return 'Table65536';
}

function spreadsheetTableNameExists(
  content: WorkSpreadsheetContent,
  name: string,
  exceptTableId?: string,
): boolean {
  const normalized = name.toLocaleLowerCase();
  return (
    content.sheets.some((sheet) =>
      (sheet.tables ?? []).some(
        (table) =>
          table.id !== exceptTableId &&
          [table.name, table.displayName].some(
            (candidate) => candidate?.toLocaleLowerCase() === normalized,
          ),
      ),
    ) ||
    (content.namedRanges ?? []).some(
      (range) => range.name.trim().toLocaleLowerCase() === normalized,
    )
  );
}

function validSpreadsheetTableName(name: string): boolean {
  return (
    isValidSpreadsheetDefinedName(name) &&
    !name.startsWith('_xlnm.') &&
    !['r', 'c'].includes(name.toLocaleLowerCase())
  );
}

function validSpreadsheetTableStyle(style: WorkSpreadsheetTableStyle): boolean {
  if (style.family === 'none') return true;
  if (!Number.isInteger(style.number)) return false;
  if (style.family === 'light') return style.number >= 1 && style.number <= 21;
  if (style.family === 'medium') return style.number >= 1 && style.number <= 28;
  return style.number >= 1 && style.number <= 11;
}

function spreadsheetTableColumnNames(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  headerRow: boolean,
): string[] {
  const observed = new Set<string>();
  const names: string[] = [];
  for (let column = range.column[0]; column <= range.column[1]; column += 1) {
    const raw = headerRow
      ? spreadsheetCellText(sheet, range.row[0], column).trim()
      : '';
    const base = validTableColumnName(raw)
      ? raw
      : `Column${column - range.column[0] + 1}`;
    let candidate = base;
    let suffix = 2;
    while (observed.has(candidate.toLocaleLowerCase())) {
      const suffixText = String(suffix);
      candidate = `${truncateTableColumnName(
        base,
        255 - suffixText.length,
      )}${suffixText}`;
      suffix += 1;
    }
    observed.add(candidate.toLocaleLowerCase());
    names.push(candidate);
  }
  return names;
}

function validTableColumnName(name: string): boolean {
  const characters = Array.from(name);
  return (
    name.trim() === name &&
    characters.length >= 1 &&
    characters.length <= 255 &&
    !characters.some(
      (character) =>
        /\p{Cc}/u.test(character) ||
        character === '\uFFFE' ||
        character === '\uFFFF',
    )
  );
}

function truncateTableColumnName(value: string, maximum: number): string {
  return Array.from(value).slice(0, maximum).join('');
}

function spreadsheetCellText(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): string {
  const cell =
    sheet.data?.[row]?.[column] ??
    sheet.celldata?.find(
      (candidate) => candidate.r === row && candidate.c === column,
    )?.v;
  const value = cell?.m ?? cell?.v;
  return value === undefined || value === null ? '' : String(value);
}

function stampSpreadsheetTableHeaders(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  names: readonly string[],
): WorkSpreadsheetSheet {
  if (sheet.data !== undefined) {
    const data = cloneSparseMatrixShell(sheet.data);
    const sourceRow = sheet.data[range.row[0]];
    const row = cloneSparseRow(sourceRow);
    for (const [offset, name] of names.entries()) {
      const column = range.column[0] + offset;
      row[column] = stampHeaderCell(sourceRow?.[column], name);
    }
    data[range.row[0]] = row;
    return { ...sheet, data };
  }
  const byCoordinate = new Map(
    (sheet.celldata ?? []).map((entry) => [`${entry.r}_${entry.c}`, entry]),
  );
  for (const [offset, name] of names.entries()) {
    const column = range.column[0] + offset;
    const key = `${range.row[0]}_${column}`;
    const source = byCoordinate.get(key);
    byCoordinate.set(key, {
      r: range.row[0],
      c: column,
      v: stampHeaderCell(source?.v, name),
    });
  }
  return { ...sheet, celldata: [...byCoordinate.values()] };
}

function cloneSparseMatrixShell(
  source: NonNullable<WorkSpreadsheetSheet['data']>,
): NonNullable<WorkSpreadsheetSheet['data']> {
  const data: NonNullable<WorkSpreadsheetSheet['data']> = [];
  data.length = source.length;
  for (const [row, value] of sparseArrayEntries(source)) data[row] = value;
  return data;
}

function cloneSparseRow(
  source: NonNullable<WorkSpreadsheetSheet['data']>[number] | undefined,
): NonNullable<WorkSpreadsheetSheet['data']>[number] {
  const row: NonNullable<WorkSpreadsheetSheet['data']>[number] = [];
  if (!source) return row;
  row.length = source.length;
  for (const column of sparseArrayIndexes(source)) row[column] = source[column];
  return row;
}

function stampHeaderCell(source: Cell | null | undefined, name: string): Cell {
  return { ...(source ?? {}), m: name, v: name };
}

function spreadsheetTableIntersectsMerge(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): boolean {
  return Object.values(sheet.config?.merge ?? {}).some((merge) =>
    spreadsheetCellRangesIntersect(range, {
      row: [merge.r, merge.r + Math.max(1, merge.rs) - 1],
      column: [merge.c, merge.c + Math.max(1, merge.cs) - 1],
    }),
  );
}

function replaceSpreadsheetTableSheet(
  content: WorkSpreadsheetContent,
  nextSheet: WorkSpreadsheetSheet,
): WorkSpreadsheetContent {
  return {
    ...content,
    sheets: content.sheets.map((sheet) =>
      sheet.id === nextSheet.id ? nextSheet : sheet,
    ),
  };
}

function spreadsheetTableHasStructuredReferences(
  content: WorkSpreadsheetContent,
  table: WorkSpreadsheetTable,
): boolean {
  const aliases = [table.name, table.displayName].filter(
    (value): value is string => Boolean(value),
  );
  const patterns = aliases.map(
    (alias) => new RegExp(`${escapeRegExp(alias)}\\s*\\[`, 'i'),
  );
  const matches = (value: unknown) =>
    typeof value === 'string' &&
    patterns.some((pattern) => pattern.test(value));
  for (const sheet of content.sheets) {
    for (const [, row] of sparseArrayEntries(sheet.data)) {
      for (const [, cell] of sparseArrayEntries(row)) {
        if (matches(cell?.f)) return true;
      }
    }
    for (const entry of sheet.celldata ?? []) {
      if (matches(entry.v?.f)) return true;
    }
    if (
      Object.values(sheet.formulaMetadata?.sourceFormulas ?? {}).some(
        matches,
      ) ||
      matches(JSON.stringify(sheet.dataValidationRanges ?? [])) ||
      matches(JSON.stringify(sheet.luckysheet_conditionformat_save ?? [])) ||
      matches(JSON.stringify(sheet.charts ?? []))
    ) {
      return true;
    }
  }
  return (content.namedRanges ?? []).some((range) => matches(range.reference));
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function cloneRange(range: SpreadsheetCellRange): SpreadsheetCellRange {
  return { row: [...range.row], column: [...range.column] };
}

function tableError(
  code: SpreadsheetTableErrorCode,
): Extract<SpreadsheetTableValidation, { ok: false }> {
  const messages: Record<SpreadsheetTableErrorCode, string> = {
    'auto-filter-overlap': '表格区域不能与工作表自动筛选区域重叠。',
    'invalid-column-name': '表格列名必须唯一且不超过 255 个字符。',
    'invalid-name': '表格名称必须是有效且不类似单元格引用的标识符。',
    'invalid-range': '表格至少需要一行数据。',
    'invalid-style': '表格样式身份无效。',
    'merged-range': '请先取消表格区域内的合并单元格。',
    'name-conflict': '表格名称已被其他表格或名称使用。',
    'out-of-bounds': '表格区域超出当前工作表边界。',
    'pivot-table': '数据透视表工作表不能创建普通表格。',
    'protected-range': '受保护或只读区域不能创建表格。',
    'range-too-large': `一次最多可创建 ${MAX_SPREADSHEET_TABLE_CELLS.toLocaleString()} 个单元格的表格。`,
    'sheet-not-found': '找不到目标工作表。',
    'table-overlap': '表格区域不能与其他表格重叠。',
  };
  return { code, message: messages[code], ok: false };
}
