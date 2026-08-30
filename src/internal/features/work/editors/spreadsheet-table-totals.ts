import type { Cell } from '@fortune-sheet/core';
import {
  editableSpreadsheetFormula,
  formulaHasExternalReference,
  spreadsheetFormulaFunctions,
} from '../work-spreadsheet-formulas';
import type {
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
  WorkSpreadsheetTableColumn,
  WorkSpreadsheetTableTotalsFunction,
} from '../work-types';

export const MAX_SPREADSHEET_TABLE_TOTALS_FORMULA_LENGTH = 8_192;
export const MAX_SPREADSHEET_TABLE_TOTALS_LABEL_LENGTH = 255;
export const DEFAULT_SPREADSHEET_TABLE_TOTALS_LABEL = 'Total';

/** The closed function set exposed by the Table Design totals menu. */
export const SPREADSHEET_TABLE_TOTALS_FUNCTIONS = Object.freeze([
  'sum',
  'average',
  'count',
  'countNums',
  'max',
  'min',
  'stdDev',
  'stdDevP',
  'var',
  'varP',
  'custom',
] as const satisfies readonly WorkSpreadsheetTableTotalsFunction[]);

export interface SpreadsheetTableTotalsColumnPatch {
  totalsFormula?: string | null;
  totalsFunction?: WorkSpreadsheetTableTotalsFunction | null;
  totalsLabel?: string | null;
}

const UNSAFE_TOTALS_FUNCTIONS = new Set([
  'CALL',
  'DDE',
  'DDE.REQUEST',
  'DDE.POKE',
  'EXEC',
  'HYPERLINK',
  'INDIRECT',
  'OFFSET',
  'REGISTER.ID',
  'RTD',
  'WEBSERVICE',
]);

const SUBTOTAL_CODES: Readonly<
  Record<Exclude<WorkSpreadsheetTableTotalsFunction, 'custom'>, number>
> = {
  average: 101,
  count: 103,
  countNums: 102,
  max: 104,
  min: 105,
  stdDev: 107,
  stdDevP: 108,
  sum: 109,
  var: 110,
  varP: 111,
};

const OOXML_TOTALS_FUNCTIONS: Readonly<
  Record<Exclude<WorkSpreadsheetTableTotalsFunction, 'custom'>, string>
> = {
  average: 'average',
  count: 'count',
  countNums: 'countNums',
  max: 'max',
  min: 'min',
  stdDev: 'stdDev',
  stdDevP: 'stdDevp',
  sum: 'sum',
  var: 'var',
  varP: 'varp',
};

const TOTALS_FUNCTION_LABELS: Readonly<
  Record<WorkSpreadsheetTableTotalsFunction, string>
> = {
  sum: '求和',
  average: '平均值',
  count: '计数',
  countNums: '数值计数',
  max: '最大值',
  min: '最小值',
  stdDev: '标准差',
  stdDevP: '总体标准差',
  var: '方差',
  varP: '总体方差',
  custom: '自定义',
};

export function spreadsheetTableTotalsFunctionLabel(
  value: WorkSpreadsheetTableTotalsFunction | undefined,
): string {
  return value ? TOTALS_FUNCTION_LABELS[value] : '不汇总';
}

export function normalizeSpreadsheetTableTotalsFunction(
  value: unknown,
): WorkSpreadsheetTableTotalsFunction | undefined {
  if (typeof value !== 'string') return undefined;
  if (value === 'none' || value === '') return undefined;
  return (SPREADSHEET_TABLE_TOTALS_FUNCTIONS as readonly string[]).includes(
    value,
  )
    ? (value as WorkSpreadsheetTableTotalsFunction)
    : undefined;
}

/** Normalize a table totals formula to the editable leading-equals form. */
export function normalizeSpreadsheetTableTotalsFormula(
  value: unknown,
): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  const formula = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
  if (
    formula.startsWith('==') ||
    formula.length > MAX_SPREADSHEET_TABLE_TOTALS_FORMULA_LENGTH ||
    /[\u0000-\u001f\u007f]/u.test(formula) ||
    formulaHasExternalReference(formula) ||
    spreadsheetFormulaFunctions(formula).some((name) =>
      UNSAFE_TOTALS_FUNCTIONS.has(name),
    )
  ) {
    return undefined;
  }
  const editable = editableSpreadsheetFormula(formula).trim();
  return editable || undefined;
}

export function normalizeSpreadsheetTableTotalsLabel(
  value: unknown,
): string | undefined {
  if (typeof value !== 'string') return undefined;
  const label = value.trim();
  if (
    !label ||
    Array.from(label).length > MAX_SPREADSHEET_TABLE_TOTALS_LABEL_LENGTH ||
    /[\u0000-\u001f\u007f]/u.test(label)
  ) {
    return undefined;
  }
  return label;
}

export function spreadsheetTableTotalsFunctionFromOoxml(
  value: unknown,
): WorkSpreadsheetTableTotalsFunction | undefined {
  if (typeof value !== 'string') return undefined;
  const normalized = value.trim().toLowerCase();
  if (normalized === 'custom') return 'custom';
  const entry = Object.entries(OOXML_TOTALS_FUNCTIONS).find(
    ([, token]) => token.toLowerCase() === normalized,
  );
  return entry?.[0] as WorkSpreadsheetTableTotalsFunction | undefined;
}

export function spreadsheetTableTotalsFunctionToOoxml(
  value: WorkSpreadsheetTableTotalsFunction | undefined,
): string | undefined {
  if (!value || value === 'custom')
    return value === 'custom' ? 'custom' : undefined;
  return OOXML_TOTALS_FUNCTIONS[value];
}

/**
 * Returns the formula that should be placed in one totals-row cell. Native
 * aggregate functions use SUBTOTAL so filtered-out rows stay excluded; a
 * custom rule uses its explicit formula.
 */
export function spreadsheetTableTotalsFormula(
  table: WorkSpreadsheetTable,
  columnOffset: number,
): string | undefined {
  const column = table.columns[columnOffset];
  if (!column) return undefined;
  const custom = normalizeSpreadsheetTableTotalsFormula(column.totalsFormula);
  if (custom) return custom;
  const functionName = normalizeSpreadsheetTableTotalsFunction(
    column.totalsFunction,
  );
  if (!functionName || functionName === 'custom') return undefined;
  const code = SUBTOTAL_CODES[functionName];
  if (!code) return undefined;
  return `=SUBTOTAL(${code},${table.name}[${escapeStructuredColumnName(column.name)}])`;
}

export function spreadsheetTableTotalsColumnPatch(
  columns: readonly WorkSpreadsheetTableColumn[],
  patches:
    | Readonly<Record<string, SpreadsheetTableTotalsColumnPatch>>
    | readonly (SpreadsheetTableTotalsColumnPatch | null | undefined)[]
    | undefined,
): WorkSpreadsheetTableColumn[] | null {
  if (patches === undefined) return columns.map((column) => ({ ...column }));
  const entries = Array.isArray(patches)
    ? patches.flatMap((patch, offset) =>
        patch ? [[String(offset), patch] as const] : [],
      )
    : Object.entries(patches);
  const next = columns.map((column) => ({ ...column }));
  for (const [rawOffset, patch] of entries) {
    const offset = Number(rawOffset);
    if (
      !Number.isSafeInteger(offset) ||
      offset < 0 ||
      offset >= next.length ||
      !patch ||
      typeof patch !== 'object' ||
      Array.isArray(patch)
    ) {
      return null;
    }
    const current = next[offset];
    if (!current) return null;
    const candidate: WorkSpreadsheetTableColumn = { ...current };
    if (Object.hasOwn(patch, 'totalsFunction')) {
      if (patch.totalsFunction === null || patch.totalsFunction === '') {
        delete candidate.totalsFunction;
      } else {
        const functionName = normalizeSpreadsheetTableTotalsFunction(
          patch.totalsFunction,
        );
        if (!functionName) return null;
        candidate.totalsFunction = functionName;
      }
    }
    if (Object.hasOwn(patch, 'totalsLabel')) {
      if (patch.totalsLabel === null || patch.totalsLabel === '') {
        delete candidate.totalsLabel;
      } else {
        const label = normalizeSpreadsheetTableTotalsLabel(patch.totalsLabel);
        if (!label) return null;
        candidate.totalsLabel = label;
      }
    }
    if (Object.hasOwn(patch, 'totalsFormula')) {
      if (patch.totalsFormula === null || patch.totalsFormula === '') {
        delete candidate.totalsFormula;
      } else {
        const formula = normalizeSpreadsheetTableTotalsFormula(
          patch.totalsFormula,
        );
        if (!formula) return null;
        candidate.totalsFormula = formula;
      }
    }
    if (candidate.totalsFormula) {
      // A formula is the source of truth for a custom totals rule. Native
      // function metadata cannot describe the same cell and is therefore
      // rejected instead of silently changing the user's formula.
      if (
        candidate.totalsFunction !== undefined &&
        candidate.totalsFunction !== 'custom'
      ) {
        return null;
      }
      candidate.totalsFunction = 'custom';
    } else if (candidate.totalsFunction === 'custom') {
      // `custom` without a formula is not a usable rule.
      return null;
    }
    if (
      candidate.totalsFunction !== undefined &&
      candidate.totalsFunction !== 'custom'
    ) {
      delete candidate.totalsFormula;
    }
    if (candidate.totalsFunction !== undefined && candidate.totalsLabel) {
      // A totals label occupies the same cell as a totals formula. Keep the
      // explicit function/formula and discard the ambiguous label.
      delete candidate.totalsLabel;
    }
    next[offset] = candidate;
  }
  return next;
}

/** Apply the native default label used when a totals row is first enabled. */
export function withDefaultSpreadsheetTableTotalsLabel(
  columns: readonly WorkSpreadsheetTableColumn[],
  totalsRow: boolean,
): WorkSpreadsheetTableColumn[] {
  const next = columns.map((column) => ({ ...column }));
  if (!totalsRow || !next.length) return next;
  const hasExplicitLabel = next.some((column) => column.totalsLabel);
  if (hasExplicitLabel) return next;
  const first = next[0];
  if (first && !first.totalsFunction && !first.totalsFormula) {
    first.totalsLabel = DEFAULT_SPREADSHEET_TABLE_TOTALS_LABEL;
  }
  return next;
}

/** True when a proposed totals row would overwrite a populated cell. */
export function spreadsheetTableTotalsRowHasContent(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  row = table.totalsRow ? table.range.row[1] : table.range.row[1] + 1,
): boolean {
  for (
    let column = table.range.column[0];
    column <= table.range.column[1];
    column += 1
  ) {
    if (
      !spreadsheetTableTotalsCellIsEmpty(
        spreadsheetTableTotalsCellAt(sheet, row, column),
      )
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Synchronize totals-row cells after a table metadata change. Explicitly
 * changed column metadata owns the target cell; untouched manual cells remain
 * authoritative. The helper works with Fortune's dense and sparse forms.
 */
export function synchronizeSpreadsheetTableTotalsRow(
  sheet: WorkSpreadsheetSheet,
  previousTable: WorkSpreadsheetTable | undefined,
  nextTable: WorkSpreadsheetTable,
  options: { force?: boolean } = {},
): WorkSpreadsheetSheet {
  if (!nextTable.totalsRow) return sheet;
  const force = options.force === true;
  const previousRow = previousTable?.range.row[1];
  const nextRow = nextTable.range.row[1];
  const width = nextTable.range.column[1] - nextTable.range.column[0] + 1;
  const updates = new Map<number, Cell | null>();
  for (let offset = 0; offset < width; offset += 1) {
    const column = nextTable.range.column[0] + offset;
    const nextColumn = nextTable.columns[offset];
    if (!nextColumn) continue;
    const previousColumn = previousTable?.columns[offset];
    const oldFormula = previousTable
      ? spreadsheetTableTotalsFormula(previousTable, offset)
      : undefined;
    const newFormula = spreadsheetTableTotalsFormula(nextTable, offset);
    const oldLabel = normalizeSpreadsheetTableTotalsLabel(
      previousColumn?.totalsLabel,
    );
    const newLabel = normalizeSpreadsheetTableTotalsLabel(
      nextColumn.totalsLabel,
    );
    const changed =
      force ||
      previousTable === undefined ||
      previousTable.name !== nextTable.name ||
      previousTable.range.row[1] !== nextTable.range.row[1] ||
      previousColumn?.totalsFunction !== nextColumn.totalsFunction ||
      previousColumn?.totalsFormula !== nextColumn.totalsFormula ||
      previousColumn?.totalsLabel !== nextColumn.totalsLabel ||
      previousColumn?.name !== nextColumn.name;
    const current = spreadsheetTableTotalsCellAt(sheet, nextRow, column);
    const previousCell =
      previousRow === undefined
        ? null
        : spreadsheetTableTotalsCellAt(sheet, previousRow, column);
    const styleSource = current ?? previousCell;
    // Formula metadata has precedence over a label. A manually edited cell
    // remains authoritative even when the table name or header changes; only
    // an empty cell or the previous generated value is replaceable.
    const empty = spreadsheetTableTotalsCellIsEmpty(current);
    const currentFormula = normalizeSpreadsheetTableTotalsFormula(current?.f);
    const generatedFormula = Boolean(
      oldFormula && currentFormula && formulasEqual(currentFormula, oldFormula),
    );
    const generatedLabel = Boolean(
      oldLabel && !current?.f && spreadsheetTableCellText(current) === oldLabel,
    );
    const replaceable =
      force || empty || (changed && (generatedFormula || generatedLabel));
    if (newFormula && replaceable) {
      updates.set(
        column,
        spreadsheetCellWithTotalsFormula(styleSource, newFormula),
      );
    } else if (newLabel && replaceable) {
      updates.set(
        column,
        spreadsheetCellWithTotalsLabel(styleSource, newLabel),
      );
    } else if (
      !newFormula &&
      !newLabel &&
      changed &&
      (generatedFormula || generatedLabel)
    ) {
      updates.set(column, clearSpreadsheetTableTotalsCell(current));
    }
  }
  return applySpreadsheetTableTotalsCellUpdates(sheet, nextRow, updates);
}

/** Reconcile metadata after a user edits a totals-row cell directly. */
export function reconcileSpreadsheetTableTotalsColumns(
  sheet: WorkSpreadsheetSheet,
  table: WorkSpreadsheetTable,
  editedOffsets?: ReadonlySet<number>,
): WorkSpreadsheetTableColumn[] {
  return table.columns.map((column, offset) => {
    const candidate: WorkSpreadsheetTableColumn = { ...column };
    const declaredFunction = normalizeSpreadsheetTableTotalsFunction(
      candidate.totalsFunction,
    );
    const declaredFormula = normalizeSpreadsheetTableTotalsFormula(
      candidate.totalsFormula,
    );
    if (declaredFunction) candidate.totalsFunction = declaredFunction;
    else delete candidate.totalsFunction;
    if (declaredFormula) candidate.totalsFormula = declaredFormula;
    else delete candidate.totalsFormula;
    const label = normalizeSpreadsheetTableTotalsLabel(candidate.totalsLabel);
    if (label) candidate.totalsLabel = label;
    else delete candidate.totalsLabel;
    if (!table.totalsRow) return candidate;
    const cell = spreadsheetTableTotalsCellAt(
      sheet,
      table.range.row[1],
      table.range.column[0] + offset,
    );
    // Without an explicit edit marker, metadata remains authoritative. This
    // is important when an imported table has a totals definition but the
    // worksheet XML omits cached totals cells. Directly edited cells are
    // reconciled below and may intentionally become manual/custom.
    if (editedOffsets && !editedOffsets.has(offset)) return candidate;
    const cellFormula = normalizeSpreadsheetTableTotalsFormula(cell?.f);
    const expected = spreadsheetTableTotalsFormula(
      {
        ...table,
        columns: [{ ...candidate }],
      },
      0,
    );
    if (cellFormula) {
      if (expected && formulasEqual(cellFormula, expected)) {
        return candidate;
      }
      delete candidate.totalsFunction;
      delete candidate.totalsFormula;
      candidate.totalsFunction = 'custom';
      candidate.totalsFormula = cellFormula;
      return candidate;
    }
    // A direct value edit replaces any previous formula/function ownership.
    // Empty cells are also treated as an explicit removal when this offset is
    // marked as edited.
    delete candidate.totalsFunction;
    delete candidate.totalsFormula;
    const cellLabel = normalizeSpreadsheetTableTotalsLabel(
      spreadsheetTableCellText(cell),
    );
    if (cellLabel) {
      if (!candidate.totalsLabel || candidate.totalsLabel === cellLabel) {
        candidate.totalsLabel = cellLabel;
      } else {
        delete candidate.totalsLabel;
      }
    } else {
      delete candidate.totalsLabel;
    }
    return candidate;
  });
}

function escapeStructuredColumnName(value: string): string {
  return value.replaceAll(']', ']]');
}

function formulasEqual(left: string, right: string): boolean {
  return (
    left.trim().replace(/\s+/g, '').toLocaleLowerCase() ===
    right.trim().replace(/\s+/g, '').toLocaleLowerCase()
  );
}

function spreadsheetTableTotalsCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  return (
    sheet.data?.[row]?.[column] ??
    sheet.celldata?.find((entry) => entry.r === row && entry.c === column)?.v ??
    null
  );
}

function spreadsheetTableTotalsCellIsEmpty(
  cell: Cell | null | undefined,
): boolean {
  return (
    !cell?.f &&
    (cell?.v === undefined || cell.v === null || cell.v === '') &&
    (cell?.m === undefined || cell.m === null || cell.m === '')
  );
}

function spreadsheetTableCellText(cell: Cell | null | undefined): string {
  const value = cell?.m ?? cell?.v;
  return value === undefined || value === null ? '' : String(value);
}

function spreadsheetCellWithTotalsFormula(
  source: Cell | null | undefined,
  formula: string,
): Cell {
  const { f: _formula, m: _display, v: _value, ...presentation } = source ?? {};
  return { ...presentation, f: formula };
}

function spreadsheetCellWithTotalsLabel(
  source: Cell | null | undefined,
  label: string,
): Cell {
  const { f: _formula, ...withoutFormula } = source ?? {};
  return { ...withoutFormula, m: label, v: label };
}

function clearSpreadsheetTableTotalsCell(
  source: Cell | null | undefined,
): Cell | null {
  if (!source) return null;
  const { f: _formula, m: _display, v: _value, ...presentation } = source;
  return Object.keys(presentation).length ? presentation : null;
}

function applySpreadsheetTableTotalsCellUpdates(
  sheet: WorkSpreadsheetSheet,
  rowIndex: number,
  updates: ReadonlyMap<number, Cell | null>,
): WorkSpreadsheetSheet {
  if (!updates.size) return sheet;
  if (sheet.data !== undefined) {
    const data = sheet.data.slice();
    const source = data[rowIndex];
    const row = source ? source.slice() : [];
    for (const [column, cell] of updates) row[column] = cell;
    data[rowIndex] = row;
    return { ...sheet, data };
  }
  const entries = [...(sheet.celldata ?? [])];
  const indexes = new Map(
    entries.map((entry, index) => [`${entry.r}:${entry.c}`, index]),
  );
  for (const [column, cell] of updates) {
    const key = `${rowIndex}:${column}`;
    const index = indexes.get(key);
    if (cell === null) {
      if (index !== undefined) entries.splice(index, 1);
      if (index !== undefined) {
        indexes.clear();
        entries.forEach((entry, nextIndex) => {
          indexes.set(`${entry.r}:${entry.c}`, nextIndex);
        });
      }
      continue;
    }
    const entry = { r: rowIndex, c: column, v: cell };
    if (index === undefined) {
      indexes.set(key, entries.length);
      entries.push(entry);
    } else {
      entries[index] = entry;
    }
  }
  entries.sort((left, right) => left.r - right.r || left.c - right.c);
  return { ...sheet, celldata: entries };
}
