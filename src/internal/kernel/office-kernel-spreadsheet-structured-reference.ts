import type {
  OfficeKernelSpreadsheetInputSheet,
  OfficeKernelSpreadsheetInputTable,
} from './office-kernel-spreadsheet-protocol';

export interface SpreadsheetStructuredRowSelection {
  all: boolean;
  headers: boolean;
  data: boolean;
  totals: boolean;
  current: boolean;
}

export interface SpreadsheetStructuredReference {
  tableName?: string;
  firstColumn?: string;
  lastColumn?: string;
  rows: SpreadsheetStructuredRowSelection;
}

export interface SpreadsheetStructuredReferenceArea {
  sheetId: string;
  sheetName: string;
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

export type SpreadsheetStructuredReferenceErrorKind =
  | 'invalid'
  | 'unsupported'
  | 'missing-table'
  | 'missing-column';

export class SpreadsheetStructuredReferenceError extends Error {
  readonly kind: SpreadsheetStructuredReferenceErrorKind;

  constructor(kind: SpreadsheetStructuredReferenceErrorKind, message: string) {
    super(message);
    this.name = 'SpreadsheetStructuredReferenceError';
    this.kind = kind;
  }
}

/** Parse one bounded table-reference token using the OOXML column grammar. */
export function parseSpreadsheetStructuredReference(
  reference: string,
): SpreadsheetStructuredReference {
  const open = reference.indexOf('[');
  if (open < 0) throw invalidReference(reference);
  const tableName = reference.slice(0, open) || undefined;
  const content = outerGroup(reference.slice(open));
  if (content === null) throw invalidReference(reference);

  const rows: SpreadsheetStructuredRowSelection = {
    all: false,
    headers: false,
    data: false,
    totals: false,
    current: false,
  };
  let firstColumn: string | undefined;
  let lastColumn: string | undefined;
  if (content.startsWith('@')) {
    rows.current = true;
    const column = parseCurrentColumn(content.slice(1), reference);
    firstColumn = column;
    lastColumn = column;
  } else if (content.startsWith('[')) {
    [firstColumn, lastColumn] = parseNestedSelection(content, reference, rows);
  } else {
    const item = tableItem(content);
    if (item) applyTableItem(rows, item, reference);
    else {
      const column = parsePlainColumn(content, reference);
      firstColumn = column;
      lastColumn = column;
    }
  }
  if (!rows.all && !rows.headers && !rows.data && !rows.totals && !rows.current)
    rows.data = true;
  return { tableName, firstColumn, lastColumn, rows };
}

/**
 * Catalog the semantic tables in a workbook and resolve references to bounded
 * zero-based rectangular areas. The catalog deliberately contains no cells;
 * values remain in the calculation evaluator so dependencies stay observable.
 */
export class SpreadsheetStructuredReferenceCatalog {
  private readonly definitions: Array<
    OfficeKernelSpreadsheetInputTable & {
      sheetId: string;
      sheetName: string;
    }
  > = [];
  private readonly byName = new Map<string, number>();
  private readonly bySheet = new Map<string, number[]>();

  constructor(readonly sheets: readonly OfficeKernelSpreadsheetInputSheet[]) {
    for (const sheet of sheets) {
      const indexes: number[] = [];
      for (const table of sheet.tables ?? []) {
        const definition = {
          ...table,
          sheetId: sheet.id,
          sheetName: sheet.name,
        };
        const index = this.definitions.length;
        this.definitions.push(definition);
        indexes.push(index);
        for (const alias of [table.name, table.displayName]) {
          if (!alias) continue;
          const key = alias.toLocaleLowerCase();
          if (!this.byName.has(key)) this.byName.set(key, index);
        }
      }
      this.bySheet.set(sheet.id, indexes);
    }
  }

  resolve(
    qualifier: string | undefined,
    reference: string,
    currentSheet: OfficeKernelSpreadsheetInputSheet,
    currentColumn: number,
    currentRow: number,
  ): SpreadsheetStructuredReferenceArea[] {
    const parsed = parseSpreadsheetStructuredReference(reference);
    const table = this.resolveTable(
      parsed,
      reference,
      currentSheet,
      currentColumn,
      currentRow,
    );
    if (
      qualifier &&
      table.sheetName.toLocaleLowerCase() !== qualifier.toLocaleLowerCase()
    ) {
      throw new SpreadsheetStructuredReferenceError(
        'missing-table',
        `Spreadsheet table '${table.name}' is not on worksheet '${qualifier}'.`,
      );
    }
    if (parsed.rows.current && currentSheet.id !== table.sheetId) {
      throw new SpreadsheetStructuredReferenceError(
        'unsupported',
        `Structured reference #This Row requires the current formula cell to be on table '${table.name}'.`,
      );
    }
    const [startColumn, endColumn] = this.resolveColumns(table, parsed);
    return this.resolveRows(table, parsed.rows, currentRow).map(
      ([startRow, endRow]) => ({
        sheetId: table.sheetId,
        sheetName: table.sheetName,
        startRow,
        endRow,
        startColumn,
        endColumn,
      }),
    );
  }

  private resolveTable(
    parsed: SpreadsheetStructuredReference,
    reference: string,
    currentSheet: OfficeKernelSpreadsheetInputSheet,
    currentColumn: number,
    currentRow: number,
  ) {
    if (parsed.tableName) {
      const index = this.byName.get(parsed.tableName.toLocaleLowerCase());
      if (index === undefined)
        throw new SpreadsheetStructuredReferenceError(
          'missing-table',
          `Spreadsheet table '${parsed.tableName}' does not exist.`,
        );
      const table = this.definitions[index];
      if (!table) throw invalidReference(reference);
      return table;
    }
    const matching = (this.bySheet.get(currentSheet.id) ?? [])
      .map((index) => this.definitions[index])
      .filter(
        (table): table is NonNullable<typeof table> =>
          table !== undefined &&
          currentRow >= table.startRow &&
          currentRow <= table.endRow &&
          currentColumn >= table.startColumn &&
          currentColumn <= table.endColumn,
      );
    if (!matching.length)
      throw new SpreadsheetStructuredReferenceError(
        'missing-table',
        `Table-local structured reference '${reference}' requires the current formula cell to be inside a Spreadsheet table.`,
      );
    if (matching.length > 1)
      throw new SpreadsheetStructuredReferenceError(
        'unsupported',
        `Table-local structured reference '${reference}' is ambiguous at the current formula cell.`,
      );
    return matching[0]!;
  }

  private resolveColumns(
    table: (typeof this.definitions)[number],
    parsed: SpreadsheetStructuredReference,
  ): [number, number] {
    let first = 0;
    let last = table.columns.length - 1;
    if (parsed.firstColumn !== undefined && parsed.lastColumn !== undefined) {
      first = table.columns.findIndex(
        (column) =>
          column.toLocaleLowerCase() ===
          parsed.firstColumn!.toLocaleLowerCase(),
      );
      last = table.columns.findIndex(
        (column) =>
          column.toLocaleLowerCase() === parsed.lastColumn!.toLocaleLowerCase(),
      );
      if (first < 0)
        throw new SpreadsheetStructuredReferenceError(
          'missing-column',
          `Spreadsheet table '${table.name}' has no column '${parsed.firstColumn}'.`,
        );
      if (last < 0)
        throw new SpreadsheetStructuredReferenceError(
          'missing-column',
          `Spreadsheet table '${table.name}' has no column '${parsed.lastColumn}'.`,
        );
      if (first > last)
        throw new SpreadsheetStructuredReferenceError(
          'unsupported',
          `Structured-reference column range '${parsed.firstColumn}:${parsed.lastColumn}' is reversed.`,
        );
    }
    if (first < 0 || last < first || table.startColumn + last > table.endColumn)
      throw invalidReference(table.name);
    return [table.startColumn + first, table.startColumn + last];
  }

  private resolveRows(
    table: (typeof this.definitions)[number],
    rows: SpreadsheetStructuredRowSelection,
    currentRow: number,
  ): Array<[number, number]> {
    const selected: Array<[number, number]> = [];
    if (rows.all) selected.push([table.startRow, table.endRow]);
    if (rows.headers) {
      if (!table.headerRow) throw missingRows(table, '#Headers');
      selected.push([table.startRow, table.startRow]);
    }
    if (rows.data) selected.push(this.dataRows(table));
    if (rows.totals) {
      if (!table.totalsRow) throw missingRows(table, '#Totals');
      selected.push([table.endRow, table.endRow]);
    }
    if (rows.current) {
      const [start, end] = this.dataRows(table);
      if (currentRow < start || currentRow > end)
        throw new SpreadsheetStructuredReferenceError(
          'unsupported',
          `Structured reference #This Row requires the current formula row to be inside table '${table.name}'.`,
        );
      selected.push([currentRow, currentRow]);
    }
    selected.sort((left, right) => left[0] - right[0]);
    const merged: Array<[number, number]> = [];
    for (const [start, end] of selected) {
      const previous = merged.at(-1);
      if (previous && start <= previous[1] + 1)
        previous[1] = Math.max(previous[1], end);
      else merged.push([start, end]);
    }
    if (!merged.length)
      throw new SpreadsheetStructuredReferenceError(
        'unsupported',
        'Structured reference selects no table rows.',
      );
    return merged;
  }

  private dataRows(table: (typeof this.definitions)[number]): [number, number] {
    const start = table.startRow + (table.headerRow ? 1 : 0);
    const end = table.endRow - (table.totalsRow ? 1 : 0);
    if (start > end)
      throw new SpreadsheetStructuredReferenceError(
        'unsupported',
        `Spreadsheet table '${table.name}' has no data rows.`,
      );
    return [start, end];
  }
}

/** Expand structured tokens to A1 ranges before invoking the scalar parser. */
export function expandSpreadsheetStructuredReferences(
  formula: string,
  catalog: SpreadsheetStructuredReferenceCatalog,
  currentSheet: OfficeKernelSpreadsheetInputSheet,
  currentRow: number,
  currentColumn: number,
): string {
  const hasEquals = formula.startsWith('=');
  const source = hasEquals ? formula.slice(1) : formula;
  let output = '';
  let cursor = 0;
  while (cursor < source.length) {
    const character = source[cursor] ?? '';
    if (character === '"') {
      const end = quotedStringEnd(source, cursor);
      output += source.slice(cursor, end);
      cursor = end;
      continue;
    }
    const token = scanStructuredReference(source, cursor);
    if (!token) {
      output += character;
      cursor += 1;
      continue;
    }
    const areas = catalog.resolve(
      token.qualifier,
      token.reference,
      currentSheet,
      currentColumn,
      currentRow,
    );
    if (areas.length > 1) {
      throw new SpreadsheetStructuredReferenceError(
        'unsupported',
        `Structured reference '${token.reference}' resolves to disjoint row areas; the JavaScript fallback requires a contiguous range.`,
      );
    }
    const ranges = areas.map((area) => {
      const prefix =
        area.sheetId === currentSheet.id
          ? ''
          : `${quoteSheetName(area.sheetName)}!`;
      const start = cellAddress(area.startRow, area.startColumn);
      const end = cellAddress(area.endRow, area.endColumn);
      return `${prefix}${start}${start === end ? '' : `:${end}`}`;
    });
    output += ranges.length === 1 ? ranges[0]! : `(${ranges.join(',')})`;
    cursor = token.end;
  }
  return hasEquals ? `=${output}` : output;
}

interface StructuredToken {
  end: number;
  qualifier?: string;
  reference: string;
}

function scanStructuredReference(
  source: string,
  start: number,
): StructuredToken | null {
  const qualified = scanQualifier(source, start);
  let cursor = qualified?.end ?? start;
  const qualifier = qualified?.name;
  if (source[cursor] === '[') {
    const end = matchingBracket(source, cursor);
    if (end === null) return null;
    const content = source.slice(cursor + 1, end - 1);
    if (
      !content.startsWith('@') &&
      !content.startsWith('#') &&
      !content.startsWith('[')
    )
      return null;
    return { end, qualifier, reference: source.slice(cursor, end) };
  }
  const nameStart = cursor;
  if (!isNameStart(source[cursor] ?? '')) return null;
  cursor += 1;
  while (cursor < source.length && isNameContinue(source[cursor]!)) cursor += 1;
  if (source[cursor] !== '[') return null;
  const end = matchingBracket(source, cursor);
  if (end === null)
    throw invalidReference(source.slice(nameStart, source.length));
  return { end, qualifier, reference: source.slice(nameStart, end) };
}

function scanQualifier(
  source: string,
  start: number,
): { name: string; end: number } | null {
  if (source[start] === "'") {
    let cursor = start + 1;
    let decoded = '';
    while (cursor < source.length) {
      const character = source[cursor]!;
      if (character === "'") {
        if (source[cursor + 1] === "'") {
          decoded += "'";
          cursor += 2;
          continue;
        }
        if (source[cursor + 1] === '!')
          return { name: decoded, end: cursor + 2 };
        return null;
      }
      decoded += character;
      cursor += 1;
    }
    return null;
  }
  let cursor = start;
  while (cursor < source.length && isQualifierCharacter(source[cursor]!))
    cursor += 1;
  if (source[cursor] !== '!' || cursor === start) return null;
  return { name: source.slice(start, cursor), end: cursor + 1 };
}

function matchingBracket(source: string, start: number): number | null {
  let depth = 0;
  for (let cursor = start; cursor < source.length; cursor += 1) {
    const character = source[cursor]!;
    if (character === "'") {
      cursor += source[cursor + 1] === "'" ? 1 : 0;
      continue;
    }
    if (character === '[') depth += 1;
    else if (character === ']') {
      depth -= 1;
      if (depth === 0) return cursor + 1;
      if (depth < 0) return null;
    }
  }
  return null;
}

function outerGroup(value: string): string | null {
  if (!value.startsWith('[')) return null;
  const end = matchingBracket(value, 0);
  return end === value.length ? value.slice(1, -1) : null;
}

function bracketAtom(value: string): { atom: string; consumed: number } | null {
  if (!value.startsWith('[')) return null;
  const end = matchingBracket(value, 0);
  return end === null ? null : { atom: value.slice(1, end - 1), consumed: end };
}

function parseCurrentColumn(value: string, reference: string): string {
  if (value.startsWith('[')) {
    const atom = bracketAtom(value);
    if (!atom || atom.consumed !== value.length)
      throw invalidReference(reference);
    const column = decodeAtom(atom.atom);
    if (!column) throw invalidReference(reference);
    return column;
  }
  return parsePlainColumn(value, reference);
}

function parseNestedSelection(
  content: string,
  reference: string,
  rows: SpreadsheetStructuredRowSelection,
): [string | undefined, string | undefined] {
  const atoms: string[] = [];
  const separators: string[] = [];
  let cursor = 0;
  while (cursor < content.length) {
    const atom = bracketAtom(content.slice(cursor));
    if (!atom) throw invalidReference(reference);
    atoms.push(atom.atom);
    cursor += atom.consumed;
    if (cursor === content.length) break;
    const separator = content[cursor];
    if (separator !== ',' && separator !== ':')
      throw invalidReference(reference);
    separators.push(separator);
    cursor += 1;
  }
  const columns: Array<{ index: number; name: string }> = [];
  atoms.forEach((atom, index) => {
    const item = tableItem(atom);
    if (item) applyTableItem(rows, item, reference);
    else {
      const column = decodeAtom(atom);
      if (!column) throw invalidReference(reference);
      columns.push({ index, name: column });
    }
  });
  if (columns.length === 0) {
    if (separators.includes(':')) throw invalidReference(reference);
    return [undefined, undefined];
  }
  if (columns.length === 1) {
    if (separators.includes(':')) throw invalidReference(reference);
    return [columns[0]!.name, columns[0]!.name];
  }
  const first = columns[0]!;
  const last = columns[1]!;
  if (
    columns.length === 2 &&
    last.index === first.index + 1 &&
    separators[first.index] === ':' &&
    separators.every(
      (separator, index) => index === first.index || separator === ',',
    )
  )
    return [first.name, last.name];
  throw new SpreadsheetStructuredReferenceError(
    'unsupported',
    'Disjoint structured-reference columns are not supported.',
  );
}

type TableItem = 'all' | 'headers' | 'data' | 'totals' | 'current';

function tableItem(value: string): TableItem | undefined {
  const normalized = value.toLocaleLowerCase();
  if (normalized === '#all') return 'all';
  if (normalized === '#headers') return 'headers';
  if (normalized === '#data') return 'data';
  if (normalized === '#totals') return 'totals';
  if (normalized === '#this row') return 'current';
  return undefined;
}

function applyTableItem(
  rows: SpreadsheetStructuredRowSelection,
  item: TableItem,
  reference: string,
): void {
  rows[item === 'current' ? 'current' : item] = true;
  if (rows.current && (rows.all || rows.headers || rows.data || rows.totals))
    throw new SpreadsheetStructuredReferenceError(
      'unsupported',
      `Structured reference '${reference}' cannot combine #This Row with another item.`,
    );
}

function parsePlainColumn(value: string, reference: string): string {
  if (!value || /[[\],:]/u.test(value)) throw invalidReference(reference);
  const column = decodeAtom(value);
  if (!column) throw invalidReference(reference);
  return column;
}

function decodeAtom(value: string): string {
  let output = '';
  for (let cursor = 0; cursor < value.length; cursor += 1) {
    const character = value[cursor]!;
    if (character === "'") {
      const escaped = value[cursor + 1];
      if (escaped === undefined) throw invalidReference(value);
      output += escaped;
      cursor += 1;
    } else output += character;
  }
  return output;
}

function missingRows(
  table: { name: string },
  item: string,
): SpreadsheetStructuredReferenceError {
  return new SpreadsheetStructuredReferenceError(
    'unsupported',
    `Structured reference ${item} requires table '${table.name}' to contain that row.`,
  );
}

function invalidReference(
  reference: string,
): SpreadsheetStructuredReferenceError {
  return new SpreadsheetStructuredReferenceError(
    'invalid',
    `Structured reference '${reference}' is not in a supported canonical form.`,
  );
}

function quotedStringEnd(source: string, start: number): number {
  for (let cursor = start + 1; cursor < source.length; cursor += 1) {
    if (source[cursor] !== '"') continue;
    if (source[cursor + 1] === '"') {
      cursor += 1;
      continue;
    }
    return cursor + 1;
  }
  return source.length;
}

function cellAddress(row: number, column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return `${label}${row + 1}`;
}

function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/u.test(name)
    ? name
    : `'${name.replaceAll("'", "''")}'`;
}

function isNameStart(value: string): boolean {
  return /^[A-Za-z_\\?]$/u.test(value);
}

function isNameContinue(value: string): boolean {
  return /^[A-Za-z0-9_.\\?]$/u.test(value);
}

function isQualifierCharacter(value: string): boolean {
  return /^[A-Za-z0-9_.\\?[\]:$]$/u.test(value);
}
