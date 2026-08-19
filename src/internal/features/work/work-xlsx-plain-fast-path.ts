import type { SpreadsheetImportWorkbookMetadata } from './work-spreadsheet-import-worker-protocol';

const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const SPREADSHEET_NAMESPACE =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOCUMENT_RELATIONSHIP = `${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument`;
const WORKSHEET_RELATIONSHIP = `${OFFICE_RELATIONSHIPS_NAMESPACE}/worksheet`;
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const RELATIONSHIPS_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.relationships+xml';
const MAX_XLSX_COLUMN = 16_383;
const MAX_XLSX_ROW = 1_048_575;
export const PLAIN_XLSX_COLUMN_BITS = 14;
export const PLAIN_XLSX_COLUMN_MASK = (1 << PLAIN_XLSX_COLUMN_BITS) - 1;
export const PLAIN_XLSX_CELL_CHUNK_SIZE = 4_096;
export const PLAIN_XLSX_ROW_CHUNK_SIZE = 256;
export const PLAIN_XLSX_CELL_NUMBER = 0;
export const PLAIN_XLSX_CELL_BOOLEAN = 1;
export const PLAIN_XLSX_CELL_ERROR = 2;
export const PLAIN_XLSX_CELL_TEXT = 3;

export interface PlainXlsxPackageParts {
  contentTypes: string;
  packagePaths: readonly string[];
  rootRelationships: string;
  workbook: string;
  workbookRelationships: string;
}

export interface PlainXlsxWorkbookSheet {
  name: string;
  partPath: string;
}

export interface PlainXlsxWorkbookPlan {
  sheets: PlainXlsxWorkbookSheet[];
  workbook: SpreadsheetImportWorkbookMetadata;
}

export interface PlainXlsxCellChunk {
  coordinates: Uint32Array;
  kinds: Uint8Array;
  numericValues: Float64Array;
  startRow: number;
  textValues: string[];
}

export interface PlainXlsxWorksheetResult {
  columnCount: number;
  populatedCellCount: number;
  properties: Record<string, unknown>;
  rowCount: number;
}

export interface PlainXlsxRangeExtent {
  columnCount: number;
  rowCount: number;
}

interface ParsedTag {
  attributes: Record<string, string>;
  name: string;
  selfClosing: boolean;
}

interface WorkbookSheet {
  hidden: 0 | 1 | 2;
  name: string;
  relationshipId: string;
  sheetId: string;
}

interface CellPosition {
  column: number;
  row: number;
}

type PlainXlsxCellValue =
  | {
      end: number;
      kind:
        | typeof PLAIN_XLSX_CELL_NUMBER
        | typeof PLAIN_XLSX_CELL_BOOLEAN
        | typeof PLAIN_XLSX_CELL_ERROR;
      numericValue: number;
    }
  | {
      end: number;
      kind: typeof PLAIN_XLSX_CELL_TEXT;
      textValue: string;
    };

/**
 * Authenticates the deliberately small OOXML subset handled by the fast
 * worksheet parser. Any package part or workbook construct outside this
 * closed set returns null and leaves SheetJS as the authority.
 */
export function createPlainXlsxWorkbookPlan(
  parts: PlainXlsxPackageParts,
): PlainXlsxWorkbookPlan | null {
  const sheets = parseWorkbookSheets(parts.workbook);
  if (!sheets?.length) return null;
  const relationships = parseRelationships(parts.workbookRelationships);
  if (!relationships || relationships.length !== sheets.length) return null;
  const rootRelationships = parseRelationships(parts.rootRelationships);
  if (
    !rootRelationships ||
    rootRelationships.length !== 1 ||
    rootRelationships[0]?.Type !== OFFICE_DOCUMENT_RELATIONSHIP ||
    rootRelationships[0]?.Target !== 'xl/workbook.xml' ||
    rootRelationships[0]?.TargetMode !== undefined
  ) {
    return null;
  }

  const relationshipById = new Map(
    relationships.map((relationship) => [relationship.Id, relationship]),
  );
  if (relationshipById.size !== relationships.length) return null;
  const plannedSheets: PlainXlsxWorkbookSheet[] = [];
  for (const sheet of sheets) {
    const relationship = relationshipById.get(sheet.relationshipId);
    if (
      !relationship ||
      relationship.Type !== WORKSHEET_RELATIONSHIP ||
      relationship.TargetMode !== undefined
    ) {
      return null;
    }
    const partPath = plainWorksheetPartPath(relationship.Target);
    if (!partPath) return null;
    plannedSheets.push({ name: sheet.name, partPath });
  }
  if (
    new Set(plannedSheets.map((sheet) => sheet.partPath)).size !== sheets.length
  )
    return null;

  const expectedPaths = new Set([
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    ...plannedSheets.map((sheet) => sheet.partPath),
  ]);
  if (!sameStringSet(parts.packagePaths, expectedPaths)) return null;
  if (!validPlainContentTypes(parts.contentTypes, plannedSheets)) return null;

  const sheetNames = sheets.map((sheet) => sheet.name);
  return {
    sheets: plannedSheets,
    workbook: {
      Custprops: {},
      Props: { SheetNames: sheetNames, Worksheets: sheets.length },
      SheetNames: sheetNames,
      Workbook: {
        Names: [],
        Sheets: sheets.map((sheet) => ({
          Hidden: sheet.hidden,
          id: sheet.relationshipId,
          name: sheet.name,
          sheetId: sheet.sheetId,
          sheetid: sheet.sheetId,
        })),
      },
    } as SpreadsheetImportWorkbookMetadata,
  };
}

/**
 * Parses primitive dense cells with a low-allocation cursor. Row chunks may be
 * emitted provisionally; callers must use them only when the final result is
 * non-null because a later unsupported element invalidates the whole sheet.
 */
export function streamPlainXlsxWorksheet(
  source: string,
  visit: (chunk: PlainXlsxCellChunk) => void,
  start?: (extent: PlainXlsxRangeExtent) => void,
): PlainXlsxWorksheetResult | null {
  let cursor = xmlBodyStart(source);
  const worksheetOpen = `<worksheet xmlns="${SPREADSHEET_NAMESPACE}">`;
  if (!source.startsWith(worksheetOpen, cursor)) return null;
  cursor = skipXmlWhitespace(source, cursor + worksheetOpen.length);

  const dimensionPrefix = '<dimension ref="';
  if (!source.startsWith(dimensionPrefix, cursor)) return null;
  const dimensionEnd = source.indexOf('"/>', cursor + dimensionPrefix.length);
  if (dimensionEnd < 0) return null;
  const reference = source.slice(cursor + dimensionPrefix.length, dimensionEnd);
  const range = decodePlainRange(reference);
  if (!range) return null;
  cursor = skipXmlWhitespace(source, dimensionEnd + 3);
  if (!source.startsWith('<sheetData>', cursor)) return null;
  cursor += '<sheetData>'.length;
  start?.({
    columnCount: range.end.column + 1,
    rowCount: range.end.row + 1,
  });

  let chunkCoordinates: number[] = [];
  let chunkKinds: number[] = [];
  let chunkNumericValues: number[] = [];
  let chunkTextValues: string[] = [];
  let chunkStart = -1;
  let lastRow = -1;
  let populatedCellCount = 0;
  const flush = () => {
    if (chunkStart < 0 || !chunkCoordinates.length) return;
    visit({
      coordinates: Uint32Array.from(chunkCoordinates),
      kinds: Uint8Array.from(chunkKinds),
      numericValues: Float64Array.from(chunkNumericValues),
      startRow: chunkStart,
      textValues: chunkTextValues,
    });
    chunkCoordinates = [];
    chunkKinds = [];
    chunkNumericValues = [];
    chunkTextValues = [];
    chunkStart = -1;
  };

  while (true) {
    cursor = skipXmlWhitespace(source, cursor);
    if (source.startsWith('</sheetData>', cursor)) {
      cursor += '</sheetData>'.length;
      break;
    }
    const rowStart = readExactPositiveInteger(
      source,
      cursor,
      '<row r="',
      '">',
      MAX_XLSX_ROW + 1,
    );
    if (!rowStart) return null;
    const row = rowStart.value - 1;
    if (
      row <= lastRow ||
      row > MAX_XLSX_ROW ||
      row < range.start.row ||
      row > range.end.row
    )
      return null;
    cursor = rowStart.end;
    const nextChunkStart =
      Math.floor(row / PLAIN_XLSX_ROW_CHUNK_SIZE) * PLAIN_XLSX_ROW_CHUNK_SIZE;
    if (chunkStart !== nextChunkStart) {
      flush();
      chunkStart = nextChunkStart;
    }
    let lastColumn = -1;
    let rowCellCount = 0;

    while (true) {
      cursor = skipXmlWhitespace(source, cursor);
      if (source.startsWith('</row>', cursor)) {
        cursor += '</row>'.length;
        break;
      }
      if (!source.startsWith('<c r="', cursor)) return null;
      const addressStart = cursor + '<c r="'.length;
      const addressEnd = source.indexOf('"', addressStart);
      if (addressEnd < 0) return null;
      const column = decodePlainCellColumnForRow(
        source,
        addressStart,
        addressEnd,
        row,
      );
      if (
        column === null ||
        column <= lastColumn ||
        column < range.start.column ||
        column > range.end.column
      ) {
        return null;
      }
      const cell = readPlainCell(source, addressEnd + 1);
      if (!cell) return null;
      if (chunkCoordinates.length >= PLAIN_XLSX_CELL_CHUNK_SIZE) {
        flush();
        chunkStart = nextChunkStart;
      }
      chunkCoordinates.push(
        ((row - chunkStart) << PLAIN_XLSX_COLUMN_BITS) | column,
      );
      chunkKinds.push(cell.kind);
      if (cell.kind === PLAIN_XLSX_CELL_TEXT) {
        chunkTextValues.push(cell.textValue);
      } else {
        chunkNumericValues.push(cell.numericValue);
      }
      cursor = cell.end;
      lastColumn = column;
      rowCellCount += 1;
    }
    if (!rowCellCount) return null;
    populatedCellCount += rowCellCount;
    lastRow = row;
  }

  cursor = skipXmlWhitespace(source, cursor);
  if (!source.startsWith('</worksheet>', cursor)) return null;
  cursor = skipXmlWhitespace(source, cursor + '</worksheet>'.length);
  if (cursor !== source.length || populatedCellCount === 0) return null;
  flush();
  return {
    columnCount: range.end.column + 1,
    populatedCellCount,
    properties: { '!ref': reference },
    rowCount: Math.max(lastRow + 1, range.end.row + 1),
  };
}

export function plainXlsxRangeExtent(
  reference: string,
): PlainXlsxRangeExtent | null {
  const range = decodePlainRange(reference);
  return range
    ? {
        columnCount: range.end.column + 1,
        rowCount: range.end.row + 1,
      }
    : null;
}

function readPlainCell(
  source: string,
  cursor: number,
): PlainXlsxCellValue | null {
  if (source.startsWith('>', cursor)) {
    const value = readElementText(source, cursor + 1, '<v>', '</v></c>');
    if (!value || !PLAIN_NUMBER.test(value.text)) return null;
    const number = Number(value.text);
    return Number.isFinite(number)
      ? {
          end: value.end,
          kind: PLAIN_XLSX_CELL_NUMBER,
          numericValue: number,
        }
      : null;
  }
  if (source.startsWith(' t="inlineStr">', cursor)) {
    const contentStart = cursor + ' t="inlineStr">'.length;
    const textPrefix = source.startsWith('<is><t>', contentStart)
      ? '<is><t>'
      : source.startsWith('<is><t xml:space="preserve">', contentStart)
        ? '<is><t xml:space="preserve">'
        : null;
    if (!textPrefix) return null;
    const value = readElementText(
      source,
      contentStart,
      textPrefix,
      '</t></is></c>',
    );
    if (!value) return null;
    const text = decodeStrictXmlText(value.text);
    return text === null
      ? null
      : { end: value.end, kind: PLAIN_XLSX_CELL_TEXT, textValue: text };
  }
  if (source.startsWith(' t="b">', cursor)) {
    const value = readElementText(
      source,
      cursor + ' t="b">'.length,
      '<v>',
      '</v></c>',
    );
    if (!value || (value.text !== '0' && value.text !== '1')) return null;
    return {
      end: value.end,
      kind: PLAIN_XLSX_CELL_BOOLEAN,
      numericValue: value.text === '1' ? 1 : 0,
    };
  }
  if (source.startsWith(' t="e">', cursor)) {
    const value = readElementText(
      source,
      cursor + ' t="e">'.length,
      '<v>',
      '</v></c>',
    );
    if (!value) return null;
    const error = XLSX_ERROR_VALUES[value.text];
    return error === undefined
      ? null
      : {
          end: value.end,
          kind: PLAIN_XLSX_CELL_ERROR,
          numericValue: error,
        };
  }
  return null;
}

const PLAIN_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;

const XLSX_ERROR_VALUES: Readonly<Record<string, number>> = {
  '#NULL!': 0x00,
  '#DIV/0!': 0x07,
  '#VALUE!': 0x0f,
  '#REF!': 0x17,
  '#NAME?': 0x1d,
  '#NUM!': 0x24,
  '#N/A': 0x2a,
  '#GETTING_DATA': 0x2b,
};

function readElementText(
  source: string,
  cursor: number,
  opening: string,
  closing: string,
): { end: number; text: string } | null {
  if (!source.startsWith(opening, cursor)) return null;
  const start = cursor + opening.length;
  const end = source.indexOf(closing, start);
  return end < 0
    ? null
    : { end: end + closing.length, text: source.slice(start, end) };
}

function parseWorkbookSheets(source: string): WorkbookSheet[] | null {
  const workbook = rootContent(source, 'workbook', {
    xmlns: SPREADSHEET_NAMESPACE,
    'xmlns:r': OFFICE_RELATIONSHIPS_NAMESPACE,
  });
  if (workbook === null) return null;
  const sheetsSource = rootContent(workbook, 'sheets', {});
  if (sheetsSource === null) return null;
  const tags = emptyChildTags(sheetsSource, 'sheet');
  if (!tags?.length) return null;
  const names = new Set<string>();
  const relationshipIds = new Set<string>();
  const sheetIds = new Set<string>();
  const sheets: WorkbookSheet[] = [];
  for (const tag of tags) {
    if (!hasOnlyKeys(tag.attributes, ['name', 'sheetId', 'r:id', 'state']))
      return null;
    const name = tag.attributes.name;
    const relationshipId = tag.attributes['r:id'];
    const sheetId = tag.attributes.sheetId;
    const hidden = plainSheetHiddenValue(tag.attributes.state);
    if (
      !name ||
      !relationshipId ||
      !sheetId ||
      hidden === null ||
      names.has(name) ||
      relationshipIds.has(relationshipId) ||
      sheetIds.has(sheetId)
    ) {
      return null;
    }
    names.add(name);
    relationshipIds.add(relationshipId);
    sheetIds.add(sheetId);
    sheets.push({ hidden, name, relationshipId, sheetId });
  }
  return sheets;
}

function validPlainContentTypes(
  source: string,
  sheets: readonly PlainXlsxWorkbookSheet[],
): boolean {
  const content = rootContent(source, 'Types', {
    xmlns: CONTENT_TYPES_NAMESPACE,
  });
  if (content === null) return false;
  const tags = emptyChildTags(content);
  if (!tags) return false;
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const tag of tags) {
    if (tag.name === 'Default') {
      if (!hasOnlyKeys(tag.attributes, ['Extension', 'ContentType']))
        return false;
      const extension = tag.attributes.Extension;
      const contentType = tag.attributes.ContentType;
      if (!extension || !contentType || defaults.has(extension)) return false;
      defaults.set(extension, contentType);
    } else if (tag.name === 'Override') {
      if (!hasOnlyKeys(tag.attributes, ['PartName', 'ContentType']))
        return false;
      const partName = tag.attributes.PartName;
      const contentType = tag.attributes.ContentType;
      if (!partName || !contentType || overrides.has(partName)) return false;
      overrides.set(partName, contentType);
    } else return false;
  }
  if (
    defaults.size !== 2 ||
    defaults.get('rels') !== RELATIONSHIPS_CONTENT_TYPE ||
    defaults.get('xml') !== 'application/xml'
  ) {
    return false;
  }
  const expectedOverrides = new Map<string, string>([
    ['/xl/workbook.xml', WORKBOOK_CONTENT_TYPE],
    ...sheets.map(
      (sheet) => [`/${sheet.partPath}`, WORKSHEET_CONTENT_TYPE] as const,
    ),
  ]);
  if (overrides.size !== expectedOverrides.size) return false;
  for (const [path, contentType] of expectedOverrides) {
    if (overrides.get(path) !== contentType) return false;
  }
  return true;
}

function parseRelationships(
  source: string,
): Array<Record<string, string>> | null {
  const content = rootContent(source, 'Relationships', {
    xmlns: PACKAGE_RELATIONSHIPS_NAMESPACE,
  });
  if (content === null) return null;
  const tags = emptyChildTags(content, 'Relationship');
  if (!tags) return null;
  const relationships: Array<Record<string, string>> = [];
  for (const tag of tags) {
    if (!hasOnlyKeys(tag.attributes, ['Id', 'Type', 'Target', 'TargetMode']))
      return null;
    if (!tag.attributes.Id || !tag.attributes.Type || !tag.attributes.Target)
      return null;
    relationships.push(tag.attributes);
  }
  return relationships;
}

function rootContent(
  source: string,
  name: string,
  attributes: Readonly<Record<string, string>>,
): string | null {
  const body = stripSmallXmlDocument(source);
  const tagEnd = body.indexOf('>');
  if (tagEnd < 0) return null;
  const tag = parseTag(body.slice(0, tagEnd + 1));
  if (
    !tag ||
    tag.selfClosing ||
    tag.name !== name ||
    !sameAttributes(tag.attributes, attributes)
  ) {
    return null;
  }
  const closing = `</${name}>`;
  if (!body.endsWith(closing)) return null;
  return body.slice(tagEnd + 1, -closing.length).trim();
}

function emptyChildTags(
  source: string,
  expectedName?: string,
): ParsedTag[] | null {
  const tags: ParsedTag[] = [];
  let cursor = 0;
  while (cursor < source.length) {
    cursor = skipXmlWhitespace(source, cursor);
    if (cursor === source.length) break;
    const end = source.indexOf('>', cursor);
    if (end < 0) return null;
    const tag = parseTag(source.slice(cursor, end + 1));
    if (!tag?.selfClosing || (expectedName && tag.name !== expectedName))
      return null;
    tags.push(tag);
    cursor = end + 1;
  }
  return tags;
}

function parseTag(source: string): ParsedTag | null {
  const match = /^<([A-Za-z_][\w:.-]*)([\s\S]*?)(\/?)>$/.exec(source);
  if (!match) return null;
  const attributes: Record<string, string> = {};
  const attributeSource = match[2] ?? '';
  let cursor = 0;
  while (cursor < attributeSource.length) {
    const attribute = /^\s+([A-Za-z_][\w:.-]*)="([^"]*)"/.exec(
      attributeSource.slice(cursor),
    );
    if (!attribute || Object.hasOwn(attributes, attribute[1])) return null;
    const value = decodeStrictXmlText(attribute[2]);
    if (value === null) return null;
    attributes[attribute[1]] = value;
    cursor += attribute[0].length;
  }
  return { attributes, name: match[1], selfClosing: match[3] === '/' };
}

function decodeStrictXmlText(value: string): string | null {
  if (value.includes('<')) return null;
  if (!value.includes('&')) return value;
  let result = '';
  let cursor = 0;
  while (cursor < value.length) {
    const start = value.indexOf('&', cursor);
    if (start < 0) return result + value.slice(cursor);
    result += value.slice(cursor, start);
    const end = value.indexOf(';', start + 1);
    if (end < 0) return null;
    const entity = value.slice(start + 1, end);
    const decoded = decodeXmlEntity(entity);
    if (decoded === null) return null;
    result += decoded;
    cursor = end + 1;
  }
  return result;
}

function decodeXmlEntity(entity: string): string | null {
  const named: Readonly<Record<string, string>> = {
    amp: '&',
    apos: "'",
    gt: '>',
    lt: '<',
    quot: '"',
  };
  if (Object.hasOwn(named, entity)) return named[entity] ?? null;
  const codePoint = /^#x[\da-f]+$/i.test(entity)
    ? Number.parseInt(entity.slice(2), 16)
    : /^#\d+$/.test(entity)
      ? Number.parseInt(entity.slice(1), 10)
      : Number.NaN;
  if (!isXmlCodePoint(codePoint)) return null;
  try {
    return String.fromCodePoint(codePoint);
  } catch {
    return null;
  }
}

function isXmlCodePoint(value: number): boolean {
  return (
    Number.isSafeInteger(value) &&
    (value === 0x09 ||
      value === 0x0a ||
      value === 0x0d ||
      (value >= 0x20 && value <= 0xd7ff) ||
      (value >= 0xe000 && value <= 0xfffd) ||
      (value >= 0x10000 && value <= 0x10ffff))
  );
}

function decodePlainRange(
  reference: string,
): { end: CellPosition; start: CellPosition } | null {
  const separator = reference.indexOf(':');
  const start = decodePlainCellAddress(
    separator < 0 ? reference : reference.slice(0, separator),
  );
  const end = decodePlainCellAddress(
    separator < 0 ? reference : reference.slice(separator + 1),
  );
  return start &&
    end &&
    start.row <= end.row &&
    start.column <= end.column &&
    (separator < 0 || reference.indexOf(':', separator + 1) < 0)
    ? { end, start }
    : null;
}

function decodePlainCellAddress(address: string): CellPosition | null {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(address);
  if (!match) return null;
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  let row = Number(match[2]);
  column -= 1;
  row -= 1;
  return column <= MAX_XLSX_COLUMN && row <= MAX_XLSX_ROW
    ? { column, row }
    : null;
}

function decodePlainCellColumnForRow(
  source: string,
  start: number,
  end: number,
  expectedRow: number,
): number | null {
  let cursor = start;
  let column = 0;
  let columnDigits = 0;
  while (cursor < end && columnDigits < 3) {
    const code = source.charCodeAt(cursor);
    if (code < 65 || code > 90) break;
    column = column * 26 + code - 64;
    columnDigits += 1;
    cursor += 1;
  }
  if (columnDigits === 0 || column > MAX_XLSX_COLUMN + 1 || cursor >= end) {
    return null;
  }

  const firstRowDigit = source.charCodeAt(cursor);
  if (firstRowDigit < 49 || firstRowDigit > 57) return null;
  let row = firstRowDigit - 48;
  let rowDigits = 1;
  cursor += 1;
  while (cursor < end) {
    const code = source.charCodeAt(cursor);
    if (code < 48 || code > 57 || rowDigits >= 7) return null;
    row = row * 10 + code - 48;
    rowDigits += 1;
    cursor += 1;
  }
  return row === expectedRow + 1 && row <= MAX_XLSX_ROW + 1 ? column - 1 : null;
}

function readExactPositiveInteger(
  source: string,
  cursor: number,
  prefix: string,
  suffix: string,
  maximum = Number.MAX_SAFE_INTEGER,
): { end: number; value: number } | null {
  if (!source.startsWith(prefix, cursor)) return null;
  const start = cursor + prefix.length;
  const end = source.indexOf(suffix, start);
  if (end <= start) return null;
  const firstDigit = source.charCodeAt(start);
  if (firstDigit < 49 || firstDigit > 57) return null;
  let value = firstDigit - 48;
  for (let index = start + 1; index < end; index += 1) {
    const code = source.charCodeAt(index);
    if (code < 48 || code > 57) return null;
    value = value * 10 + code - 48;
    if (value > maximum) return null;
  }
  return value <= maximum ? { end: end + suffix.length, value } : null;
}

function stripSmallXmlDocument(source: string): string {
  let body = source.trim();
  if (body.startsWith('<?xml')) {
    const end = body.indexOf('?>');
    if (end < 0) return '';
    body = body.slice(end + 2).trim();
  }
  return body;
}

function xmlBodyStart(source: string): number {
  let cursor = skipXmlWhitespace(source, 0);
  if (source.startsWith('<?xml', cursor)) {
    const end = source.indexOf('?>', cursor + 5);
    if (end < 0) return source.length;
    cursor = skipXmlWhitespace(source, end + 2);
  }
  return cursor;
}

function skipXmlWhitespace(source: string, start: number): number {
  let cursor = start;
  while (cursor < source.length) {
    const code = source.charCodeAt(cursor);
    if (code !== 9 && code !== 10 && code !== 13 && code !== 32) break;
    cursor += 1;
  }
  return cursor;
}

function plainWorksheetPartPath(target: string | undefined): string | null {
  return target && /^worksheets\/[A-Za-z0-9._-]+\.xml$/.test(target)
    ? `xl/${target}`
    : null;
}

function plainSheetHiddenValue(value: string | undefined): 0 | 1 | 2 | null {
  if (value === undefined || value === 'visible') return 0;
  if (value === 'hidden') return 1;
  if (value === 'veryHidden') return 2;
  return null;
}

function sameStringSet(
  values: readonly string[],
  expected: ReadonlySet<string>,
): boolean {
  const actual = new Set(values);
  if (actual.size !== values.length || actual.size !== expected.size)
    return false;
  for (const value of actual) if (!expected.has(value)) return false;
  return true;
}

function hasOnlyKeys(
  value: Readonly<Record<string, string>>,
  allowed: readonly string[],
): boolean {
  const keys = Object.keys(value);
  return keys.every((key) => allowed.includes(key));
}

function sameAttributes(
  actual: Readonly<Record<string, string>>,
  expected: Readonly<Record<string, string>>,
): boolean {
  const keys = Object.keys(actual);
  return (
    keys.length === Object.keys(expected).length &&
    keys.every((key) => actual[key] === expected[key])
  );
}
