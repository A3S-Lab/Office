import {
  applyDocumentPageGeometry,
  documentPaperSizeForGeometry,
  type WorkDocumentPageGeometry,
} from './work-document-page-size';
import type { DocumentWindowingOptions } from './work-document-windowing';
import type { WorkDocumentSectionLayout } from './work-types';

export const LARGE_SIMPLE_DOCX_MINIMUM_LOGICAL_BLOCKS = 20_000;

const WORD_TWIPS_PER_MILLIMETER = 1_440 / 25.4;
const UNSUPPORTED_SIMPLE_WORD_ELEMENT =
  /<(?:\/)?(?!w:(?:t(?:bl(?:Grid|Layout|Pr|W)?|c(?:Pr|W)?|r)?|p(?:g(?:Mar|Sz))?|r|sectPr|body|document|gridCol)(?=[\s/>]))[A-Za-z_][\w.-]*:[A-Za-z_][\w.-]*(?=[\s/>])/;
const SIMPLE_WORD_ENVELOPE_ELEMENT = /<w:(document|body|sectPr)(?=[\s/>])/g;
const WORD_PARAGRAPH_CLOSE = '</w:p>';
const WORD_TABLE_CLOSE = '</w:tbl>';
const WORD_TABLE_CELL_CLOSE = '</w:tc>';
const WORD_TABLE_ROW_CLOSE = '</w:tr>';
const WORD_TEXT_CLOSE = '</w:t>';

export interface LargeSimpleDocxParseOptions {
  minimumLogicalBlocks?: number;
  windowing?: Partial<DocumentWindowingOptions>;
}

export interface LargeSimpleDocxStreamResult {
  layout: WorkDocumentSectionLayout;
  logicalBlockCount: number;
  paragraphCount: number;
  tableRowCount: number;
}

export interface LargeSimpleDocxStreamSink {
  paragraph: (text: string) => void;
  tableEnd: () => void;
  tableRow: (row: LargeSimpleDocxTableRow) => void;
  tableStart: () => void;
}

export interface LargeSimpleDocxTableRow {
  cellParagraphCounts: readonly number[];
  texts: readonly string[];
}

/**
 * Parses logical blocks into a sink so a Worker can stream bounded batches
 * instead of structured-cloning one complete 100,000-row object graph.
 */
export function streamLargeSimpleDocxDocumentXml(
  source: string,
  options: LargeSimpleDocxParseOptions,
  sink: LargeSimpleDocxStreamSink,
): LargeSimpleDocxStreamResult | null {
  const minimumLogicalBlocks = positiveInteger(
    options.minimumLogicalBlocks,
    LARGE_SIMPLE_DOCX_MINIMUM_LOGICAL_BLOCKS,
  );
  const eligibilityStartedAt = largeDocxImportNow();
  if (!simpleDocxXmlHasSupportedStructure(source)) {
    recordLargeDocxImportMeasure(
      'a3s-office.document.large-simple-eligibility',
      eligibilityStartedAt,
      largeDocxImportNow(),
      { accepted: false },
    );
    return null;
  }
  recordLargeDocxImportMeasure(
    'a3s-office.document.large-simple-eligibility',
    eligibilityStartedAt,
    largeDocxImportNow(),
    { accepted: true },
  );

  const envelopeStartedAt = largeDocxImportNow();
  const body = elementContentRange(source, 'body');
  if (body === null) return null;
  const layout = simpleDocumentLayout(source, body.start, body.end);
  recordLargeDocxImportMeasure(
    'a3s-office.document.large-simple-envelope',
    envelopeStartedAt,
    largeDocxImportNow(),
  );
  let topLevelBlockCount = 0;
  let paragraphCount = 0;
  let standaloneParagraphCount = 0;
  let tableRowCount = 0;
  const contentStartedAt = largeDocxImportNow();
  let blockCursor = body.start;
  let paragraphStart = findWordStartTag(source, 'p', blockCursor, body.end);
  let tableStart = findWordStartTag(source, 'tbl', blockCursor, body.end);
  while (paragraphStart >= 0 || tableStart >= 0) {
    const blockStart = firstIndex(paragraphStart, tableStart);
    const tagEnd = source.indexOf('>', blockStart);
    if (tagEnd < 0 || tagEnd >= body.end) return null;
    if (blockStart === paragraphStart) {
      const contentEnd = source.indexOf(WORD_PARAGRAPH_CLOSE, tagEnd + 1);
      if (contentEnd < 0 || contentEnd > body.end) return null;
      sink.paragraph(simpleText(source, tagEnd + 1, contentEnd));
      topLevelBlockCount += 1;
      paragraphCount += 1;
      standaloneParagraphCount += 1;
      blockCursor = contentEnd + WORD_PARAGRAPH_CLOSE.length;
      paragraphStart = findWordStartTag(source, 'p', blockCursor, body.end);
      if (tableStart >= 0 && tableStart < blockCursor) {
        tableStart = findWordStartTag(source, 'tbl', blockCursor, body.end);
      }
      continue;
    }
    const contentEnd = source.indexOf(WORD_TABLE_CLOSE, tagEnd + 1);
    if (contentEnd < 0 || contentEnd > body.end) return null;
    if (findWordStartTag(source, 'tbl', tagEnd + 1, contentEnd) >= 0) {
      return null;
    }
    sink.tableStart();
    const table = streamSimpleTableRows(
      source,
      tagEnd + 1,
      contentEnd,
      sink.tableRow,
    );
    if (!table) return null;
    sink.tableEnd();
    topLevelBlockCount += 1;
    paragraphCount += table.paragraphCount;
    tableRowCount += table.rowCount;
    blockCursor = contentEnd + WORD_TABLE_CLOSE.length;
    tableStart = findWordStartTag(source, 'tbl', blockCursor, body.end);
    if (paragraphStart >= 0 && paragraphStart < blockCursor) {
      paragraphStart = findWordStartTag(source, 'p', blockCursor, body.end);
    }
  }
  if (!topLevelBlockCount) return null;
  recordLargeDocxImportMeasure(
    'a3s-office.document.large-simple-content',
    contentStartedAt,
    largeDocxImportNow(),
    {
      paragraphs: paragraphCount,
      tableRows: tableRowCount,
    },
  );
  const logicalBlockCount = tableRowCount + standaloneParagraphCount;
  if (logicalBlockCount < minimumLogicalBlocks) return null;
  return { layout, logicalBlockCount, paragraphCount, tableRowCount };
}

export function simpleDocxXmlIsEligibleForLargeImport(
  source: string,
  minimumLogicalBlocks = LARGE_SIMPLE_DOCX_MINIMUM_LOGICAL_BLOCKS,
): boolean {
  if (!source || source.includes('<!DOCTYPE') || source.includes('<![CDATA[')) {
    return false;
  }
  const requiredBlocks = positiveInteger(minimumLogicalBlocks, 1);
  const counts = simpleDocxElementCounts(source, requiredBlocks);
  return Boolean(
    counts &&
      counts.document === 1 &&
      counts.body === 1 &&
      counts.sectPr === 1 &&
      Math.max(counts.paragraph, counts.row) >= requiredBlocks,
  );
}

function simpleDocxXmlHasSupportedStructure(source: string): boolean {
  if (!source || source.includes('<!DOCTYPE') || source.includes('<![CDATA[')) {
    return false;
  }
  if (UNSUPPORTED_SIMPLE_WORD_ELEMENT.test(source)) return false;
  const counts = simpleDocxEnvelopeElementCounts(source);
  return counts.document === 1 && counts.body === 1 && counts.sectPr === 1;
}

function streamSimpleTableRows(
  source: string,
  start: number,
  end: number,
  visit: (row: LargeSimpleDocxTableRow) => void,
): {
  paragraphCount: number;
  rowCount: number;
} | null {
  let paragraphCount = 0;
  let rowCount = 0;
  let rowCursor = start;
  while (rowCursor < end) {
    const rowStart = findWordStartTag(source, 'tr', rowCursor, end);
    if (rowStart < 0) break;
    const rowTagEnd = source.indexOf('>', rowStart);
    if (rowTagEnd < 0 || rowTagEnd >= end) return null;
    const rowEnd = source.indexOf(WORD_TABLE_ROW_CLOSE, rowTagEnd + 1);
    if (rowEnd < 0 || rowEnd > end) return null;
    const cellParagraphCounts: number[] = [];
    const texts: string[] = [];
    let cellCursor = rowTagEnd + 1;
    while (cellCursor < rowEnd) {
      const cellStart = findWordStartTag(source, 'tc', cellCursor, rowEnd);
      if (cellStart < 0) break;
      const cellTagEnd = source.indexOf('>', cellStart);
      if (cellTagEnd < 0 || cellTagEnd >= rowEnd) return null;
      const cellEnd = source.indexOf(WORD_TABLE_CELL_CLOSE, cellTagEnd + 1);
      if (cellEnd < 0 || cellEnd > rowEnd) return null;
      let cellParagraphCount = 0;
      let paragraphCursor = cellTagEnd + 1;
      while (paragraphCursor < cellEnd) {
        const paragraphStart = findWordStartTag(
          source,
          'p',
          paragraphCursor,
          cellEnd,
        );
        if (paragraphStart < 0) break;
        const paragraphTagEnd = source.indexOf('>', paragraphStart);
        if (paragraphTagEnd < 0 || paragraphTagEnd >= cellEnd) return null;
        const paragraphEnd = source.indexOf(
          WORD_PARAGRAPH_CLOSE,
          paragraphTagEnd + 1,
        );
        if (paragraphEnd < 0 || paragraphEnd > cellEnd) return null;
        texts.push(simpleText(source, paragraphTagEnd + 1, paragraphEnd));
        cellParagraphCount += 1;
        paragraphCount += 1;
        paragraphCursor = paragraphEnd + WORD_PARAGRAPH_CLOSE.length;
      }
      if (!cellParagraphCount) {
        texts.push('');
        cellParagraphCount = 1;
        paragraphCount += 1;
      }
      cellParagraphCounts.push(cellParagraphCount);
      cellCursor = cellEnd + WORD_TABLE_CELL_CLOSE.length;
    }
    if (!cellParagraphCounts.length) return null;
    visit({ cellParagraphCounts, texts });
    rowCount += 1;
    rowCursor = rowEnd + WORD_TABLE_ROW_CLOSE.length;
  }
  return rowCount ? { paragraphCount, rowCount } : null;
}

function simpleText(source: string, start: number, end: number): string {
  let result = '';
  let cursor = start;
  while (cursor < end) {
    const textStart = findWordStartTag(source, 't', cursor, end);
    if (textStart < 0) break;
    const tagEnd = source.indexOf('>', textStart);
    if (tagEnd < 0 || tagEnd >= end) return '';
    const textEnd = source.indexOf(WORD_TEXT_CLOSE, tagEnd + 1);
    if (textEnd < 0 || textEnd > end) return '';
    result += decodeXmlText(source.slice(tagEnd + 1, textEnd));
    cursor = textEnd + WORD_TEXT_CLOSE.length;
  }
  return result;
}

function simpleDocumentLayout(
  source: string,
  bodyStart = 0,
  bodyEnd = source.length,
): WorkDocumentSectionLayout {
  const base = initialSimpleDocumentLayout();
  const sectionCandidate = source.lastIndexOf('<w:sectPr', bodyEnd);
  const sectionStart =
    sectionCandidate >= bodyStart &&
    wordTagNameBoundary(
      source.charCodeAt(sectionCandidate + '<w:sectPr'.length),
    )
      ? sectionCandidate
      : bodyStart;
  const pageSizeTag = startTag(source, 'pgSz', sectionStart, bodyEnd);
  const width = wordIntegerAttribute(pageSizeTag, 'w');
  const height = wordIntegerAttribute(pageSizeTag, 'h');
  const orientation: WorkDocumentSectionLayout['orientation'] =
    wordAttribute(pageSizeTag, 'orient') === 'landscape'
      ? 'landscape'
      : 'portrait';
  const geometry: WorkDocumentPageGeometry | null =
    width && height
      ? {
          width,
          height,
          orientation,
        }
      : null;
  const pageMarginsTag = startTag(source, 'pgMar', sectionStart, bodyEnd);
  const margins = {
    top: wordMarginMillimeters(pageMarginsTag, 'top', base.margins.top),
    right: wordMarginMillimeters(pageMarginsTag, 'right', base.margins.right),
    bottom: wordMarginMillimeters(
      pageMarginsTag,
      'bottom',
      base.margins.bottom,
    ),
    left: wordMarginMillimeters(pageMarginsTag, 'left', base.margins.left),
  };
  const layout: WorkDocumentSectionLayout = {
    ...base,
    margins,
    orientation,
    ...(geometry
      ? {
          pageSize: documentPaperSizeForGeometry(geometry),
          pageGeometry: geometry,
        }
      : {}),
  };
  return geometry ? applyDocumentPageGeometry(layout, geometry) : layout;
}

function initialSimpleDocumentLayout(): WorkDocumentSectionLayout {
  return {
    pageSize: 'a4',
    orientation: 'portrait',
    margins: { top: 25, right: 23, bottom: 25, left: 23 },
    columns: { count: 1, spacing: 12, separator: false },
    breakAfter: 'nextPage',
    headerText: undefined,
    footerText: undefined,
    showPageNumbers: false,
    pageNumberStart: undefined,
    pageChrome: {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: { headerHtml: '', footerHtml: '', showPageNumber: false },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    },
  };
}

function elementContentRange(
  source: string,
  localName: string,
): { end: number; start: number } | null {
  const start = findWordStartTag(source, localName, 0);
  if (start < 0) return null;
  const tagEnd = source.indexOf('>', start);
  if (tagEnd < 0) return null;
  const end = source.indexOf(`</w:${localName}>`, tagEnd + 1);
  return end < 0 ? null : { end, start: tagEnd + 1 };
}

function startTag(
  source: string,
  localName: string,
  searchStart = 0,
  searchEnd = source.length,
): string {
  const start = findWordStartTag(source, localName, searchStart, searchEnd);
  if (start < 0) return '';
  const end = source.indexOf('>', start);
  return end < 0 || end >= searchEnd ? '' : source.slice(start, end + 1);
}

function wordAttribute(source: string, localName: string): string | null {
  const escaped = localName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (
    new RegExp(`(?:[A-Za-z_][\\w.-]*:)?${escaped}="([^"]*)"`).exec(
      source,
    )?.[1] ?? null
  );
}

function wordIntegerAttribute(
  source: string,
  localName: string,
): number | null {
  const value = Number(wordAttribute(source, localName));
  return Number.isSafeInteger(value) && value > 0 ? value : null;
}

function wordMarginMillimeters(
  source: string,
  name: string,
  fallback: number,
): number {
  const value = wordAttribute(source, name);
  if (value === null) return fallback;
  const twips = Number(value);
  return Number.isFinite(twips) && twips >= 0
    ? Math.round((twips / WORD_TWIPS_PER_MILLIMETER) * 100) / 100
    : fallback;
}

function simpleDocxElementCounts(
  source: string,
  requiredBlocks: number,
): {
  body: number;
  document: number;
  paragraph: number;
  row: number;
  sectPr: number;
} | null {
  if (UNSUPPORTED_SIMPLE_WORD_ELEMENT.test(source)) return null;
  const row = countWordStartTags(source, 'tr');
  const envelope = simpleDocxEnvelopeElementCounts(source);
  return {
    body: envelope.body,
    document: envelope.document,
    paragraph: row >= requiredBlocks ? 0 : countWordStartTags(source, 'p'),
    row,
    sectPr: envelope.sectPr,
  };
}

function simpleDocxEnvelopeElementCounts(source: string): {
  body: number;
  document: number;
  sectPr: number;
} {
  let body = 0;
  let document = 0;
  let sectPr = 0;
  for (const match of source.matchAll(SIMPLE_WORD_ENVELOPE_ELEMENT)) {
    if (match[1] === 'body') body += 1;
    else if (match[1] === 'document') document += 1;
    else sectPr += 1;
  }
  return { body, document, sectPr };
}

/**
 * Finds an exact WordprocessingML start tag without allocating a RegExp match
 * object or mistaking `p` for `pgMar` and `tbl` for `tblPr`.
 */
function findWordStartTag(
  source: string,
  localName: string,
  start: number,
  end = source.length,
): number {
  const prefix = `<w:${localName}`;
  let index = source.indexOf(prefix, start);
  while (index >= 0 && index < end) {
    if (wordTagNameBoundary(source.charCodeAt(index + prefix.length))) {
      return index;
    }
    index = source.indexOf(prefix, index + prefix.length);
  }
  return -1;
}

function firstIndex(left: number, right: number): number {
  if (left < 0) return right;
  if (right < 0) return left;
  return Math.min(left, right);
}

function countWordStartTags(source: string, localName: string): number {
  const prefix = `<w:${localName}`;
  let count = 0;
  let index = source.indexOf(prefix);
  while (index >= 0) {
    if (wordTagNameBoundary(source.charCodeAt(index + prefix.length))) {
      count += 1;
    }
    index = source.indexOf(prefix, index + prefix.length);
  }
  return count;
}

function wordTagNameBoundary(code: number): boolean {
  return (
    code === 32 ||
    code === 9 ||
    code === 10 ||
    code === 13 ||
    code === 47 ||
    code === 62
  );
}

function decodeXmlText(value: string): string {
  if (!value.includes('&')) return value;
  return value.replace(
    /&(?:#(\d+)|#x([\da-f]+)|(amp|apos|gt|lt|quot));/gi,
    (entity, decimal: string, hexadecimal: string, named: string) => {
      if (decimal) return safeCodePoint(Number.parseInt(decimal, 10), entity);
      if (hexadecimal) {
        return safeCodePoint(Number.parseInt(hexadecimal, 16), entity);
      }
      return (
        {
          amp: '&',
          apos: "'",
          gt: '>',
          lt: '<',
          quot: '"',
        }[named.toLowerCase()] ?? entity
      );
    },
  );
}

function safeCodePoint(value: number, fallback: string): string {
  if (!Number.isSafeInteger(value) || value < 0 || value > 0x10ffff) {
    return fallback;
  }
  try {
    return String.fromCodePoint(value);
  } catch {
    return fallback;
  }
}

function positiveInteger(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0 ? number : fallback;
}

function largeDocxImportNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function recordLargeDocxImportMeasure(
  name: string,
  start: number,
  end: number,
  detail?: Record<string, number | boolean>,
): void {
  try {
    globalThis.performance?.measure(name, { detail, end, start });
  } catch {
    // User Timing is diagnostic only and must never affect file import.
  }
}
