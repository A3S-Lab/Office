import type {
  ContentControlLimits,
  StaticContentParagraphRecord,
} from './work-docx-note-comment-content-control-model';
import {
  hasContentControlRelationshipReference,
  hasUnsupportedContentControlSemanticChild,
  isDocxWordElement,
  wordDirectChildren,
} from './work-docx-note-comment-content-control-xml';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

const MAX_TABLE_DEPTH = 16;
const MAX_TABLES = 65_536;
const MAX_CELLS = 262_144;
const MAX_GRID_SPAN = 32_767;

type DocumentRole = 'generated' | 'source';

interface StaticContentStructureRecord {
  element: Element;
  key: string;
}

export interface StaticContentBlockRecord {
  element: Element;
  extensionScopes: StaticContentStructureRecord[];
  fingerprint: string;
  kind: 'paragraph' | 'table';
  paragraphs: StaticContentParagraphRecord[];
}

type ParagraphReader = (
  paragraph: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
) => StaticContentParagraphRecord | null;

interface ParsedCell {
  fingerprint: unknown;
  paragraphs: StaticContentParagraphRecord[];
  scopes: StaticContentStructureRecord[];
}

interface ParsedTable {
  fingerprint: unknown;
  paragraphs: StaticContentParagraphRecord[];
  scopes: StaticContentStructureRecord[];
}

export function paragraphContentBlock(
  paragraph: StaticContentParagraphRecord,
): StaticContentBlockRecord {
  return {
    element: paragraph.element,
    extensionScopes: [],
    fingerprint: JSON.stringify(['p', paragraph.text]),
    kind: 'paragraph',
    paragraphs: [paragraph],
  };
}

export function readStaticTableBlock(
  table: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
  readParagraph: ParagraphReader,
): StaticContentBlockRecord | null {
  if (
    !isWord(table, 'tbl') ||
    hasContentControlRelationshipReference(table) ||
    wordDescendants(table, 'sdt').length
  ) {
    return null;
  }
  const parsed = parseTable(table, role, limits, readParagraph, 0, 'table');
  return parsed
    ? {
        element: table,
        extensionScopes: parsed.scopes,
        fingerprint: JSON.stringify(parsed.fingerprint),
        kind: 'table',
        paragraphs: parsed.paragraphs,
      }
    : null;
}

function parseTable(
  table: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
  readParagraph: ParagraphReader,
  depth: number,
  path: string,
): ParsedTable | null {
  if (depth > MAX_TABLE_DEPTH || !isWord(table, 'tbl')) return null;
  incrementLimit(limits, role, 'Tables', MAX_TABLES, 'table');
  const children = semanticChildren(table);
  if (!children) return null;
  const properties = singleLeadingChild(children, 'tblPr');
  const withoutProperties = properties ? children.slice(1) : children;
  const grid = singleLeadingChild(withoutProperties, 'tblGrid');
  const rows = grid ? withoutProperties.slice(1) : withoutProperties;
  if (!rows.length || rows.some((row) => !isWord(row, 'tr'))) return null;
  if (grid && !validGrid(grid)) return null;

  const paragraphs: StaticContentParagraphRecord[] = [];
  const scopes: StaticContentStructureRecord[] = [
    { element: table, key: path },
    ...(properties ? [{ element: properties, key: `${path}/properties` }] : []),
  ];
  const rowFingerprints: unknown[] = [];
  for (const [rowIndex, row] of rows.entries()) {
    const parsed = parseRow(
      row,
      role,
      limits,
      readParagraph,
      depth,
      `${path}/row:${rowIndex}`,
    );
    if (!parsed) return null;
    rowFingerprints.push(parsed.fingerprint);
    paragraphs.push(...parsed.paragraphs);
    scopes.push(...parsed.scopes);
  }
  return {
    fingerprint: ['table', rowFingerprints],
    paragraphs,
    scopes,
  };
}

function parseRow(
  row: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
  readParagraph: ParagraphReader,
  depth: number,
  path: string,
): ParsedTable | null {
  const children = semanticChildren(row);
  if (!children) return null;
  const properties = singleLeadingChild(children, 'trPr');
  const cells = properties ? children.slice(1) : children;
  if (!cells.length || cells.some((cell) => !isWord(cell, 'tc'))) return null;
  const paragraphs: StaticContentParagraphRecord[] = [];
  const scopes: StaticContentStructureRecord[] = [
    { element: row, key: path },
    ...(properties ? [{ element: properties, key: `${path}/properties` }] : []),
  ];
  const gridOffset = rowGridOffset(properties);
  if (!gridOffset) return null;
  const cellFingerprints: unknown[] = [];
  for (const [cellIndex, cell] of cells.entries()) {
    const parsed = parseCell(
      cell,
      role,
      limits,
      readParagraph,
      depth,
      `${path}/cell:${cellIndex}`,
    );
    if (!parsed) return null;
    cellFingerprints.push(parsed.fingerprint);
    paragraphs.push(...parsed.paragraphs);
    scopes.push(...parsed.scopes);
  }
  return {
    fingerprint: ['row', gridOffset, cellFingerprints],
    paragraphs,
    scopes,
  };
}

function parseCell(
  cell: Element,
  role: DocumentRole,
  limits: ContentControlLimits,
  readParagraph: ParagraphReader,
  depth: number,
  path: string,
): ParsedCell | null {
  incrementLimit(limits, role, 'Cells', MAX_CELLS, 'table-cell');
  const children = semanticChildren(cell);
  if (!children) return null;
  const properties = singleLeadingChild(children, 'tcPr');
  const blocks = properties ? children.slice(1) : children;
  if (
    !blocks.length ||
    blocks.some((block) => !isWord(block, 'p') && !isWord(block, 'tbl'))
  ) {
    return null;
  }
  const paragraphs: StaticContentParagraphRecord[] = [];
  const scopes: StaticContentStructureRecord[] = [
    { element: cell, key: path },
    ...(properties ? [{ element: properties, key: `${path}/properties` }] : []),
  ];
  const merge = cellMergeFingerprint(properties);
  if (!merge) return null;
  const blockFingerprints: unknown[] = [];
  for (const [blockIndex, block] of blocks.entries()) {
    if (block.localName === 'p') {
      const paragraph = readParagraph(block, role, limits);
      if (!paragraph) return null;
      paragraphs.push(paragraph);
      blockFingerprints.push(['p', paragraph.text]);
      continue;
    }
    const nested = parseTable(
      block,
      role,
      limits,
      readParagraph,
      depth + 1,
      `${path}/block:${blockIndex}/table`,
    );
    if (!nested) return null;
    paragraphs.push(...nested.paragraphs);
    scopes.push(...nested.scopes);
    blockFingerprints.push(nested.fingerprint);
  }
  return {
    fingerprint: ['cell', merge, blockFingerprints],
    paragraphs,
    scopes,
  };
}

function semanticChildren(element: Element): Element[] | null {
  if (hasUnsupportedContentControlSemanticChild(element)) return null;
  return Array.from(element.children).filter(isDocxWordElement);
}

function singleLeadingChild(
  children: readonly Element[],
  localName: string,
): Element | null {
  const matches = children.filter((child) => child.localName === localName);
  if (
    matches.length > 1 ||
    (matches.length === 1 && children[0] !== matches[0])
  ) {
    return null;
  }
  return matches[0] ?? null;
}

function validGrid(grid: Element): boolean {
  const children = semanticChildren(grid);
  return Boolean(children?.every((child) => child.localName === 'gridCol'));
}

function rowGridOffset(properties: Element | null): [number, number] | null {
  const before = unsignedWordValue(properties, 'gridBefore', 0, 0);
  const after = unsignedWordValue(properties, 'gridAfter', 0, 0);
  return before === null || after === null ? null : [before, after];
}

function cellMergeFingerprint(
  properties: Element | null,
): [number, string, string] | null {
  const span = unsignedWordValue(properties, 'gridSpan', 1, 1);
  const horizontal = mergeValue(properties, 'hMerge');
  const vertical = mergeValue(properties, 'vMerge');
  return span === null || !horizontal || !vertical
    ? null
    : [span, horizontal, vertical];
}

function unsignedWordValue(
  properties: Element | null,
  localName: string,
  fallback: number,
  minimum: number,
): number | null {
  if (!properties) return fallback;
  const elements = wordDirectChildren(properties, localName);
  if (!elements.length) return fallback;
  if (elements.length !== 1) return null;
  const value = wordAttribute(elements[0], 'val');
  if (!value || !/^[0-9]{1,5}$/u.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) &&
    number >= minimum &&
    number <= MAX_GRID_SPAN
    ? number
    : null;
}

function mergeValue(
  properties: Element | null,
  localName: string,
): string | null {
  if (!properties) return 'none';
  const elements = wordDirectChildren(properties, localName);
  if (!elements.length) return 'none';
  if (elements.length !== 1) return null;
  const value = wordAttribute(elements[0], 'val')?.trim().toLowerCase();
  return !value || value === 'continue'
    ? 'continue'
    : value === 'restart'
      ? 'restart'
      : null;
}

function wordAttribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        isWordNamespace(xmlAttributeNamespace(element, item)),
    )?.value ?? null
  );
}

function wordDescendants(element: Element, localName: string): Element[] {
  return Array.from(element.querySelectorAll('*')).filter((child) =>
    isWord(child, localName),
  );
}

function isWord(element: Element, localName: string): boolean {
  return element.localName === localName && isDocxWordElement(element);
}

function isWordNamespace(namespace: string | null): boolean {
  return Boolean(
    namespace &&
      (namespace ===
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main' ||
        namespace === 'http://purl.oclc.org/ooxml/wordprocessingml/main'),
  );
}

function incrementLimit(
  limits: ContentControlLimits,
  role: DocumentRole,
  suffix: 'Cells' | 'Tables',
  maximum: number,
  label: string,
): void {
  const key = `${role}${suffix}` as keyof ContentControlLimits;
  limits[key] += 1;
  if (limits[key] > maximum) {
    throw new Error(
      `${role === 'source' ? 'Registered source' : 'Generated'} DOCX exceeds the stable note/comment ${label} limit.`,
    );
  }
}
