import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';
import {
  type DocumentTableBorderEdge,
  documentTableBordersFromElement,
  renderDocumentTableBorders,
} from './work-document-table-borders';

export type ImportedDocxTableCellVerticalAlign = 'top' | 'middle' | 'bottom';
export type ImportedDocxTableCellBorderStyle =
  | 'solid'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'none';

export interface ImportedDocxTableCellMarker {
  marker: string;
  backgroundColor?: string;
  verticalAlign?: ImportedDocxTableCellVerticalAlign;
  borders?: ImportedDocxTableCellBorders;
}

export interface ImportedDocxTableCellMarkers {
  cells: ImportedDocxTableCellMarker[];
}

interface ImportedDocxCellBorder {
  color: string;
  style: ImportedDocxTableCellBorderStyle;
  width: number;
}

type ImportedDocxTableCellBorders = Partial<
  Record<DocumentTableBorderEdge, ImportedDocxCellBorder>
>;

const TABLE_BORDER_EDGES: readonly DocumentTableBorderEdge[] = [
  'top',
  'right',
  'bottom',
  'left',
];

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TABLE_CELL_MARKER_PATTERN = /__A3S_WORK_TABLE_CELL_\d+__/g;

export function markDocxTableCells(
  document: Document,
): ImportedDocxTableCellMarkers {
  const cells: ImportedDocxTableCellMarker[] = [];
  for (const cell of descendants(document, 'tc')) {
    const properties = directChild(cell, 'tcPr');
    const shading = properties ? directChild(properties, 'shd') : undefined;
    const verticalAlignment = properties
      ? directChild(properties, 'vAlign')
      : undefined;
    const borders = importedTableCellBorders(cell, properties);
    const backgroundColor = ooxmlColor(attribute(shading ?? cell, 'fill'));
    const verticalAlign = tableVerticalAlign(
      attribute(verticalAlignment ?? cell, 'val'),
    );
    if (!backgroundColor && !verticalAlign && !borders) continue;
    const paragraph = firstTableCellParagraph(document, cell);
    if (!paragraph) continue;
    const marker = `__A3S_WORK_TABLE_CELL_${cells.length + 1}__`;
    insertCellMarker(document, paragraph, marker);
    cells.push({
      marker,
      ...(backgroundColor ? { backgroundColor } : {}),
      ...(verticalAlign ? { verticalAlign } : {}),
      ...(borders ? { borders } : {}),
    });
  }
  return { cells };
}

export function applyImportedDocxTableCellMarkers(
  document: Document,
  markers: ImportedDocxTableCellMarkers,
): void {
  const cellByMarker = new Map(
    markers.cells.map((cell) => [cell.marker, cell]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_TABLE_CELL_')) continue;
    const cell = node.parentElement?.closest('td, th');
    node.data = node.data.replace(TABLE_CELL_MARKER_PATTERN, (marker) => {
      const format = cellByMarker.get(marker);
      if (cell instanceof HTMLElement && format) applyCellFormat(cell, format);
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxTableCellMarkers(
  markers: ImportedDocxTableCellMarkers,
): boolean {
  return markers.cells.length > 0;
}

function applyCellFormat(
  cell: HTMLElement,
  format: ImportedDocxTableCellMarker,
): void {
  if (format.backgroundColor) {
    cell.dataset.officeCellFill = format.backgroundColor;
    cell.style.backgroundColor = format.backgroundColor;
  }
  if (format.verticalAlign) {
    cell.dataset.officeCellVerticalAlign = format.verticalAlign;
    cell.style.verticalAlign = format.verticalAlign;
  }
  if (format.borders) {
    const borders = documentTableBordersFromElement(cell, {
      color: '#cfd5df',
      style: 'solid',
      width: 1,
    });
    for (const edge of TABLE_BORDER_EDGES) {
      const border = format.borders[edge];
      if (border) borders[edge] = { ...border };
    }
    const rendered = renderDocumentTableBorders(borders);
    for (const [name, value] of Object.entries(rendered)) {
      if (name !== 'style') cell.setAttribute(name, value);
    }
    cell.style.cssText = `${cell.style.cssText}; ${rendered.style}`;
  }
}

function importedTableCellBorders(
  cell: Element,
  properties: Element | undefined,
): ImportedDocxTableCellBorders | null {
  const tableBorders = tableBordersForCell(cell);
  const cellBorders = borderEdges(
    properties ? directChild(properties, 'tcBorders') : undefined,
  );
  const borders: ImportedDocxTableCellBorders = {};
  for (const edge of TABLE_BORDER_EDGES) {
    const border = cellBorders[edge] ?? tableBorders[edge];
    if (border) borders[edge] = border;
  }
  return Object.keys(borders).length ? borders : null;
}

function tableBordersForCell(cell: Element): ImportedDocxTableCellBorders {
  const table = closestAncestor(cell, 'tbl');
  const properties = table ? directChild(table, 'tblPr') : undefined;
  const borders = properties
    ? borderElementMap(directChild(properties, 'tblBorders'))
    : new Map<string, Element>();
  if (!table || !borders.size) return {};

  const row = closestAncestor(cell, 'tr');
  if (!row) return {};
  const rows = descendants(table, 'tr').filter(
    (candidate) => closestAncestor(candidate, 'tbl') === table,
  );
  const cells = descendants(row, 'tc').filter(
    (candidate) => closestAncestor(candidate, 'tr') === row,
  );
  const rowIndex = rows.indexOf(row);
  const cellIndex = cells.indexOf(cell);
  if (rowIndex < 0 || cellIndex < 0) return {};

  return compactBorders({
    top: parseBorder(borders.get(rowIndex === 0 ? 'top' : 'insideH')),
    right: parseBorder(
      borderElement(
        borders,
        cellIndex === cells.length - 1 ? ['right', 'end'] : ['insideV'],
      ),
    ),
    bottom: parseBorder(
      borders.get(rowIndex === rows.length - 1 ? 'bottom' : 'insideH'),
    ),
    left: parseBorder(
      borderElement(borders, cellIndex === 0 ? ['left', 'start'] : ['insideV']),
    ),
  });
}

function borderEdges(
  borders: Element | undefined,
): ImportedDocxTableCellBorders {
  const edges = borderElementMap(borders);
  return compactBorders({
    top: parseBorder(edges.get('top')),
    right: parseBorder(borderElement(edges, ['right', 'end'])),
    bottom: parseBorder(edges.get('bottom')),
    left: parseBorder(borderElement(edges, ['left', 'start'])),
  });
}

function borderElementMap(borders: Element | undefined): Map<string, Element> {
  return new Map(
    borders
      ? directChildren(borders).map((edge) => [edge.localName, edge])
      : [],
  );
}

function borderElement(
  borders: ReadonlyMap<string, Element>,
  names: readonly string[],
): Element | undefined {
  for (const name of names) {
    const edge = borders.get(name);
    if (edge) return edge;
  }
  return undefined;
}

function compactBorders(
  borders: Partial<
    Record<DocumentTableBorderEdge, ImportedDocxCellBorder | null>
  >,
): ImportedDocxTableCellBorders {
  const result: ImportedDocxTableCellBorders = {};
  for (const edge of TABLE_BORDER_EDGES) {
    const border = borders[edge];
    if (border) result[edge] = border;
  }
  return result;
}

function parseBorder(edge: Element | undefined): ImportedDocxCellBorder | null {
  if (!edge) return null;
  const style = tableBorderStyle(attribute(edge, 'val'));
  if (!style) return null;
  if (style === 'none') return { color: '#000000', style, width: 0 };
  const color = ooxmlColor(attribute(edge, 'color')) ?? '#000000';
  const size = Number(attribute(edge, 'sz'));
  const width = Number.isFinite(size) && size > 0 ? size / 6 : 1;
  return {
    color,
    style,
    width: Math.max(0.5, Math.min(6, Math.round(width * 2) / 2)),
  };
}

function tableBorderStyle(
  value: string | null,
): ImportedDocxTableCellBorderStyle | null {
  const normalized = value?.trim();
  if (!normalized) return null;
  if (normalized === 'nil' || normalized === 'none') return 'none';
  if (normalized === 'double' || normalized === 'triple') return 'double';
  if (normalized === 'dotted') return 'dotted';
  if (/dash|dotDash|dotDotDash/i.test(normalized)) return 'dashed';
  return 'solid';
}

function tableVerticalAlign(
  value: string | null,
): ImportedDocxTableCellVerticalAlign | null {
  if (value === 'center') return 'middle';
  if (value === 'top' || value === 'bottom') return value;
  return null;
}

function ooxmlColor(value: string | null): string | null {
  const normalized = value?.trim().replace(/^#/, '').toLowerCase();
  if (!normalized || normalized === 'auto') return null;
  if (/^[0-9a-f]{3}$/.test(normalized)) {
    return `#${normalized
      .split('')
      .map((part) => `${part}${part}`)
      .join('')}`;
  }
  return /^[0-9a-f]{6}$/.test(normalized) ? `#${normalized}` : null;
}

function firstTableCellParagraph(
  document: Document,
  cell: Element,
): Element | null {
  const existing =
    directChild(cell, 'p') ??
    descendants(cell, 'p').find(
      (paragraph) => closestAncestor(paragraph, 'tc') === cell,
    );
  if (existing) return existing;
  const paragraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  cell.append(paragraph);
  return paragraph;
}

function insertCellMarker(
  document: Document,
  paragraph: Element,
  marker: string,
): void {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChild(paragraph, 'pPr');
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function textNodes(root: ParentNode): Text[] {
  const document = root.ownerDocument;
  const walker = document?.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}
