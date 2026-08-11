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
import {
  type DocxThemeResolver,
  type DocxThemeSource,
  docxThemeColor,
  resolveDocxThemeResolver,
} from './work-docx-theme';
import {
  type DocxTableStyleLayer,
  type DocxTableStyleSource,
  docxTableCellStyleLayers,
  resolveDocxTableStyleResolver,
} from './work-docx-table-styles';
import {
  renderDocumentTableCellMarginOverrides,
  type DocumentTableCellMarginOverrides,
  type DocumentTableCellMarginSide,
} from './work-document-table-geometry';
import {
  type DocxThemeColorReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';

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
  themeFill?: DocxThemeColorReference;
  verticalAlign?: ImportedDocxTableCellVerticalAlign;
  borders?: ImportedDocxTableCellBorders;
  margins?: DocumentTableCellMarginOverrides;
}

export interface ImportedDocxTableCellMarkers {
  cells: ImportedDocxTableCellMarker[];
}

interface ImportedDocxCellBorder {
  color: string;
  style: ImportedDocxTableCellBorderStyle;
  width: number;
  theme?: DocxThemeColorReference;
}

interface ImportedDocxCellColor {
  color: string;
  theme?: DocxThemeColorReference;
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
const PIXELS_PER_TWIP = 96 / 1440;

export function markDocxTableCells(
  document: Document,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxTableCellMarkers {
  const cells: ImportedDocxTableCellMarker[] = [];
  const theme = resolveDocxThemeResolver(themeSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  for (const cell of descendants(document, 'tc')) {
    const properties = directChild(cell, 'tcPr');
    const table = closestAncestor(cell, 'tbl');
    const directTableProperties = table
      ? directChild(table, 'tblPr')
      : undefined;
    const layers: DocxTableStyleLayer[] = [
      ...docxTableCellStyleLayers(cell, tableStyles),
      ...(directTableProperties
        ? [{ tableProperties: directTableProperties }]
        : []),
      ...(properties ? [{ cellProperties: properties }] : []),
    ];
    const borders = importedTableCellBorders(cell, layers, theme);
    const background = importedTableCellBackgroundColor(layers, theme);
    const verticalAlign = importedTableCellVerticalAlign(layers);
    const margins = importedTableCellMargins(layers);
    if (!background && !verticalAlign && !borders && !margins) continue;
    const paragraph = firstTableCellParagraph(document, cell);
    if (!paragraph) continue;
    const marker = `__A3S_WORK_TABLE_CELL_${cells.length + 1}__`;
    insertCellMarker(document, paragraph, marker);
    cells.push({
      marker,
      ...(background ? { backgroundColor: background.color } : {}),
      ...(background?.theme ? { themeFill: background.theme } : {}),
      ...(verticalAlign ? { verticalAlign } : {}),
      ...(borders ? { borders } : {}),
      ...(margins ? { margins } : {}),
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
  const themeFill = serializeDocxThemeReference(format.themeFill ?? null);
  if (themeFill) cell.dataset.officeCellThemeFill = themeFill;
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
      const theme = serializeDocxThemeReference(border?.theme ?? null);
      if (theme)
        cell.dataset[`officeCellBorderTheme${capitalize(edge)}`] = theme;
    }
    const rendered = renderDocumentTableBorders(borders);
    for (const [name, value] of Object.entries(rendered)) {
      if (name !== 'style') cell.setAttribute(name, value);
    }
    cell.style.cssText = `${cell.style.cssText}; ${rendered.style}`;
  }
  if (format.margins) {
    const rendered = renderDocumentTableCellMarginOverrides(format.margins);
    for (const [name, value] of Object.entries(rendered)) {
      if (name !== 'style') cell.setAttribute(name, value);
    }
    if (rendered.style) {
      cell.style.cssText = `${cell.style.cssText}; ${rendered.style}`;
    }
  }
}

function importedTableCellMargins(
  layers: readonly DocxTableStyleLayer[],
): DocumentTableCellMarginOverrides | null {
  const margins: DocumentTableCellMarginOverrides = {};
  for (const layer of layers) {
    const marginElement = layer.cellProperties
      ? directChild(layer.cellProperties, 'tcMar')
      : undefined;
    if (!marginElement) continue;
    const children = new Map(
      directChildren(marginElement).map((margin) => [margin.localName, margin]),
    );
    assignCellMargin(margins, 'top', children.get('top'));
    assignCellMargin(
      margins,
      'right',
      children.get('right') ?? children.get('end'),
    );
    assignCellMargin(margins, 'bottom', children.get('bottom'));
    assignCellMargin(
      margins,
      'left',
      children.get('left') ?? children.get('start'),
    );
  }
  return Object.keys(margins).length ? margins : null;
}

function assignCellMargin(
  target: DocumentTableCellMarginOverrides,
  side: DocumentTableCellMarginSide,
  element: Element | undefined,
): void {
  if (!element) return;
  const type = attribute(element, 'type');
  if (type && type !== 'dxa') return;
  const twips = Number(attribute(element, 'w'));
  if (!Number.isFinite(twips) || twips < 0) return;
  target[side] = Math.round(twips * PIXELS_PER_TWIP * 100) / 100;
}

function importedTableCellBorders(
  cell: Element,
  layers: readonly DocxTableStyleLayer[],
  theme: DocxThemeResolver,
): ImportedDocxTableCellBorders | null {
  const borders: ImportedDocxTableCellBorders = {};
  for (const layer of layers) {
    if (layer.tableProperties) {
      mergeBorders(
        borders,
        tableBordersForCell(cell, layer.tableProperties, theme),
      );
    }
    if (layer.cellProperties) {
      mergeBorders(
        borders,
        borderEdges(directChild(layer.cellProperties, 'tcBorders'), theme),
      );
    }
  }
  return Object.keys(borders).length ? borders : null;
}

function tableBordersForCell(
  cell: Element,
  properties: Element,
  theme: DocxThemeResolver,
): ImportedDocxTableCellBorders {
  const table = closestAncestor(cell, 'tbl');
  const borders = borderElementMap(directChild(properties, 'tblBorders'));
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
    top: parseBorder(borders.get(rowIndex === 0 ? 'top' : 'insideH'), theme),
    right: parseBorder(
      borderElement(
        borders,
        cellIndex === cells.length - 1 ? ['right', 'end'] : ['insideV'],
      ),
      theme,
    ),
    bottom: parseBorder(
      borders.get(rowIndex === rows.length - 1 ? 'bottom' : 'insideH'),
      theme,
    ),
    left: parseBorder(
      borderElement(borders, cellIndex === 0 ? ['left', 'start'] : ['insideV']),
      theme,
    ),
  });
}

function mergeBorders(
  target: ImportedDocxTableCellBorders,
  source: ImportedDocxTableCellBorders,
): void {
  for (const edge of TABLE_BORDER_EDGES) {
    const border = source[edge];
    if (border) target[edge] = border;
  }
}

function borderEdges(
  borders: Element | undefined,
  theme: DocxThemeResolver,
): ImportedDocxTableCellBorders {
  const edges = borderElementMap(borders);
  return compactBorders({
    top: parseBorder(edges.get('top'), theme),
    right: parseBorder(borderElement(edges, ['right', 'end']), theme),
    bottom: parseBorder(edges.get('bottom'), theme),
    left: parseBorder(borderElement(edges, ['left', 'start']), theme),
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

function parseBorder(
  edge: Element | undefined,
  theme: DocxThemeResolver,
): ImportedDocxCellBorder | null {
  if (!edge) return null;
  const style = tableBorderStyle(attribute(edge, 'val'));
  if (!style) return null;
  if (style === 'none') return { color: '#000000', style, width: 0 };
  const themed = docxThemeColor(
    theme,
    attribute(edge, 'themeColor'),
    attribute(edge, 'themeTint'),
    attribute(edge, 'themeShade'),
  );
  const color = themed
    ? `#${themed}`
    : (ooxmlColor(attribute(edge, 'color')) ?? '#000000');
  const size = Number(attribute(edge, 'sz'));
  const width = Number.isFinite(size) && size > 0 ? size / 6 : 1;
  return {
    color,
    style,
    width: Math.max(0.5, Math.min(6, Math.round(width * 2) / 2)),
    ...(themed
      ? {
          theme: themeReference(
            edge,
            'themeColor',
            'themeTint',
            'themeShade',
            color,
          ),
        }
      : {}),
  };
}

function importedTableCellBackgroundColor(
  layers: readonly DocxTableStyleLayer[],
  theme: DocxThemeResolver,
): ImportedDocxCellColor | null {
  let color: ImportedDocxCellColor | null = null;
  for (const layer of layers) {
    const tableShading = layer.tableProperties
      ? directChild(layer.tableProperties, 'shd')
      : undefined;
    const cellShading = layer.cellProperties
      ? directChild(layer.cellProperties, 'shd')
      : undefined;
    if (tableShading) color = tableCellBackgroundColor(tableShading, theme);
    if (cellShading) color = tableCellBackgroundColor(cellShading, theme);
  }
  return color;
}

function tableCellBackgroundColor(
  shading: Element,
  theme: DocxThemeResolver,
): ImportedDocxCellColor | null {
  const themed = docxThemeColor(
    theme,
    attribute(shading, 'themeFill'),
    attribute(shading, 'themeFillTint'),
    attribute(shading, 'themeFillShade'),
  );
  if (themed) {
    const color = `#${themed}`;
    return {
      color,
      theme: themeReference(
        shading,
        'themeFill',
        'themeFillTint',
        'themeFillShade',
        color,
      ),
    };
  }
  const color = ooxmlColor(attribute(shading, 'fill'));
  return color ? { color } : null;
}

function themeReference(
  element: Element,
  themeName: string,
  tintName: string,
  shadeName: string,
  resolved: string,
): DocxThemeColorReference {
  const theme = attribute(element, themeName)?.trim() ?? '';
  const tint = normalizedByteHex(attribute(element, tintName));
  const shade = normalizedByteHex(attribute(element, shadeName));
  return {
    theme,
    resolved,
    ...(tint ? { tint } : {}),
    ...(shade ? { shade } : {}),
  };
}

function normalizedByteHex(value: string | null): string | undefined {
  const normalized = value?.trim().toUpperCase();
  return normalized && /^[0-9A-F]{2}$/.test(normalized)
    ? normalized
    : undefined;
}

function importedTableCellVerticalAlign(
  layers: readonly DocxTableStyleLayer[],
): ImportedDocxTableCellVerticalAlign | null {
  let alignment: ImportedDocxTableCellVerticalAlign | null = null;
  for (const layer of layers) {
    const element = layer.cellProperties
      ? directChild(layer.cellProperties, 'vAlign')
      : undefined;
    if (!element) continue;
    const value = tableVerticalAlign(attribute(element, 'val'));
    if (value) alignment = value;
  }
  return alignment;
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

function capitalize(value: string): string {
  return `${value.slice(0, 1).toUpperCase()}${value.slice(1)}`;
}
