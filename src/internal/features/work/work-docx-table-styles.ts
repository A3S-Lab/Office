import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export interface DocxTableStyleLayer {
  tableProperties?: Element;
  cellProperties?: Element;
  paragraphProperties?: Element;
  runProperties?: Element;
}

interface DocxTableStyle extends DocxTableStyleLayer {
  basedOn?: string;
  conditional: ReadonlyMap<string, DocxTableStyleLayer>;
}

export interface DocxTableStyleResolver {
  defaultStyleId?: string;
  tableStyles: ReadonlyMap<string, DocxTableStyle>;
}

export type DocxTableStyleSource =
  | Document
  | DocxTableStyleResolver
  | null
  | undefined;

interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHorizontalBand: boolean;
  noVerticalBand: boolean;
}

interface TableCellPosition {
  rowIndex: number;
  rowCount: number;
  columnStart: number;
  columnEnd: number;
  columnCount: number;
}

const MAX_STYLE_INHERITANCE_DEPTH = 64;
const DEFAULT_TABLE_LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  noHorizontalBand: false,
  noVerticalBand: true,
};
const TABLE_LOOK_BITS = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHorizontalBand: 0x0200,
  noVerticalBand: 0x0400,
} as const;

export function createDocxTableStyleResolver(
  stylesDocument?: Document | null,
): DocxTableStyleResolver {
  if (!stylesDocument) return { tableStyles: new Map() };

  const tableStyles = new Map<string, DocxTableStyle>();
  let defaultStyleId: string | undefined;
  for (const style of descendants(stylesDocument, 'style')) {
    if (wordAttribute(style, 'type') !== 'table') continue;
    const styleId = wordAttribute(style, 'styleId')?.trim();
    if (!styleId) continue;
    const basedOn = directChild(style, 'basedOn');
    const conditional = new Map<string, DocxTableStyleLayer>();
    for (const conditionalStyle of directChildren(style, 'tblStylePr')) {
      const type = wordAttribute(conditionalStyle, 'type')?.trim();
      if (type) conditional.set(type, styleLayer(conditionalStyle));
    }
    tableStyles.set(styleId, {
      ...styleLayer(style),
      basedOn: basedOn
        ? wordAttribute(basedOn, 'val')?.trim() || undefined
        : undefined,
      conditional,
    });
    if (!defaultStyleId && onOffAttribute(style, 'default')) {
      defaultStyleId = styleId;
    }
  }
  return { defaultStyleId, tableStyles };
}

export function resolveDocxTableStyleResolver(
  source: DocxTableStyleSource,
): DocxTableStyleResolver {
  return source && 'tableStyles' in source
    ? source
    : createDocxTableStyleResolver(source);
}

export function docxTablePropertySources(
  table: Element,
  resolver: DocxTableStyleResolver,
): readonly Element[] {
  const directTableProperties = directChild(table, 'tblPr');
  const styleReference = directTableProperties
    ? directChild(directTableProperties, 'tblStyle')
    : undefined;
  const styleId =
    (styleReference
      ? wordAttribute(styleReference, 'val')?.trim()
      : undefined) || resolver.defaultStyleId;
  const sources = styleId
    ? tableStyleChain(resolver.tableStyles, styleId)
        .map(({ tableProperties }) => tableProperties)
        .filter((properties): properties is Element => Boolean(properties))
    : [];
  if (directTableProperties) sources.push(directTableProperties);
  return sources;
}

export function docxTableCellStyleLayers(
  cell: Element,
  resolver: DocxTableStyleResolver,
): readonly DocxTableStyleLayer[] {
  const table = closestAncestor(cell, 'tbl');
  if (!table) return [];
  const directTableProperties = directChild(table, 'tblPr');
  const styleReference = directTableProperties
    ? directChild(directTableProperties, 'tblStyle')
    : undefined;
  const styleId =
    (styleReference
      ? wordAttribute(styleReference, 'val')?.trim()
      : undefined) || resolver.defaultStyleId;
  if (!styleId) return [];

  const styles = tableStyleChain(resolver.tableStyles, styleId);
  if (!styles.length) return [];
  const unconditionalLayers = styles.map(styleLayerFromStyle);
  const propertySources = unconditionalLayers
    .map(({ tableProperties }) => tableProperties)
    .filter((properties): properties is Element => Boolean(properties));
  if (directTableProperties) propertySources.push(directTableProperties);
  const position = tableCellPosition(table, cell);
  if (!position) return unconditionalLayers;
  const look = resolveTableLook(propertySources);
  const rowBandSize = resolveBandSize(propertySources, 'tblStyleRowBandSize');
  const columnBandSize = resolveBandSize(
    propertySources,
    'tblStyleColBandSize',
  );
  const layers = [...unconditionalLayers];
  for (const condition of applicableConditions(
    position,
    look,
    rowBandSize,
    columnBandSize,
  )) {
    for (const style of styles) {
      const layer = style.conditional.get(condition);
      if (layer) layers.push(layer);
    }
  }
  return layers;
}

export function docxTableRunPropertySources(
  element: Element,
  resolver: DocxTableStyleResolver,
): Element[] {
  const cell = closestAncestor(element, 'tc');
  if (!cell) return [];
  const sources: Element[] = [];
  for (const layer of docxTableCellStyleLayers(cell, resolver)) {
    const paragraphRunProperties = layer.paragraphProperties
      ? directChild(layer.paragraphProperties, 'rPr')
      : undefined;
    if (paragraphRunProperties) sources.push(paragraphRunProperties);
    if (layer.runProperties) sources.push(layer.runProperties);
  }
  return sources;
}

export function docxTableParagraphPropertySources(
  element: Element,
  resolver: DocxTableStyleResolver,
): Element[] {
  const cell = closestAncestor(element, 'tc');
  if (!cell) return [];
  return docxTableCellStyleLayers(cell, resolver)
    .map(({ paragraphProperties }) => paragraphProperties)
    .filter((properties): properties is Element => Boolean(properties));
}

function styleLayer(parent: ParentNode): DocxTableStyleLayer {
  return {
    tableProperties: directChild(parent, 'tblPr'),
    cellProperties: directChild(parent, 'tcPr'),
    paragraphProperties: directChild(parent, 'pPr'),
    runProperties: directChild(parent, 'rPr'),
  };
}

function styleLayerFromStyle(style: DocxTableStyle): DocxTableStyleLayer {
  return {
    tableProperties: style.tableProperties,
    cellProperties: style.cellProperties,
    paragraphProperties: style.paragraphProperties,
    runProperties: style.runProperties,
  };
}

function tableStyleChain(
  styles: ReadonlyMap<string, DocxTableStyle>,
  styleId: string,
): DocxTableStyle[] {
  const inherited: DocxTableStyle[] = [];
  const visited = new Set<string>();
  let currentStyleId: string | undefined = styleId;
  let depth = 0;
  while (
    currentStyleId &&
    depth < MAX_STYLE_INHERITANCE_DEPTH &&
    !visited.has(currentStyleId)
  ) {
    depth += 1;
    visited.add(currentStyleId);
    const style = styles.get(currentStyleId);
    if (!style) break;
    inherited.push(style);
    currentStyleId = style.basedOn;
  }
  return inherited.reverse();
}

function applicableConditions(
  position: TableCellPosition,
  look: TableLook,
  rowBandSize: number,
  columnBandSize: number,
): string[] {
  const conditions = ['wholeTable'];
  const isFirstRow = position.rowIndex === 0;
  const isLastRow = position.rowIndex === position.rowCount - 1;
  const isFirstColumn = position.columnStart === 0;
  const isLastColumn = position.columnEnd === position.columnCount - 1;

  if (!look.noVerticalBand) {
    const offset = look.firstColumn ? 1 : 0;
    const bandIndex = position.columnStart - offset;
    if (bandIndex >= 0) {
      conditions.push(
        Math.floor(bandIndex / columnBandSize) % 2 === 0
          ? 'band1Vert'
          : 'band2Vert',
      );
    }
  }
  if (!look.noHorizontalBand) {
    const offset = look.firstRow ? 1 : 0;
    const bandIndex = position.rowIndex - offset;
    if (bandIndex >= 0) {
      conditions.push(
        Math.floor(bandIndex / rowBandSize) % 2 === 0
          ? 'band1Horz'
          : 'band2Horz',
      );
    }
  }
  if (look.firstColumn && isFirstColumn) conditions.push('firstCol');
  if (look.lastColumn && isLastColumn) conditions.push('lastCol');
  if (look.firstRow && isFirstRow) conditions.push('firstRow');
  if (look.lastRow && isLastRow) conditions.push('lastRow');

  if (look.firstRow && isFirstRow && look.firstColumn && isFirstColumn) {
    conditions.push('nwCell');
  }
  if (look.firstRow && isFirstRow && look.lastColumn && isLastColumn) {
    conditions.push('neCell');
  }
  if (look.lastRow && isLastRow && look.firstColumn && isFirstColumn) {
    conditions.push('swCell');
  }
  if (look.lastRow && isLastRow && look.lastColumn && isLastColumn) {
    conditions.push('seCell');
  }
  return conditions;
}

function resolveTableLook(propertySources: readonly Element[]): TableLook {
  const look = { ...DEFAULT_TABLE_LOOK };
  for (const properties of propertySources) {
    const element = directChild(properties, 'tblLook');
    if (!element) continue;
    const encoded = wordAttribute(element, 'val')?.trim();
    if (encoded && /^[0-9a-f]{1,4}$/i.test(encoded)) {
      const value = Number.parseInt(encoded, 16);
      look.firstRow = Boolean(value & TABLE_LOOK_BITS.firstRow);
      look.lastRow = Boolean(value & TABLE_LOOK_BITS.lastRow);
      look.firstColumn = Boolean(value & TABLE_LOOK_BITS.firstColumn);
      look.lastColumn = Boolean(value & TABLE_LOOK_BITS.lastColumn);
      look.noHorizontalBand = Boolean(value & TABLE_LOOK_BITS.noHorizontalBand);
      look.noVerticalBand = Boolean(value & TABLE_LOOK_BITS.noVerticalBand);
    }
    assignLookFlag(look, 'firstRow', element, 'firstRow');
    assignLookFlag(look, 'lastRow', element, 'lastRow');
    assignLookFlag(look, 'firstColumn', element, 'firstColumn');
    assignLookFlag(look, 'lastColumn', element, 'lastColumn');
    assignLookFlag(look, 'noHorizontalBand', element, 'noHBand');
    assignLookFlag(look, 'noVerticalBand', element, 'noVBand');
  }
  return look;
}

function assignLookFlag(
  look: TableLook,
  key: keyof TableLook,
  element: Element,
  attributeName: string,
): void {
  const value = onOffValue(wordAttribute(element, attributeName));
  if (value !== undefined) look[key] = value;
}

function resolveBandSize(
  propertySources: readonly Element[],
  propertyName: string,
): number {
  let size = 1;
  for (const properties of propertySources) {
    const element = directChild(properties, propertyName);
    const value = Number(element ? wordAttribute(element, 'val') : undefined);
    if (Number.isInteger(value) && value > 0 && value <= 64) size = value;
  }
  return size;
}

function tableCellPosition(
  table: Element,
  targetCell: Element,
): TableCellPosition | null {
  const rows = descendants(table, 'tr').filter(
    (row) => closestAncestor(row, 'tbl') === table,
  );
  const targetRow = closestAncestor(targetCell, 'tr');
  if (!targetRow) return null;
  const rowIndex = rows.indexOf(targetRow);
  if (rowIndex < 0) return null;
  const cells = rowCells(targetRow);
  const cellIndex = cells.indexOf(targetCell);
  if (cellIndex < 0) return null;
  const columnStart =
    rowGridOffset(targetRow, 'gridBefore') +
    cells
      .slice(0, cellIndex)
      .reduce((total, cell) => total + cellGridSpan(cell), 0);
  const columnEnd = columnStart + cellGridSpan(targetCell) - 1;
  const columnCount = Math.max(
    1,
    ...rows.map(
      (row) =>
        rowGridOffset(row, 'gridBefore') +
        rowCells(row).reduce((total, cell) => total + cellGridSpan(cell), 0) +
        rowGridOffset(row, 'gridAfter'),
    ),
  );
  return {
    rowIndex,
    rowCount: rows.length,
    columnStart,
    columnEnd,
    columnCount,
  };
}

function rowCells(row: Element): Element[] {
  return descendants(row, 'tc').filter(
    (cell) => closestAncestor(cell, 'tr') === row,
  );
}

function cellGridSpan(cell: Element): number {
  const properties = directChild(cell, 'tcPr');
  const gridSpan = properties ? directChild(properties, 'gridSpan') : undefined;
  const value = Number(gridSpan ? wordAttribute(gridSpan, 'val') : undefined);
  return Number.isInteger(value) && value > 0 && value <= 256 ? value : 1;
}

function rowGridOffset(row: Element, propertyName: string): number {
  const properties = directChild(row, 'trPr');
  const element = properties
    ? directChild(properties, propertyName)
    : undefined;
  const value = Number(element ? wordAttribute(element, 'val') : undefined);
  return Number.isInteger(value) && value > 0 && value <= 256 ? value : 0;
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function onOffAttribute(element: Element, name: string): boolean {
  return onOffValue(wordAttribute(element, name)) ?? false;
}

function onOffValue(value: string | null | undefined): boolean | undefined {
  if (value === null || value === undefined) return undefined;
  const normalized = value.trim().toLowerCase();
  return normalized !== '0' && normalized !== 'false' && normalized !== 'off';
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}
