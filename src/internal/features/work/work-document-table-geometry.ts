export type DocumentTableLayoutAlgorithm = 'autofit' | 'fixed';
export type DocumentTablePreferredWidthType = 'auto' | 'percent' | 'pixels';
export type DocumentTableAlignment = 'left' | 'center' | 'right';
export type DocumentTableLayoutMode = 'window' | 'contents' | 'fixed';
export type DocumentTableCellMarginSide = 'top' | 'right' | 'bottom' | 'left';

export interface DocumentTablePreferredWidth {
  type: DocumentTablePreferredWidthType;
  value: number | null;
}

export interface DocumentTableCellMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export type DocumentTableCellMarginOverrides =
  Partial<DocumentTableCellMargins>;

export interface DocumentTableGeometry {
  layout: DocumentTableLayoutAlgorithm;
  width: DocumentTablePreferredWidth;
  alignment: DocumentTableAlignment;
  indent: number;
  cellMargins: DocumentTableCellMargins;
}

const MAX_TABLE_SIZE = 100_000;
const MAX_CELL_MARGIN = 1_000;
const TABLE_MARGIN_DATASET_KEYS = {
  top: 'officeTableCellMarginTop',
  right: 'officeTableCellMarginRight',
  bottom: 'officeTableCellMarginBottom',
  left: 'officeTableCellMarginLeft',
} as const satisfies Record<DocumentTableCellMarginSide, keyof DOMStringMap>;
const CELL_MARGIN_DATASET_KEYS = {
  top: 'officeCellMarginTop',
  right: 'officeCellMarginRight',
  bottom: 'officeCellMarginBottom',
  left: 'officeCellMarginLeft',
} as const satisfies Record<DocumentTableCellMarginSide, keyof DOMStringMap>;
const TABLE_MARGIN_CSS_PROPERTIES = {
  top: '--work-table-cell-margin-top',
  right: '--work-table-cell-margin-right',
  bottom: '--work-table-cell-margin-bottom',
  left: '--work-table-cell-margin-left',
} as const;
const CELL_PADDING_STYLE_PROPERTIES = {
  top: 'paddingTop',
  right: 'paddingRight',
  bottom: 'paddingBottom',
  left: 'paddingLeft',
} as const satisfies Record<
  DocumentTableCellMarginSide,
  'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
>;

export const DEFAULT_DOCUMENT_TABLE_CELL_MARGINS: DocumentTableCellMargins = {
  top: 0,
  right: 7.2,
  bottom: 0,
  left: 7.2,
};

export const DEFAULT_DOCUMENT_TABLE_GEOMETRY: DocumentTableGeometry = {
  layout: 'autofit',
  width: { type: 'percent', value: 100 },
  alignment: 'left',
  indent: 0,
  cellMargins: DEFAULT_DOCUMENT_TABLE_CELL_MARGINS,
};

export function normalizeDocumentTableGeometry(
  value: unknown,
  legacyLayoutMode?: unknown,
): DocumentTableGeometry {
  if (!isRecord(value)) {
    return geometryForLegacyLayoutMode(legacyLayoutMode);
  }
  const layout = normalizeDocumentTableLayoutAlgorithm(value.layout);
  const width = normalizeDocumentTablePreferredWidth(value.width);
  const alignment = normalizeDocumentTableAlignment(value.alignment);
  const indent = normalizeDimension(value.indent, MAX_TABLE_SIZE);
  const cellMargins = normalizeDocumentTableCellMargins(value.cellMargins);
  return {
    layout: layout ?? DEFAULT_DOCUMENT_TABLE_GEOMETRY.layout,
    width: width ?? clonePreferredWidth(DEFAULT_DOCUMENT_TABLE_GEOMETRY.width),
    alignment: alignment ?? DEFAULT_DOCUMENT_TABLE_GEOMETRY.alignment,
    indent: indent ?? DEFAULT_DOCUMENT_TABLE_GEOMETRY.indent,
    cellMargins:
      cellMargins ?? cloneCellMargins(DEFAULT_DOCUMENT_TABLE_CELL_MARGINS),
  };
}

export function normalizeDocumentTableLayoutMode(
  value: unknown,
): DocumentTableLayoutMode {
  return value === 'contents' || value === 'fixed' ? value : 'window';
}

export function normalizeDocumentTableLayoutAlgorithm(
  value: unknown,
): DocumentTableLayoutAlgorithm | null {
  return value === 'autofit' || value === 'fixed' ? value : null;
}

export function normalizeDocumentTableAlignment(
  value: unknown,
): DocumentTableAlignment | null {
  return value === 'center' || value === 'right' || value === 'left'
    ? value
    : null;
}

export function normalizeDocumentTableCellMargins(
  value: unknown,
): DocumentTableCellMargins | null {
  if (!isRecord(value)) return null;
  const margins = normalizeDocumentTableCellMarginOverrides(value);
  if (!margins || !hasAllMarginSides(margins)) return null;
  return margins as DocumentTableCellMargins;
}

export function normalizeDocumentTableCellMarginOverrides(
  value: unknown,
): DocumentTableCellMarginOverrides | null {
  if (!isRecord(value)) return null;
  const margins: DocumentTableCellMarginOverrides = {};
  for (const side of marginSides()) {
    const normalized = normalizeDimension(value[side], MAX_CELL_MARGIN);
    if (normalized !== null) margins[side] = normalized;
  }
  return Object.keys(margins).length ? margins : null;
}

export function documentTableLayoutMode(
  geometry: DocumentTableGeometry,
): DocumentTableLayoutMode {
  if (geometry.layout === 'autofit' && geometry.width.type === 'auto') {
    return 'contents';
  }
  if (
    geometry.layout === 'autofit' &&
    geometry.width.type === 'percent' &&
    geometry.width.value === 100
  ) {
    return 'window';
  }
  return 'fixed';
}

export function documentTableGeometryForLayoutMode(
  mode: DocumentTableLayoutMode,
  current: DocumentTableGeometry,
  renderedTableWidth?: number,
): DocumentTableGeometry {
  if (mode === 'contents') {
    return {
      ...current,
      layout: 'autofit',
      width: { type: 'auto', value: null },
    };
  }
  if (mode === 'window') {
    return {
      ...current,
      layout: 'autofit',
      width: { type: 'percent', value: 100 },
    };
  }
  const rendered = normalizePositiveDimension(
    renderedTableWidth,
    MAX_TABLE_SIZE,
  );
  return {
    ...current,
    layout: 'fixed',
    width: rendered
      ? { type: 'pixels', value: rendered }
      : current.width.type === 'pixels'
        ? clonePreferredWidth(current.width)
        : { type: 'auto', value: null },
  };
}

export function documentTableGeometryFromElement(
  table: HTMLElement,
): DocumentTableGeometry {
  const rawLayout = table.dataset.officeTableLayout;
  const exactLayout = normalizeDocumentTableLayoutAlgorithm(rawLayout);
  const legacyMode =
    rawLayout === 'window' || rawLayout === 'contents' || rawLayout === 'fixed'
      ? rawLayout
      : undefined;
  const width = preferredWidthFromElement(table, legacyMode);
  const margins = tableCellMarginsFromElement(table);
  const alignment = normalizeDocumentTableAlignment(
    table.dataset.officeTableAlignment,
  );
  const indent = normalizeDimension(
    table.dataset.officeTableIndent,
    MAX_TABLE_SIZE,
  );
  const hasGeometryAttributes = Boolean(
    rawLayout ||
      table.dataset.officeTableWidthType ||
      table.dataset.officeTableWidth ||
      table.dataset.officeTableAlignment ||
      table.dataset.officeTableIndent ||
      margins,
  );
  if (!hasGeometryAttributes) {
    return cloneGeometry(DEFAULT_DOCUMENT_TABLE_GEOMETRY);
  }
  return {
    layout: exactLayout ?? (legacyMode === 'fixed' ? 'fixed' : 'autofit'),
    width,
    alignment: alignment ?? DEFAULT_DOCUMENT_TABLE_GEOMETRY.alignment,
    indent: indent ?? DEFAULT_DOCUMENT_TABLE_GEOMETRY.indent,
    cellMargins:
      margins ?? cloneCellMargins(DEFAULT_DOCUMENT_TABLE_CELL_MARGINS),
  };
}

export function renderDocumentTableGeometry(
  value: unknown,
  legacyLayoutMode?: unknown,
): Record<string, string> {
  const geometry = normalizeDocumentTableGeometry(value, legacyLayoutMode);
  return {
    'data-office-table-layout': geometry.layout,
    'data-office-table-width-type': geometry.width.type,
    ...(geometry.width.value === null
      ? {}
      : { 'data-office-table-width': formatNumber(geometry.width.value) }),
    'data-office-table-alignment': geometry.alignment,
    'data-office-table-indent': formatNumber(geometry.indent),
    'data-office-table-cell-margin-top': formatNumber(geometry.cellMargins.top),
    'data-office-table-cell-margin-right': formatNumber(
      geometry.cellMargins.right,
    ),
    'data-office-table-cell-margin-bottom': formatNumber(
      geometry.cellMargins.bottom,
    ),
    'data-office-table-cell-margin-left': formatNumber(
      geometry.cellMargins.left,
    ),
    style: documentTableGeometryStyle(geometry),
  };
}

export function applyDocumentTableGeometryToElement(
  table: HTMLElement,
  value: unknown,
  legacyLayoutMode?: unknown,
): void {
  const geometry = normalizeDocumentTableGeometry(value, legacyLayoutMode);
  const rendered = renderDocumentTableGeometry(geometry);
  for (const attributeName of tableGeometryAttributeNames()) {
    table.removeAttribute(attributeName);
  }
  for (const [name, attributeValue] of Object.entries(rendered)) {
    if (name !== 'style') table.setAttribute(name, attributeValue);
  }
  table.style.removeProperty('table-layout');
  table.style.removeProperty('width');
  table.style.removeProperty('margin-left');
  table.style.removeProperty('margin-right');
  for (const property of Object.values(TABLE_MARGIN_CSS_PROPERTIES)) {
    table.style.removeProperty(property);
  }
  for (const declaration of rendered.style
    .split(';')
    .map((part) => part.trim())
    .filter(Boolean)) {
    const separator = declaration.indexOf(':');
    if (separator < 0) continue;
    table.style.setProperty(
      declaration.slice(0, separator).trim(),
      declaration.slice(separator + 1).trim(),
    );
  }
  if (geometry.width.type !== 'auto') {
    table.style.removeProperty('min-width');
  }
}

export function documentTableCellMarginOverridesFromElement(
  cell: HTMLElement,
): DocumentTableCellMarginOverrides | null {
  const margins: DocumentTableCellMarginOverrides = {};
  for (const side of marginSides()) {
    const datasetValue = cell.dataset[CELL_MARGIN_DATASET_KEYS[side]];
    const styleValue = cell.style[CELL_PADDING_STYLE_PROPERTIES[side]];
    const normalized =
      normalizeDimension(datasetValue, MAX_CELL_MARGIN) ??
      cssPixels(styleValue);
    if (normalized !== null) margins[side] = normalized;
  }
  return Object.keys(margins).length ? margins : null;
}

export function renderDocumentTableCellMarginOverrides(
  value: unknown,
): Record<string, string> {
  const margins = normalizeDocumentTableCellMarginOverrides(value);
  if (!margins) return {};
  const attributes: Record<string, string> = {};
  const styles: string[] = [];
  for (const side of marginSides()) {
    const margin = margins[side];
    if (margin === undefined) continue;
    attributes[`data-office-cell-margin-${side}`] = formatNumber(margin);
    styles.push(`padding-${side}: ${formatNumber(margin)}px`);
  }
  return { ...attributes, style: styles.join('; ') };
}

function preferredWidthFromElement(
  table: HTMLElement,
  legacyMode: DocumentTableLayoutMode | undefined,
): DocumentTablePreferredWidth {
  const type = normalizePreferredWidthType(table.dataset.officeTableWidthType);
  const value = normalizePositiveDimension(
    table.dataset.officeTableWidth,
    MAX_TABLE_SIZE,
  );
  if (type === 'auto') return { type, value: null };
  if (type && value !== null) return { type, value };

  const styleWidth = cssPreferredWidth(table.style.width);
  if (styleWidth) return styleWidth;
  if (legacyMode === 'contents') return { type: 'auto', value: null };
  if (legacyMode === 'fixed') {
    const columnWidth = firstRowColumnWidth(table);
    return columnWidth === null
      ? { type: 'auto', value: null }
      : { type: 'pixels', value: columnWidth };
  }
  return clonePreferredWidth(DEFAULT_DOCUMENT_TABLE_GEOMETRY.width);
}

function tableCellMarginsFromElement(
  table: HTMLElement,
): DocumentTableCellMargins | null {
  const margins: DocumentTableCellMarginOverrides = {};
  for (const side of marginSides()) {
    const datasetValue = table.dataset[TABLE_MARGIN_DATASET_KEYS[side]];
    const cssValue = table.style.getPropertyValue(
      TABLE_MARGIN_CSS_PROPERTIES[side],
    );
    const normalized =
      normalizeDimension(datasetValue, MAX_CELL_MARGIN) ?? cssPixels(cssValue);
    if (normalized !== null) margins[side] = normalized;
  }
  return hasAllMarginSides(margins)
    ? (margins as DocumentTableCellMargins)
    : null;
}

function documentTableGeometryStyle(geometry: DocumentTableGeometry): string {
  const styles = [
    `table-layout: ${geometry.layout === 'fixed' ? 'fixed' : 'auto'}`,
  ];
  if (geometry.width.type === 'percent' && geometry.width.value !== null) {
    styles.push(`width: ${formatNumber(geometry.width.value)}%`);
  } else if (
    geometry.width.type === 'pixels' &&
    geometry.width.value !== null
  ) {
    styles.push(`width: ${formatNumber(geometry.width.value)}px`);
  }
  if (geometry.alignment === 'center') {
    styles.push('margin-left: auto', 'margin-right: auto');
  } else if (geometry.alignment === 'right') {
    styles.push('margin-left: auto', 'margin-right: 0px');
  } else {
    styles.push(
      `margin-left: ${formatNumber(geometry.indent)}px`,
      'margin-right: auto',
    );
  }
  for (const side of marginSides()) {
    styles.push(
      `${TABLE_MARGIN_CSS_PROPERTIES[side]}: ${formatNumber(
        geometry.cellMargins[side],
      )}px`,
    );
  }
  return styles.join('; ');
}

function geometryForLegacyLayoutMode(value: unknown): DocumentTableGeometry {
  const mode = normalizeDocumentTableLayoutMode(value);
  if (mode === 'contents') {
    return {
      ...cloneGeometry(DEFAULT_DOCUMENT_TABLE_GEOMETRY),
      width: { type: 'auto', value: null },
    };
  }
  if (mode === 'fixed') {
    return {
      ...cloneGeometry(DEFAULT_DOCUMENT_TABLE_GEOMETRY),
      layout: 'fixed',
      width: { type: 'auto', value: null },
    };
  }
  return cloneGeometry(DEFAULT_DOCUMENT_TABLE_GEOMETRY);
}

function normalizeDocumentTablePreferredWidth(
  value: unknown,
): DocumentTablePreferredWidth | null {
  if (!isRecord(value)) return null;
  const type = normalizePreferredWidthType(value.type);
  if (!type) return null;
  if (type === 'auto') return { type, value: null };
  const width = normalizePositiveDimension(value.value, MAX_TABLE_SIZE);
  return width === null ? null : { type, value: width };
}

function normalizePreferredWidthType(
  value: unknown,
): DocumentTablePreferredWidthType | null {
  return value === 'auto' || value === 'percent' || value === 'pixels'
    ? value
    : null;
}

function normalizeDimension(value: unknown, maximum: number): number | null {
  const numeric =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value)
        : Number.NaN;
  if (!Number.isFinite(numeric) || numeric < 0 || numeric > maximum) {
    return null;
  }
  return Math.round(numeric * 100) / 100;
}

function normalizePositiveDimension(
  value: unknown,
  maximum: number,
): number | null {
  const normalized = normalizeDimension(value, maximum);
  return normalized !== null && normalized > 0 ? normalized : null;
}

function cssPixels(value: string): number | null {
  const match = /^(-?\d+(?:\.\d+)?)px$/i.exec(value.trim());
  return match ? normalizeDimension(match[1], MAX_CELL_MARGIN) : null;
}

function cssPreferredWidth(value: string): DocumentTablePreferredWidth | null {
  const normalized = value.trim();
  const percentage = /^(\d+(?:\.\d+)?)%$/.exec(normalized);
  if (percentage) {
    const width = normalizePositiveDimension(percentage[1], MAX_TABLE_SIZE);
    return width === null ? null : { type: 'percent', value: width };
  }
  const pixels = /^(\d+(?:\.\d+)?)px$/.exec(normalized);
  if (pixels) {
    const width = normalizePositiveDimension(pixels[1], MAX_TABLE_SIZE);
    return width === null ? null : { type: 'pixels', value: width };
  }
  return normalized === 'auto' ? { type: 'auto', value: null } : null;
}

function firstRowColumnWidth(table: HTMLElement): number | null {
  const row = table.querySelector('tr');
  if (!row) return null;
  let total = 0;
  for (const cell of Array.from(
    row.querySelectorAll(':scope > th, :scope > td'),
  )) {
    const raw = cell.getAttribute('colwidth');
    if (!raw) return null;
    const widths = raw
      .split(',')
      .map((part) => normalizePositiveDimension(part, MAX_TABLE_SIZE));
    if (!widths.length || widths.some((width) => width === null)) return null;
    total += widths.reduce<number>((sum, width) => sum + Number(width), 0);
  }
  return total > 0 ? Math.round(total * 100) / 100 : null;
}

function tableGeometryAttributeNames(): string[] {
  return [
    'data-office-table-layout',
    'data-office-table-width-type',
    'data-office-table-width',
    'data-office-table-alignment',
    'data-office-table-indent',
    ...marginSides().map((side) => `data-office-table-cell-margin-${side}`),
  ];
}

function marginSides(): readonly DocumentTableCellMarginSide[] {
  return ['top', 'right', 'bottom', 'left'];
}

function hasAllMarginSides(value: DocumentTableCellMarginOverrides): boolean {
  return marginSides().every((side) => value[side] !== undefined);
}

function cloneGeometry(value: DocumentTableGeometry): DocumentTableGeometry {
  return {
    ...value,
    width: clonePreferredWidth(value.width),
    cellMargins: cloneCellMargins(value.cellMargins),
  };
}

function clonePreferredWidth(
  value: DocumentTablePreferredWidth,
): DocumentTablePreferredWidth {
  return { ...value };
}

function cloneCellMargins(
  value: DocumentTableCellMargins,
): DocumentTableCellMargins {
  return { ...value };
}

function formatNumber(value: number): string {
  return String(Math.round(value * 100) / 100);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
