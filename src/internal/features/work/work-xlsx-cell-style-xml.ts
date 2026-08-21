import type { XlsxCellBorderLine } from './work-xlsx-cell-borders';
import { directChild, directChildren } from './work-ooxml-package';
import type { SpreadsheetUnderlineStyle } from './work-spreadsheet-underline';

const borderChildOrder = [
  'start',
  'end',
  'left',
  'right',
  'top',
  'bottom',
  'diagonal',
  'vertical',
  'horizontal',
  'extLst',
] as const;

export interface XlsxDirectAlignmentStyle {
  horizontal?: 'center' | 'left' | 'right';
  textRotation?: number;
  vertical?: 'bottom' | 'center' | 'top';
  wrapText?: boolean;
}

export function ensureXlsxStyleCollection(
  document: Document,
  name: string,
  anchors: readonly string[],
): Element {
  const root = document.documentElement;
  const existing = directChild(root, name);
  if (existing) return existing;
  const collection = document.createElementNS(root.namespaceURI, name);
  root.insertBefore(
    collection,
    directChildren(root).find((child) => anchors.includes(child.localName)) ??
      null,
  );
  return collection;
}

export function defaultXlsxFill(
  document: Document,
  patternType: string,
): Element {
  const fill = document.createElementNS(
    document.documentElement.namespaceURI,
    'fill',
  );
  const pattern = document.createElementNS(
    document.documentElement.namespaceURI,
    'patternFill',
  );
  pattern.setAttribute('patternType', patternType);
  fill.append(pattern);
  return fill;
}

export function defaultXlsxBorder(document: Document): Element {
  const border = document.createElementNS(
    document.documentElement.namespaceURI,
    'border',
  );
  for (const name of ['left', 'right', 'top', 'bottom', 'diagonal']) {
    border.append(
      document.createElementNS(document.documentElement.namespaceURI, name),
    );
  }
  return border;
}

export function setXlsxBorderLine(
  document: Document,
  border: Element,
  name: 'bottom' | 'diagonal' | 'left' | 'right' | 'top',
  line: XlsxCellBorderLine | null,
): void {
  removeXlsxChildren(border, name);
  const element = document.createElementNS(
    document.documentElement.namespaceURI,
    name,
  );
  if (line) {
    element.setAttribute('style', line.style);
    const color = xlsxRgbColor(line.color);
    if (color) {
      const child = document.createElementNS(
        document.documentElement.namespaceURI,
        'color',
      );
      child.setAttribute('rgb', color);
      element.append(child);
    }
  }
  insertXlsxOrderedChild(border, element, borderChildOrder);
}

export function writeXlsxAlignment(
  document: Document,
  xf: Element,
  style: XlsxDirectAlignmentStyle,
): void {
  let alignment = directChild(xf, 'alignment');
  if (!alignment) {
    alignment = document.createElementNS(
      document.documentElement.namespaceURI,
      'alignment',
    );
    xf.insertBefore(
      alignment,
      directChildren(xf).find((child) =>
        ['protection', 'extLst'].includes(child.localName),
      ) ?? null,
    );
  }
  if (style.horizontal !== undefined)
    alignment.setAttribute('horizontal', style.horizontal);
  if (style.vertical !== undefined)
    alignment.setAttribute('vertical', style.vertical);
  if (style.wrapText !== undefined)
    alignment.setAttribute('wrapText', style.wrapText ? '1' : '0');
  if (
    style.textRotation !== undefined &&
    Number.isInteger(style.textRotation) &&
    ((style.textRotation >= 0 && style.textRotation <= 180) ||
      style.textRotation === 255)
  )
    alignment.setAttribute('textRotation', String(style.textRotation));
}

export function setXlsxValueChild(
  document: Document,
  parent: Element,
  name: string,
  value: string,
  order: readonly string[],
): void {
  removeXlsxChildren(parent, name);
  const child = document.createElementNS(
    document.documentElement.namespaceURI,
    name,
  );
  child.setAttribute('val', value);
  insertXlsxOrderedChild(parent, child, order);
}

export function setXlsxColorChild(
  document: Document,
  parent: Element,
  color: string,
  order: readonly string[],
): void {
  removeXlsxChildren(parent, 'color');
  const child = document.createElementNS(
    document.documentElement.namespaceURI,
    'color',
  );
  child.setAttribute('rgb', color);
  insertXlsxOrderedChild(parent, child, order);
}

export function setXlsxToggleChild(
  document: Document,
  parent: Element,
  name: string,
  enabled: boolean,
  order: readonly string[],
): void {
  removeXlsxChildren(parent, name);
  if (!enabled) return;
  const child = document.createElementNS(
    document.documentElement.namespaceURI,
    name,
  );
  child.setAttribute('val', '1');
  insertXlsxOrderedChild(parent, child, order);
}

export function setXlsxUnderlineChild(
  document: Document,
  parent: Element,
  style: SpreadsheetUnderlineStyle,
  order: readonly string[],
): void {
  removeXlsxChildren(parent, 'u');
  if (style === 'none') return;
  const child = document.createElementNS(
    document.documentElement.namespaceURI,
    'u',
  );
  child.setAttribute('val', style);
  insertXlsxOrderedChild(parent, child, order);
}

export function xlsxRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const color = value.trim().replace('#', '').toUpperCase();
  if (/^[0-9A-F]{6}$/.test(color)) return `FF${color}`;
  if (/^[0-9A-F]{8}$/.test(color)) return color;
  if (/^[0-9A-F]{3}$/.test(color)) {
    return `FF${[...color].map((character) => character.repeat(2)).join('')}`;
  }
  return null;
}

function removeXlsxChildren(parent: Element, name: string): void {
  for (const child of directChildren(parent, name)) child.remove();
}

function insertXlsxOrderedChild(
  parent: Element,
  child: Element,
  order: readonly string[],
): void {
  const requested = order.indexOf(child.localName);
  const anchor = directChildren(parent).find(
    (candidate) => order.indexOf(candidate.localName) > requested,
  );
  parent.insertBefore(child, anchor ?? null);
}
