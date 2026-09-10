import {
  normalizeDocumentTableCellTextDirection,
  serializeDocumentCellFormatting,
  type DocumentTableCellHMerge,
  type DocumentTableCellVMerge,
  type DocumentTableCellTextDirection,
} from './work-document-cell-format-changes';
import {
  parseDocxHMergeValue,
  parseDocxVMergeValue,
} from './work-document-table-cell-formatting';
import { parseDocxCnfStyleElement } from './work-document-cnf-style';
import { normalizeTableColor } from './work-document-table-borders';
import {
  importedDocxCellFormattingBorders,
  type DocumentCellFormattingBorders,
} from './work-document-table-formatting-borders';
import { normalizeDocumentTableVerticalAlign } from './work-document-table-cell-formatting';
import {
  normalizeDocumentTableCellMarginOverrides,
  type DocumentTableCellMarginOverrides,
  type DocumentTableCellMarginSide,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute, directChildren } from './work-ooxml-package';

const MAX_REVISION_DATE_LENGTH = 64;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const SUPPORTED_PRIOR_CHILDREN = new Set([
  'vAlign',
  'shd',
  'tcMar',
  'tcW',
  'noWrap',
  'textDirection',
  'tcFitText',
  'hideMark',
  'cnfStyle',
  'hMerge',
  'vMerge',
  'tcBorders',
]);
const SOLID_SHADING_VALUES = new Set(['clear', 'nil', 'none', '']);
const MARGIN_SIDES = new Set([
  'top',
  'right',
  'bottom',
  'left',
  'start',
  'end',
]);
const PIXELS_PER_TWIP = 96 / 1440;

export interface SupportedDocxCellFormattingChange {
  id: string;
  author: string;
  date: string;
  before: string;
}

/**
 * Relationship-free `w:tcPrChange` whose prior snapshot contains only
 * `w:vAlign`, solid direct-color `w:shd`, `w:tcMar`, `w:tcW`, `w:noWrap`,
 * `w:textDirection`, `w:tcFitText`, `w:hideMark`, `w:cnfStyle`, `w:hMerge`, `w:vMerge`, and/or
 * direct-color `w:tcBorders`. Broader cell property sets stay on the opaque OMML path.
 */
export function isSupportedDocxCellFormattingChange(change: Element): boolean {
  return supportedCellFormattingChange(change) !== null;
}

export function supportedDocxCellFormattingChangeFromProperties(
  cellProperties: Element | null | undefined,
): SupportedDocxCellFormattingChange | null {
  if (!cellProperties) return null;
  const changes = directChildren(cellProperties, 'tcPrChange').filter(
    (element) => element.namespaceURI === cellProperties.namespaceURI,
  );
  if (changes.length !== 1) return null;
  return supportedCellFormattingChange(changes[0] ?? null);
}

export function applyDocumentCellFormattingChangeToElement(
  cell: HTMLElement,
  change: SupportedDocxCellFormattingChange,
): void {
  cell.dataset.documentChange = 'true';
  cell.dataset.changeKind = 'cell-formatting';
  cell.dataset.changeId = change.id;
  cell.dataset.changeAuthor = change.author;
  cell.dataset.changeDate = change.date;
  cell.dataset.changeBefore = change.before;
}

function supportedCellFormattingChange(
  change: Element | null,
): SupportedDocxCellFormattingChange | null {
  if (
    !change ||
    change.localName !== 'tcPrChange' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  if (
    hasRelationshipBindings(change) ||
    hasUnsupportedRevisionAttributes(change)
  ) {
    return null;
  }
  const id = wordAttribute(change, 'id')?.trim() ?? '';
  const author = wordAttribute(change, 'author')?.trim() ?? '';
  const rawDate = wordAttribute(change, 'date');
  if (
    !/^\+?\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (rawDate !== null && rawDate.length > MAX_REVISION_DATE_LENGTH) ||
    (rawDate !== null && !Number.isFinite(Date.parse(rawDate)))
  ) {
    return null;
  }
  const priors = Array.from(change.children).filter(
    (child) =>
      child.localName === 'tcPr' && child.namespaceURI === change.namespaceURI,
  );
  if (priors.length !== 1) return null;
  const prior = priors[0];
  if (!prior || hasRelationshipBindings(prior)) return null;
  const children = Array.from(prior.children);
  if (!children.length) return null;
  if (
    children.some((child) => {
      if (
        child.namespaceURI !== change.namespaceURI ||
        !SUPPORTED_PRIOR_CHILDREN.has(child.localName)
      ) {
        return true;
      }
      if (
        child.localName === 'tcMar' ||
        child.localName === 'tcW' ||
        child.localName === 'noWrap' ||
        child.localName === 'textDirection' ||
        child.localName === 'tcFitText' ||
        child.localName === 'hideMark' ||
        child.localName === 'cnfStyle' ||
        child.localName === 'hMerge' ||
        child.localName === 'vMerge' ||
        child.localName === 'tcBorders'
      ) {
        return false;
      }
      return child.children.length > 0;
    })
  ) {
    return null;
  }
  const names = children.map((child) => child.localName);
  if (new Set(names).size !== names.length) return null;
  const snapshot: {
    verticalAlign?: string;
    fill?: string;
    margins?: DocumentTableCellMarginOverrides;
    width?: DocumentTablePreferredWidth;
    noWrap?: boolean;
    textDirection?: DocumentTableCellTextDirection;
    fitText?: boolean;
    hideMark?: boolean;
    cnfStyle?: string;
    hMerge?: DocumentTableCellHMerge;
    vMerge?: DocumentTableCellVMerge;
    borders?: DocumentCellFormattingBorders;
  } = {};
  for (const child of children) {
    if (child.localName === 'vAlign') {
      const verticalAlign = normalizeDocumentTableVerticalAlign(
        mapVAlignValue(attribute(child, 'val')),
      );
      if (!verticalAlign) return null;
      snapshot.verticalAlign = verticalAlign;
      continue;
    }
    if (child.localName === 'shd') {
      const fill = solidShadingFill(child);
      if (!fill) return null;
      snapshot.fill = fill;
      continue;
    }
    if (child.localName === 'tcMar') {
      const margins = importedCellMargins(child);
      if (!margins) return null;
      snapshot.margins = margins;
      continue;
    }
    if (child.localName === 'tcW') {
      const width = importedPreferredWidth(child);
      if (!width) return null;
      snapshot.width = width;
      continue;
    }
    if (child.localName === 'noWrap') {
      snapshot.noWrap = onOffValue(child);
      continue;
    }
    if (child.localName === 'textDirection') {
      const textDirection = normalizeDocumentTableCellTextDirection(
        attribute(child, 'val'),
      );
      if (!textDirection) return null;
      snapshot.textDirection = textDirection;
      continue;
    }
    if (child.localName === 'tcFitText') {
      snapshot.fitText = onOffValue(child);
      continue;
    }
    if (child.localName === 'hideMark') {
      snapshot.hideMark = onOffValue(child);
      continue;
    }
    if (child.localName === 'cnfStyle') {
      const cnfStyle = parseDocxCnfStyleElement(child);
      if (!cnfStyle) return null;
      snapshot.cnfStyle = cnfStyle;
      continue;
    }
    if (child.localName === 'hMerge') {
      const hMerge = parseDocxHMergeValue(attribute(child, 'val'));
      if (!hMerge) return null;
      snapshot.hMerge = hMerge;
      continue;
    }
    if (child.localName === 'vMerge') {
      const vMerge = parseDocxVMergeValue(attribute(child, 'val'));
      if (!vMerge) return null;
      snapshot.vMerge = vMerge;
      continue;
    }
    if (child.localName === 'tcBorders') {
      const borders = importedDocxCellFormattingBorders(child, attribute);
      if (!borders) return null;
      snapshot.borders = borders;
    }
  }
  return {
    id: `docx-cell-format-change-${id}`,
    author,
    date: normalizeRevisionDate(rawDate),
    before: serializeDocumentCellFormatting(snapshot),
  };
}

function importedCellMargins(
  element: Element,
): DocumentTableCellMarginOverrides | null {
  const children = Array.from(element.children);
  if (!children.length) return null;
  if (
    children.some(
      (child) =>
        child.namespaceURI !== element.namespaceURI ||
        !MARGIN_SIDES.has(child.localName) ||
        child.children.length > 0,
    )
  ) {
    return null;
  }
  const names = children.map((child) => child.localName);
  if (new Set(names).size !== names.length) return null;
  const margins: DocumentTableCellMarginOverrides = {};
  for (const child of children) {
    const side = marginSide(child.localName);
    if (!side) return null;
    const type = attribute(child, 'type');
    if (type !== null && type !== 'dxa') return null;
    const pixels = twipsToPixels(Number(attribute(child, 'w')));
    if (pixels === null) return null;
    if (margins[side] !== undefined) return null;
    margins[side] = pixels;
  }
  return normalizeDocumentTableCellMarginOverrides(margins);
}

function importedPreferredWidth(
  width: Element,
): DocumentTablePreferredWidth | null {
  const type = attribute(width, 'type');
  if (type === 'auto' || type === 'nil') return { type: 'auto', value: null };
  if (type === 'pct') {
    const value = percentageValue(attribute(width, 'w'));
    return value === null ? null : { type: 'percent', value };
  }
  if (type === 'dxa') {
    const pixels = twipsToPixels(Number(attribute(width, 'w')));
    if (pixels === null || pixels <= 0) return null;
    return { type: 'pixels', value: pixels };
  }
  return null;
}

function percentageValue(value: string | null): number | null {
  const normalized = value?.trim();
  if (!normalized) return null;
  const percentage = normalized.endsWith('%')
    ? Number(normalized.slice(0, -1))
    : Number(normalized) / 50;
  if (!Number.isFinite(percentage) || percentage <= 0) return null;
  return Math.round(percentage * 100) / 100;
}

function marginSide(localName: string): DocumentTableCellMarginSide | null {
  if (localName === 'top') return 'top';
  if (localName === 'bottom') return 'bottom';
  if (localName === 'left' || localName === 'start') return 'left';
  if (localName === 'right' || localName === 'end') return 'right';
  return null;
}

function twipsToPixels(value: number): number | null {
  if (!Number.isFinite(value) || value < 0) return null;
  return Math.round(value * PIXELS_PER_TWIP * 100) / 100;
}

function solidShadingFill(element: Element): string | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) => xmlAttributeNamespace(element, candidate) === namespace,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if (
    names.has('themeFill') ||
    names.has('themeFillTint') ||
    names.has('themeFillShade') ||
    names.has('themeColor') ||
    names.has('color') ||
    names.has('themeTint') ||
    names.has('themeShade')
  ) {
    return null;
  }
  if (![...names].every((name) => name === 'val' || name === 'fill')) {
    return null;
  }
  const val = attribute(element, 'val');
  if (val !== null && !SOLID_SHADING_VALUES.has(val.trim().toLowerCase())) {
    return null;
  }
  const fill = attribute(element, 'fill')?.trim() ?? '';
  if (!/^[0-9A-Fa-f]{6}$/.test(fill)) return null;
  return normalizeTableColor(`#${fill}`);
}

function mapVAlignValue(value: string | null): string | null {
  if (!value) return null;
  if (value === 'center') return 'middle';
  return value;
}

function onOffValue(element: Element): boolean {
  const value = attribute(element, 'val');
  if (value === null || value === '') return true;
  return !(
    value === '0' ||
    value === 'false' ||
    value === 'off' ||
    value === 'False'
  );
}

function wordAttribute(element: Element, localName: string): string | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === localName &&
      xmlAttributeNamespace(element, candidate) === namespace,
  );
  return matches.length === 1 ? (matches[0]?.value ?? null) : null;
}

function hasUnsupportedRevisionAttributes(element: Element): boolean {
  const namespace = element.namespaceURI;
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !REVISION_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
  );
}

function hasRelationshipBindings(element: Element): boolean {
  for (const attributeNode of Array.from(element.attributes)) {
    const namespace =
      attributeNode.namespaceURI ||
      xmlAttributeNamespace(element, attributeNode) ||
      '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return true;
  }
  for (const descendant of Array.from(element.querySelectorAll('*'))) {
    for (const attributeNode of Array.from(descendant.attributes)) {
      const namespace =
        attributeNode.namespaceURI ||
        xmlAttributeNamespace(descendant, attributeNode) ||
        '';
      if (RELATIONSHIP_NAMESPACES.has(namespace)) return true;
    }
  }
  return false;
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}
