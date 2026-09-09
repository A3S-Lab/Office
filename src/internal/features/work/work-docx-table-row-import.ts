import {
  applyDocumentTableRowIdentityToElement,
  normalizeDocumentTableRowIdentity,
  type WorkDocumentTableRowIdentity,
} from './work-document-table-row-identity';
import {
  applyDocumentRowPropertyRevisionOmmlToElement,
  serializePreservableDocxRowPropertyRevision,
} from './work-document-table-property-revision';
import { parseDocxRowCnfStyleElement } from './work-document-row-cnf-style';
import {
  applyDocumentRowFormattingChangeToElement,
  supportedDocxRowFormattingChangeFromProperties,
  type SupportedDocxRowFormattingChange,
} from './work-docx-row-format-change-import';
import { normalizeDocumentParagraphId } from './work-document-paragraph-identity';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { attribute, descendants, directChild } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

export interface ImportedDocxTableRowMarker {
  marker: string;
  cantSplit?: boolean;
  repeatHeader?: boolean;
  hidden?: boolean;
  alignment?: 'left' | 'center' | 'right';
  gridBefore?: number;
  gridAfter?: number;
  widthBefore?: { type: 'auto' | 'percent' | 'pixels'; value: number | null };
  widthAfter?: { type: 'auto' | 'percent' | 'pixels'; value: number | null };
  cnfStyle?: string;
  rowId?: string;
  rowHeight?: number;
  rowHeightRule?: 'atLeast' | 'exact';
  rowTextId?: string;
  propertyRevisionOmml?: string;
  formattingChange?: SupportedDocxRowFormattingChange;
}

export interface ImportedDocxTableRowMarkers {
  rows: ImportedDocxTableRowMarker[];
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const TABLE_ROW_MARKER_PATTERN = /__A3S_WORK_TABLE_ROW_\d+__/g;
const PIXELS_PER_TWIP = 96 / 1440;

export function markDocxTableRows(
  document: Document,
): ImportedDocxTableRowMarkers {
  const wordRows = descendants(document, 'tr').filter((row) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(row.namespaceURI ?? ''),
  );
  const identities = new Map<Element, WorkDocumentTableRowIdentity>();
  const identityCounts = wordIdentityIdCounts(document);
  for (const row of wordRows) {
    const identity = wordTableRowIdentity(row);
    if (!identity) continue;
    identities.set(row, identity);
  }
  const rows: ImportedDocxTableRowMarker[] = [];
  for (const row of wordRows) {
    const properties = directChild(row, 'trPr');
    const cantSplit = properties
      ? directChild(properties, 'cantSplit')
      : undefined;
    const repeatHeader = properties
      ? directChild(properties, 'tblHeader')
      : undefined;
    const hidden = properties ? directChild(properties, 'hidden') : undefined;
    const jc = properties ? directChild(properties, 'jc') : undefined;
    const alignment = jc
      ? normalizeRowAlignment(attribute(jc, 'val'))
      : undefined;
    const gridBeforeElement = properties
      ? directChild(properties, 'gridBefore')
      : undefined;
    const gridBeforeRaw = gridBeforeElement
      ? Number(attribute(gridBeforeElement, 'val'))
      : null;
    const gridBefore =
      gridBeforeRaw !== null &&
      Number.isInteger(gridBeforeRaw) &&
      gridBeforeRaw >= 0
        ? gridBeforeRaw
        : undefined;
    const gridAfterElement = properties
      ? directChild(properties, 'gridAfter')
      : undefined;
    const gridAfterRaw = gridAfterElement
      ? Number(attribute(gridAfterElement, 'val'))
      : null;
    const gridAfter =
      gridAfterRaw !== null &&
      Number.isInteger(gridAfterRaw) &&
      gridAfterRaw >= 0
        ? gridAfterRaw
        : undefined;
    const widthBeforeElement = properties
      ? directChild(properties, 'wBefore')
      : undefined;
    const widthBefore = widthBeforeElement
      ? importedRowPreferredWidth(widthBeforeElement)
      : undefined;
    const widthAfterElement = properties
      ? directChild(properties, 'wAfter')
      : undefined;
    const widthAfter = widthAfterElement
      ? importedRowPreferredWidth(widthAfterElement)
      : undefined;
    const cnfStyleElement = properties
      ? directChild(properties, 'cnfStyle')
      : undefined;
    const cnfStyleParsed = cnfStyleElement
      ? parseDocxRowCnfStyleElement(cnfStyleElement)
      : null;
    const cnfStyle = cnfStyleParsed ?? undefined;
    const height = properties ? directChild(properties, 'trHeight') : undefined;
    const rowHeight = height
      ? twipsToPixels(Number(attribute(height, 'val')))
      : null;
    const rowHeightRule = tableRowHeightRule(attribute(height ?? row, 'hRule'));
    const rowIdentity = identities.get(row);
    const uniqueIdentity =
      rowIdentity && identityCounts.get(rowIdentity.rowId) === 1
        ? rowIdentity
        : null;
    const formattingChange =
      supportedDocxRowFormattingChangeFromProperties(properties);
    const propertyRevisionOmml = formattingChange
      ? undefined
      : (serializePreservableDocxRowPropertyRevision(properties) ?? undefined);
    if (
      !cantSplit &&
      !repeatHeader &&
      !hidden &&
      !alignment &&
      gridBefore === undefined &&
      gridAfter === undefined &&
      widthBefore === undefined &&
      widthAfter === undefined &&
      cnfStyle === undefined &&
      rowHeight === null &&
      !uniqueIdentity &&
      !propertyRevisionOmml &&
      !formattingChange
    ) {
      continue;
    }
    const paragraph = firstTableRowParagraph(document, row);
    if (!paragraph) continue;
    const marker = `__A3S_WORK_TABLE_ROW_${rows.length + 1}__`;
    insertRowMarker(document, paragraph, marker);
    rows.push({
      marker,
      ...(cantSplit ? { cantSplit: onOffValue(cantSplit) } : {}),
      ...(repeatHeader ? { repeatHeader: onOffValue(repeatHeader) } : {}),
      ...(hidden ? { hidden: onOffValue(hidden) } : {}),
      ...(alignment ? { alignment } : {}),
      ...(gridBefore !== undefined ? { gridBefore } : {}),
      ...(gridAfter !== undefined ? { gridAfter } : {}),
      ...(widthBefore !== undefined ? { widthBefore } : {}),
      ...(widthAfter !== undefined ? { widthAfter } : {}),
      ...(cnfStyle !== undefined ? { cnfStyle } : {}),
      ...(uniqueIdentity ?? {}),
      ...(rowHeight !== null ? { rowHeight } : {}),
      ...(rowHeight !== null && rowHeightRule ? { rowHeightRule } : {}),
      ...(propertyRevisionOmml ? { propertyRevisionOmml } : {}),
      ...(formattingChange ? { formattingChange } : {}),
    });
  }
  return { rows };
}

function wordIdentityIdCounts(document: Document): Map<string, number> {
  const counts = new Map<string, number>();
  for (const localName of ['p', 'tr']) {
    for (const element of descendants(document, localName)) {
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
        continue;
      }
      const paragraphId = normalizeDocumentParagraphId(
        word2010Attribute(element, 'paraId'),
      );
      if (!paragraphId) continue;
      counts.set(paragraphId, (counts.get(paragraphId) ?? 0) + 1);
    }
  }
  return counts;
}

export function applyImportedDocxTableRowMarkers(
  document: Document,
  markers: ImportedDocxTableRowMarkers,
): void {
  const rowByMarker = new Map(markers.rows.map((row) => [row.marker, row]));
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_TABLE_ROW_')) continue;
    const row = node.parentElement?.closest('tr');
    node.data = node.data.replace(TABLE_ROW_MARKER_PATTERN, (marker) => {
      const properties = rowByMarker.get(marker);
      if (row instanceof HTMLTableRowElement && properties) {
        applyDocumentTableRowIdentityToElement(row, properties);
        setBooleanAttribute(
          row,
          'data-office-cant-split',
          properties.cantSplit,
        );
        setBooleanAttribute(
          row,
          'data-office-repeat-header',
          properties.repeatHeader,
        );
        setBooleanAttribute(row, 'data-office-row-hidden', properties.hidden);
        if (properties.alignment) {
          row.dataset.officeRowAlignment = properties.alignment;
        }
        if (properties.gridBefore !== undefined) {
          row.dataset.officeRowGridBefore = String(properties.gridBefore);
        }
        if (properties.gridAfter !== undefined) {
          row.dataset.officeRowGridAfter = String(properties.gridAfter);
        }
        if (properties.widthBefore) {
          row.dataset.officeRowWidthBeforeType = properties.widthBefore.type;
          if (
            properties.widthBefore.type !== 'auto' &&
            properties.widthBefore.value !== null
          ) {
            row.dataset.officeRowWidthBefore = String(
              properties.widthBefore.value,
            );
          }
        }
        if (properties.widthAfter) {
          row.dataset.officeRowWidthAfterType = properties.widthAfter.type;
          if (
            properties.widthAfter.type !== 'auto' &&
            properties.widthAfter.value !== null
          ) {
            row.dataset.officeRowWidthAfter = String(
              properties.widthAfter.value,
            );
          }
        }
        if (properties.cnfStyle !== undefined) {
          row.dataset.officeRowCnfStyle = properties.cnfStyle;
        }
        if (properties.rowHeight !== undefined) {
          row.dataset.officeRowHeight = String(properties.rowHeight);
          row.style.height = `${properties.rowHeight}px`;
          row.dataset.officeRowHeightRule =
            properties.rowHeightRule ?? 'atLeast';
        }
        applyDocumentRowPropertyRevisionOmmlToElement(
          row,
          properties.propertyRevisionOmml,
        );
        if (properties.formattingChange) {
          applyDocumentRowFormattingChangeToElement(
            row,
            properties.formattingChange,
          );
        }
      }
      return '';
    });
  }
  document.body.normalize();
}

export function hasImportedDocxTableRowMarkers(
  markers: ImportedDocxTableRowMarkers,
): boolean {
  return markers.rows.length > 0;
}

function firstTableRowParagraph(
  document: Document,
  row: Element,
): Element | null {
  const cell = directChild(row, 'tc');
  if (!cell) return null;
  const existing =
    directChild(cell, 'p') ??
    descendants(cell, 'p').find(
      (paragraph) => closestAncestor(paragraph, 'tr') === row,
    );
  if (existing) return existing;
  const namespace = cell.namespaceURI ?? WORD_NAMESPACE;
  const prefix = cell.prefix ? `${cell.prefix}:` : '';
  const paragraph = document.createElementNS(namespace, `${prefix}p`);
  cell.append(paragraph);
  return paragraph;
}

function insertRowMarker(
  document: Document,
  paragraph: Element,
  marker: string,
): void {
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChild(paragraph, 'pPr');
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
}

function wordTableRowIdentity(
  row: Element,
): WorkDocumentTableRowIdentity | null {
  return normalizeDocumentTableRowIdentity({
    rowId: word2010Attribute(row, 'paraId'),
    rowTextId: word2010Attribute(row, 'textId'),
  });
}

function word2010Attribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

function onOffValue(element: Element): boolean {
  const value = (attribute(element, 'val') ?? attribute(element, 'w:val'))
    ?.trim()
    .toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function normalizeRowAlignment(
  value: string | null,
): 'left' | 'center' | 'right' | undefined {
  if (!value) return undefined;
  if (value === 'start' || value === 'left') return 'left';
  if (value === 'end' || value === 'right') return 'right';
  if (value === 'center') return 'center';
  return undefined;
}

function setBooleanAttribute(
  element: HTMLElement,
  name: string,
  value: boolean | undefined,
): void {
  if (value !== undefined) element.setAttribute(name, String(value));
}

function tableRowHeightRule(value: string | null): 'atLeast' | 'exact' | null {
  if (value === 'exact') return 'exact';
  return value === 'atLeast' ? 'atLeast' : null;
}

function twipsToPixels(value: number): number | null {
  if (!Number.isFinite(value) || value <= 0) return null;
  return Math.round(value * PIXELS_PER_TWIP * 100) / 100;
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

function importedRowPreferredWidth(
  width: Element,
): { type: 'auto' | 'percent' | 'pixels'; value: number | null } | undefined {
  const type = attribute(width, 'type');
  if (type === 'auto' || type === 'nil') return { type: 'auto', value: null };
  if (type === 'pct') {
    const raw = attribute(width, 'w')?.trim();
    if (!raw) return undefined;
    const percentage = raw.endsWith('%')
      ? Number(raw.slice(0, -1))
      : Number(raw) / 50;
    if (!Number.isFinite(percentage) || percentage <= 0) return undefined;
    return {
      type: 'percent',
      value: Math.round(percentage * 100) / 100,
    };
  }
  if (type === 'dxa') {
    const pixels = twipsToPixels(Number(attribute(width, 'w')));
    if (pixels === null || pixels <= 0) return undefined;
    return { type: 'pixels', value: pixels };
  }
  return undefined;
}
