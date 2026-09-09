import JSZip from 'jszip';
import type { DocumentTablePreferredWidth } from './work-document-table-geometry';
import {
  parseDocumentTableFormatting,
  normalizeDocumentTableOverlap,
  normalizeDocumentTableStyleId,
  normalizeDocumentTableCellSpacing,
  type DocumentTableOverlap,
} from './work-document-table-format-changes';
import {
  documentTableLookBitmask,
  parseDocumentTableLookDataset,
  type DocumentTableLook,
} from './work-document-table-look';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxTableFormattingChangePatch {
  id: number;
  author: string;
  date: string;
  before: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_TABLE_FORMATTING_CHANGE_PATCHES = 4_096;
const TWIPS_PER_PIXEL = 1440 / 96;

export class DocxTableFormattingChangePatchCollector {
  readonly patches: Array<DocxTableFormattingChangePatch | null> = [];
  readonly bidiVisual: boolean[] = [];
  readonly fills: Array<string | null> = [];
  readonly looks: Array<DocumentTableLook | null> = [];
  readonly overlaps: Array<DocumentTableOverlap | null> = [];
  readonly styleIds: Array<string | null> = [];
  readonly cellSpacings: Array<number | null> = [];

  record(element: HTMLTableElement, id: number): void {
    this.bidiVisual.push(element.dataset.officeTableBidiVisual === 'true');
    const fill = element.dataset.officeTableFill?.trim() ?? '';
    this.fills.push(/^#[0-9A-Fa-f]{6}$/i.test(fill) ? fill.toLowerCase() : null);
    this.looks.push(
      parseDocumentTableLookDataset(element.dataset.officeTableLook),
    );
    this.overlaps.push(
      normalizeDocumentTableOverlap(element.dataset.officeTableOverlap),
    );
    this.styleIds.push(
      normalizeDocumentTableStyleId(element.dataset.officeTableStyleId),
    );
    this.cellSpacings.push(
      normalizeDocumentTableCellSpacing(
        element.dataset.officeTableCellSpacing,
      ),
    );
    if (
      element.dataset.changeKind !== 'table-formatting' ||
      element.getAttribute('data-document-change') !== 'true'
    ) {
      this.patches.push(null);
      return;
    }
    const author = element.dataset.changeAuthor?.trim() ?? '';
    const date = normalizedRevisionDate(element.dataset.changeDate);
    const before = element.dataset.changeBefore ?? '';
    if (
      !author ||
      author.length > 255 ||
      !parseDocumentTableFormatting(before)
    ) {
      throw new Error('Document contains an invalid table-formatting revision.');
    }
    if (this.patches.filter(Boolean).length >= MAX_TABLE_FORMATTING_CHANGE_PATCHES) {
      throw new Error('Document exceeds the table-formatting revision limit.');
    }
    this.patches.push({ id, author, date, before });
  }
}

export async function patchDocxTableFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly (DocxTableFormattingChangePatch | null)[],
  bidiVisual: readonly boolean[] = [],
  fills: readonly (string | null)[] = [],
  looks: readonly (DocumentTableLook | null)[] = [],
  overlaps: readonly (DocumentTableOverlap | null)[] = [],
  styleIds: readonly (string | null)[] = [],
  cellSpacings: readonly (number | null)[] = [],
): Promise<ArrayBuffer> {
  if (
    !patches.some(Boolean) &&
    !bidiVisual.some(Boolean) &&
    !fills.some(Boolean) &&
    !looks.some(Boolean) &&
    !overlaps.some(Boolean) &&
    !styleIds.some(Boolean) &&
    !cellSpacings.some((value) => value !== null && value !== undefined)
  ) {
    return buffer;
  }
  if (patches.filter(Boolean).length > MAX_TABLE_FORMATTING_CHANGE_PATCHES) {
    throw new Error('Document exceeds the table-formatting revision limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(
    decodeXmlBytes(
      await entry.async('uint8array'),
      'generated DOCX word/document.xml',
    ),
    'generated DOCX word/document.xml',
  );
  const tables = descendants(document, 'tbl').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  let changed = false;
  let index = 0;
  for (const table of tables) {
    const patch = patches[index] ?? null;
    const tableBidiVisual = bidiVisual[index] === true;
    const tableFill = fills[index] ?? null;
    const tableLook = looks[index] ?? null;
    const tableOverlap = overlaps[index] ?? null;
    const tableStyleId = styleIds[index] ?? null;
    const tableCellSpacing = cellSpacings[index] ?? null;
    index += 1;
    if (patch) {
      setTableFormattingChange(document, table, patch);
      changed = true;
    }
    if (setTableBidiVisual(document, table, tableBidiVisual)) {
      changed = true;
    }
    if (setTableFill(document, table, tableFill)) {
      changed = true;
    }
    if (setTableLook(document, table, tableLook)) {
      changed = true;
    }
    if (setTableOverlap(document, table, tableOverlap)) {
      changed = true;
    }
    if (setTableStyleId(document, table, tableStyleId)) {
      changed = true;
    }
    if (setTableCellSpacing(document, table, tableCellSpacing)) {
      changed = true;
    }
  }
  if (index !== patches.length) {
    throw new Error(
      `DOCX table-formatting revision patch count mismatch (${patches.length} patches, ${index} tables).`,
    );
  }
  if (bidiVisual.length && bidiVisual.length !== patches.length) {
    throw new Error(
      `DOCX table bidiVisual patch count mismatch (${bidiVisual.length} flags, ${patches.length} tables).`,
    );
  }
  if (fills.length && fills.length !== patches.length) {
    throw new Error(
      `DOCX table fill patch count mismatch (${fills.length} fills, ${patches.length} tables).`,
    );
  }
  if (looks.length && looks.length !== patches.length) {
    throw new Error(
      `DOCX table look patch count mismatch (${looks.length} looks, ${patches.length} tables).`,
    );
  }
  if (overlaps.length && overlaps.length !== patches.length) {
    throw new Error(
      `DOCX table overlap patch count mismatch (${overlaps.length} overlaps, ${patches.length} tables).`,
    );
  }
  if (styleIds.length && styleIds.length !== patches.length) {
    throw new Error(
      `DOCX table styleId patch count mismatch (${styleIds.length} styleIds, ${patches.length} tables).`,
    );
  }
  if (cellSpacings.length && cellSpacings.length !== patches.length) {
    throw new Error(
      `DOCX table cellSpacing patch count mismatch (${cellSpacings.length} cellSpacings, ${patches.length} tables).`,
    );
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function setTableFormattingChange(
  document: Document,
  table: Element,
  patch: DocxTableFormattingChangePatch,
): void {
  const formatting = parseDocumentTableFormatting(patch.before);
  if (!formatting) {
    throw new Error('Document contains an invalid table-formatting revision.');
  }
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblPrChange' &&
      child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', String(patch.id));
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', patch.author);
  if (patch.date) {
    change.setAttributeNS(WORD_NAMESPACE, 'w:date', patch.date);
  }
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  if (formatting.layout) {
    const layout = document.createElementNS(WORD_NAMESPACE, 'w:tblLayout');
    layout.setAttributeNS(WORD_NAMESPACE, 'w:type', formatting.layout);
    prior.append(layout);
  }
  if (formatting.alignment) {
    const jc = document.createElementNS(WORD_NAMESPACE, 'w:jc');
    jc.setAttributeNS(WORD_NAMESPACE, 'w:val', formatting.alignment);
    prior.append(jc);
  }
  if (formatting.width) {
    prior.append(createPreferredWidthElement(document, formatting.width));
  }
  if (formatting.indent !== undefined) {
    prior.append(createIndentElement(document, formatting.indent));
  }
  if (formatting.cellMargins) {
    prior.append(createTableCellMarginsElement(document, formatting.cellMargins));
  }
  if (formatting.bidiVisual !== undefined) {
    const bidiVisual = document.createElementNS(WORD_NAMESPACE, 'w:bidiVisual');
    if (!formatting.bidiVisual) {
      bidiVisual.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(bidiVisual);
  }
  if (formatting.fill !== undefined) {
    const shading = document.createElementNS(WORD_NAMESPACE, 'w:shd');
    shading.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
    shading.setAttributeNS(
      WORD_NAMESPACE,
      'w:fill',
      formatting.fill.replace(/^#/, '').toUpperCase(),
    );
    prior.append(shading);
  }
  if (formatting.look) {
    prior.append(createTblLookElement(document, formatting.look));
  }
  if (formatting.overlap) {
    prior.append(createTblOverlapElement(document, formatting.overlap));
  }
  if (formatting.styleId) {
    prior.append(createTblStyleElement(document, formatting.styleId));
  }
  if (formatting.cellSpacing !== undefined) {
    prior.append(createTblCellSpacingElement(document, formatting.cellSpacing));
  }
  change.append(prior);
  properties.append(change);
}

function setTableFill(
  document: Document,
  table: Element,
  fill: string | null,
): boolean {
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (!fill) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'shd' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!fill) return true;
  const shading = document.createElementNS(WORD_NAMESPACE, 'w:shd');
  shading.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
  shading.setAttributeNS(
    WORD_NAMESPACE,
    'w:fill',
    fill.replace(/^#/, '').toUpperCase(),
  );
  properties.append(shading);
  return true;
}

function setTableBidiVisual(
  document: Document,
  table: Element,
  bidiVisual: boolean,
): boolean {
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (!bidiVisual) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'bidiVisual' &&
      child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!bidiVisual) return true;
  properties.append(document.createElementNS(WORD_NAMESPACE, 'w:bidiVisual'));
  return true;
}

function setTableCellSpacing(
  document: Document,
  table: Element,
  cellSpacing: number | null,
): boolean {
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (cellSpacing === null) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblCellSpacing' &&
      child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (cellSpacing === null) return true;
  properties.append(createTblCellSpacingElement(document, cellSpacing));
  return true;
}

function createTblCellSpacingElement(
  document: Document,
  cellSpacing: number,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblCellSpacing');
  element.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:w',
    String(Math.round(cellSpacing * TWIPS_PER_PIXEL)),
  );
  return element;
}

function setTableStyleId(
  document: Document,
  table: Element,
  styleId: string | null,
): boolean {
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (!styleId) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblStyle' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!styleId) return true;
  properties.append(createTblStyleElement(document, styleId));
  return true;
}

function createTblStyleElement(
  document: Document,
  styleId: string,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblStyle');
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', styleId);
  return element;
}

function setTableOverlap(
  document: Document,
  table: Element,
  overlap: DocumentTableOverlap | null,
): boolean {
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (!overlap) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblOverlap' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!overlap) return true;
  properties.append(createTblOverlapElement(document, overlap));
  return true;
}

function createTblOverlapElement(
  document: Document,
  overlap: DocumentTableOverlap,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblOverlap');
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', overlap);
  return element;
}

function setTableLook(
  document: Document,
  table: Element,
  look: DocumentTableLook | null,
): boolean {
  let properties = directChild(table, 'tblPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (!look) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    table.insertBefore(properties, table.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblLook' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!look) return true;
  properties.append(createTblLookElement(document, look));
  return true;
}

function createTblLookElement(
  document: Document,
  look: DocumentTableLook,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblLook');
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    documentTableLookBitmask(look),
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:firstRow',
    look.firstRow ? '1' : '0',
  );
  element.setAttributeNS(WORD_NAMESPACE, 'w:lastRow', look.lastRow ? '1' : '0');
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:firstColumn',
    look.firstColumn ? '1' : '0',
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:lastColumn',
    look.lastColumn ? '1' : '0',
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:noHBand',
    look.noHorizontalBand ? '1' : '0',
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:noVBand',
    look.noVerticalBand ? '1' : '0',
  );
  return element;
}

function createPreferredWidthElement(
  document: Document,
  width: DocumentTablePreferredWidth,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblW');
  if (width.type === 'auto') {
    element.setAttributeNS(WORD_NAMESPACE, 'w:type', 'auto');
    element.setAttributeNS(WORD_NAMESPACE, 'w:w', '0');
    return element;
  }
  if (width.type === 'percent') {
    element.setAttributeNS(WORD_NAMESPACE, 'w:type', 'pct');
    element.setAttributeNS(
      WORD_NAMESPACE,
      'w:w',
      String(Math.round((width.value ?? 0) * 50)),
    );
    return element;
  }
  element.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:w',
    String(Math.round((width.value ?? 0) * TWIPS_PER_PIXEL)),
  );
  return element;
}

function createIndentElement(document: Document, indent: number): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblInd');
  element.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:w',
    String(Math.round(indent * TWIPS_PER_PIXEL)),
  );
  return element;
}

function createTableCellMarginsElement(
  document: Document,
  margins: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  },
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblCellMar');
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const value = margins[side];
    if (value === undefined) continue;
    const edge = document.createElementNS(WORD_NAMESPACE, `w:${side}`);
    edge.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
    edge.setAttributeNS(
      WORD_NAMESPACE,
      'w:w',
      String(Math.round(value * TWIPS_PER_PIXEL)),
    );
    element.append(edge);
  }
  return element;
}

function normalizedRevisionDate(value: string | undefined): string {
  if (!value?.trim()) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}
