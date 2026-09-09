import JSZip from 'jszip';
import { normalizeDocumentRowCnfStyle } from './work-document-row-cnf-style';
import { parseDocumentRowFormatting } from './work-document-row-format-changes';
import {
  normalizeDocumentTablePreferredWidth,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxRowFormattingChangePatch {
  id: number;
  author: string;
  date: string;
  before: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_ROW_FORMATTING_CHANGE_PATCHES = 4_096;
const TWIPS_PER_PIXEL = 1440 / 96;

export class DocxRowFormattingChangePatchCollector {
  readonly patches: Array<DocxRowFormattingChangePatch | null> = [];
  readonly hidden: boolean[] = [];
  readonly alignments: Array<'left' | 'center' | 'right' | null> = [];
  readonly gridBefore: Array<number | null> = [];
  readonly gridAfter: Array<number | null> = [];
  readonly widthBefore: Array<DocumentTablePreferredWidth | null> = [];
  readonly widthAfter: Array<DocumentTablePreferredWidth | null> = [];
  readonly cnfStyles: Array<string | null> = [];

  record(element: HTMLTableRowElement, id: number): void {
    this.hidden.push(element.dataset.officeRowHidden === 'true');
    const alignment = element.dataset.officeRowAlignment;
    this.alignments.push(
      alignment === 'left' || alignment === 'center' || alignment === 'right'
        ? alignment
        : null,
    );
    const gridBeforeRaw = element.dataset.officeRowGridBefore;
    const gridBeforeValue =
      gridBeforeRaw === undefined || gridBeforeRaw === ''
        ? null
        : Number(gridBeforeRaw);
    this.gridBefore.push(
      gridBeforeValue !== null &&
        Number.isInteger(gridBeforeValue) &&
        gridBeforeValue > 0
        ? gridBeforeValue
        : null,
    );
    const gridAfterRaw = element.dataset.officeRowGridAfter;
    const gridAfterValue =
      gridAfterRaw === undefined || gridAfterRaw === ''
        ? null
        : Number(gridAfterRaw);
    this.gridAfter.push(
      gridAfterValue !== null &&
        Number.isInteger(gridAfterValue) &&
        gridAfterValue > 0
        ? gridAfterValue
        : null,
    );
    this.widthBefore.push(
      preferredWidthFromRowElement(
        element,
        'officeRowWidthBeforeType',
        'officeRowWidthBefore',
      ),
    );
    this.widthAfter.push(
      preferredWidthFromRowElement(
        element,
        'officeRowWidthAfterType',
        'officeRowWidthAfter',
      ),
    );
    this.cnfStyles.push(
      normalizeDocumentRowCnfStyle(element.dataset.officeRowCnfStyle),
    );
    if (
      element.dataset.changeKind !== 'row-formatting' ||
      element.getAttribute('data-document-change') !== 'true'
    ) {
      this.patches.push(null);
      return;
    }
    const author = element.dataset.changeAuthor?.trim() ?? '';
    const date = normalizedRevisionDate(element.dataset.changeDate);
    const before = element.dataset.changeBefore ?? '';
    if (!author || author.length > 255 || !parseDocumentRowFormatting(before)) {
      throw new Error('Document contains an invalid row-formatting revision.');
    }
    if (
      this.patches.filter(Boolean).length >= MAX_ROW_FORMATTING_CHANGE_PATCHES
    ) {
      throw new Error('Document exceeds the row-formatting revision limit.');
    }
    this.patches.push({ id, author, date, before });
  }
}

export async function patchDocxRowFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly (DocxRowFormattingChangePatch | null)[],
  hidden: readonly boolean[] = [],
  alignments: readonly ('left' | 'center' | 'right' | null)[] = [],
  gridBefore: readonly (number | null)[] = [],
  gridAfter: readonly (number | null)[] = [],
  widthBefore: readonly (DocumentTablePreferredWidth | null)[] = [],
  widthAfter: readonly (DocumentTablePreferredWidth | null)[] = [],
  cnfStyles: readonly (string | null)[] = [],
): Promise<ArrayBuffer> {
  if (
    !patches.some(Boolean) &&
    !hidden.some(Boolean) &&
    !alignments.some(Boolean) &&
    !gridBefore.some((value) => value !== null && value > 0) &&
    !gridAfter.some((value) => value !== null && value > 0) &&
    !widthBefore.some((value) => value !== null) &&
    !widthAfter.some((value) => value !== null) &&
    !cnfStyles.some((value) => value !== null)
  ) {
    return buffer;
  }
  if (patches.filter(Boolean).length > MAX_ROW_FORMATTING_CHANGE_PATCHES) {
    throw new Error('Document exceeds the row-formatting revision limit.');
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
  const rows = descendants(document, 'tr').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  let changed = false;
  let index = 0;
  for (const row of rows) {
    const patch = patches[index] ?? null;
    const rowHidden = hidden[index] === true;
    const alignment = alignments[index] ?? null;
    const rowGridBefore = gridBefore[index] ?? null;
    const rowGridAfter = gridAfter[index] ?? null;
    const rowWidthBefore = widthBefore[index] ?? null;
    const rowWidthAfter = widthAfter[index] ?? null;
    const rowCnfStyle = cnfStyles[index] ?? null;
    index += 1;
    if (patch) {
      setRowFormattingChange(document, row, patch);
      changed = true;
    }
    if (rowHidden) {
      setRowHidden(document, row, true);
      changed = true;
    }
    if (alignment) {
      setRowAlignment(document, row, alignment);
      changed = true;
    }
    if (rowGridBefore !== null && rowGridBefore > 0) {
      setRowGridBefore(document, row, rowGridBefore);
      changed = true;
    }
    if (rowGridAfter !== null && rowGridAfter > 0) {
      setRowGridAfter(document, row, rowGridAfter);
      changed = true;
    }
    if (rowWidthBefore) {
      setRowPreferredWidth(document, row, 'wBefore', rowWidthBefore);
      changed = true;
    }
    if (rowWidthAfter) {
      setRowPreferredWidth(document, row, 'wAfter', rowWidthAfter);
      changed = true;
    }
    if (rowCnfStyle) {
      setRowCnfStyle(document, row, rowCnfStyle);
      changed = true;
    }
  }
  if (patches.length && index !== patches.length) {
    throw new Error(
      `DOCX row-formatting revision patch count mismatch (${patches.length} patches, ${index} rows).`,
    );
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function setRowHidden(document: Document, row: Element, hidden: boolean): void {
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'hidden' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!hidden) return;
  properties.append(document.createElementNS(WORD_NAMESPACE, 'w:hidden'));
}

function setRowGridBefore(
  document: Document,
  row: Element,
  gridBefore: number,
): void {
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'gridBefore' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (gridBefore <= 0) return;
  const element = document.createElementNS(WORD_NAMESPACE, 'w:gridBefore');
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', String(gridBefore));
  properties.append(element);
}

function setRowGridAfter(
  document: Document,
  row: Element,
  gridAfter: number,
): void {
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'gridAfter' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (gridAfter <= 0) return;
  const element = document.createElementNS(WORD_NAMESPACE, 'w:gridAfter');
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', String(gridAfter));
  properties.append(element);
}

function setRowAlignment(
  document: Document,
  row: Element,
  alignment: 'left' | 'center' | 'right',
): void {
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'jc' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const jc = document.createElementNS(WORD_NAMESPACE, 'w:jc');
  jc.setAttributeNS(WORD_NAMESPACE, 'w:val', alignment);
  properties.append(jc);
}

function setRowFormattingChange(
  document: Document,
  row: Element,
  patch: DocxRowFormattingChangePatch,
): void {
  const formatting = parseDocumentRowFormatting(patch.before);
  if (!formatting) {
    throw new Error('Document contains an invalid row-formatting revision.');
  }
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'trPrChange' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', String(patch.id));
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', patch.author);
  if (patch.date) {
    change.setAttributeNS(WORD_NAMESPACE, 'w:date', patch.date);
  }
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  if (formatting.cantSplit !== undefined) {
    const cantSplit = document.createElementNS(WORD_NAMESPACE, 'w:cantSplit');
    if (!formatting.cantSplit) {
      cantSplit.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(cantSplit);
  }
  if (formatting.repeatHeader !== undefined) {
    const header = document.createElementNS(WORD_NAMESPACE, 'w:tblHeader');
    if (!formatting.repeatHeader) {
      header.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(header);
  }
  if (formatting.height) {
    const height = document.createElementNS(WORD_NAMESPACE, 'w:trHeight');
    height.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      String(
        Math.max(1, Math.round(formatting.height.value * TWIPS_PER_PIXEL)),
      ),
    );
    if (formatting.height.rule === 'exact') {
      height.setAttributeNS(WORD_NAMESPACE, 'w:hRule', 'exact');
    } else {
      height.setAttributeNS(WORD_NAMESPACE, 'w:hRule', 'atLeast');
    }
    prior.append(height);
  }
  if (formatting.hidden !== undefined) {
    const hidden = document.createElementNS(WORD_NAMESPACE, 'w:hidden');
    if (!formatting.hidden) {
      hidden.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(hidden);
  }
  if (formatting.alignment !== undefined) {
    const jc = document.createElementNS(WORD_NAMESPACE, 'w:jc');
    jc.setAttributeNS(WORD_NAMESPACE, 'w:val', formatting.alignment);
    prior.append(jc);
  }
  if (formatting.gridBefore !== undefined) {
    const gridBefore = document.createElementNS(WORD_NAMESPACE, 'w:gridBefore');
    gridBefore.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      String(formatting.gridBefore),
    );
    prior.append(gridBefore);
  }
  if (formatting.gridAfter !== undefined) {
    const gridAfter = document.createElementNS(WORD_NAMESPACE, 'w:gridAfter');
    gridAfter.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      String(formatting.gridAfter),
    );
    prior.append(gridAfter);
  }
  if (formatting.widthBefore !== undefined) {
    prior.append(
      createPreferredWidthElement(document, 'wBefore', formatting.widthBefore),
    );
  }
  if (formatting.widthAfter !== undefined) {
    prior.append(
      createPreferredWidthElement(document, 'wAfter', formatting.widthAfter),
    );
  }
  if (formatting.cnfStyle !== undefined) {
    const cnfStyle = document.createElementNS(WORD_NAMESPACE, 'w:cnfStyle');
    cnfStyle.setAttributeNS(WORD_NAMESPACE, 'w:val', formatting.cnfStyle);
    prior.append(cnfStyle);
  }
  change.append(prior);
  properties.append(change);
}

function normalizedRevisionDate(value: string | undefined): string {
  if (!value?.trim()) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

function setRowCnfStyle(
  document: Document,
  row: Element,
  cnfStyle: string,
): void {
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'cnfStyle' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const element = document.createElementNS(WORD_NAMESPACE, 'w:cnfStyle');
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', cnfStyle);
  properties.append(element);
}

function setRowPreferredWidth(
  document: Document,
  row: Element,
  localName: 'wBefore' | 'wAfter',
  width: DocumentTablePreferredWidth,
): void {
  let properties = directChild(row, 'trPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    row.insertBefore(properties, row.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === localName && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  properties.append(createPreferredWidthElement(document, localName, width));
}

function createPreferredWidthElement(
  document: Document,
  localName: 'wBefore' | 'wAfter',
  width: DocumentTablePreferredWidth,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
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
    String(Math.max(1, Math.round((width.value ?? 0) * TWIPS_PER_PIXEL))),
  );
  return element;
}

function preferredWidthFromRowElement(
  element: HTMLTableRowElement,
  typeKey: string,
  valueKey: string,
): DocumentTablePreferredWidth | null {
  const type = element.dataset[typeKey];
  if (type === 'auto') return { type: 'auto', value: null };
  if (type === 'percent' || type === 'pixels') {
    const value = Number(element.dataset[valueKey]);
    return normalizeDocumentTablePreferredWidth({ type, value });
  }
  return null;
}
