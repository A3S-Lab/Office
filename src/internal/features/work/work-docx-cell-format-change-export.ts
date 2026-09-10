import JSZip from 'jszip';
import { parseDocumentCellFormatting } from './work-document-cell-format-changes';
import { normalizeDocumentCnfStyle } from './work-document-cnf-style';
import {
  documentTableBordersFromElement,
  normalizeDocumentTableBorderStyle,
  normalizeDocumentTableBorderWidth,
  normalizeTableColor,
  type DocumentTableBorder,
} from './work-document-table-borders';
import type { DocumentTablePreferredWidth } from './work-document-table-geometry';
import {
  DOCUMENT_CELL_FORMATTING_BORDER_EDGES,
  docxSzFromSnapshotBorderWidth,
  mapSnapshotBorderStyleToDocx,
  normalizeDocumentCellFormattingBorders,
  orderedDocumentCellFormattingBorders,
  type DocumentCellFormattingBorders,
} from './work-document-table-formatting-borders';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxCellFormattingChangePatch {
  id: number;
  author: string;
  date: string;
  before: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_CELL_FORMATTING_CHANGE_PATCHES = 4_096;
const TWIPS_PER_PIXEL = 1440 / 96;

export class DocxCellFormattingChangePatchCollector {
  readonly patches: Array<DocxCellFormattingChangePatch | null> = [];
  readonly noWrap: boolean[] = [];
  readonly textDirection: Array<string | null> = [];
  readonly fitText: boolean[] = [];
  readonly hideMark: boolean[] = [];
  readonly cnfStyles: Array<string | null> = [];
  readonly borders: Array<DocumentCellFormattingBorders | null> = [];

  record(element: HTMLTableCellElement, id: number): void {
    this.noWrap.push(element.dataset.officeCellNoWrap === 'true');
    this.textDirection.push(
      element.dataset.officeCellTextDirection?.trim() || null,
    );
    this.fitText.push(element.dataset.officeCellFitText === 'true');
    this.hideMark.push(element.dataset.officeCellHideMark === 'true');
    this.cnfStyles.push(
      normalizeDocumentCnfStyle(element.dataset.officeCellCnfStyle),
    );
    this.borders.push(cellFormattingBordersFromElement(element));
    if (
      element.dataset.changeKind !== 'cell-formatting' ||
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
      !parseDocumentCellFormatting(before)
    ) {
      throw new Error('Document contains an invalid cell-formatting revision.');
    }
    if (
      this.patches.filter(Boolean).length >= MAX_CELL_FORMATTING_CHANGE_PATCHES
    ) {
      throw new Error('Document exceeds the cell-formatting revision limit.');
    }
    this.patches.push({ id, author, date, before });
  }
}

export async function patchDocxCellFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly (DocxCellFormattingChangePatch | null)[],
  noWrap: readonly boolean[] = [],
  textDirection: readonly (string | null)[] = [],
  fitText: readonly boolean[] = [],
  hideMark: readonly boolean[] = [],
  cnfStyles: readonly (string | null)[] = [],
  borders: readonly (DocumentCellFormattingBorders | null)[] = [],
): Promise<ArrayBuffer> {
  if (
    !patches.some(Boolean) &&
    !noWrap.some(Boolean) &&
    !textDirection.some(Boolean) &&
    !fitText.some(Boolean) &&
    !hideMark.some(Boolean) &&
    !cnfStyles.some((value) => value !== null) &&
    !borders.some(Boolean)
  ) {
    return buffer;
  }
  if (patches.filter(Boolean).length > MAX_CELL_FORMATTING_CHANGE_PATCHES) {
    throw new Error('Document exceeds the cell-formatting revision limit.');
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
  const cells = descendants(document, 'tc').filter(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  let changed = false;
  let index = 0;
  for (const cell of cells) {
    const patch = patches[index] ?? null;
    const cellNoWrap = noWrap[index] === true;
    const cellTextDirection = textDirection[index] ?? null;
    const cellFitText = fitText[index] === true;
    const cellHideMark = hideMark[index] === true;
    const cellCnfStyle = cnfStyles[index] ?? null;
    const cellBorders = borders[index] ?? null;
    index += 1;
    if (patch) {
      setCellFormattingChange(document, cell, patch);
      changed = true;
    }
    if (cellNoWrap) {
      setCellNoWrap(document, cell, true);
      changed = true;
    }
    if (cellTextDirection) {
      setCellTextDirection(document, cell, cellTextDirection);
      changed = true;
    }
    if (cellFitText) {
      setCellFitText(document, cell, true);
      changed = true;
    }
    if (cellHideMark) {
      setCellHideMark(document, cell, true);
      changed = true;
    }
    if (cellCnfStyle) {
      setCellCnfStyle(document, cell, cellCnfStyle);
      changed = true;
    }
    if (setCellBorders(document, cell, cellBorders)) {
      changed = true;
    }
  }
  if (patches.length && index !== patches.length) {
    throw new Error(
      `DOCX cell-formatting revision patch count mismatch (${patches.length} patches, ${index} cells).`,
    );
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function setCellNoWrap(
  document: Document,
  cell: Element,
  noWrap: boolean,
): void {
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'noWrap' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!noWrap) return;
  properties.append(document.createElementNS(WORD_NAMESPACE, 'w:noWrap'));
}

function setCellTextDirection(
  document: Document,
  cell: Element,
  textDirection: string,
): void {
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'textDirection' &&
      child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const element = document.createElementNS(WORD_NAMESPACE, 'w:textDirection');
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', textDirection);
  properties.append(element);
}

function setCellFitText(
  document: Document,
  cell: Element,
  fitText: boolean,
): void {
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tcFitText' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!fitText) return;
  properties.append(document.createElementNS(WORD_NAMESPACE, 'w:tcFitText'));
}

function setCellHideMark(
  document: Document,
  cell: Element,
  hideMark: boolean,
): void {
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'hideMark' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!hideMark) return;
  properties.append(document.createElementNS(WORD_NAMESPACE, 'w:hideMark'));
}

function setCellCnfStyle(
  document: Document,
  cell: Element,
  cnfStyle: string,
): void {
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
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

function setCellFormattingChange(
  document: Document,
  cell: Element,
  patch: DocxCellFormattingChangePatch,
): void {
  const formatting = parseDocumentCellFormatting(patch.before);
  if (!formatting) {
    throw new Error('Document contains an invalid cell-formatting revision.');
  }
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tcPrChange' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', String(patch.id));
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', patch.author);
  if (patch.date) {
    change.setAttributeNS(WORD_NAMESPACE, 'w:date', patch.date);
  }
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  if (formatting.verticalAlign !== undefined) {
    const vAlign = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
    vAlign.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      formatting.verticalAlign === 'middle'
        ? 'center'
        : formatting.verticalAlign,
    );
    prior.append(vAlign);
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
  if (formatting.margins !== undefined) {
    prior.append(createCellMarginsElement(document, formatting.margins));
  }
  if (formatting.width !== undefined) {
    prior.append(createPreferredWidthElement(document, formatting.width));
  }
  if (formatting.noWrap !== undefined) {
    const noWrap = document.createElementNS(WORD_NAMESPACE, 'w:noWrap');
    if (!formatting.noWrap) {
      noWrap.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(noWrap);
  }
  if (formatting.textDirection !== undefined) {
    const textDirection = document.createElementNS(
      WORD_NAMESPACE,
      'w:textDirection',
    );
    textDirection.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      formatting.textDirection,
    );
    prior.append(textDirection);
  }
  if (formatting.fitText !== undefined) {
    const fitText = document.createElementNS(WORD_NAMESPACE, 'w:tcFitText');
    if (!formatting.fitText) {
      fitText.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(fitText);
  }
  if (formatting.hideMark !== undefined) {
    const hideMark = document.createElementNS(WORD_NAMESPACE, 'w:hideMark');
    if (!formatting.hideMark) {
      hideMark.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    prior.append(hideMark);
  }
  if (formatting.cnfStyle !== undefined) {
    const cnfStyle = document.createElementNS(WORD_NAMESPACE, 'w:cnfStyle');
    cnfStyle.setAttributeNS(WORD_NAMESPACE, 'w:val', formatting.cnfStyle);
    prior.append(cnfStyle);
  }
  if (formatting.borders !== undefined) {
    prior.append(createTcBordersElement(document, formatting.borders));
  }
  change.append(prior);
  properties.append(change);
}

function setCellBorders(
  document: Document,
  cell: Element,
  borders: DocumentCellFormattingBorders | null,
): boolean {
  const normalized = normalizeDocumentCellFormattingBorders(borders);
  let properties = directChild(cell, 'tcPr');
  if (!properties || properties.namespaceURI !== WORD_NAMESPACE) {
    if (!normalized) return false;
    properties = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    cell.insertBefore(properties, cell.firstChild);
  }
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tcBorders' && child.namespaceURI === WORD_NAMESPACE,
  )) {
    existing.remove();
  }
  if (!normalized) return true;
  properties.append(createTcBordersElement(document, normalized));
  return true;
}

function createTcBordersElement(
  document: Document,
  borders: DocumentCellFormattingBorders,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tcBorders');
  const ordered = orderedDocumentCellFormattingBorders(borders);
  for (const edge of DOCUMENT_CELL_FORMATTING_BORDER_EDGES) {
    const border = ordered[edge];
    if (!border) continue;
    const child = document.createElementNS(WORD_NAMESPACE, `w:${edge}`);
    child.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      mapSnapshotBorderStyleToDocx(border.style),
    );
    child.setAttributeNS(
      WORD_NAMESPACE,
      'w:sz',
      border.style === 'none' ? '0' : docxSzFromSnapshotBorderWidth(border.width),
    );
    child.setAttributeNS(WORD_NAMESPACE, 'w:space', '0');
    child.setAttributeNS(
      WORD_NAMESPACE,
      'w:color',
      border.style === 'none'
        ? 'auto'
        : border.color.replace(/^#/, '').toUpperCase(),
    );
    element.append(child);
  }
  return element;
}

function cellFormattingBordersFromElement(
  cell: HTMLTableCellElement,
): DocumentCellFormattingBorders | null {
  if (!hasDocumentTableCellBorderPresentation(cell)) return null;
  const fallback = defaultCellBorderFromElement(cell);
  const full = documentTableBordersFromElement(cell, fallback);
  return normalizeDocumentCellFormattingBorders(full);
}

function hasDocumentTableCellBorderPresentation(
  cell: HTMLTableCellElement,
): boolean {
  return Boolean(
    Object.keys(cell.dataset).some((key) =>
      key.startsWith('officeCellBorder'),
    ) ||
      cell.style.borderStyle ||
      cell.style.borderTopStyle ||
      cell.style.borderRightStyle ||
      cell.style.borderBottomStyle ||
      cell.style.borderLeftStyle,
  );
}

function defaultCellBorderFromElement(
  cell: HTMLTableCellElement,
): DocumentTableBorder {
  const legacyStyle = normalizeDocumentTableBorderStyle(
    cell.dataset.officeCellBorderStyle || cell.style.borderStyle,
  );
  return {
    color:
      normalizeTableColor(
        cell.dataset.officeCellBorderColor || cell.style.borderColor,
      ) ?? '#cfd5df',
    style: legacyStyle ?? 'solid',
    width:
      legacyStyle === 'none'
        ? 0
        : (normalizeDocumentTableBorderWidth(
            cell.dataset.officeCellBorderWidth || cell.style.borderWidth,
          ) ?? 1),
  };
}

function createPreferredWidthElement(
  document: Document,
  width: DocumentTablePreferredWidth,
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tcW');
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

function createCellMarginsElement(
  document: Document,
  margins: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  },
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tcMar');
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
