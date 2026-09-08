import JSZip from 'jszip';
import { parseDocumentCellFormatting } from './work-document-cell-format-changes';
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

export class DocxCellFormattingChangePatchCollector {
  readonly patches: Array<DocxCellFormattingChangePatch | null> = [];

  record(element: HTMLTableCellElement, id: number): void {
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
): Promise<ArrayBuffer> {
  if (!patches.some(Boolean)) return buffer;
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
    const patch = patches[index++] ?? null;
    if (!patch) continue;
    setCellFormattingChange(document, cell, patch);
    changed = true;
  }
  if (index !== patches.length) {
    throw new Error(
      `DOCX cell-formatting revision patch count mismatch (${patches.length} patches, ${index} cells).`,
    );
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
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
      child.localName === 'tcPrChange' &&
      child.namespaceURI === WORD_NAMESPACE,
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
  change.append(prior);
  properties.append(change);
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
  const TWIPS_PER_PIXEL = 1440 / 96;
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
