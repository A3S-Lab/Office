import JSZip from 'jszip';
import { parseDocumentRowFormatting } from './work-document-row-format-changes';
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

  record(element: HTMLTableRowElement, id: number): void {
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
    if (
      !author ||
      author.length > 255 ||
      !parseDocumentRowFormatting(before)
    ) {
      throw new Error('Document contains an invalid row-formatting revision.');
    }
    if (this.patches.filter(Boolean).length >= MAX_ROW_FORMATTING_CHANGE_PATCHES) {
      throw new Error('Document exceeds the row-formatting revision limit.');
    }
    this.patches.push({ id, author, date, before });
  }
}

export async function patchDocxRowFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly (DocxRowFormattingChangePatch | null)[],
): Promise<ArrayBuffer> {
  if (!patches.some(Boolean)) return buffer;
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
    const patch = patches[index++] ?? null;
    if (!patch) continue;
    setRowFormattingChange(document, row, patch);
    changed = true;
  }
  if (index !== patches.length) {
    throw new Error(
      `DOCX row-formatting revision patch count mismatch (${patches.length} patches, ${index} rows).`,
    );
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
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
      child.localName === 'trPrChange' &&
      child.namespaceURI === WORD_NAMESPACE,
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
      String(Math.max(1, Math.round(formatting.height.value * TWIPS_PER_PIXEL))),
    );
    if (formatting.height.rule === 'exact') {
      height.setAttributeNS(WORD_NAMESPACE, 'w:hRule', 'exact');
    } else {
      height.setAttributeNS(WORD_NAMESPACE, 'w:hRule', 'atLeast');
    }
    prior.append(height);
  }
  change.append(prior);
  properties.append(change);
}

function normalizedRevisionDate(value: string | undefined): string {
  if (!value?.trim()) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}
