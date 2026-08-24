import {
  documentIndexEntryHtml,
  documentIndexHtml,
  type WorkDocumentIndexEntryDraft,
  type WorkDocumentIndexGeneratedEntry,
  type WorkDocumentIndexOptions,
} from './work-document-index';
import {
  parseDocumentIndexEntryInstruction,
  parseDocumentIndexInstruction,
} from './work-document-index-fields';
import {
  type DocxFieldOccurrence,
  docxFieldOccurrenceIsInlineEditable,
  docxFieldOccurrences,
} from './work-docx-field-instructions';
import { attribute, descendants, directChild } from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export interface ImportedDocxIndexMarkers {
  entries: ImportedDocxIndexEntryMarker[];
  indexes: ImportedDocxIndexMarker[];
}

export interface ImportedDocxIndexEntryMarker {
  start: string;
  end: string;
  id: string;
  value: WorkDocumentIndexEntryDraft;
}

export interface ImportedDocxIndexMarker {
  marker: string;
  id: string;
  options: WorkDocumentIndexOptions;
  entries: WorkDocumentIndexGeneratedEntry[];
}

export function markDocxIndexes(document: Document): ImportedDocxIndexMarkers {
  const fields = docxFieldOccurrences(document);
  const entries = fields.flatMap((field, index) => {
    const parsed = parseDocumentIndexEntryInstruction(field.instruction);
    if (!parsed.supported || !docxFieldOccurrenceIsInlineEditable(field)) {
      return [];
    }
    const marker: ImportedDocxIndexEntryMarker = {
      start: `__A3S_WORK_INDEX_ENTRY_START_${index + 1}__`,
      end: `__A3S_WORK_INDEX_ENTRY_END_${index + 1}__`,
      id: `index-entry-import-${index + 1}`,
      value: parsed.value,
    };
    insertFieldBoundaryMarkers(document, field, marker.start, marker.end);
    return [marker];
  });

  const indexes: ImportedDocxIndexMarker[] = [];
  for (const field of fields) {
    if (!supportedDocxIndexField(field)) continue;
    const parsed = parseDocumentIndexInstruction(field.instruction);
    if (!parsed.supported) continue;
    const control = closestAncestor(field.start, 'sdt');
    if (!control || !document.documentElement.contains(control)) continue;
    const index = indexes.length + 1;
    const marker: ImportedDocxIndexMarker = {
      marker: `__A3S_WORK_DOCUMENT_INDEX_${index}__`,
      id: `document-index-import-${index}`,
      options: {
        ...parsed.options,
        leader: importedIndexLeader(control, parsed.options),
      },
      entries: importedIndexEntries(control, index),
    };
    control.replaceWith(markerParagraph(document, marker.marker));
    indexes.push(marker);
  }
  return { entries, indexes };
}

export function applyImportedDocxIndexMarkers(
  document: Document,
  markers: ImportedDocxIndexMarkers,
): void {
  for (const entry of markers.entries) {
    const html = documentIndexEntryHtml({ id: entry.id, ...entry.value });
    const parsed = document.createRange().createContextualFragment(html);
    const replacement = parsed.firstElementChild;
    if (!(replacement instanceof HTMLElement)) continue;
    replaceMarkerRange(document.body, entry.start, entry.end, replacement);
  }

  const importedEntries = Array.from(
    document.body.querySelectorAll<HTMLElement>('[data-document-index-entry]'),
  );
  for (const index of markers.indexes) {
    const text = textNodes(document.body).find((node) =>
      node.data.includes(index.marker),
    );
    if (!text) continue;
    const entries = index.entries.map((entry, entryIndex) => {
      const targets = importedEntries.filter(
        (candidate) =>
          normalizedTerm(candidate.dataset.indexMainEntry) ===
            normalizedTerm(entry.mainEntry) &&
          normalizedTerm(candidate.dataset.indexSubEntry) ===
            normalizedTerm(entry.subEntry),
      );
      const targetIds = targets.flatMap((candidate) => {
        const id = candidate.dataset.indexEntryId?.trim();
        return id ? [id] : [];
      });
      return {
        ...entry,
        pages: entry.pages.map((page) => ({
          ...page,
          targetIds: targetIds.length
            ? targetIds
            : [`index-entry-import-${index.id}-${entryIndex + 1}`],
        })),
      };
    });
    const fragment = document.createRange().createContextualFragment(
      documentIndexHtml({
        id: index.id,
        options: index.options,
        entries,
      }),
    );
    closestHtmlBlock(text.parentElement)?.replaceWith(fragment);
  }
}

export function hasImportedDocxIndexMarkers(
  markers: ImportedDocxIndexMarkers,
): boolean {
  return markers.entries.length > 0 || markers.indexes.length > 0;
}

export function supportedDocxIndexEntryField(
  field: DocxFieldOccurrence,
): boolean {
  return (
    parseDocumentIndexEntryInstruction(field.instruction).supported &&
    docxFieldOccurrenceIsInlineEditable(field)
  );
}

export function supportedDocxIndexField(field: DocxFieldOccurrence): boolean {
  return (
    parseDocumentIndexInstruction(field.instruction).supported &&
    field.complete &&
    !field.nested &&
    !field.inDeletion &&
    Boolean(closestAncestor(field.start, 'sdt'))
  );
}

function importedIndexEntries(
  control: Element,
  indexNumber: number,
): WorkDocumentIndexGeneratedEntry[] {
  const entries: WorkDocumentIndexGeneratedEntry[] = [];
  let currentMain = '';
  let row = 0;
  for (const paragraph of descendants(control, 'p')) {
    const style = directChild(
      directChild(paragraph, 'pPr') ?? paragraph,
      'pStyle',
    );
    const styleName = attribute(style ?? paragraph, 'val') ?? '';
    const level = /^Index([12])$/i.exec(styleName)?.[1];
    if (!level) continue;
    row += 1;
    const parsed = importedIndexParagraph(paragraph, indexNumber, row);
    if (!parsed) continue;
    if (level === '1') {
      currentMain = parsed.term;
      if (!parsed.pages.length && !parsed.crossReference) continue;
      entries.push({
        mainEntry: parsed.term,
        subEntry: '',
        crossReference: parsed.crossReference,
        pages: parsed.pages,
      });
      continue;
    }
    if (!currentMain) continue;
    entries.push({
      mainEntry: currentMain,
      subEntry: parsed.term,
      crossReference: parsed.crossReference,
      pages: parsed.pages,
    });
  }
  return entries;
}

function importedIndexParagraph(
  paragraph: Element,
  indexNumber: number,
  row: number,
): {
  term: string;
  crossReference: string;
  pages: WorkDocumentIndexGeneratedEntry['pages'];
} | null {
  const textElements = descendants(paragraph, 't');
  const values = textElements.map((element) => element.textContent ?? '');
  const combined = values.join('').replace(/\s+/g, ' ').trim();
  if (!combined) return null;
  const crossReference = /^(.+?),\s*(?:see|参见)\s+(.+)$/i.exec(combined);
  if (crossReference?.[1] && crossReference[2]) {
    return {
      term: crossReference[1].trim(),
      crossReference: crossReference[2].trim(),
      pages: [],
    };
  }
  const pageElements = textElements.filter((element) =>
    /^\d{1,6}$/.test((element.textContent ?? '').trim()),
  );
  const pages = pageElements.map((element, pageIndex) => {
    const properties = directChild(
      closestAncestor(element, 'r') ?? element,
      'rPr',
    );
    return {
      pageNumber: Number((element.textContent ?? '').trim()),
      pageBold: Boolean(properties && directChild(properties, 'b')),
      pageItalic: Boolean(properties && directChild(properties, 'i')),
      targetIds: [`index-entry-import-${indexNumber}-${row}-${pageIndex + 1}`],
    };
  });
  const pageText = pageElements.map((element) => element.textContent ?? '');
  let term = combined;
  for (const value of pageText) {
    term = term.replace(
      new RegExp(`(?:,\\s*)?${escapeRegExp(value)}\\s*$`),
      '',
    );
  }
  term = term.replace(/[,\s]+$/, '').trim();
  return term ? { term, crossReference: '', pages } : null;
}

function importedIndexLeader(
  control: Element,
  options: WorkDocumentIndexOptions,
): WorkDocumentIndexOptions['leader'] {
  if (!options.rightAlignPageNumbers) return 'none';
  const tab = descendants(control, 'tab').find(
    (candidate) =>
      candidate.parentElement?.localName === 'tabs' &&
      attribute(candidate, 'val') === 'right',
  );
  const leader = attribute(tab ?? control, 'leader');
  if (leader === 'hyphen') return 'dash';
  if (leader === 'underscore') return 'underline';
  if (!leader || leader === 'none') return 'none';
  return 'dot';
}

function insertFieldBoundaryMarkers(
  document: Document,
  field: DocxFieldOccurrence,
  start: string,
  end: string,
): void {
  if (
    field.start.localName === 'fldChar' &&
    field.end.localName === 'fldChar'
  ) {
    field.start.parentNode?.insertBefore(
      markerText(document, start),
      field.start,
    );
    field.end.parentNode?.insertBefore(
      markerText(document, end),
      field.end.nextSibling,
    );
    return;
  }
  field.start.parentNode?.insertBefore(markerRun(document, start), field.start);
  field.end.parentNode?.insertBefore(
    markerRun(document, end),
    field.end.nextSibling,
  );
}

function markerParagraph(document: Document, marker: string): Element {
  const paragraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  paragraph.append(markerRun(document, marker));
  return paragraph;
}

function markerRun(document: Document, value: string): Element {
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  run.append(markerText(document, value));
  return run;
}

function markerText(document: Document, value: string): Element {
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = value;
  return text;
}

function replaceMarkerRange(
  root: HTMLElement,
  start: string,
  end: string,
  replacement: HTMLElement,
): void {
  const nodes = textNodes(root);
  const startNode = nodes.find((node) => node.data.includes(start));
  const endNode = nodes.find((node) => node.data.includes(end));
  if (!startNode || !endNode) return;
  const range = root.ownerDocument.createRange();
  range.setStart(startNode, startNode.data.indexOf(start));
  range.setEnd(endNode, endNode.data.indexOf(end) + end.length);
  range.deleteContents();
  range.insertNode(replacement);
}

function textNodes(root: ParentNode): Text[] {
  const walker = root.ownerDocument?.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
  );
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}

function closestHtmlBlock(element: Element | null): Element | null {
  let current = element;
  while (current && current.parentElement?.tagName.toLowerCase() !== 'body') {
    current = current.parentElement;
  }
  return current;
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function normalizedTerm(value: string | undefined): string {
  return (value ?? '')
    .normalize('NFKC')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
