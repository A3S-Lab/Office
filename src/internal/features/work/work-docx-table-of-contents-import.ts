import { normalizeDocumentParagraphId } from './work-document-paragraph-identity';
import {
  documentTableOfContentsHtml,
  parseDocumentTableOfContentsInstruction,
  type WorkDocumentTableOfContentsEntry,
  type WorkDocumentTableOfContentsOptions,
} from './work-document-table-of-contents';
import {
  docxFieldOccurrences,
  type DocxFieldOccurrence,
} from './work-docx-field-instructions';
import { attribute, descendants, directChild } from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export interface ImportedDocxTableOfContentsMarkers {
  tables: ImportedDocxTableOfContentsMarker[];
}

export interface ImportedDocxTableOfContentsMarker {
  marker: string;
  id: string;
  options: WorkDocumentTableOfContentsOptions;
  entries: WorkDocumentTableOfContentsEntry[];
}

export function markDocxTablesOfContents(
  document: Document,
): ImportedDocxTableOfContentsMarkers {
  const tables: ImportedDocxTableOfContentsMarker[] = [];
  for (const element of descendants(document, 'sdt')) {
    if (!document.documentElement.contains(element)) continue;
    const field = docxFieldOccurrences(element).find(
      supportedDocxTableOfContentsField,
    );
    if (!field) continue;
    const parsed = parseDocumentTableOfContentsInstruction(field.instruction);
    if (!parsed.supported) continue;
    const tableIndex = tables.length + 1;
    const marker: ImportedDocxTableOfContentsMarker = {
      marker: `__A3S_WORK_TABLE_OF_CONTENTS_${tableIndex}__`,
      id: `docx-table-of-contents-${tableIndex}`,
      options: {
        ...parsed.options,
        leader: importedTableOfContentsLeader(element, parsed.options),
      },
      entries: importedTableOfContentsEntries(element, tableIndex),
    };
    element.replaceWith(markerParagraph(document, marker.marker));
    tables.push(marker);
  }
  return { tables };
}

export function applyImportedDocxTableOfContentsMarkers(
  document: Document,
  markers: ImportedDocxTableOfContentsMarkers,
): void {
  const headings = importedHtmlHeadings(document);
  for (const table of markers.tables) {
    const usedHeadingIndexes = new Set<number>();
    const text = textNodes(document.body).find((node) =>
      node.data.includes(table.marker),
    );
    if (!text) continue;
    const entries = table.entries.map((entry) => {
      const headingIndex = headings.findIndex(
        (heading, index) =>
          !usedHeadingIndexes.has(index) &&
          heading.level === entry.level &&
          normalizedTitle(heading.title) === normalizedTitle(entry.title),
      );
      if (headingIndex < 0) return entry;
      usedHeadingIndexes.add(headingIndex);
      const heading = headings[headingIndex];
      return heading ? { ...entry, targetId: heading.targetId } : entry;
    });
    const fragment = document.createRange().createContextualFragment(
      documentTableOfContentsHtml({
        id: table.id,
        options: table.options,
        entries,
      }),
    );
    const block = closestHtmlBlock(text.parentElement);
    if (block) block.replaceWith(fragment);
  }
}

export function hasImportedDocxTableOfContentsMarkers(
  markers: ImportedDocxTableOfContentsMarkers,
): boolean {
  return markers.tables.length > 0;
}

export function supportedDocxTableOfContentsField(
  field: DocxFieldOccurrence,
): boolean {
  const parsed = parseDocumentTableOfContentsInstruction(field.instruction);
  return (
    parsed.supported &&
    field.complete &&
    !field.nested &&
    !field.inDeletion &&
    Boolean(closestAncestor(field.start, 'sdt'))
  );
}

function importedTableOfContentsEntries(
  element: Element,
  tableIndex: number,
): WorkDocumentTableOfContentsEntry[] {
  const paragraphs = descendants(element, 'p');
  return paragraphs.flatMap((paragraph, index) => {
    const style = directChild(
      directChild(paragraph, 'pPr') ?? paragraph,
      'pStyle',
    );
    const level = Number(
      /^TOC([1-9])$/i.exec(attribute(style ?? paragraph, 'val') ?? '')?.[1],
    );
    if (!level) return [];
    const values = descendants(paragraph, 't')
      .map((text) => text.textContent ?? '')
      .filter(Boolean);
    const last = values.at(-1)?.trim() ?? '';
    const hasPageNumber = /^\d{1,6}$/.test(last);
    const title = (hasPageNumber ? values.slice(0, -1) : values)
      .join('')
      .replace(/\s+/g, ' ')
      .trim();
    if (!title) return [];
    const hyperlink = descendants(paragraph, 'hyperlink').find((candidate) =>
      Boolean(attribute(candidate, 'anchor')),
    );
    const anchor = attribute(hyperlink ?? paragraph, 'anchor')?.trim();
    const targetId = /^heading-[a-z0-9._:-]{1,128}$/i.test(anchor ?? '')
      ? (anchor as string)
      : `heading-import-${tableIndex}-${index + 1}`;
    return [
      {
        targetId,
        title,
        level,
        pageNumber: hasPageNumber ? Number(last) : 1,
      },
    ];
  });
}

function importedTableOfContentsLeader(
  element: Element,
  options: WorkDocumentTableOfContentsOptions,
): WorkDocumentTableOfContentsOptions['leader'] {
  if (!options.showPageNumbers || !options.rightAlignPageNumbers) return 'none';
  const tab = descendants(element, 'tab').find(
    (candidate) =>
      candidate.parentElement?.localName === 'tabs' &&
      attribute(candidate, 'val') === 'right',
  );
  const leader = attribute(tab ?? element, 'leader');
  if (leader === 'hyphen') return 'dash';
  if (leader === 'underscore') return 'underline';
  if (!leader || leader === 'none') return 'none';
  return 'dot';
}

function markerParagraph(document: Document, marker: string): Element {
  const paragraph = document.createElementNS(WORD_NAMESPACE, 'w:p');
  const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.append(run);
  return paragraph;
}

function importedHtmlHeadings(document: Document): Array<{
  title: string;
  level: number;
  targetId: string;
}> {
  return Array.from(
    document.body.querySelectorAll<HTMLElement>(
      'h1,h2,h3,h4,h5,h6,p[data-office-outline-level]',
    ),
  ).flatMap((heading, index) => {
    const headingLevel = /^H([1-6])$/.exec(heading.tagName)?.[1];
    const outlineLevel = Number(heading.dataset.officeOutlineLevel);
    const level = headingLevel
      ? Number(headingLevel)
      : Number.isInteger(outlineLevel) && outlineLevel >= 0 && outlineLevel <= 8
        ? outlineLevel + 1
        : null;
    if (level === null) return [];
    const identity = normalizeDocumentParagraphId(
      heading.dataset.officeParagraphId,
    );
    return [
      {
        title: heading.textContent ?? '',
        level,
        targetId: identity
          ? `heading-${identity.toLowerCase()}`
          : `heading-imported-${index + 1}`,
      },
    ];
  });
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

function normalizedTitle(value: string): string {
  return value
    .normalize('NFKC')
    .replace(/\s+/g, ' ')
    .trim()
    .toLocaleLowerCase();
}
