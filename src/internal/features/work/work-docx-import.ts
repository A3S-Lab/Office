import JSZip from 'jszip';
import { normalizeDocumentBookmarkReferencesHtml } from './work-document-bookmark-references';
import { normalizeDocumentBookmarksHtml } from './work-document-bookmarks';
import { normalizeDocumentCaptionsHtml } from './work-document-captions';
import { normalizeDocumentCitationsHtml } from './work-document-citations';
import { normalizeDocumentFieldsHtml } from './work-document-fields';
import { normalizeDocumentNotesHtml } from './work-document-notes';
import {
  documentContentLayoutProperties,
  documentInitialSectionLayout,
  documentSectionDomAttributes,
} from './work-document-section';
import { readDocxBibliography } from './work-docx-bibliography';
import {
  applyImportedDocxBookmarkMarkers,
  hasImportedDocxBookmarkMarkers,
  type ImportedDocxBookmarkMarkers,
  markDocxBookmarks,
} from './work-docx-bookmark-import';
import {
  applyImportedDocxCaptionMarkers,
  hasImportedDocxCaptionMarkers,
  type ImportedDocxCaptionMarkers,
  markDocxCaptionFields,
} from './work-docx-caption-import';
import {
  applyImportedDocxChangeMarkers,
  hasImportedDocxChangeMarkers,
  type ImportedDocxChangeMarkers,
  markDocxTextChanges,
} from './work-docx-change-import';
import {
  applyImportedDocxCitationMarkers,
  hasImportedDocxCitationMarkers,
  type ImportedDocxCitationMarkers,
  markDocxCitationFields,
} from './work-docx-citation-import';
import { importDocxColumns } from './work-docx-column-import';
import {
  applyImportedDocxCommentMarkers,
  hasImportedDocxCommentMarkers,
  type ImportedDocxCommentMarkers,
  markDocxComments,
} from './work-docx-comment-import';
import {
  applyImportedDocxFieldMarkers,
  hasImportedDocxFieldMarkers,
  type ImportedDocxFieldMarkers,
  markDocxBodyFields,
} from './work-docx-field-import';
import {
  applyImportedDocxImageLayoutMarkers,
  hasImportedDocxImageLayoutMarkers,
  type ImportedDocxImageLayoutMarkers,
  markDocxImageLayouts,
} from './work-docx-image-layout-import';
import {
  applyImportedDocxListMarkers,
  hasImportedDocxListMarkers,
  type ImportedDocxListMarkers,
  markDocxLists,
} from './work-docx-list-import';
import {
  extractMammothDocumentNotes,
  placeMammothDocumentNotes,
} from './work-docx-note-import';
import {
  documentUsesOddEvenPageChrome,
  importSectionPageChrome,
} from './work-docx-page-chrome-import';
import { importDocxPageColor } from './work-docx-page-color';
import {
  applyImportedDocxParagraphAlignmentMarkers,
  hasImportedDocxParagraphAlignmentMarkers,
  type ImportedDocxParagraphAlignmentMarkers,
  markDocxParagraphAlignments,
} from './work-docx-paragraph-alignment-import';
import {
  applyImportedDocxParagraphDirectionMarkers,
  hasImportedDocxParagraphDirectionMarkers,
  type ImportedDocxParagraphDirectionMarkers,
  markDocxParagraphDirections,
} from './work-docx-paragraph-direction-import';
import {
  applyImportedDocxParagraphIndentMarkers,
  hasImportedDocxParagraphIndentMarkers,
  type ImportedDocxParagraphIndentMarkers,
  markDocxParagraphIndents,
} from './work-docx-paragraph-indent-import';
import {
  applyImportedDocxParagraphIdentityMarkers,
  hasImportedDocxParagraphIdentityMarkers,
  type ImportedDocxParagraphIdentityMarkers,
  markDocxParagraphIdentities,
} from './work-docx-paragraph-identity-import';
import {
  applyImportedDocxParagraphPaginationMarkers,
  hasImportedDocxParagraphPaginationMarkers,
  type ImportedDocxParagraphPaginationMarkers,
  markDocxParagraphPagination,
} from './work-docx-paragraph-pagination-import';
import {
  applyImportedDocxParagraphSpacingMarkers,
  hasImportedDocxParagraphSpacingMarkers,
  type ImportedDocxParagraphSpacingMarkers,
  markDocxParagraphSpacing,
} from './work-docx-paragraph-spacing-import';
import { createDocxParagraphStyleResolver } from './work-docx-paragraph-styles';
import {
  applyImportedDocxRunFormattingMarkers,
  hasImportedDocxRunFormattingMarkers,
  type ImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from './work-docx-run-formatting-import';
import {
  applyImportedDocxParagraphTabStopMarkers,
  hasImportedDocxParagraphTabStopMarkers,
  type ImportedDocxParagraphTabStopMarkers,
  markDocxParagraphTabStops,
} from './work-docx-tab-stop-import';
import {
  applyImportedDocxTableCellMarkers,
  hasImportedDocxTableCellMarkers,
  type ImportedDocxTableCellMarkers,
  markDocxTableCells,
} from './work-docx-table-cell-import';
import {
  applyImportedDocxTableRowMarkers,
  hasImportedDocxTableRowMarkers,
  type ImportedDocxTableRowMarkers,
  markDocxTableRows,
} from './work-docx-table-row-import';
import {
  applyImportedDocxTableSizingMarkers,
  hasImportedDocxTableSizingMarkers,
  type ImportedDocxTableSizingMarkers,
  markDocxTableSizing,
} from './work-docx-table-sizing-import';
import { createDocxTableStyleResolver } from './work-docx-table-styles';
import {
  attribute,
  descendants,
  directChild,
  firstDescendant,
  OoxmlPackage,
} from './work-ooxml-package';
import type {
  WorkDocumentContent,
  WorkDocumentGrid,
  WorkDocumentGridType,
  WorkDocumentMargins,
  WorkDocumentSectionBreakType,
  WorkDocumentSectionLayout,
} from './work-types';

type ImportedDocumentLayout = Omit<WorkDocumentContent, 'type' | 'html'>;

export interface PreparedDocxImport {
  conversionBuffer: ArrayBuffer;
  sections: Array<{ id: string; layout: WorkDocumentSectionLayout }>;
  pageColor?: string;
  captionMarkers: ImportedDocxCaptionMarkers;
  bookmarkMarkers: ImportedDocxBookmarkMarkers;
  changeMarkers: ImportedDocxChangeMarkers;
  commentMarkers: ImportedDocxCommentMarkers;
  fieldMarkers: ImportedDocxFieldMarkers;
  citationMarkers: ImportedDocxCitationMarkers;
  listMarkers: ImportedDocxListMarkers;
  imageLayoutMarkers: ImportedDocxImageLayoutMarkers;
  paragraphIdentityMarkers: ImportedDocxParagraphIdentityMarkers;
  paragraphAlignmentMarkers: ImportedDocxParagraphAlignmentMarkers;
  paragraphDirectionMarkers: ImportedDocxParagraphDirectionMarkers;
  paragraphIndentMarkers: ImportedDocxParagraphIndentMarkers;
  paragraphPaginationMarkers: ImportedDocxParagraphPaginationMarkers;
  paragraphSpacingMarkers: ImportedDocxParagraphSpacingMarkers;
  runFormattingMarkers: ImportedDocxRunFormattingMarkers;
  tabStopMarkers: ImportedDocxParagraphTabStopMarkers;
  tableCellMarkers: ImportedDocxTableCellMarkers;
  tableRowMarkers: ImportedDocxTableRowMarkers;
  tableSizingMarkers: ImportedDocxTableSizingMarkers;
  bibliography?: WorkDocumentContent['bibliography'];
  trackChanges: boolean;
}

const TWIPS_PER_MILLIMETER = 1440 / 25.4;
const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export async function prepareDocxImport(
  buffer: ArrayBuffer,
): Promise<PreparedDocxImport> {
  const archive = await OoxmlPackage.load(buffer);
  const bibliography = (await readDocxBibliography(archive)).bibliography;
  const fallback = documentInitialSectionLayout({
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  });
  if (!archive.has('word/document.xml')) {
    return {
      conversionBuffer: buffer,
      sections: [{ id: 'document-section-1', layout: fallback }],
      captionMarkers: { captions: [], references: [] },
      bookmarkMarkers: { bookmarks: [], references: [] },
      changeMarkers: { changes: [] },
      commentMarkers: { comments: [], ranges: [] },
      fieldMarkers: { fields: [] },
      citationMarkers: { citations: [], bibliographies: [] },
      listMarkers: { lists: [] },
      imageLayoutMarkers: { images: [] },
      paragraphIdentityMarkers: { paragraphs: [] },
      paragraphAlignmentMarkers: { paragraphs: [] },
      paragraphDirectionMarkers: { paragraphs: [] },
      paragraphIndentMarkers: { paragraphs: [] },
      paragraphPaginationMarkers: { paragraphs: [] },
      paragraphSpacingMarkers: { paragraphs: [] },
      runFormattingMarkers: { runs: [] },
      tabStopMarkers: { paragraphs: [], inlineTabs: [] },
      tableCellMarkers: { cells: [] },
      tableRowMarkers: { rows: [] },
      tableSizingMarkers: { tables: [] },
      bibliography,
      trackChanges: false,
    };
  }

  const document = await archive.xml('word/document.xml');
  const pageColor = importDocxPageColor(document);
  const paragraphIdentityMarkers = markDocxParagraphIdentities(document);
  const numbering = archive.has('word/numbering.xml')
    ? await archive.xml('word/numbering.xml')
    : null;
  const commentMarkers = await markDocxComments(document, archive);
  const changeMarkers = markDocxTextChanges(document);
  const captionMarkers = markDocxCaptionFields(document);
  const bookmarkMarkers = markDocxBookmarks(document);
  const citationMarkers = markDocxCitationFields(document);
  const fieldMarkers = markDocxBodyFields(document);
  const listMarkers = markDocxLists(document, numbering);
  const imageLayoutMarkers = markDocxImageLayouts(document);
  const paragraphStylesDocument = archive.has('word/styles.xml')
    ? await archive.xml('word/styles.xml')
    : null;
  const paragraphStyles = createDocxParagraphStyleResolver(
    paragraphStylesDocument,
  );
  const tableStyles = createDocxTableStyleResolver(paragraphStylesDocument);
  const paragraphDirectionMarkers = markDocxParagraphDirections(
    document,
    paragraphStyles,
    tableStyles,
  );
  const paragraphAlignmentMarkers = markDocxParagraphAlignments(
    document,
    paragraphStyles,
    tableStyles,
  );
  const paragraphIndentMarkers = markDocxParagraphIndents(
    document,
    paragraphStyles,
    tableStyles,
  );
  const paragraphSpacingMarkers = markDocxParagraphSpacing(
    document,
    paragraphStyles,
    tableStyles,
  );
  const paragraphPaginationMarkers = markDocxParagraphPagination(
    document,
    paragraphStyles,
    tableStyles,
  );
  const themePath = archive
    .paths('word/theme/')
    .find((path) => /\/theme\d*\.xml$/i.test(path));
  const themeDocument = themePath ? await archive.xml(themePath) : null;
  const runFormattingMarkers = markDocxRunFormatting(
    document,
    paragraphStyles,
    themeDocument,
    tableStyles,
  );
  const tabStopMarkers = markDocxParagraphTabStops(
    document,
    paragraphStyles,
    tableStyles,
  );
  const tableCellMarkers = markDocxTableCells(
    document,
    themeDocument,
    tableStyles,
  );
  const tableRowMarkers = markDocxTableRows(document);
  const tableSizingMarkers = markDocxTableSizing(document, tableStyles);
  const settings = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const trackChanges =
    Boolean(settings && firstDescendant(settings, 'trackRevisions')) ||
    changeMarkers.changes.length > 0;
  const sectionElements = effectiveSectionProperties(document);
  if (!sectionElements.length) {
    return {
      conversionBuffer:
        hasImportedDocxCaptionMarkers(captionMarkers) ||
        hasImportedDocxBookmarkMarkers(bookmarkMarkers) ||
        hasImportedDocxChangeMarkers(changeMarkers) ||
        hasImportedDocxCommentMarkers(commentMarkers) ||
        hasImportedDocxCitationMarkers(citationMarkers) ||
        hasImportedDocxFieldMarkers(fieldMarkers) ||
        hasImportedDocxListMarkers(listMarkers) ||
        hasImportedDocxImageLayoutMarkers(imageLayoutMarkers) ||
        hasImportedDocxParagraphIdentityMarkers(paragraphIdentityMarkers) ||
        hasImportedDocxParagraphAlignmentMarkers(paragraphAlignmentMarkers) ||
        hasImportedDocxParagraphDirectionMarkers(paragraphDirectionMarkers) ||
        hasImportedDocxParagraphIndentMarkers(paragraphIndentMarkers) ||
        hasImportedDocxParagraphSpacingMarkers(paragraphSpacingMarkers) ||
        hasImportedDocxParagraphPaginationMarkers(paragraphPaginationMarkers) ||
        hasImportedDocxRunFormattingMarkers(runFormattingMarkers) ||
        hasImportedDocxParagraphTabStopMarkers(tabStopMarkers) ||
        hasImportedDocxTableCellMarkers(tableCellMarkers) ||
        hasImportedDocxTableRowMarkers(tableRowMarkers) ||
        hasImportedDocxTableSizingMarkers(tableSizingMarkers)
          ? await writeDocumentXml(buffer, document)
          : buffer,
      sections: [{ id: 'document-section-1', layout: fallback }],
      pageColor,
      captionMarkers,
      bookmarkMarkers,
      changeMarkers,
      commentMarkers,
      fieldMarkers,
      citationMarkers,
      listMarkers,
      imageLayoutMarkers,
      paragraphIdentityMarkers,
      paragraphAlignmentMarkers,
      paragraphDirectionMarkers,
      paragraphIndentMarkers,
      paragraphPaginationMarkers,
      paragraphSpacingMarkers,
      runFormattingMarkers,
      tabStopMarkers,
      tableCellMarkers,
      tableRowMarkers,
      tableSizingMarkers,
      bibliography,
      trackChanges,
    };
  }
  const relationships = await archive.relationships('word/document.xml');
  const oddEvenPageChrome = documentUsesOddEvenPageChrome(settings);
  const sections: PreparedDocxImport['sections'] = [];
  let previous = fallback;
  for (const [index, element] of sectionElements.entries()) {
    const layout = await parseSectionLayout(
      element,
      archive,
      relationships,
      previous,
      oddEvenPageChrome,
    );
    sections.push({ id: `document-section-${index + 1}`, layout });
    previous = layout;
  }
  if (sections.length > 1) addSectionMarkers(document, sectionElements);
  return {
    conversionBuffer:
      sections.length > 1 ||
      hasImportedDocxCaptionMarkers(captionMarkers) ||
      hasImportedDocxBookmarkMarkers(bookmarkMarkers) ||
      hasImportedDocxChangeMarkers(changeMarkers) ||
      hasImportedDocxCommentMarkers(commentMarkers) ||
      hasImportedDocxCitationMarkers(citationMarkers) ||
      hasImportedDocxFieldMarkers(fieldMarkers) ||
      hasImportedDocxListMarkers(listMarkers) ||
      hasImportedDocxImageLayoutMarkers(imageLayoutMarkers) ||
      hasImportedDocxParagraphIdentityMarkers(paragraphIdentityMarkers) ||
      hasImportedDocxParagraphAlignmentMarkers(paragraphAlignmentMarkers) ||
      hasImportedDocxParagraphDirectionMarkers(paragraphDirectionMarkers) ||
      hasImportedDocxParagraphIndentMarkers(paragraphIndentMarkers) ||
      hasImportedDocxParagraphSpacingMarkers(paragraphSpacingMarkers) ||
      hasImportedDocxParagraphPaginationMarkers(paragraphPaginationMarkers) ||
      hasImportedDocxRunFormattingMarkers(runFormattingMarkers) ||
      hasImportedDocxParagraphTabStopMarkers(tabStopMarkers) ||
      hasImportedDocxTableCellMarkers(tableCellMarkers) ||
      hasImportedDocxTableRowMarkers(tableRowMarkers) ||
      hasImportedDocxTableSizingMarkers(tableSizingMarkers)
        ? await writeDocumentXml(buffer, document)
        : buffer,
    sections,
    pageColor,
    captionMarkers,
    bookmarkMarkers,
    changeMarkers,
    commentMarkers,
    fieldMarkers,
    citationMarkers,
    listMarkers,
    imageLayoutMarkers,
    paragraphIdentityMarkers,
    paragraphAlignmentMarkers,
    paragraphDirectionMarkers,
    paragraphIndentMarkers,
    paragraphPaginationMarkers,
    paragraphSpacingMarkers,
    runFormattingMarkers,
    tabStopMarkers,
    tableCellMarkers,
    tableRowMarkers,
    tableSizingMarkers,
    bibliography,
    trackChanges,
  };
}

export function applyDocxSectionsToHtml(
  html: string,
  sections: PreparedDocxImport['sections'],
  captionMarkers: ImportedDocxCaptionMarkers = { captions: [], references: [] },
  bookmarkMarkers: ImportedDocxBookmarkMarkers = {
    bookmarks: [],
    references: [],
  },
  changeMarkers: ImportedDocxChangeMarkers = { changes: [] },
  commentMarkers: ImportedDocxCommentMarkers = { comments: [], ranges: [] },
  fieldMarkers: ImportedDocxFieldMarkers = { fields: [] },
  citationMarkers: ImportedDocxCitationMarkers = {
    citations: [],
    bibliographies: [],
  },
  listMarkers: ImportedDocxListMarkers = { lists: [] },
  imageLayoutMarkers: ImportedDocxImageLayoutMarkers = { images: [] },
  paragraphIdentityMarkers: ImportedDocxParagraphIdentityMarkers = {
    paragraphs: [],
  },
  paragraphAlignmentMarkers: ImportedDocxParagraphAlignmentMarkers = {
    paragraphs: [],
  },
  runFormattingMarkers: ImportedDocxRunFormattingMarkers = {
    runs: [],
  },
  paragraphDirectionMarkers: ImportedDocxParagraphDirectionMarkers = {
    paragraphs: [],
  },
  paragraphIndentMarkers: ImportedDocxParagraphIndentMarkers = {
    paragraphs: [],
  },
  paragraphSpacingMarkers: ImportedDocxParagraphSpacingMarkers = {
    paragraphs: [],
  },
  paragraphPaginationMarkers: ImportedDocxParagraphPaginationMarkers = {
    paragraphs: [],
  },
  bibliography?: WorkDocumentContent['bibliography'],
  tabStopMarkers: ImportedDocxParagraphTabStopMarkers = {
    paragraphs: [],
    inlineTabs: [],
  },
  tableCellMarkers: ImportedDocxTableCellMarkers = { cells: [] },
  tableRowMarkers: ImportedDocxTableRowMarkers = { rows: [] },
  tableSizingMarkers: ImportedDocxTableSizingMarkers = { tables: [] },
): string {
  const document = new DOMParser().parseFromString(html, 'text/html');
  applyImportedDocxRunFormattingMarkers(document, runFormattingMarkers);
  applyImportedDocxBookmarkMarkers(document, bookmarkMarkers);
  applyImportedDocxCaptionMarkers(document, captionMarkers);
  applyImportedDocxCitationMarkers(document, citationMarkers);
  applyImportedDocxFieldMarkers(document, fieldMarkers);
  applyImportedDocxListMarkers(document, listMarkers);
  applyImportedDocxImageLayoutMarkers(document, imageLayoutMarkers);
  applyImportedDocxParagraphDirectionMarkers(
    document,
    paragraphDirectionMarkers,
  );
  applyImportedDocxParagraphIndentMarkers(document, paragraphIndentMarkers);
  applyImportedDocxParagraphSpacingMarkers(document, paragraphSpacingMarkers);
  applyImportedDocxParagraphPaginationMarkers(
    document,
    paragraphPaginationMarkers,
  );
  applyImportedDocxParagraphTabStopMarkers(document, tabStopMarkers);
  applyImportedDocxParagraphAlignmentMarkers(
    document,
    paragraphAlignmentMarkers,
  );
  applyImportedDocxTableCellMarkers(document, tableCellMarkers);
  applyImportedDocxTableSizingMarkers(document, tableSizingMarkers);
  applyImportedDocxTableRowMarkers(document, tableRowMarkers);
  applyImportedDocxChangeMarkers(document, changeMarkers);
  applyImportedDocxCommentMarkers(document, commentMarkers);
  applyImportedDocxParagraphIdentityMarkers(document, paragraphIdentityMarkers);
  const notes = extractMammothDocumentNotes(document);
  const sourceNodes = Array.from(document.body.childNodes);
  document.body.replaceChildren();
  let sectionIndex = 0;
  let section = createHtmlSection(document, sections[sectionIndex]);

  for (const node of sourceNodes) {
    if (isSectionMarker(node, sectionIndex)) {
      ensureSectionContent(section);
      document.body.append(section);
      sectionIndex += 1;
      section = createHtmlSection(
        document,
        sections[sectionIndex] ?? sections.at(-1),
      );
      continue;
    }
    section.append(node);
  }
  ensureSectionContent(section);
  document.body.append(section);
  while (sectionIndex + 1 < sections.length) {
    sectionIndex += 1;
    const missing = createHtmlSection(document, sections[sectionIndex]);
    ensureSectionContent(missing);
    document.body.append(missing);
  }
  placeMammothDocumentNotes(document, notes);
  return normalizeDocumentBookmarkReferencesHtml(
    normalizeDocumentBookmarksHtml(
      normalizeDocumentCitationsHtml(
        normalizeDocumentFieldsHtml(
          normalizeDocumentCaptionsHtml(
            normalizeDocumentNotesHtml(document.body.innerHTML),
          ),
        ),
        bibliography,
      ),
    ),
  );
}

export async function readDocxLayout(
  buffer: ArrayBuffer,
): Promise<ImportedDocumentLayout> {
  const prepared = await prepareDocxImport(buffer);
  return {
    ...documentContentLayoutProperties(prepared.sections[0].layout),
    ...(prepared.pageColor ? { pageColor: prepared.pageColor } : {}),
  };
}

async function parseSectionLayout(
  section: Element,
  archive: OoxmlPackage,
  relationships: Awaited<ReturnType<OoxmlPackage['relationships']>>,
  previous: WorkDocumentSectionLayout,
  oddEvenPageChrome: boolean,
): Promise<WorkDocumentSectionLayout> {
  const pageSize = firstDescendant(section, 'pgSz');
  const width = numberAttribute(pageSize, 'w');
  const height = numberAttribute(pageSize, 'h');
  const orientation =
    attribute(pageSize ?? section, 'orient') === 'landscape' ||
    (width > 0 && height > 0 && width > height)
      ? 'landscape'
      : pageSize
        ? 'portrait'
        : previous.orientation;
  const shortEdge = Math.min(width || 11_906, height || 16_838);
  const size = pageSize
    ? Math.abs(shortEdge - 12_240) < Math.abs(shortEdge - 11_906)
      ? 'letter'
      : 'a4'
    : previous.pageSize;
  const marginsElement = firstDescendant(section, 'pgMar');
  const columnsElement = firstDescendant(section, 'cols');
  const documentGridElement = directChild(section, 'docGrid');
  const pageChrome = await importSectionPageChrome(
    section,
    archive,
    relationships,
    previous,
    oddEvenPageChrome,
  );
  const pageNumberStart = numberAttribute(
    firstDescendant(section, 'pgNumType'),
    'start',
  );
  return {
    pageSize: size,
    orientation,
    margins: marginsElement
      ? parseMargins(marginsElement, previous.margins)
      : { ...previous.margins },
    columns: columnsElement
      ? importDocxColumns(columnsElement, previous.columns)
      : { ...previous.columns },
    ...(documentGridElement
      ? { documentGrid: parseDocumentGrid(documentGridElement) }
      : previous.documentGrid
        ? { documentGrid: { ...previous.documentGrid } }
        : {}),
    breakAfter: parseSectionBreak(firstDescendant(section, 'type')),
    ...pageChrome,
    pageNumberStart: pageNumberStart > 0 ? pageNumberStart : undefined,
  };
}

function parseDocumentGrid(element: Element): WorkDocumentGrid {
  const sourceType = attribute(element, 'type');
  const type: WorkDocumentGridType =
    sourceType === 'lines' ||
    sourceType === 'linesAndChars' ||
    sourceType === 'snapToChars'
      ? sourceType
      : 'default';
  const sourceLinePitch = numberAttribute(element, 'linePitch');
  return {
    type,
    linePitch:
      sourceLinePitch > 0 ? Number((sourceLinePitch / 20).toFixed(2)) : 18,
  };
}

function addSectionMarkers(
  document: Document,
  sectionElements: Element[],
): void {
  for (let index = 0; index < sectionElements.length - 1; index += 1) {
    const paragraph = closestAncestor(sectionElements[index], 'p');
    if (!paragraph?.parentNode) continue;
    const marker = document.createElementNS(WORD_NAMESPACE, 'w:p');
    const run = document.createElementNS(WORD_NAMESPACE, 'w:r');
    const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
    text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
    text.textContent = sectionMarker(index);
    run.append(text);
    marker.append(run);
    paragraph.parentNode.insertBefore(marker, paragraph.nextSibling);
  }
}

async function writeDocumentXml(
  buffer: ArrayBuffer,
  document: Document,
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  zip.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

function effectiveSectionProperties(document: Document): Element[] {
  return descendants(document, 'sectPr').filter(
    (element) => !closestAncestor(element, 'sectPrChange'),
  );
}

function createHtmlSection(
  document: Document,
  section: PreparedDocxImport['sections'][number] | undefined,
): HTMLElement {
  const element = document.createElement('section');
  const fallback = documentInitialSectionLayout({
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  });
  const layout = section?.layout ?? fallback;
  const id = section?.id ?? 'document-section';
  for (const [name, value] of Object.entries(
    documentSectionDomAttributes(layout, id),
  )) {
    element.setAttribute(name, value);
  }
  return element;
}

function ensureSectionContent(section: HTMLElement) {
  if (!section.childNodes.length) section.innerHTML = '<p></p>';
}

function isSectionMarker(node: ChildNode, index: number): boolean {
  return (
    node instanceof HTMLElement &&
    node.textContent?.trim() === sectionMarker(index)
  );
}

function sectionMarker(index: number): string {
  return `__A3S_WORK_DOCUMENT_SECTION_${index + 1}__`;
}

function parseMargins(
  element: Element,
  fallback: WorkDocumentMargins,
): WorkDocumentMargins {
  return {
    top: twipsToMillimeters(numberAttribute(element, 'top'), fallback.top),
    right: twipsToMillimeters(
      numberAttribute(element, 'right'),
      fallback.right,
    ),
    bottom: twipsToMillimeters(
      numberAttribute(element, 'bottom'),
      fallback.bottom,
    ),
    left: twipsToMillimeters(numberAttribute(element, 'left'), fallback.left),
  };
}

function parseSectionBreak(
  element: Element | undefined,
): WorkDocumentSectionBreakType {
  if (!element) return 'nextPage';
  const value = attribute(element, 'val');
  if (
    value === 'continuous' ||
    value === 'evenPage' ||
    value === 'oddPage' ||
    value === 'nextColumn'
  )
    return value;
  return 'nextPage';
}

function numberAttribute(element: Element | undefined, name: string): number {
  if (!element) return 0;
  const value = Number(attribute(element, name));
  return Number.isFinite(value) ? value : 0;
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function twipsToMillimeters(value: number, fallback: number): number {
  if (value <= 0) return fallback;
  return Math.round((value / TWIPS_PER_MILLIMETER) * 10) / 10;
}
