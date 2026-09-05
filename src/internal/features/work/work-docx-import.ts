import JSZip from 'jszip';
import { normalizeDocumentBookmarkReferencesHtml } from './work-document-bookmark-references';
import { normalizeDocumentBookmarksHtml } from './work-document-bookmarks';
import { normalizeDocumentCaptionsHtml } from './work-document-captions';
import { normalizeDocumentCitationsHtml } from './work-document-citations';
import { normalizeDocumentFieldsHtml } from './work-document-fields';
import { normalizeDocumentIndexesHtml } from './work-document-index';
import { normalizeDocumentNotesHtml } from './work-document-notes';
import {
  documentPageMarginBody,
  documentPageMarginsForLayout,
} from './work-document-page-margins';
import { applyDocumentPageGeometry } from './work-document-page-size';
import {
  documentContentLayoutProperties,
  documentInitialSectionLayout,
  documentSectionDomAttributes,
} from './work-document-section';
import { normalizeDocumentTableOfContentsHtml } from './work-document-table-of-contents';
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
  applyImportedDocxContentControlMarkers,
  hasImportedDocxContentControlMarkers,
  type ImportedDocxContentControlMarkers,
  markDocxContentControls,
} from './work-docx-content-control-import';
import {
  applyImportedDocxEquationMarkers,
  type ImportedDocxEquationMarkers,
  markDocxPackageEquations,
} from './work-docx-equation-import';
import {
  applyImportedDocxFieldMarkers,
  hasImportedDocxFieldMarkers,
  type ImportedDocxFieldMarkers,
  markDocxBodyFields,
} from './work-docx-field-import';
import {
  applyImportedDocxImageLayoutMarkers,
  createImportedDocxImageLayoutMarkerState,
  hasImportedDocxImageLayoutMarkers,
  type ImportedDocxImageLayoutMarkers,
  markDocxImageLayouts,
} from './work-docx-image-layout-import';
import {
  applyImportedDocxIndexMarkers,
  hasImportedDocxIndexMarkers,
  type ImportedDocxIndexMarkers,
  markDocxIndexes,
} from './work-docx-index-import';
import {
  applyImportedDocxListMarkers,
  hasImportedDocxListMarkers,
  type ImportedDocxListMarkers,
  markDocxLists,
} from './work-docx-list-import';
import { markDocxNoteImageLayouts } from './work-docx-note-image-import';
import {
  extractMammothDocumentNotes,
  placeMammothDocumentNotes,
} from './work-docx-note-import';
import {
  applyImportedDocxNumberingChangeMarkers,
  hasImportedDocxNumberingChangeMarkers,
  type ImportedDocxNumberingChangeMarkers,
  markDocxNumberingChanges,
} from './work-docx-numbering-change-import';
import { parseDocxPageBorders } from './work-docx-page-borders-import';
import {
  documentUsesOddEvenPageChrome,
  importSectionPageChrome,
} from './work-docx-page-chrome-import';
import { importDocxPageColor } from './work-docx-page-color';
import {
  type InspectedDocxPageMarginSettings,
  inspectDocxPageMarginSettings,
  parseDocxPageMargins,
} from './work-docx-page-margins-import';
import {
  parseDocxPageGeometry,
  parseDocxPaperSource,
} from './work-docx-page-size-import';
import {
  applyImportedDocxParagraphAlignmentMarkers,
  hasImportedDocxParagraphAlignmentMarkers,
  type ImportedDocxParagraphAlignmentMarkers,
  markDocxParagraphAlignments,
} from './work-docx-paragraph-alignment-import';
import {
  applyImportedDocxParagraphBorderMarkers,
  hasImportedDocxParagraphBorderMarkers,
  type ImportedDocxParagraphBorderMarkers,
  markDocxParagraphBorders,
} from './work-docx-paragraph-borders-import';
import {
  applyImportedDocxParagraphDirectionMarkers,
  hasImportedDocxParagraphDirectionMarkers,
  type ImportedDocxParagraphDirectionMarkers,
  markDocxParagraphDirections,
} from './work-docx-paragraph-direction-import';
import {
  applyImportedDocxParagraphFormattingChangeMarkers,
  hasImportedDocxParagraphFormattingChangeMarkers,
  type ImportedDocxParagraphFormattingChangeMarkers,
  markDocxParagraphFormattingChanges,
} from './work-docx-paragraph-format-change-import';
import {
  applyImportedDocxParagraphIdentityMarkers,
  hasImportedDocxParagraphIdentityMarkers,
  type ImportedDocxParagraphIdentityMarkers,
  markDocxParagraphIdentities,
} from './work-docx-paragraph-identity-import';
import {
  applyImportedDocxParagraphIndentMarkers,
  hasImportedDocxParagraphIndentMarkers,
  type ImportedDocxParagraphIndentMarkers,
  markDocxParagraphIndents,
} from './work-docx-paragraph-indent-import';
import {
  applyImportedDocxParagraphMarkChangeMarkers,
  hasImportedDocxParagraphMarkChangeMarkers,
  type ImportedDocxParagraphMarkChangeMarkers,
  markDocxParagraphMarkChanges,
} from './work-docx-paragraph-mark-change-import';
import {
  applyImportedDocxParagraphPaginationMarkers,
  hasImportedDocxParagraphPaginationMarkers,
  type ImportedDocxParagraphPaginationMarkers,
  markDocxParagraphPagination,
} from './work-docx-paragraph-pagination-import';
import {
  applyImportedDocxParagraphShadingMarkers,
  hasImportedDocxParagraphShadingMarkers,
  type ImportedDocxParagraphShadingMarkers,
  markDocxParagraphShading,
} from './work-docx-paragraph-shading-import';
import {
  applyImportedDocxParagraphSpacingMarkers,
  hasImportedDocxParagraphSpacingMarkers,
  type ImportedDocxParagraphSpacingMarkers,
  markDocxParagraphSpacing,
} from './work-docx-paragraph-spacing-import';
import { createDocxParagraphStyleResolver } from './work-docx-paragraph-styles';
import {
  applyImportedDocxRunFormattingMarkers,
  createImportedDocxRunFormattingMarkerState,
  hasImportedDocxRunFormattingMarkers,
  type ImportedDocxRunFormattingMarkers,
  markDocxRunFormattingIntoState,
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
  applyImportedDocxTableOfContentsMarkers,
  hasImportedDocxTableOfContentsMarkers,
  type ImportedDocxTableOfContentsMarkers,
  markDocxTablesOfContents,
} from './work-docx-table-of-contents-import';
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
  applyImportedDocxTextBoxMarkers,
  hasImportedDocxTextBoxMarkers,
  type ImportedDocxTextBoxMarkers,
  markDocxTextBoxes,
} from './work-docx-text-box-import';
import {
  createDocxThemeResolver,
  type DocxThemeResolver,
} from './work-docx-theme';
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
  tableOfContentsMarkers: ImportedDocxTableOfContentsMarkers;
  indexMarkers: ImportedDocxIndexMarkers;
  equationMarkers: ImportedDocxEquationMarkers;
  citationMarkers: ImportedDocxCitationMarkers;
  listMarkers: ImportedDocxListMarkers;
  numberingChangeMarkers: ImportedDocxNumberingChangeMarkers;
  imageLayoutMarkers: ImportedDocxImageLayoutMarkers;
  textBoxMarkers: ImportedDocxTextBoxMarkers;
  contentControlMarkers: ImportedDocxContentControlMarkers;
  paragraphIdentityMarkers: ImportedDocxParagraphIdentityMarkers;
  paragraphFormattingChangeMarkers: ImportedDocxParagraphFormattingChangeMarkers;
  paragraphMarkChangeMarkers: ImportedDocxParagraphMarkChangeMarkers;
  paragraphAlignmentMarkers: ImportedDocxParagraphAlignmentMarkers;
  paragraphDirectionMarkers: ImportedDocxParagraphDirectionMarkers;
  paragraphIndentMarkers: ImportedDocxParagraphIndentMarkers;
  paragraphPaginationMarkers: ImportedDocxParagraphPaginationMarkers;
  paragraphBorderMarkers: ImportedDocxParagraphBorderMarkers;
  paragraphShadingMarkers: ImportedDocxParagraphShadingMarkers;
  paragraphSpacingMarkers: ImportedDocxParagraphSpacingMarkers;
  runFormattingMarkers: ImportedDocxRunFormattingMarkers;
  tabStopMarkers: ImportedDocxParagraphTabStopMarkers;
  tableCellMarkers: ImportedDocxTableCellMarkers;
  tableRowMarkers: ImportedDocxTableRowMarkers;
  tableSizingMarkers: ImportedDocxTableSizingMarkers;
  bibliography?: WorkDocumentContent['bibliography'];
  trackChanges: boolean;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

export async function prepareDocxImport(
  buffer: ArrayBuffer,
  sourcePackage?: OoxmlPackage,
): Promise<PreparedDocxImport> {
  const archive = sourcePackage ?? (await OoxmlPackage.load(buffer));
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
      tableOfContentsMarkers: { tables: [] },
      indexMarkers: { entries: [], indexes: [] },
      equationMarkers: { equations: [] },
      citationMarkers: { citations: [], bibliographies: [] },
      listMarkers: { lists: [] },
      numberingChangeMarkers: { groups: [] },
      imageLayoutMarkers: { images: [] },
      textBoxMarkers: { textBoxes: [] },
      contentControlMarkers: { controls: [], unsupported: 0 },
      paragraphIdentityMarkers: { paragraphs: [] },
      paragraphFormattingChangeMarkers: { paragraphs: [] },
      paragraphMarkChangeMarkers: { paragraphs: [] },
      paragraphAlignmentMarkers: { paragraphs: [] },
      paragraphDirectionMarkers: { paragraphs: [] },
      paragraphIndentMarkers: { paragraphs: [] },
      paragraphPaginationMarkers: { paragraphs: [] },
      paragraphBorderMarkers: {
        paragraphs: [],
        invalidCount: 0,
        spoofedCount: 0,
      },
      paragraphShadingMarkers: {
        paragraphs: [],
        invalidCount: 0,
        spoofedCount: 0,
      },
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
  const paragraphMarkChangeMarkers = markDocxParagraphMarkChanges(document);
  const tableOfContentsMarkers = markDocxTablesOfContents(document);
  const indexMarkers = markDocxIndexes(document);
  const contentControlMarkers = markDocxContentControls(document);
  const textBoxMarkers = markDocxTextBoxes(document);
  const paragraphIdentityMarkers = markDocxParagraphIdentities(document);
  const numbering = archive.has('word/numbering.xml')
    ? await archive.xml('word/numbering.xml')
    : null;
  const commentMarkers = await markDocxComments(document, archive);
  const changeMarkers = markDocxTextChanges(document);
  const captionMarkers = markDocxCaptionFields(document);
  const bookmarkMarkers = markDocxBookmarks(document);
  const citationMarkers = markDocxCitationFields(document);
  const fieldMarkers = markDocxBodyFields(document, bookmarkMarkers.bookmarks);
  const listMarkers = markDocxLists(document, numbering);
  const numberingChangeMarkers = markDocxNumberingChanges(document);
  const imageLayoutMarkerState = createImportedDocxImageLayoutMarkerState();
  const imageLayoutMarkers = markDocxImageLayouts(
    document,
    imageLayoutMarkerState,
  );
  const noteImageMarkers = await markDocxNoteImageLayouts(
    archive,
    imageLayoutMarkerState,
  );
  imageLayoutMarkers.images.push(...noteImageMarkers.markers.images);
  const equationMarkers = await markDocxPackageEquations(
    archive,
    document,
    noteImageMarkers.parts,
  );
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
  const settings = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const pageMarginSettings = inspectDocxPageMarginSettings(settings);
  const theme = createDocxThemeResolver(themeDocument, settings);
  const paragraphBorderMarkers = markDocxParagraphBorders(
    document,
    paragraphStyles,
    theme,
    tableStyles,
  );
  const paragraphShadingMarkers = markDocxParagraphShading(
    document,
    paragraphStyles,
    theme,
    tableStyles,
  );
  const runFormattingMarkerState = createImportedDocxRunFormattingMarkerState([
    document,
    ...equationMarkers.parts.map((part) => part.document),
  ]);
  markDocxRunFormattingIntoState(
    document,
    runFormattingMarkerState,
    paragraphStyles,
    theme,
    tableStyles,
  );
  for (const part of equationMarkers.parts) {
    markDocxRunFormattingIntoState(
      part.document,
      runFormattingMarkerState,
      paragraphStyles,
      theme,
      tableStyles,
    );
  }
  const runFormattingMarkers = runFormattingMarkerState.markers;
  const paragraphFormattingChangeMarkers = markDocxParagraphFormattingChanges(
    document,
    paragraphStyles,
    theme,
    tableStyles,
  );
  const tabStopMarkers = markDocxParagraphTabStops(
    document,
    paragraphStyles,
    tableStyles,
  );
  const tableCellMarkers = markDocxTableCells(document, theme, tableStyles);
  const tableRowMarkers = markDocxTableRows(document);
  const tableSizingMarkers = markDocxTableSizing(document, tableStyles);
  const trackChanges =
    Boolean(settings && firstDescendant(settings, 'trackRevisions')) ||
    changeMarkers.changes.length > 0 ||
    paragraphMarkChangeMarkers.paragraphs.length > 0 ||
    numberingChangeMarkers.groups.length > 0 ||
    paragraphFormattingChangeMarkers.paragraphs.length > 0 ||
    runFormattingMarkers.runs.some((run) => Boolean(run.change));
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
        hasImportedDocxTableOfContentsMarkers(tableOfContentsMarkers) ||
        hasImportedDocxIndexMarkers(indexMarkers) ||
        equationMarkers.changed ||
        hasImportedDocxListMarkers(listMarkers) ||
        hasImportedDocxNumberingChangeMarkers(numberingChangeMarkers) ||
        hasImportedDocxImageLayoutMarkers(imageLayoutMarkers) ||
        hasImportedDocxTextBoxMarkers(textBoxMarkers) ||
        hasImportedDocxContentControlMarkers(contentControlMarkers) ||
        hasImportedDocxParagraphIdentityMarkers(paragraphIdentityMarkers) ||
        hasImportedDocxParagraphFormattingChangeMarkers(
          paragraphFormattingChangeMarkers,
        ) ||
        hasImportedDocxParagraphMarkChangeMarkers(paragraphMarkChangeMarkers) ||
        hasImportedDocxParagraphAlignmentMarkers(paragraphAlignmentMarkers) ||
        hasImportedDocxParagraphDirectionMarkers(paragraphDirectionMarkers) ||
        hasImportedDocxParagraphIndentMarkers(paragraphIndentMarkers) ||
        hasImportedDocxParagraphSpacingMarkers(paragraphSpacingMarkers) ||
        hasImportedDocxParagraphBorderMarkers(paragraphBorderMarkers) ||
        hasImportedDocxParagraphShadingMarkers(paragraphShadingMarkers) ||
        hasImportedDocxParagraphPaginationMarkers(paragraphPaginationMarkers) ||
        hasImportedDocxRunFormattingMarkers(runFormattingMarkers) ||
        hasImportedDocxParagraphTabStopMarkers(tabStopMarkers) ||
        hasImportedDocxTableCellMarkers(tableCellMarkers) ||
        hasImportedDocxTableRowMarkers(tableRowMarkers) ||
        hasImportedDocxTableSizingMarkers(tableSizingMarkers)
          ? await writeDocumentXml(buffer, document, equationMarkers.parts)
          : buffer,
      sections: [{ id: 'document-section-1', layout: fallback }],
      pageColor,
      captionMarkers,
      bookmarkMarkers,
      changeMarkers,
      commentMarkers,
      fieldMarkers,
      tableOfContentsMarkers,
      indexMarkers,
      equationMarkers: equationMarkers.markers,
      citationMarkers,
      listMarkers,
      numberingChangeMarkers,
      imageLayoutMarkers,
      textBoxMarkers,
      contentControlMarkers,
      paragraphIdentityMarkers,
      paragraphFormattingChangeMarkers,
      paragraphMarkChangeMarkers,
      paragraphAlignmentMarkers,
      paragraphDirectionMarkers,
      paragraphIndentMarkers,
      paragraphPaginationMarkers,
      paragraphBorderMarkers,
      paragraphShadingMarkers,
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
      theme,
      pageMarginSettings,
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
      hasImportedDocxTableOfContentsMarkers(tableOfContentsMarkers) ||
      hasImportedDocxIndexMarkers(indexMarkers) ||
      equationMarkers.changed ||
      hasImportedDocxListMarkers(listMarkers) ||
      hasImportedDocxNumberingChangeMarkers(numberingChangeMarkers) ||
      hasImportedDocxImageLayoutMarkers(imageLayoutMarkers) ||
      hasImportedDocxTextBoxMarkers(textBoxMarkers) ||
      hasImportedDocxContentControlMarkers(contentControlMarkers) ||
      hasImportedDocxParagraphIdentityMarkers(paragraphIdentityMarkers) ||
      hasImportedDocxParagraphFormattingChangeMarkers(
        paragraphFormattingChangeMarkers,
      ) ||
      hasImportedDocxParagraphMarkChangeMarkers(paragraphMarkChangeMarkers) ||
      hasImportedDocxParagraphAlignmentMarkers(paragraphAlignmentMarkers) ||
      hasImportedDocxParagraphDirectionMarkers(paragraphDirectionMarkers) ||
      hasImportedDocxParagraphIndentMarkers(paragraphIndentMarkers) ||
      hasImportedDocxParagraphSpacingMarkers(paragraphSpacingMarkers) ||
      hasImportedDocxParagraphBorderMarkers(paragraphBorderMarkers) ||
      hasImportedDocxParagraphShadingMarkers(paragraphShadingMarkers) ||
      hasImportedDocxParagraphPaginationMarkers(paragraphPaginationMarkers) ||
      hasImportedDocxRunFormattingMarkers(runFormattingMarkers) ||
      hasImportedDocxParagraphTabStopMarkers(tabStopMarkers) ||
      hasImportedDocxTableCellMarkers(tableCellMarkers) ||
      hasImportedDocxTableRowMarkers(tableRowMarkers) ||
      hasImportedDocxTableSizingMarkers(tableSizingMarkers)
        ? await writeDocumentXml(buffer, document, equationMarkers.parts)
        : buffer,
    sections,
    pageColor,
    captionMarkers,
    bookmarkMarkers,
    changeMarkers,
    commentMarkers,
    fieldMarkers,
    tableOfContentsMarkers,
    indexMarkers,
    equationMarkers: equationMarkers.markers,
    citationMarkers,
    listMarkers,
    numberingChangeMarkers,
    imageLayoutMarkers,
    textBoxMarkers,
    contentControlMarkers,
    paragraphIdentityMarkers,
    paragraphFormattingChangeMarkers,
    paragraphMarkChangeMarkers,
    paragraphAlignmentMarkers,
    paragraphDirectionMarkers,
    paragraphIndentMarkers,
    paragraphPaginationMarkers,
    paragraphBorderMarkers,
    paragraphShadingMarkers,
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
  tableOfContentsMarkers: ImportedDocxTableOfContentsMarkers = { tables: [] },
  indexMarkers: ImportedDocxIndexMarkers = { entries: [], indexes: [] },
  equationMarkers: ImportedDocxEquationMarkers = { equations: [] },
  citationMarkers: ImportedDocxCitationMarkers = {
    citations: [],
    bibliographies: [],
  },
  listMarkers: ImportedDocxListMarkers = { lists: [] },
  numberingChangeMarkers: ImportedDocxNumberingChangeMarkers = { groups: [] },
  imageLayoutMarkers: ImportedDocxImageLayoutMarkers = { images: [] },
  paragraphIdentityMarkers: ImportedDocxParagraphIdentityMarkers = {
    paragraphs: [],
  },
  paragraphFormattingChangeMarkers: ImportedDocxParagraphFormattingChangeMarkers = {
    paragraphs: [],
  },
  paragraphMarkChangeMarkers: ImportedDocxParagraphMarkChangeMarkers = {
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
  paragraphBorderMarkers: ImportedDocxParagraphBorderMarkers = {
    paragraphs: [],
    invalidCount: 0,
    spoofedCount: 0,
  },
  paragraphShadingMarkers: ImportedDocxParagraphShadingMarkers = {
    paragraphs: [],
    invalidCount: 0,
    spoofedCount: 0,
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
  textBoxMarkers: ImportedDocxTextBoxMarkers = { textBoxes: [] },
  contentControlMarkers: ImportedDocxContentControlMarkers = {
    controls: [],
    unsupported: 0,
  },
): string {
  const document = new DOMParser().parseFromString(html, 'text/html');
  applyImportedDocxRunFormattingMarkers(document, runFormattingMarkers);
  applyImportedDocxBookmarkMarkers(document, bookmarkMarkers);
  applyImportedDocxCaptionMarkers(document, captionMarkers);
  applyImportedDocxCitationMarkers(document, citationMarkers);
  applyImportedDocxFieldMarkers(document, fieldMarkers);
  applyImportedDocxEquationMarkers(document, equationMarkers);
  applyImportedDocxListMarkers(document, listMarkers);
  applyImportedDocxNumberingChangeMarkers(document, numberingChangeMarkers);
  applyImportedDocxImageLayoutMarkers(document, imageLayoutMarkers);
  applyImportedDocxParagraphDirectionMarkers(
    document,
    paragraphDirectionMarkers,
  );
  applyImportedDocxParagraphIndentMarkers(document, paragraphIndentMarkers);
  applyImportedDocxParagraphSpacingMarkers(document, paragraphSpacingMarkers);
  applyImportedDocxParagraphBorderMarkers(document, paragraphBorderMarkers);
  applyImportedDocxParagraphShadingMarkers(document, paragraphShadingMarkers);
  applyImportedDocxParagraphPaginationMarkers(
    document,
    paragraphPaginationMarkers,
  );
  applyImportedDocxParagraphTabStopMarkers(document, tabStopMarkers);
  applyImportedDocxParagraphAlignmentMarkers(
    document,
    paragraphAlignmentMarkers,
  );
  applyImportedDocxParagraphFormattingChangeMarkers(
    document,
    paragraphFormattingChangeMarkers,
  );
  applyImportedDocxParagraphMarkChangeMarkers(
    document,
    paragraphMarkChangeMarkers,
  );
  applyImportedDocxTableCellMarkers(document, tableCellMarkers);
  applyImportedDocxTableSizingMarkers(document, tableSizingMarkers);
  applyImportedDocxTableRowMarkers(document, tableRowMarkers);
  applyImportedDocxChangeMarkers(document, changeMarkers);
  applyImportedDocxCommentMarkers(document, commentMarkers);
  applyImportedDocxParagraphIdentityMarkers(document, paragraphIdentityMarkers);
  applyImportedDocxTextBoxMarkers(document, textBoxMarkers);
  applyImportedDocxTableOfContentsMarkers(document, tableOfContentsMarkers);
  applyImportedDocxIndexMarkers(document, indexMarkers);
  applyImportedDocxContentControlMarkers(document, contentControlMarkers);
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
  return normalizeDocumentIndexesHtml(
    normalizeDocumentBookmarkReferencesHtml(
      normalizeDocumentBookmarksHtml(
        normalizeDocumentCitationsHtml(
          normalizeDocumentTableOfContentsHtml(
            normalizeDocumentFieldsHtml(
              normalizeDocumentCaptionsHtml(
                normalizeDocumentNotesHtml(document.body.innerHTML),
              ),
            ),
          ),
          bibliography,
        ),
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
  theme: DocxThemeResolver,
  pageMarginSettings: InspectedDocxPageMarginSettings,
): Promise<WorkDocumentSectionLayout> {
  const parsedPageGeometry = parseDocxPageGeometry(
    section,
    previous.pageGeometry,
  );
  const pageGeometry =
    parsedPageGeometry === null ? previous.pageGeometry : parsedPageGeometry;
  const parsedPaperSource = parseDocxPaperSource(section, previous.paperSource);
  const paperSource =
    parsedPaperSource === null ? previous.paperSource : parsedPaperSource;
  const columnsElement = firstDescendant(section, 'cols');
  const documentGridElement = directChild(section, 'docGrid');
  const pageBorders = parseDocxPageBorders(section, theme);
  const parsedPageMargins = parseDocxPageMargins(
    section,
    pageMarginSettings,
    documentPageMarginsForLayout(previous),
  );
  const pageMargins =
    parsedPageMargins === null
      ? documentPageMarginsForLayout(previous)
      : parsedPageMargins;
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
  const layout: WorkDocumentSectionLayout = {
    pageSize: previous.pageSize,
    orientation: previous.orientation,
    margins: documentPageMarginBody(pageMargins, previous.margins),
    columns: columnsElement
      ? importDocxColumns(columnsElement, previous.columns)
      : { ...previous.columns },
    ...(documentGridElement
      ? { documentGrid: parseDocumentGrid(documentGridElement) }
      : previous.documentGrid
        ? { documentGrid: { ...previous.documentGrid } }
        : {}),
    ...(pageBorders ? { pageBorders } : {}),
    ...(pageMargins ? { pageMargins } : {}),
    ...(pageGeometry ? { pageGeometry } : {}),
    ...(paperSource ? { paperSource } : {}),
    breakAfter: parseSectionBreak(firstDescendant(section, 'type')),
    ...pageChrome,
    pageNumberStart: pageNumberStart > 0 ? pageNumberStart : undefined,
  };
  return pageGeometry
    ? applyDocumentPageGeometry(layout, pageGeometry)
    : layout;
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
  additionalParts: readonly { document: Document; path: string }[] = [],
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  zip.file('word/document.xml', serializeDocxConversionXml(document));
  for (const part of additionalParts) {
    zip.file(part.path, serializeDocxConversionXml(part.document));
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}

function serializeDocxConversionXml(document: Document): string {
  const dialects = new Map<string, string>([
    [
      'http://purl.oclc.org/ooxml/wordprocessingml/main',
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    ],
    [
      'http://purl.oclc.org/ooxml/officeDocument/math',
      'http://schemas.openxmlformats.org/officeDocument/2006/math',
    ],
    [
      'http://purl.oclc.org/ooxml/officeDocument/relationships',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    ],
    [
      'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
      'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    ],
    [
      'http://purl.oclc.org/ooxml/drawingml/main',
      'http://schemas.openxmlformats.org/drawingml/2006/main',
    ],
    [
      'http://purl.oclc.org/ooxml/drawingml/picture',
      'http://schemas.openxmlformats.org/drawingml/2006/picture',
    ],
  ] as const);
  return new XMLSerializer()
    .serializeToString(document)
    .replace(
      /(\sxmlns(?::[A-Za-z_][A-Za-z0-9_.-]*)?\s*=\s*)(["'])([^"']+)\2/g,
      (_declaration, assignment: string, quote: string, namespace: string) =>
        `${assignment}${quote}${dialects.get(namespace) ?? namespace}${quote}`,
    );
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
