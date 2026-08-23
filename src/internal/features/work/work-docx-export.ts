import type {
  FileChild,
  IRunOptions,
  ISectionOptions,
  ParagraphChild,
} from 'docx';
import { normalizeDocumentBookmarkReferencesHtml } from './work-document-bookmark-references';
import { normalizeDocumentBookmarksHtml } from './work-document-bookmarks';
import { documentCharacterPositionHalfPointsFromElement } from './work-document-character-position';
import { documentCharacterSpacingTwipsFromElement } from './work-document-character-spacing';
import { normalizeDocumentHref } from './work-document-links';
import {
  collectDocumentNotes,
  documentNoteKey,
  documentNoteKind,
  type WorkDocumentNote,
  type WorkDocumentNoteKind,
} from './work-document-notes';
import { normalizeDocumentPageChrome } from './work-document-page-chrome';
import { documentPageMarginsForLayout } from './work-document-page-margins';
import { documentPageGeometryForLayout } from './work-document-page-size';
import { documentSections } from './work-document-section';
import { patchDocxBibliography } from './work-docx-bibliography';
import {
  DocxBookmarkPatchCollector,
  patchDocxBookmarks,
} from './work-docx-bookmarks';
import {
  docxCaptionParagraph,
  docxCrossReferenceRuns,
} from './work-docx-caption-export';
import {
  docxBibliographyParagraph,
  docxCitationRun,
} from './work-docx-citation-export';
import { docxSectionColumns } from './work-docx-column-export';
import { createDocxCommentRecords } from './work-docx-comment-export';
import { patchDocxCommentMetadata } from './work-docx-comment-metadata';
import { docxCharacterPositionValue } from './work-docx-character-position';
import {
  docxCharacterSpacingValue,
  patchDocxExplicitZeroCharacterSpacing,
} from './work-docx-character-spacing';
import { patchDocxDocumentLayout } from './work-docx-document-layout';
import {
  DocxEquationPatchCollector,
  patchDocxEquations,
} from './work-docx-equation-export';
import {
  cssColorToHex,
  cssFontFamily,
  cssFontSize,
  dataBoolean,
  domDirection,
  paragraphAlignment,
  paragraphBidirectional,
  paragraphDirectionOptions,
  paragraphIndent,
  paragraphPaginationOptions,
  paragraphSpacingOptions,
  paragraphTabStops,
} from './work-docx-export-formatting';
import { imageToDocx } from './work-docx-export-image';
import { docxDocumentFieldRun } from './work-docx-field-export';
import {
  DocxRunFormattingChangePatchCollector,
  patchDocxRunFormattingChanges,
} from './work-docx-format-change-export';
import {
  DocxImageCropPatchCollector,
  patchDocxImageCrops,
} from './work-docx-image-crop';
import {
  DocxImageIdentityPatchCollector,
  patchDocxImageIdentities,
} from './work-docx-image-identity';
import {
  DocxImageLayerPatchCollector,
  patchDocxImageLayers,
} from './work-docx-image-layer';
import {
  DocxImageWrapPatchCollector,
  patchDocxImageWraps,
} from './work-docx-image-wrap';
import {
  type DocxListExportContext,
  listToDocxParagraphs,
} from './work-docx-list-export';
import {
  assignDocxCommentThreads,
  assignDocxNoteIds,
} from './work-docx-note-comment-identity';
import { patchDocxNoteImageRelationships } from './work-docx-note-image-relationships';
import { patchDocxNumberingRestartRules } from './work-docx-numbering';
import { patchDocxPageColor } from './work-docx-page-color';
import {
  DocxParagraphDefaultCollapsedPatchCollector,
  patchDocxParagraphDefaultCollapsed,
} from './work-docx-paragraph-default-collapsed';
import {
  DocxParagraphFormattingChangePatchCollector,
  patchDocxParagraphFormattingChanges,
} from './work-docx-paragraph-format-change-export';
import {
  DocxParagraphIdentityPatchCollector,
  patchDocxParagraphIdentities,
} from './work-docx-paragraph-identity';
import {
  DocxParagraphBorderPatchCollector,
  documentParagraphBordersDocxOptions,
  patchDocxParagraphBorders,
} from './work-docx-paragraph-borders-export';
import { documentParagraphShadingDocxOptions } from './work-docx-paragraph-shading-export';
import { documentTableCellDocxOptions } from './work-docx-table-cell-export';
import {
  documentTableCellSizingDocxOptions,
  documentTableRowSizingDocxOptions,
  documentTableSizingDocxOptions,
} from './work-docx-table-sizing-export';
import {
  DocxThemePatchCollector,
  parseDocxThemeReference,
  patchDocxThemeReferences,
} from './work-docx-theme-reference';
import { preserveDocxSourcePackage } from './work-ooxml-package-preservation';
import { normalizeDocumentTextCase } from './work-document-text-case';
import {
  DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE,
  documentUnderlineFormattingFromElement,
  type WorkDocumentUnderlineFormatting,
} from './work-document-underline';
import {
  DOCUMENT_STRIKE_STYLE_ATTRIBUTE,
  documentStrikeFormattingFromElement,
} from './work-document-strike';
import type {
  WorkDocumentComment,
  WorkDocumentContent,
  WorkDocumentSectionBreakType,
  WorkDocumentSectionLayout,
} from './work-types';

interface DocxNoteContext extends DocxListExportContext {
  ids: Map<string, number>;
  changeIds: Map<string, number>;
  nextChangeId: number;
  commentIds: Map<string, number>;
  commentRangeCounts: Map<string, number>;
  commentRangeSeen: Map<string, number>;
  themePatches: DocxThemePatchCollector;
  bookmarkPatches: DocxBookmarkPatchCollector;
  imageCropPatches: DocxImageCropPatchCollector;
  imageIdentityPatches: DocxImageIdentityPatchCollector;
  imageLayerPatches: DocxImageLayerPatchCollector;
  imageWrapPatches: DocxImageWrapPatchCollector;
  paragraphBorderPatches: DocxParagraphBorderPatchCollector;
  paragraphDefaultCollapsedPatches: DocxParagraphDefaultCollapsedPatchCollector;
  paragraphIdentityPatches: DocxParagraphIdentityPatchCollector;
  equationPatches: DocxEquationPatchCollector;
  formattingChangePatches: DocxRunFormattingChangePatchCollector;
  paragraphFormattingChangePatches: DocxParagraphFormattingChangePatchCollector;
  hasExplicitZeroCharacterSpacing: boolean;
}

interface DocxTextRevision {
  kind: 'insertion' | 'deletion';
  id: number;
  author: string;
  date: string;
}

export async function createDocxBlob(
  content: WorkDocumentContent,
  sourcePackage?: ArrayBuffer,
): Promise<Blob> {
  const docx = await import('docx');
  const normalizedContent = {
    ...content,
    html: normalizeDocumentBookmarkReferencesHtml(
      normalizeDocumentBookmarksHtml(content.html),
    ),
  };
  const noteCollection = collectDocumentNotes(normalizedContent.html);
  const comments = anchoredDocumentComments(normalizedContent);
  const commentThreads = assignDocxCommentThreads(comments);
  const noteContext: DocxNoteContext = {
    ids: assignDocxNoteIds(noteCollection.notes),
    changeIds: new Map(),
    nextChangeId: 1,
    commentIds: new Map(),
    commentRangeCounts: documentCommentRangeCounts(normalizedContent.html),
    commentRangeSeen: new Map(),
    numbering: [],
    nextNumberingReference: 1,
    numberingReferences: new Map(),
    numberingLevels: new Map(),
    numberingRestartRules: [],
    numberingRestartRulesByIdentity: new Map(),
    numberingSourceIdentities: [],
    themePatches: new DocxThemePatchCollector(
      JSON.stringify(normalizedContent),
    ),
    bookmarkPatches: new DocxBookmarkPatchCollector(normalizedContent.html),
    imageCropPatches: new DocxImageCropPatchCollector(),
    imageIdentityPatches: new DocxImageIdentityPatchCollector(),
    imageLayerPatches: new DocxImageLayerPatchCollector(),
    imageWrapPatches: new DocxImageWrapPatchCollector(),
    paragraphBorderPatches: new DocxParagraphBorderPatchCollector(
      JSON.stringify(normalizedContent),
    ),
    paragraphDefaultCollapsedPatches:
      new DocxParagraphDefaultCollapsedPatchCollector(),
    paragraphIdentityPatches: new DocxParagraphIdentityPatchCollector(
      JSON.stringify(normalizedContent),
    ),
    equationPatches: new DocxEquationPatchCollector(
      JSON.stringify(normalizedContent),
    ),
    formattingChangePatches: new DocxRunFormattingChangePatchCollector(),
    paragraphFormattingChangePatches:
      new DocxParagraphFormattingChangePatchCollector(),
    hasExplicitZeroCharacterSpacing: false,
  };
  const commentRecords = createDocxCommentRecords(
    commentThreads,
    docx,
    noteContext.commentIds,
  );
  const sections: ISectionOptions[] = [];
  let usesOddEvenPageChrome = false;
  const sourceSections = documentSections({
    ...normalizedContent,
    html: noteCollection.html,
  });
  for (const section of sourceSections) {
    const parsed = new DOMParser().parseFromString(section.html, 'text/html');
    const children: FileChild[] = [];
    for (const node of parsed.body.children) {
      const element = node as HTMLElement;
      if (element.hasAttribute('data-document-note')) continue;
      children.push(...(await blockToFileChildren(element, docx, noteContext)));
    }
    if (!children.length) children.push(new docx.Paragraph(''));
    const pageChrome = normalizeDocumentPageChrome(
      section.layout.pageChrome,
      section.layout,
    );
    usesOddEvenPageChrome ||= pageChrome.differentOddEvenPages;
    const headers = await sectionHeaders(pageChrome, docx, noteContext);
    const footers = await sectionFooters(pageChrome, docx, noteContext);
    sections.push({
      properties: sectionProperties(section.layout, docx),
      headers,
      footers,
      children,
    });
  }
  const footnotes = await createNoteRecords(
    noteCollection.notes,
    'footnote',
    docx,
    noteContext,
  );
  const endnotes = await createNoteRecords(
    noteCollection.notes,
    'endnote',
    docx,
    noteContext,
  );
  const document = new docx.Document({
    sections,
    footnotes: Object.keys(footnotes).length ? footnotes : undefined,
    endnotes: Object.keys(endnotes).length ? endnotes : undefined,
    comments: commentRecords.length ? { children: commentRecords } : undefined,
    numbering: noteContext.numbering.length
      ? { config: noteContext.numbering }
      : undefined,
    evenAndOddHeaderAndFooters: usesOddEvenPageChrome,
    features: {
      trackRevisions: Boolean(
        normalizedContent.trackChanges ||
          documentHasTrackedChanges(normalizedContent.html),
      ),
      updateFields: true,
    },
  });
  const packed = await docx.Packer.toBlob(document);
  const characterSpacingPatched = noteContext.hasExplicitZeroCharacterSpacing
    ? await patchDocxExplicitZeroCharacterSpacing(await packed.arrayBuffer())
    : await packed.arrayBuffer();
  const formattingChangesPatched = await patchDocxRunFormattingChanges(
    characterSpacingPatched,
    noteContext.formattingChangePatches.patches,
  );
  const noteImageRelationshipsPatched = await patchDocxNoteImageRelationships(
    formattingChangesPatched,
  );
  const commentMetadataPatched = await patchDocxCommentMetadata(
    noteImageRelationshipsPatched,
    commentThreads,
  );
  const numberingPatched = await patchDocxNumberingRestartRules(
    commentMetadataPatched,
    noteContext.numberingRestartRules,
  );
  const bibliographyPatched = await patchDocxBibliography(
    numberingPatched,
    normalizedContent.bibliography,
  );
  const layoutPatched = await patchDocxDocumentLayout(
    bibliographyPatched,
    sourceSections,
  );
  const paragraphBordersPatched = await patchDocxParagraphBorders(
    layoutPatched,
    noteContext.paragraphBorderPatches.patches,
  );
  const themePatched = await patchDocxThemeReferences(
    paragraphBordersPatched,
    noteContext.themePatches.patches,
  );
  const imageCropPatched = await patchDocxImageCrops(
    themePatched,
    noteContext.imageCropPatches.patches,
  );
  const imageWrapPatched = await patchDocxImageWraps(
    imageCropPatched,
    noteContext.imageWrapPatches.patches,
  );
  const imageLayerPatched = await patchDocxImageLayers(
    imageWrapPatched,
    noteContext.imageLayerPatches.patches,
  );
  const imageIdentityPatched = await patchDocxImageIdentities(
    imageLayerPatched,
    noteContext.imageIdentityPatches.patches,
  );
  const bookmarkPatched = await patchDocxBookmarks(
    imageIdentityPatched,
    noteContext.bookmarkPatches.patches,
  );
  const paragraphDefaultCollapsedPatched =
    await patchDocxParagraphDefaultCollapsed(
      bookmarkPatched,
      noteContext.paragraphDefaultCollapsedPatches.patches,
    );
  const paragraphIdentityPatched = await patchDocxParagraphIdentities(
    paragraphDefaultCollapsedPatched,
    noteContext.paragraphIdentityPatches.patches,
  );
  const paragraphFormattingChangesPatched =
    await patchDocxParagraphFormattingChanges(
      paragraphIdentityPatched,
      noteContext.paragraphFormattingChangePatches.patches,
    );
  const equationPatched = await patchDocxEquations(
    paragraphFormattingChangesPatched,
    noteContext.equationPatches.patches,
  );
  const patched = await patchDocxPageColor(
    equationPatched,
    normalizedContent.pageColor,
  );
  const preserved = sourcePackage
    ? await preserveDocxSourcePackage(patched, sourcePackage, {
        numberingIdentities: noteContext.numberingSourceIdentities,
      })
    : patched;
  return new Blob([preserved], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function sectionProperties(
  layout: WorkDocumentSectionLayout,
  docx: typeof import('docx'),
): NonNullable<ISectionOptions['properties']> {
  const landscape = layout.orientation === 'landscape';
  const dimensions = documentPageGeometryForLayout(layout);
  const pageMargins = documentPageMarginsForLayout(layout);
  return {
    type: docxSectionType(layout.breakAfter, docx),
    titlePage: normalizeDocumentPageChrome(layout.pageChrome, layout)
      .differentFirstPage,
    page: {
      size: {
        ...dimensions,
        orientation: landscape
          ? docx.PageOrientation.LANDSCAPE
          : docx.PageOrientation.PORTRAIT,
      },
      margin: {
        top: pageMargins.top,
        right: pageMargins.right,
        bottom: pageMargins.bottom,
        left: pageMargins.left,
        header: pageMargins.header,
        footer: pageMargins.footer,
        gutter: pageMargins.gutter,
      },
      pageNumbers: layout.pageNumberStart
        ? { start: layout.pageNumberStart }
        : undefined,
    },
    column: docxSectionColumns(layout, dimensions.width, docx),
  };
}

async function sectionHeaders(
  chrome: ReturnType<typeof normalizeDocumentPageChrome>,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<ISectionOptions['headers']> {
  const headers: {
    default?: InstanceType<typeof docx.Header>;
    first?: InstanceType<typeof docx.Header>;
    even?: InstanceType<typeof docx.Header>;
  } = {};
  const defaultHeader = await headerFromHtml(
    chrome.default.headerHtml,
    docx,
    noteContext,
  );
  if (defaultHeader) headers.default = defaultHeader;
  if (chrome.differentFirstPage) {
    const firstHeader = await headerFromHtml(
      chrome.first.headerHtml,
      docx,
      noteContext,
    );
    if (firstHeader) headers.first = firstHeader;
  }
  if (chrome.differentOddEvenPages) {
    const evenHeader = await headerFromHtml(
      chrome.even.headerHtml,
      docx,
      noteContext,
    );
    if (evenHeader) headers.even = evenHeader;
  }
  return Object.keys(headers).length ? headers : undefined;
}

async function sectionFooters(
  chrome: ReturnType<typeof normalizeDocumentPageChrome>,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<ISectionOptions['footers']> {
  const footers: {
    default?: InstanceType<typeof docx.Footer>;
    first?: InstanceType<typeof docx.Footer>;
    even?: InstanceType<typeof docx.Footer>;
  } = {};
  const defaultFooter = await footerFromContent(
    chrome.default,
    docx,
    noteContext,
  );
  if (defaultFooter) footers.default = defaultFooter;
  if (chrome.differentFirstPage) {
    const firstFooter = await footerFromContent(
      chrome.first,
      docx,
      noteContext,
    );
    if (firstFooter) footers.first = firstFooter;
  }
  if (chrome.differentOddEvenPages) {
    const evenFooter = await footerFromContent(chrome.even, docx, noteContext);
    if (evenFooter) footers.even = evenFooter;
  }
  return Object.keys(footers).length ? footers : undefined;
}

async function headerFromHtml(
  html: string,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<InstanceType<typeof docx.Header> | undefined> {
  const children = await pageChromeBlocks(html, docx, noteContext);
  return children.length ? new docx.Header({ children }) : undefined;
}

async function footerFromContent(
  content: ReturnType<typeof normalizeDocumentPageChrome>['default'],
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<InstanceType<typeof docx.Footer> | undefined> {
  const children = await pageChromeBlocks(
    content.footerHtml,
    docx,
    noteContext,
  );
  if (content.showPageNumber) {
    children.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        children: [new docx.TextRun({ children: [docx.PageNumber.CURRENT] })],
      }),
    );
  }
  return children.length ? new docx.Footer({ children }) : undefined;
}

function docxSectionType(
  type: WorkDocumentSectionBreakType,
  docx: typeof import('docx'),
): (typeof docx.SectionType)[keyof typeof docx.SectionType] {
  if (type === 'continuous') return docx.SectionType.CONTINUOUS;
  if (type === 'evenPage') return docx.SectionType.EVEN_PAGE;
  if (type === 'oddPage') return docx.SectionType.ODD_PAGE;
  if (type === 'nextColumn') return docx.SectionType.NEXT_COLUMN;
  return docx.SectionType.NEXT_PAGE;
}

async function blockToFileChildren(
  element: HTMLElement,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<FileChild[]> {
  const tag = element.tagName.toLowerCase();
  if (element.hasAttribute('data-page-break')) {
    return [new docx.Paragraph({ children: [new docx.PageBreak()] })];
  }
  if (element.hasAttribute('data-document-bibliography')) {
    return [
      docxBibliographyParagraph(element, docx, paragraphBidirectional(element)),
    ];
  }
  if (element.hasAttribute('data-document-caption')) {
    return [
      docxCaptionParagraph(
        element,
        await paragraphRuns(element, docx, noteContext),
        docx,
        paragraphBidirectional(element),
      ),
    ];
  }
  if (tag === 'table')
    return [await tableToDocx(element as HTMLTableElement, docx, noteContext)];
  if (tag === 'img') {
    return [
      new docx.Paragraph({
        children: [
          await imageToDocx(
            element as HTMLImageElement,
            docx,
            noteContext.imageCropPatches,
            noteContext.imageWrapPatches,
            noteContext.imageLayerPatches,
            noteContext.imageIdentityPatches,
          ),
        ],
      }),
    ];
  }
  if (tag === 'ul' || tag === 'ol') {
    return listToDocxParagraphs(element, docx, noteContext, (block) =>
      paragraphRuns(block, docx, noteContext),
    );
  }
  const runs = await paragraphRuns(element, docx, noteContext);
  const heading = paragraphHeadingLevel(tag, docx);
  return [
    new docx.Paragraph({
      children: runs.length ? runs : [new docx.TextRun('')],
      heading,
      alignment: paragraphAlignment(element, docx),
      spacing: paragraphSpacingOptions(element, Boolean(heading), docx),
      indent: paragraphIndent(element, tag),
      tabStops: paragraphTabStops(element, docx),
      ...paragraphPaginationOptions(element),
      ...paragraphDirectionOptions(element),
      ...documentParagraphBordersDocxOptions(
        element,
        noteContext.paragraphBorderPatches,
      ),
      ...documentParagraphShadingDocxOptions(element, noteContext.themePatches),
    }),
  ];
}

function paragraphHeadingLevel(tag: string, docx: typeof import('docx')) {
  if (tag === 'h1') return docx.HeadingLevel.HEADING_1;
  if (tag === 'h2') return docx.HeadingLevel.HEADING_2;
  if (tag === 'h3') return docx.HeadingLevel.HEADING_3;
  if (tag === 'h4') return docx.HeadingLevel.HEADING_4;
  if (tag === 'h5') return docx.HeadingLevel.HEADING_5;
  if (tag === 'h6') return docx.HeadingLevel.HEADING_6;
  return undefined;
}

async function pageChromeBlocks(
  html: string,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<
  Array<InstanceType<typeof docx.Paragraph> | InstanceType<typeof docx.Table>>
> {
  if (!html.trim()) return [];
  const document = new DOMParser().parseFromString(html, 'text/html');
  const children: Array<
    InstanceType<typeof docx.Paragraph> | InstanceType<typeof docx.Table>
  > = [];
  for (const element of Array.from(document.body.children)) {
    const blocks = await blockToFileChildren(
      element as HTMLElement,
      docx,
      noteContext,
    );
    for (const block of blocks) {
      if (block instanceof docx.Paragraph || block instanceof docx.Table)
        children.push(block);
    }
  }
  if (!children.length && document.body.textContent?.trim()) {
    children.push(
      new docx.Paragraph({
        children: await inlineRuns(document.body, docx, noteContext),
        ...paragraphDirectionOptions(document.body),
      }),
    );
  }
  return children;
}

async function paragraphRuns(
  element: HTMLElement,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<ParagraphChild[]> {
  const identityMarker = noteContext.paragraphIdentityPatches.marker(element);
  noteContext.paragraphDefaultCollapsedPatches.register(
    identityMarker,
    element,
  );
  const paragraphFormattingChangeMarker =
    element.hasAttribute('data-document-change') &&
    element.dataset.changeKind === 'paragraph-formatting'
      ? noteContext.paragraphFormattingChangePatches.register(
          element,
          docxRevisionId(element, noteContext),
        )
      : null;
  return [
    new docx.TextRun(identityMarker),
    ...(paragraphFormattingChangeMarker
      ? [new docx.TextRun(paragraphFormattingChangeMarker)]
      : []),
    ...(await inlineRuns(element, docx, noteContext)),
  ];
}

async function inlineRuns(
  root: HTMLElement,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<ParagraphChild[]> {
  const runs: ParagraphChild[] = [];
  const visit = async (
    node: Node,
    inherited: IRunOptions = {},
    revision?: DocxTextRevision,
  ): Promise<ParagraphChild[]> => {
    if (node.nodeType === Node.TEXT_NODE) {
      if (!node.textContent) return [];
      if (revision?.kind === 'insertion') {
        return [
          new docx.InsertedTextRun({
            ...inherited,
            id: revision.id,
            author: revision.author,
            date: revision.date,
            text: node.textContent,
          }),
        ];
      }
      if (revision?.kind === 'deletion') {
        return [
          new docx.DeletedTextRun({
            ...inherited,
            id: revision.id,
            author: revision.author,
            date: revision.date,
            text: node.textContent,
          }),
        ];
      }
      return [new docx.TextRun({ ...inherited, text: node.textContent })];
    }
    if (!(node instanceof HTMLElement)) return [];
    const tag = node.tagName.toLowerCase();
    if (node.hasAttribute('data-document-tab')) return [new docx.Tab()];
    if (node.hasAttribute('data-document-citation'))
      return [docxCitationRun(node, docx)];
    if (node.hasAttribute('data-document-field'))
      return [docxDocumentFieldRun(node, docx)];
    if (node.hasAttribute('data-document-equation')) {
      const equation = noteContext.equationPatches.marker(node);
      return [new docx.TextRun(equation ?? node.textContent ?? '')];
    }
    if (node.hasAttribute('data-document-cross-reference'))
      return docxCrossReferenceRuns(node, docx);
    if (node.hasAttribute('data-document-note-reference')) {
      const kind = documentNoteKind(node.dataset.noteKind);
      const id = node.dataset.noteId?.trim();
      const noteId =
        kind && id ? noteContext.ids.get(documentNoteKey(kind, id)) : undefined;
      if (!kind || !noteId) return [];
      return [
        kind === 'footnote'
          ? new docx.FootnoteReferenceRun(noteId)
          : (new docx.EndnoteReferenceRun(noteId) as ParagraphChild),
      ];
    }
    const textRevisionKind =
      tag === 'del' ? 'deletion' : tag === 'ins' ? 'insertion' : null;
    const change =
      node.hasAttribute('data-document-change') && textRevisionKind
        ? docxTextRevision(node, textRevisionKind, noteContext)
        : revision;
    const formattingChange =
      node.hasAttribute('data-document-change') &&
      node.dataset.changeKind === 'formatting'
        ? noteContext.formattingChangePatches.register(
            node,
            docxRevisionId(node, noteContext),
          )
        : null;
    const commentBoundary = node.hasAttribute('data-document-comment')
      ? nextDocxCommentBoundary(node.dataset.commentId, noteContext)
      : null;
    const backgroundColor = cssColorToHex(node.style.backgroundColor);
    const direction = domDirection(node);
    const resolvedColor = cssColorToHex(node.style.color) ?? inherited.color;
    const themeColorMarker = noteContext.themePatches.marker(
      'color',
      parseDocxThemeReference(node.dataset.officeThemeColor),
      resolvedColor ? `#${resolvedColor}` : null,
    );
    const resolvedShading = backgroundColor
      ? { fill: backgroundColor }
      : inherited.shading;
    const themeFillMarker = noteContext.themePatches.marker(
      'fill',
      parseDocxThemeReference(node.dataset.officeThemeFill),
      resolvedShading?.fill ? `#${resolvedShading.fill}` : null,
    );
    const explicitTextCase = normalizeDocumentTextCase(
      node.dataset.officeTextCase,
    );
    const explicitCharacterSpacing =
      documentCharacterSpacingTwipsFromElement(node);
    const explicitCharacterPosition =
      documentCharacterPositionHalfPointsFromElement(node);
    if (explicitCharacterSpacing === 0) {
      noteContext.hasExplicitZeroCharacterSpacing = true;
    }
    const characterSpacing =
      explicitCharacterSpacing === null
        ? inherited.characterSpacing
        : docxCharacterSpacingValue(explicitCharacterSpacing);
    const position =
      explicitCharacterPosition === null
        ? inherited.position
        : (docxCharacterPositionValue(
            explicitCharacterPosition,
          ) as IRunOptions['position']);
    const textCaseOptions = docxTextCaseRunOptions(explicitTextCase, inherited);
    const underline = docxUnderlineRunOptions(
      node,
      inherited.underline,
      noteContext.themePatches,
    );
    const strike = docxStrikeRunOptions(node, inherited);
    const style: IRunOptions = {
      ...inherited,
      bold: inherited.bold || tag === 'strong' || tag === 'b',
      italics: inherited.italics || tag === 'em' || tag === 'i',
      underline,
      ...strike,
      subScript: inherited.subScript || tag === 'sub',
      superScript: inherited.superScript || tag === 'sup',
      color: themeColorMarker ?? resolvedColor,
      font: cssFontFamily(node.style.fontFamily) ?? inherited.font,
      size: cssFontSize(node.style.fontSize) ?? inherited.size,
      characterSpacing,
      position,
      shading: themeFillMarker ? { fill: themeFillMarker } : resolvedShading,
      snapToGrid:
        dataBoolean(node.dataset.officeWordSnapToGrid) ?? inherited.snapToGrid,
      rightToLeft:
        direction === undefined ? inherited.rightToLeft : direction === 'rtl',
      ...textCaseOptions,
    };
    if (tag === 'br') {
      return [new docx.TextRun({ ...style, break: 1 })];
    }
    if (tag === 'img')
      return [
        await imageToDocx(
          node as HTMLImageElement,
          docx,
          noteContext.imageCropPatches,
          noteContext.imageWrapPatches,
          noteContext.imageLayerPatches,
          noteContext.imageIdentityPatches,
        ),
      ];
    if (tag === 'span' && node.dataset.documentBookmarkBoundary === 'true') {
      const marker = noteContext.bookmarkPatches.register(node);
      return marker ? [new docx.TextRun(marker)] : [];
    }
    const children: ParagraphChild[] = [];
    for (const child of node.childNodes)
      children.push(...(await visit(child, style, change)));
    const href =
      tag === 'a'
        ? normalizeDocumentHref(node.getAttribute('href') ?? '')
        : null;
    const linked = href
      ? [
          href.startsWith('#')
            ? new docx.InternalHyperlink({
                anchor: href.slice(1),
                children,
              })
            : new docx.ExternalHyperlink({ link: href, children }),
        ]
      : children;
    const result = formattingChange
      ? [
          new docx.TextRun(formattingChange.start),
          ...linked,
          new docx.TextRun(formattingChange.end),
        ]
      : linked;
    if (!commentBoundary) return result;
    return [
      ...(commentBoundary.start
        ? [new docx.CommentRangeStart(commentBoundary.id)]
        : []),
      ...result,
      ...(commentBoundary.end
        ? [
            new docx.CommentRangeEnd(commentBoundary.id),
            new docx.CommentReference(commentBoundary.id),
          ]
        : []),
    ];
  };
  const direction = domDirection(root);
  const inheritedDirection: IRunOptions =
    direction === undefined ? {} : { rightToLeft: direction === 'rtl' };
  for (const node of root.childNodes)
    runs.push(...(await visit(node, inheritedDirection)));
  return runs;
}

function docxStrikeRunOptions(
  element: HTMLElement,
  inherited: IRunOptions,
): Pick<IRunOptions, 'doubleStrike' | 'strike'> {
  if (
    element.hasAttribute(DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE) &&
    !element.hasAttribute(DOCUMENT_STRIKE_STYLE_ATTRIBUTE)
  ) {
    return {
      doubleStrike: inherited.doubleStrike,
      strike: inherited.strike,
    };
  }
  const formatting = documentStrikeFormattingFromElement(element);
  if (!formatting) {
    return {
      doubleStrike: inherited.doubleStrike,
      strike: inherited.strike,
    };
  }
  return {
    doubleStrike: formatting.style === 'double',
    strike: formatting.style === 'single',
  };
}

function docxUnderlineRunOptions(
  element: HTMLElement,
  inherited: IRunOptions['underline'],
  themePatches: DocxThemePatchCollector,
): IRunOptions['underline'] {
  if (
    element.hasAttribute(DOCUMENT_STRIKE_STYLE_ATTRIBUTE) &&
    !element.hasAttribute(DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE)
  ) {
    return inherited;
  }
  const formatting = documentUnderlineFormattingFromElement(element);
  if (!formatting) return inherited;
  const color = docxUnderlineColor(formatting, themePatches);
  return {
    type: formatting.style,
    ...(color ? { color } : {}),
  };
}

function docxUnderlineColor(
  formatting: WorkDocumentUnderlineFormatting,
  themePatches: DocxThemePatchCollector,
): string | undefined {
  const direct = formatting.color ? cssColorToHex(formatting.color) : undefined;
  return (
    themePatches.marker(
      'underline',
      formatting.themeColor ?? null,
      direct ? `#${direct}` : null,
    ) ?? direct
  );
}

function docxTextCaseRunOptions(
  explicitTextCase: ReturnType<typeof normalizeDocumentTextCase>,
  inherited: IRunOptions,
): Pick<IRunOptions, 'allCaps' | 'smallCaps'> {
  if (explicitTextCase === null) {
    if (inherited.smallCaps !== undefined) {
      return { allCaps: undefined, smallCaps: inherited.smallCaps };
    }
    return inherited.allCaps === undefined
      ? {}
      : { allCaps: inherited.allCaps, smallCaps: undefined };
  }
  if (explicitTextCase === 'all-caps') {
    return { allCaps: true, smallCaps: undefined };
  }
  if (explicitTextCase === 'small-caps') {
    return { allCaps: undefined, smallCaps: true };
  }
  return inherited.smallCaps
    ? { allCaps: undefined, smallCaps: false }
    : { allCaps: false, smallCaps: undefined };
}

function docxTextRevision(
  element: HTMLElement,
  kind: DocxTextRevision['kind'],
  context: DocxNoteContext,
): DocxTextRevision {
  const id = docxRevisionId(element, context);
  const sourceDate = element.dataset.changeDate?.trim() ?? '';
  const time = Date.parse(sourceDate);
  return {
    kind,
    id,
    author: element.dataset.changeAuthor?.trim() || 'A3S Work',
    date: Number.isFinite(time)
      ? new Date(time).toISOString()
      : new Date().toISOString(),
  };
}

function docxRevisionId(
  element: HTMLElement,
  context: DocxNoteContext,
): number {
  const key =
    element.dataset.changeId?.trim() || `change-${context.nextChangeId}`;
  let id = context.changeIds.get(key);
  if (!id) {
    id = context.nextChangeId;
    context.nextChangeId += 1;
    context.changeIds.set(key, id);
  }
  return id;
}

function documentHasTrackedChanges(html: string): boolean {
  const document = new DOMParser().parseFromString(html, 'text/html');
  return Boolean(
    document.body.querySelector(
      'ins[data-document-change], del[data-document-change], span[data-document-change][data-change-kind="formatting"], [data-document-change][data-change-kind="paragraph-formatting"]',
    ),
  );
}

function anchoredDocumentComments(
  content: WorkDocumentContent,
): WorkDocumentComment[] {
  const document = new DOMParser().parseFromString(content.html, 'text/html');
  const ids = Array.from(
    document.body.querySelectorAll<HTMLElement>(
      '[data-document-comment][data-comment-id]',
    ),
  )
    .map((element) => element.dataset.commentId?.trim() ?? '')
    .filter(Boolean);
  const stored = new Map(
    (content.comments ?? []).map((comment) => [comment.id, comment] as const),
  );
  return Array.from(new Set(ids)).map(
    (id) =>
      stored.get(id) ?? {
        id,
        author: '未知审阅者',
        date: '',
        text: '此批注的内容不可用。',
        resolved: false,
      },
  );
}

function documentCommentRangeCounts(html: string): Map<string, number> {
  const document = new DOMParser().parseFromString(html, 'text/html');
  const counts = new Map<string, number>();
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>(
      '[data-document-comment][data-comment-id]',
    ),
  )) {
    const id = element.dataset.commentId?.trim();
    if (id) counts.set(id, (counts.get(id) ?? 0) + 1);
  }
  return counts;
}

function nextDocxCommentBoundary(
  sourceId: string | undefined,
  context: DocxNoteContext,
): { id: number; start: boolean; end: boolean } | null {
  const key = sourceId?.trim() ?? '';
  const id = context.commentIds.get(key);
  const count = context.commentRangeCounts.get(key) ?? 0;
  if (id === undefined || !count) return null;
  const seen = (context.commentRangeSeen.get(key) ?? 0) + 1;
  context.commentRangeSeen.set(key, seen);
  return { id, start: seen === 1, end: seen === count };
}

async function tableToDocx(
  element: HTMLTableElement,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<InstanceType<typeof docx.Table>> {
  const rows: InstanceType<typeof docx.TableRow>[] = [];
  let inferLeadingHeader = true;
  const ownedRows = Array.from(element.rows).filter(
    (row) => row.closest('table') === element,
  );
  for (const row of ownedRows) {
    const ownedCells = Array.from(row.cells).filter(
      (cell) => cell.closest('tr') === row,
    );
    const rowIdentityMarker = ownedCells.length
      ? noteContext.paragraphIdentityPatches.rowMarker(row)
      : null;
    const cells: InstanceType<typeof docx.TableCell>[] = [];
    for (const [cellIndex, cell] of ownedCells.entries()) {
      const children: Array<
        InstanceType<typeof docx.Paragraph> | InstanceType<typeof docx.Table>
      > = [];
      if (cellIndex === 0 && rowIdentityMarker) {
        children.push(
          new docx.Paragraph({
            children: [new docx.TextRun(rowIdentityMarker)],
          }),
        );
      }
      const blocks = Array.from(cell.children);
      if (blocks.length) {
        for (const block of blocks) {
          for (const child of await blockToFileChildren(
            block as HTMLElement,
            docx,
            noteContext,
          )) {
            if (
              child instanceof docx.Paragraph ||
              child instanceof docx.Table
            ) {
              children.push(child);
            }
          }
        }
      } else {
        children.push(
          new docx.Paragraph({
            children: await inlineRuns(cell, docx, noteContext),
            ...paragraphDirectionOptions(cell),
          }),
        );
      }
      cells.push(
        new docx.TableCell({
          children: children.length ? children : [new docx.Paragraph('')],
          columnSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
          rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
          ...documentTableCellDocxOptions(cell, docx, noteContext.themePatches),
          ...documentTableCellSizingDocxOptions(cell, docx),
        }),
      );
    }
    const explicitHeader = dataBoolean(row.dataset.officeRepeatHeader);
    const inferredHeader =
      inferLeadingHeader &&
      ownedCells.length > 0 &&
      ownedCells.every((cell) => cell.tagName.toLowerCase() === 'th');
    const tableHeader = explicitHeader ?? inferredHeader;
    if (!tableHeader) inferLeadingHeader = false;
    rows.push(
      new docx.TableRow({
        children: cells,
        cantSplit: dataBoolean(row.dataset.officeCantSplit),
        tableHeader: tableHeader ? true : undefined,
        ...documentTableRowSizingDocxOptions(row, docx),
      }),
    );
  }
  return new docx.Table({
    rows,
    ...documentTableSizingDocxOptions(element, docx),
  });
}

async function createNoteRecords(
  notes: WorkDocumentNote[],
  kind: WorkDocumentNoteKind,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<
  Record<string, { children: InstanceType<typeof docx.Paragraph>[] }>
> {
  const records: Record<
    string,
    { children: InstanceType<typeof docx.Paragraph>[] }
  > = {};
  for (const note of notes) {
    if (note.kind !== kind) continue;
    const id = noteContext.ids.get(documentNoteKey(note.kind, note.id));
    if (id === undefined) continue;
    records[String(id)] = {
      children: await noteBlocks(note.html, docx, noteContext),
    };
  }
  return records;
}

async function noteBlocks(
  html: string,
  docx: typeof import('docx'),
  noteContext: DocxNoteContext,
): Promise<InstanceType<typeof docx.Paragraph>[]> {
  const blocks = await pageChromeBlocks(html, docx, noteContext);
  if (!(blocks[0] instanceof docx.Paragraph)) {
    blocks.unshift(new docx.Paragraph(''));
  }
  // The upstream type narrows note children to Paragraph even though Word's
  // block-content model and the serializer accept Table after the leading
  // reference paragraph.
  return blocks as InstanceType<typeof docx.Paragraph>[];
}
