import { normalizeDocumentBookmarkReferencesHtml } from './work-document-bookmark-references';
import { normalizeDocumentBookmarksHtml } from './work-document-bookmarks';
import { normalizeDocumentCaptionsHtml } from './work-document-captions';
import { normalizeDocumentCitationsHtml } from './work-document-citations';
import {
  DEFAULT_DOCUMENT_COLUMNS,
  normalizeDocumentColumns,
  parseDocumentColumns,
  serializeDocumentColumns,
} from './work-document-columns';
import { normalizeDocumentFieldsHtml } from './work-document-fields';
import { normalizeDocumentTableOfContentsHtml } from './work-document-table-of-contents';
import { clampDocumentMargin, documentMargins } from './work-document-layout';
import { documentModelForHtml } from './work-document-model';
import { normalizeDocumentNotesHtml } from './work-document-notes';
import {
  normalizeDocumentPageBorders,
  parseDocumentPageBorders,
  serializeDocumentPageBorders,
} from './work-document-page-borders';
import {
  documentPageMarginBody,
  normalizeDocumentPageMargins,
  parseDocumentPageMargins,
  serializeDocumentPageMargins,
} from './work-document-page-margins';
import {
  applyDocumentPageGeometry,
  normalizeDocumentPageGeometry,
  normalizeDocumentPaperSize,
  normalizeDocumentPaperSource,
  parseDocumentPageGeometry,
  parseDocumentPaperSource,
  serializeDocumentPageGeometry,
  serializeDocumentPaperSource,
} from './work-document-page-size';
import {
  documentPageChromeLegacyFields,
  normalizeDocumentPageChrome,
  parseDocumentPageChrome,
  serializeDocumentPageChrome,
} from './work-document-page-chrome';
import {
  encodeDocumentTablePropertyRevisionOmml,
  decodeDocumentTablePropertyRevisionOmml,
} from './work-document-table-property-revision';
import type {
  WorkDocumentContent,
  WorkDocumentGrid,
  WorkDocumentGridType,
  WorkDocumentLnNumRestart,
  WorkDocumentLnNumType,
  WorkDocumentPgNumFmt,
  WorkDocumentPgNumType,
  WorkDocumentSectionBreakType,
  WorkDocumentSectionFormattingChange,
  WorkDocumentSectionLayout,
  WorkDocumentSectionVerticalAlign,
} from './work-types';

export interface WorkDocumentSection {
  id: string;
  layout: WorkDocumentSectionLayout;
  html: string;
}

export interface DocumentSectionNodeAttributes {
  id: string;
  pageSize: WorkDocumentSectionLayout['pageSize'];
  orientation: WorkDocumentSectionLayout['orientation'];
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  columnCount: number;
  columnSpacing: number;
  columnSeparator: boolean;
  columnLayout: string;
  breakAfter: WorkDocumentSectionBreakType;
  headerText: string;
  footerText: string;
  showPageNumbers: boolean;
  pageNumberStart: number | null;
  pageChrome: string;
  pageBorders: string;
  pageMargins: string;
  pageGeometry: string;
  paperSource: string;
  documentGridType: WorkDocumentGridType | '';
  documentGridLinePitch: number | null;
  lnNumCountBy: number | null;
  lnNumStart: number | null;
  lnNumDistance: number | null;
  lnNumRestart: WorkDocumentLnNumRestart | '';
  pgNumFmt: WorkDocumentPgNumFmt | '';
  pgNumStart: number | null;
  formProt: boolean | null;
  noEndnote: boolean | null;
  verticalAlign: WorkDocumentSectionVerticalAlign | '';
  propertyRevisionOmml: string;
  sectionChangeKind: 'section-formatting' | null;
  sectionChangeId: string;
  sectionChangeAuthor: string;
  sectionChangeDate: string;
  sectionChangeBefore: string;
}

const SECTION_SELECTOR = 'section[data-document-section]';

export function documentInitialSectionLayout(
  content: WorkDocumentContent,
): WorkDocumentSectionLayout {
  const pageChrome = normalizeDocumentPageChrome(content.pageChrome, content);
  const legacy = documentPageChromeLegacyFields(pageChrome);
  const pageBorders = normalizeDocumentPageBorders(content.pageBorders);
  const pageMargins = normalizeDocumentPageMargins(content.pageMargins);
  const pageGeometry = normalizeDocumentPageGeometry(content.pageGeometry);
  const paperSource = normalizeDocumentPaperSource(content.paperSource);
  const margins = documentMargins(content);
  const layout: WorkDocumentSectionLayout = {
    pageSize: normalizeDocumentPaperSize(content.pageSize),
    orientation: content.orientation ?? 'portrait',
    margins: documentPageMarginBody(pageMargins, margins),
    columns: normalizeDocumentColumns(content.columns),
    breakAfter: 'nextPage',
    headerText: legacy.headerText,
    footerText: legacy.footerText,
    showPageNumbers: legacy.showPageNumbers,
    pageNumberStart: validPageNumber(content.pageNumberStart),
    pageChrome,
    ...(pageBorders ? { pageBorders } : {}),
    ...(pageMargins ? { pageMargins } : {}),
    ...(pageGeometry ? { pageGeometry } : {}),
    ...(paperSource ? { paperSource } : {}),
  };
  return pageGeometry
    ? applyDocumentPageGeometry(layout, pageGeometry)
    : layout;
}

export function normalizeDocumentHtml(content: WorkDocumentContent): string {
  const document = new DOMParser().parseFromString(content.html, 'text/html');
  const directSections = Array.from(document.body.children).filter((element) =>
    element.matches(SECTION_SELECTOR),
  );
  if (!directSections.length) {
    const section = document.createElement('section');
    applyDocumentSectionDomAttributes(
      section,
      documentInitialSectionLayout(content),
      'document-section-1',
    );
    while (document.body.firstChild) section.append(document.body.firstChild);
    if (!section.childNodes.length) section.innerHTML = '<p></p>';
    document.body.append(section);
    return normalizeDocumentSemanticHtml(
      document.body.innerHTML,
      content.bibliography,
    );
  }

  let activeSection = directSections[0] as HTMLElement;
  for (const node of Array.from(document.body.childNodes)) {
    if (node instanceof HTMLElement && node.matches(SECTION_SELECTOR)) {
      activeSection = node;
      continue;
    }
    activeSection.append(node);
  }
  let fallback = documentInitialSectionLayout(content);
  directSections.forEach((element, index) => {
    const layout = documentSectionLayoutFromElement(
      element as HTMLElement,
      fallback,
    );
    applyDocumentSectionDomAttributes(
      element as HTMLElement,
      layout,
      element.getAttribute('data-section-id') ||
        `document-section-${index + 1}`,
    );
    if (!element.childNodes.length) element.innerHTML = '<p></p>';
    fallback = layout;
  });
  return normalizeDocumentSemanticHtml(
    document.body.innerHTML,
    content.bibliography,
  );
}

export function documentSections(
  content: WorkDocumentContent,
): WorkDocumentSection[] {
  const document = new DOMParser().parseFromString(
    normalizeDocumentHtml(content),
    'text/html',
  );
  let fallback = documentInitialSectionLayout(content);
  return Array.from(
    document.body.querySelectorAll<HTMLElement>(`:scope > ${SECTION_SELECTOR}`),
  ).map((element, index) => {
    const layout = documentSectionLayoutFromElement(element, fallback);
    fallback = layout;
    return {
      id: element.dataset.sectionId || `document-section-${index + 1}`,
      layout,
      html: element.innerHTML || '<p></p>',
    };
  });
}

export function documentSectionNodeAttributes(
  layout: WorkDocumentSectionLayout,
  id: string,
): DocumentSectionNodeAttributes {
  const columns = normalizeDocumentColumns(layout.columns);
  const pageChrome = normalizeDocumentPageChrome(layout.pageChrome, layout);
  const legacy = documentPageChromeLegacyFields(pageChrome);
  return {
    id,
    pageSize: layout.pageSize,
    orientation: layout.orientation,
    marginTop: layout.margins.top,
    marginRight: layout.margins.right,
    marginBottom: layout.margins.bottom,
    marginLeft: layout.margins.left,
    columnCount: columns.count,
    columnSpacing: columns.spacing,
    columnSeparator: columns.separator,
    columnLayout: serializeDocumentColumns(columns),
    breakAfter: validBreakType(layout.breakAfter),
    headerText: legacy.headerText ?? '',
    footerText: legacy.footerText ?? '',
    showPageNumbers: Boolean(legacy.showPageNumbers),
    pageNumberStart:
      validPageNumber(layout.pgNumType?.start ?? layout.pageNumberStart) ??
      null,
    pageChrome: serializeDocumentPageChrome(pageChrome),
    pageBorders: serializeDocumentPageBorders(layout.pageBorders) ?? '',
    pageMargins: serializeDocumentPageMargins(layout.pageMargins) ?? '',
    pageGeometry: serializeDocumentPageGeometry(layout.pageGeometry) ?? '',
    paperSource: serializeDocumentPaperSource(layout.paperSource) ?? '',
    documentGridType: layout.documentGrid?.type ?? '',
    documentGridLinePitch:
      normalizedDocumentGrid(layout.documentGrid)?.linePitch ?? null,
    ...lnNumTypeNodeFields(layout.lnNumType),
    ...pgNumTypeNodeFields(layout.pgNumType, layout.pageNumberStart),
    formProt: layout.formProt ?? null,
    noEndnote: layout.noEndnote ?? null,
    verticalAlign: layout.verticalAlign ?? '',
    propertyRevisionOmml: layout.propertyRevisionOmml
      ? encodeDocumentTablePropertyRevisionOmml(layout.propertyRevisionOmml)
      : '',
    sectionChangeKind: layout.formattingChange?.kind ?? null,
    sectionChangeId: layout.formattingChange?.id ?? '',
    sectionChangeAuthor: layout.formattingChange?.author ?? '',
    sectionChangeDate: layout.formattingChange?.date ?? '',
    sectionChangeBefore: layout.formattingChange?.before ?? '',
  };
}

export function documentSectionLayoutFromNodeAttributes(
  attributes: Partial<DocumentSectionNodeAttributes>,
  fallback?: WorkDocumentSectionLayout,
): WorkDocumentSectionLayout {
  const base =
    fallback ??
    ({
      pageSize: 'a4',
      orientation: 'portrait',
      margins: { top: 25, right: 23, bottom: 25, left: 23 },
      columns: DEFAULT_DOCUMENT_COLUMNS,
      breakAfter: 'nextPage',
    } satisfies WorkDocumentSectionLayout);
  const pageChrome = parseDocumentPageChrome(
    attributes.pageChrome,
    {
      headerText: attributes.headerText,
      footerText: attributes.footerText,
      showPageNumbers: attributes.showPageNumbers,
    },
    base.pageChrome,
  );
  const legacy = documentPageChromeLegacyFields(pageChrome);
  const documentGrid = documentGridFromNodeAttributes(attributes, base);
  const lnNumType = lnNumTypeFromNodeAttributes(attributes, base);
  const pgNumType = pgNumTypeFromNodeAttributes(attributes, base);
  const pageBorders = parseDocumentPageBorders(attributes.pageBorders);
  const pageMargins = parseDocumentPageMargins(attributes.pageMargins);
  const pageGeometry = parseDocumentPageGeometry(attributes.pageGeometry);
  const paperSource = parseDocumentPaperSource(attributes.paperSource);
  const margins = {
    top: clampDocumentMargin(
      finiteNumber(attributes.marginTop, base.margins.top),
    ),
    right: clampDocumentMargin(
      finiteNumber(attributes.marginRight, base.margins.right),
    ),
    bottom: clampDocumentMargin(
      finiteNumber(attributes.marginBottom, base.margins.bottom),
    ),
    left: clampDocumentMargin(
      finiteNumber(attributes.marginLeft, base.margins.left),
    ),
  };
  const layout: WorkDocumentSectionLayout = {
    pageSize: normalizeDocumentPaperSize(attributes.pageSize, base.pageSize),
    orientation:
      attributes.orientation === 'landscape'
        ? 'landscape'
        : attributes.orientation === 'portrait'
          ? 'portrait'
          : base.orientation,
    margins: documentPageMarginBody(pageMargins, margins),
    columns: parseDocumentColumns(
      attributes.columnLayout,
      {
        count: attributes.columnCount,
        spacing: attributes.columnSpacing,
        separator: attributes.columnSeparator,
      },
      base.columns,
    ),
    breakAfter: validBreakType(attributes.breakAfter ?? base.breakAfter),
    headerText: legacy.headerText,
    footerText: legacy.footerText,
    showPageNumbers: legacy.showPageNumbers,
    pageNumberStart: validPageNumber(
      pgNumType?.start ?? attributes.pageNumberStart ?? undefined,
    ),
    pageChrome,
    ...(documentGrid ? { documentGrid } : {}),
    ...(lnNumType ? { lnNumType } : {}),
    ...(pgNumType ? { pgNumType } : {}),
    ...(attributes.formProt === true || attributes.formProt === false
      ? { formProt: attributes.formProt }
      : base.formProt !== undefined
        ? { formProt: base.formProt }
        : {}),
    ...(attributes.noEndnote === true || attributes.noEndnote === false
      ? { noEndnote: attributes.noEndnote }
      : base.noEndnote !== undefined
        ? { noEndnote: base.noEndnote }
        : {}),
    ...(attributes.verticalAlign === 'top' ||
    attributes.verticalAlign === 'center' ||
    attributes.verticalAlign === 'both' ||
    attributes.verticalAlign === 'bottom'
      ? { verticalAlign: attributes.verticalAlign }
      : base.verticalAlign !== undefined
        ? { verticalAlign: base.verticalAlign }
        : {}),
    ...(pageBorders ? { pageBorders } : {}),
    ...(pageMargins ? { pageMargins } : {}),
    ...(pageGeometry ? { pageGeometry } : {}),
    ...(paperSource ? { paperSource } : {}),
    ...(decodeDocumentTablePropertyRevisionOmml(
      attributes.propertyRevisionOmml ?? '',
    )
      ? {
          propertyRevisionOmml: decodeDocumentTablePropertyRevisionOmml(
            attributes.propertyRevisionOmml ?? '',
          )!,
        }
      : {}),
    ...(sectionFormattingChangeFromAttributes(attributes)
      ? { formattingChange: sectionFormattingChangeFromAttributes(attributes)! }
      : {}),
  };
  return pageGeometry
    ? applyDocumentPageGeometry(layout, pageGeometry)
    : layout;
}

export function documentSectionDomAttributes(
  layout: WorkDocumentSectionLayout,
  id: string,
): Record<string, string> {
  const attributes = documentSectionNodeAttributes(layout, id);
  return {
    'data-document-section': 'true',
    'data-section-id': attributes.id,
    'data-section-page-size': attributes.pageSize,
    'data-section-orientation': attributes.orientation,
    'data-section-margin-top': String(attributes.marginTop),
    'data-section-margin-right': String(attributes.marginRight),
    'data-section-margin-bottom': String(attributes.marginBottom),
    'data-section-margin-left': String(attributes.marginLeft),
    'data-section-column-count': String(attributes.columnCount),
    'data-section-column-spacing': String(attributes.columnSpacing),
    'data-section-column-separator': String(attributes.columnSeparator),
    'data-section-column-layout': attributes.columnLayout,
    'data-section-break-after': attributes.breakAfter,
    'data-section-header-text': attributes.headerText,
    'data-section-footer-text': attributes.footerText,
    'data-section-show-page-numbers': String(attributes.showPageNumbers),
    'data-section-page-number-start':
      attributes.pageNumberStart === null
        ? ''
        : String(attributes.pageNumberStart),
    'data-section-page-chrome': attributes.pageChrome,
    'data-section-page-borders': attributes.pageBorders,
    'data-section-page-margins': attributes.pageMargins,
    'data-section-page-geometry': attributes.pageGeometry,
    'data-section-paper-source': attributes.paperSource,
    'data-section-document-grid-type': attributes.documentGridType,
    'data-section-document-grid-line-pitch':
      attributes.documentGridLinePitch === null
        ? ''
        : String(attributes.documentGridLinePitch),
    'data-section-ln-num-count-by':
      attributes.lnNumCountBy === null ? '' : String(attributes.lnNumCountBy),
    'data-section-ln-num-start':
      attributes.lnNumStart === null ? '' : String(attributes.lnNumStart),
    'data-section-ln-num-distance':
      attributes.lnNumDistance === null ? '' : String(attributes.lnNumDistance),
    'data-section-ln-num-restart': attributes.lnNumRestart,
    'data-section-pg-num-fmt': attributes.pgNumFmt,
    'data-section-pg-num-start':
      attributes.pgNumStart === null ? '' : String(attributes.pgNumStart),
    'data-section-form-prot':
      attributes.formProt === null ? '' : String(attributes.formProt),
    'data-section-no-endnote':
      attributes.noEndnote === null ? '' : String(attributes.noEndnote),
    'data-section-vertical-align': attributes.verticalAlign,
    ...(attributes.propertyRevisionOmml
      ? {
          'data-section-property-revision-omml':
            attributes.propertyRevisionOmml,
        }
      : {}),
    ...(attributes.sectionChangeKind === 'section-formatting'
      ? {
          'data-document-change': 'true',
          'data-change-kind': 'section-formatting',
          'data-change-id': attributes.sectionChangeId,
          'data-change-author': attributes.sectionChangeAuthor,
          'data-change-date': attributes.sectionChangeDate,
          'data-change-before': attributes.sectionChangeBefore,
        }
      : {}),
  };
}

export function documentSectionLayoutFromElement(
  element: HTMLElement,
  fallback: WorkDocumentSectionLayout,
): WorkDocumentSectionLayout {
  return documentSectionLayoutFromNodeAttributes(
    {
      pageSize: element.dataset
        .sectionPageSize as WorkDocumentSectionLayout['pageSize'],
      orientation: element.dataset
        .sectionOrientation as WorkDocumentSectionLayout['orientation'],
      marginTop: numberValue(element.dataset.sectionMarginTop),
      marginRight: numberValue(element.dataset.sectionMarginRight),
      marginBottom: numberValue(element.dataset.sectionMarginBottom),
      marginLeft: numberValue(element.dataset.sectionMarginLeft),
      columnCount: numberValue(element.dataset.sectionColumnCount),
      columnSpacing: numberValue(element.dataset.sectionColumnSpacing),
      columnSeparator: element.dataset.sectionColumnSeparator === 'true',
      columnLayout: element.dataset.sectionColumnLayout ?? '',
      breakAfter: element.dataset
        .sectionBreakAfter as WorkDocumentSectionBreakType,
      headerText: element.dataset.sectionHeaderText ?? '',
      footerText: element.dataset.sectionFooterText ?? '',
      showPageNumbers: element.dataset.sectionShowPageNumbers === 'true',
      pageNumberStart:
        numberValue(element.dataset.sectionPageNumberStart) ?? null,
      pageChrome: element.dataset.sectionPageChrome ?? '',
      pageBorders: element.dataset.sectionPageBorders ?? '',
      pageMargins: element.dataset.sectionPageMargins ?? '',
      pageGeometry: element.dataset.sectionPageGeometry ?? '',
      paperSource: element.dataset.sectionPaperSource ?? '',
      documentGridType: element.dataset
        .sectionDocumentGridType as WorkDocumentGridType,
      documentGridLinePitch:
        numberValue(element.dataset.sectionDocumentGridLinePitch) ?? null,
      lnNumCountBy: numberValue(element.dataset.sectionLnNumCountBy) ?? null,
      lnNumStart: numberValue(element.dataset.sectionLnNumStart) ?? null,
      lnNumDistance: numberValue(element.dataset.sectionLnNumDistance) ?? null,
      lnNumRestart: (element.dataset.sectionLnNumRestart ?? '') as
        | WorkDocumentLnNumRestart
        | '',
      pgNumFmt: (element.dataset.sectionPgNumFmt ?? '') as
        | WorkDocumentPgNumFmt
        | '',
      pgNumStart: numberValue(element.dataset.sectionPgNumStart) ?? null,
      formProt:
        element.dataset.sectionFormProt === 'true'
          ? true
          : element.dataset.sectionFormProt === 'false'
            ? false
            : null,
      noEndnote:
        element.dataset.sectionNoEndnote === 'true'
          ? true
          : element.dataset.sectionNoEndnote === 'false'
            ? false
            : null,
      verticalAlign: (element.dataset.sectionVerticalAlign ?? '') as
        | WorkDocumentSectionVerticalAlign
        | '',
      propertyRevisionOmml: element.dataset.sectionPropertyRevisionOmml ?? '',
      sectionChangeKind:
        element.getAttribute('data-document-change') === 'true' &&
        element.getAttribute('data-change-kind') === 'section-formatting'
          ? 'section-formatting'
          : null,
      sectionChangeId: element.getAttribute('data-change-id') ?? '',
      sectionChangeAuthor: element.getAttribute('data-change-author') ?? '',
      sectionChangeDate: element.getAttribute('data-change-date') ?? '',
      sectionChangeBefore: element.getAttribute('data-change-before') ?? '',
    },
    fallback,
  );
}

export function syncDocumentContentFromHtml(
  content: WorkDocumentContent,
  html: string,
): WorkDocumentContent {
  const normalized = normalizeDocumentHtml({ ...content, html });
  const first =
    documentSections({ ...content, html: normalized })[0]?.layout ??
    documentInitialSectionLayout(content);
  const nextContent = { ...content };
  const model = documentModelForHtml(content.model, normalized);
  if (model) nextContent.model = model;
  else delete nextContent.model;
  return {
    ...nextContent,
    html: normalized,
    pageSize: first.pageSize,
    orientation: first.orientation,
    margins: first.margins,
    columns: first.columns,
    headerText: first.headerText,
    footerText: first.footerText,
    showPageNumbers: first.showPageNumbers,
    pageNumberStart: first.pageNumberStart,
    pageChrome: first.pageChrome,
    pageBorders: first.pageBorders,
    pageMargins: first.pageMargins,
    pageGeometry: first.pageGeometry,
    paperSource: first.paperSource,
  };
}

export function documentContentLayoutProperties(
  layout: WorkDocumentSectionLayout,
): Omit<WorkDocumentContent, 'type' | 'html'> {
  return {
    pageSize: layout.pageSize,
    orientation: layout.orientation,
    margins: layout.margins,
    columns: layout.columns,
    headerText: layout.headerText,
    footerText: layout.footerText,
    showPageNumbers: layout.showPageNumbers,
    pageNumberStart: layout.pageNumberStart,
    pageChrome: layout.pageChrome,
    pageBorders: layout.pageBorders,
    pageMargins: layout.pageMargins,
    pageGeometry: layout.pageGeometry,
    paperSource: layout.paperSource,
  };
}

function normalizeDocumentSemanticHtml(
  source: string,
  bibliography: WorkDocumentContent['bibliography'],
): string {
  return normalizeDocumentBookmarkReferencesHtml(
    normalizeDocumentBookmarksHtml(
      normalizeDocumentCitationsHtml(
        normalizeDocumentTableOfContentsHtml(
          normalizeDocumentFieldsHtml(
            normalizeDocumentCaptionsHtml(normalizeDocumentNotesHtml(source)),
          ),
        ),
        bibliography,
      ),
    ),
  );
}

function applyDocumentSectionDomAttributes(
  element: HTMLElement,
  layout: WorkDocumentSectionLayout,
  id: string,
) {
  for (const [name, value] of Object.entries(
    documentSectionDomAttributes(layout, id),
  )) {
    element.setAttribute(name, value);
  }
}

function validBreakType(
  value: string | undefined,
): WorkDocumentSectionBreakType {
  if (
    value === 'continuous' ||
    value === 'evenPage' ||
    value === 'oddPage' ||
    value === 'nextColumn'
  )
    return value;
  return 'nextPage';
}

function validPageNumber(value: number | null | undefined): number | undefined {
  return Number.isFinite(value) && Number(value) > 0
    ? Math.min(9999, Math.round(Number(value)))
    : undefined;
}

function finiteNumber(
  value: number | null | undefined,
  fallback: number,
): number {
  return Number.isFinite(value) ? Number(value) : fallback;
}

function numberValue(value: string | undefined): number | undefined {
  if (!value?.trim()) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function documentGridFromNodeAttributes(
  attributes: Partial<DocumentSectionNodeAttributes>,
  base: WorkDocumentSectionLayout,
): WorkDocumentGrid | undefined {
  if (
    attributes.documentGridType === undefined &&
    attributes.documentGridLinePitch === undefined
  ) {
    return normalizedDocumentGrid(base.documentGrid);
  }
  return normalizedDocumentGrid({
    type: attributes.documentGridType as WorkDocumentGridType,
    linePitch: Number(attributes.documentGridLinePitch),
  });
}

function normalizedDocumentGrid(
  value: WorkDocumentGrid | undefined,
): WorkDocumentGrid | undefined {
  if (!value || !validDocumentGridType(value.type)) return undefined;
  const linePitch = Number(value.linePitch);
  if (!Number.isFinite(linePitch) || linePitch <= 0) return undefined;
  return {
    type: value.type,
    linePitch: Math.min(720, Number(linePitch.toFixed(2))),
  };
}

function validDocumentGridType(value: unknown): value is WorkDocumentGridType {
  return (
    value === 'default' ||
    value === 'lines' ||
    value === 'linesAndChars' ||
    value === 'snapToChars'
  );
}

function lnNumTypeNodeFields(value: WorkDocumentLnNumType | undefined): {
  lnNumCountBy: number | null;
  lnNumStart: number | null;
  lnNumDistance: number | null;
  lnNumRestart: WorkDocumentLnNumRestart | '';
} {
  const normalized = normalizedLnNumType(value);
  return {
    lnNumCountBy: normalized?.countBy ?? null,
    lnNumStart: normalized?.start ?? null,
    lnNumDistance: normalized?.distance ?? null,
    lnNumRestart: normalized?.restart ?? '',
  };
}

function pgNumTypeNodeFields(
  value: WorkDocumentPgNumType | undefined,
  pageNumberStart?: number,
): {
  pgNumFmt: WorkDocumentPgNumFmt | '';
  pgNumStart: number | null;
} {
  const normalized = normalizedPgNumType(
    value ??
      (pageNumberStart !== undefined ? { start: pageNumberStart } : undefined),
  );
  return {
    pgNumFmt: normalized?.fmt ?? '',
    pgNumStart: normalized?.start ?? null,
  };
}

function lnNumTypeFromNodeAttributes(
  attributes: Partial<DocumentSectionNodeAttributes>,
  base: WorkDocumentSectionLayout,
): WorkDocumentLnNumType | undefined {
  if (
    attributes.lnNumCountBy === undefined &&
    attributes.lnNumStart === undefined &&
    attributes.lnNumDistance === undefined &&
    attributes.lnNumRestart === undefined
  ) {
    return normalizedLnNumType(base.lnNumType);
  }
  return normalizedLnNumType({
    ...(attributes.lnNumCountBy != null
      ? { countBy: Number(attributes.lnNumCountBy) }
      : {}),
    ...(attributes.lnNumStart != null
      ? { start: Number(attributes.lnNumStart) }
      : {}),
    ...(attributes.lnNumDistance != null
      ? { distance: Number(attributes.lnNumDistance) }
      : {}),
    ...(attributes.lnNumRestart
      ? { restart: attributes.lnNumRestart as WorkDocumentLnNumRestart }
      : {}),
  });
}

function pgNumTypeFromNodeAttributes(
  attributes: Partial<DocumentSectionNodeAttributes>,
  base: WorkDocumentSectionLayout,
): WorkDocumentPgNumType | undefined {
  if (
    attributes.pgNumFmt === undefined &&
    attributes.pgNumStart === undefined
  ) {
    return normalizedPgNumType(
      base.pgNumType ??
        (base.pageNumberStart !== undefined
          ? { start: base.pageNumberStart }
          : undefined),
    );
  }
  return normalizedPgNumType({
    ...(attributes.pgNumFmt
      ? { fmt: attributes.pgNumFmt as WorkDocumentPgNumFmt }
      : {}),
    ...(attributes.pgNumStart != null
      ? { start: Number(attributes.pgNumStart) }
      : {}),
  });
}

function normalizedPgNumType(
  value: WorkDocumentPgNumType | undefined,
): WorkDocumentPgNumType | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const next: WorkDocumentPgNumType = {};
  if (value.fmt !== undefined) {
    if (
      value.fmt !== 'decimal' &&
      value.fmt !== 'upperRoman' &&
      value.fmt !== 'lowerRoman' &&
      value.fmt !== 'upperLetter' &&
      value.fmt !== 'lowerLetter'
    ) {
      return undefined;
    }
    next.fmt = value.fmt;
  }
  if (value.start !== undefined) {
    const start = Number(value.start);
    if (!Number.isInteger(start) || start < 0 || start > 32_767) {
      return undefined;
    }
    next.start = start;
  }
  return next.fmt !== undefined || next.start !== undefined ? next : undefined;
}

function normalizedLnNumType(
  value: WorkDocumentLnNumType | undefined,
): WorkDocumentLnNumType | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const next: WorkDocumentLnNumType = {};
  if (value.countBy !== undefined) {
    const countBy = Number(value.countBy);
    if (!Number.isInteger(countBy) || countBy < 1 || countBy > 32_767) {
      return undefined;
    }
    next.countBy = countBy;
  }
  if (value.start !== undefined) {
    const start = Number(value.start);
    if (!Number.isInteger(start) || start < 0 || start > 32_767) {
      return undefined;
    }
    next.start = start;
  }
  if (value.distance !== undefined) {
    const distance = Number(value.distance);
    if (!Number.isInteger(distance) || distance < 1 || distance > 31_680) {
      return undefined;
    }
    next.distance = distance;
  }
  if (value.restart !== undefined) {
    if (
      value.restart !== 'newPage' &&
      value.restart !== 'newSection' &&
      value.restart !== 'continuous'
    ) {
      return undefined;
    }
    next.restart = value.restart;
  }
  return next.countBy !== undefined ||
    next.start !== undefined ||
    next.distance !== undefined ||
    next.restart !== undefined
    ? next
    : undefined;
}

function sectionFormattingChangeFromAttributes(
  attributes: Partial<DocumentSectionNodeAttributes>,
): WorkDocumentSectionFormattingChange | null {
  if (attributes.sectionChangeKind !== 'section-formatting') return null;
  const id =
    typeof attributes.sectionChangeId === 'string'
      ? attributes.sectionChangeId.trim()
      : '';
  const author =
    typeof attributes.sectionChangeAuthor === 'string'
      ? attributes.sectionChangeAuthor.trim()
      : '';
  const date =
    typeof attributes.sectionChangeDate === 'string'
      ? attributes.sectionChangeDate
      : '';
  const before =
    typeof attributes.sectionChangeBefore === 'string'
      ? attributes.sectionChangeBefore
      : '';
  if (!id || !author || !before) return null;
  return {
    kind: 'section-formatting',
    id,
    author,
    date,
    before,
  };
}
