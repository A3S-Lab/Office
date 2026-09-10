import { DEFAULT_DOCUMENT_COLUMNS } from './work-document-columns';
import { importDocxColumns } from './work-docx-column-import';
import {
  type DocumentSectionColumnsSnapshot,
  serializeDocumentSectionFormatting,
} from './work-document-section-format-changes';
import {
  type WorkDocumentPageMarginKey,
  type WorkDocumentPageMargins,
  normalizeDocumentPageMargins,
} from './work-document-page-margins';
import {
  MAX_WORD_PAGE_TWIPS,
  MAX_WORD_PAPER_CODE,
  MAX_WORD_PAPER_SOURCE_CODE,
  normalizeDocumentPageGeometry,
  normalizeDocumentPaperSource,
  type WorkDocumentPageGeometry,
  type WorkDocumentPaperSource,
} from './work-document-page-size';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  parseBoundedDocxInteger,
  parseDocxTwipsMeasure,
} from './work-docx-twips';
import { attribute, directChildren } from './work-ooxml-package';
import type {
  WorkDocumentGrid,
  WorkDocumentGridType,
  WorkDocumentLnNumRestart,
  WorkDocumentLnNumType,
  WorkDocumentPgNumFmt,
  WorkDocumentPgNumType,
  WorkDocumentSectionTextDirection,
  WorkDocumentSectionVerticalAlign,
} from './work-types';

const MAX_REVISION_DATE_LENGTH = 64;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);
const SUPPORTED_PRIOR_CHILDREN = new Set([
  'pgSz',
  'pgMar',
  'paperSrc',
  'cols',
  'titlePg',
  'rtlGutter',
  'docGrid',
  'lnNumType',
  'pgNumType',
  'formProt',
  'vAlign',
  'noEndnote',
  'textDirection',
  'bidi',
]);
const DOC_GRID_ATTRIBUTE_SET = new Set(['type', 'linePitch']);
const FORM_PROT_ATTRIBUTE_SET = new Set(['val']);
const NO_ENDNOTE_ATTRIBUTE_SET = new Set(['val']);
const BIDI_ATTRIBUTE_SET = new Set(['val']);
const V_ALIGN_ATTRIBUTE_SET = new Set(['val']);
const V_ALIGN_VALUES = new Set(['top', 'center', 'both', 'bottom']);
const TEXT_DIRECTION_ATTRIBUTE_SET = new Set(['val']);
const TEXT_DIRECTION_VALUES = new Set([
  'lrTb',
  'tbRl',
  'btLr',
  'lrTbV',
  'tbRlV',
  'tbLrV',
]);
const LN_NUM_TYPE_ATTRIBUTE_SET = new Set([
  'countBy',
  'start',
  'distance',
  'restart',
]);
const LN_NUM_RESTARTS = new Set(['newPage', 'newSection', 'continuous']);
const PG_NUM_TYPE_ATTRIBUTE_SET = new Set(['fmt', 'start']);
const PG_NUM_FMTS = new Set([
  'decimal',
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
]);
const DOC_GRID_TYPES = new Set([
  'default',
  'lines',
  'linesAndChars',
  'snapToChars',
]);
const MAX_DOC_GRID_LINE_PITCH_TWIPS = 14_400;
const PAGE_SIZE_ATTRIBUTE_SET = new Set(['w', 'h', 'orient', 'code']);
const PAPER_SOURCE_ATTRIBUTE_SET = new Set(['first', 'other']);
const COLUMNS_ATTRIBUTE_SET = new Set(['num', 'space', 'sep', 'equalWidth']);
const COLUMN_CHILD_ATTRIBUTE_SET = new Set(['w', 'space']);
const TWIPS_PER_MILLIMETER = 1440 / 25.4;
const PAGE_MARGIN_KEYS = [
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
] as const satisfies readonly WorkDocumentPageMarginKey[];
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const MAX_UNSIGNED_WORD_TWIPS = 31_680;
const MIN_SIGNED_WORD_TWIPS = -2_147_483_648;
const MAX_SIGNED_WORD_TWIPS = 2_147_483_647;

export interface SupportedDocxSectionFormattingChange {
  id: string;
  author: string;
  date: string;
  before: string;
}

/**
 * Relationship-free `w:sectPrChange` whose prior snapshot contains only
 * orientation-only or complete `w:pgSz` (w/h with optional orient/code), a
 * complete seven-edge `w:pgMar`, `w:paperSrc`, equal-width or unequal-width
 * `w:cols`, and/or `w:titlePg`, and/or `w:rtlGutter`, and/or bounded
 * relationship-free `w:docGrid`, and/or bounded relationship-free `w:lnNumType`,
 * and/or bounded relationship-free `w:pgNumType` (`fmt`/`start` only), and/or
 * relationship-free empty/onOff `w:formProt`, and/or relationship-free
 * `w:vAlign` with required known `w:val` (`top`/`center`/`both`/`bottom`),
 * and/or relationship-free empty/onOff `w:noEndnote`, and/or relationship-free
 * `w:textDirection` with required known `w:val`
 * (`lrTb`/`tbRl`/`btLr`/`lrTbV`/`tbRlV`/`tbLrV`), and/or relationship-free
 * empty/onOff `w:bidi`.
 * Broader section property sets stay on the opaque OMML path.
 */
export function isSupportedDocxSectionFormattingChange(
  change: Element,
): boolean {
  return supportedSectionFormattingChange(change) !== null;
}

export function supportedDocxSectionFormattingChangeFromProperties(
  sectionProperties: Element | null | undefined,
): SupportedDocxSectionFormattingChange | null {
  if (!sectionProperties) return null;
  const changes = directChildren(sectionProperties, 'sectPrChange').filter(
    (element) => element.namespaceURI === sectionProperties.namespaceURI,
  );
  if (changes.length !== 1) return null;
  return supportedSectionFormattingChange(changes[0] ?? null);
}

function supportedSectionFormattingChange(
  change: Element | null,
): SupportedDocxSectionFormattingChange | null {
  if (
    !change ||
    change.localName !== 'sectPrChange' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  if (
    hasRelationshipBindings(change) ||
    hasUnsupportedRevisionAttributes(change)
  ) {
    return null;
  }
  const id = wordAttribute(change, 'id')?.trim() ?? '';
  const author = wordAttribute(change, 'author')?.trim() ?? '';
  const rawDate = wordAttribute(change, 'date');
  if (
    !/^\+?\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (rawDate !== null && rawDate.length > MAX_REVISION_DATE_LENGTH) ||
    (rawDate !== null && !Number.isFinite(Date.parse(rawDate)))
  ) {
    return null;
  }
  const priors = Array.from(change.children).filter(
    (child) =>
      child.localName === 'sectPr' &&
      child.namespaceURI === change.namespaceURI,
  );
  if (priors.length !== 1) return null;
  const prior = priors[0];
  if (!prior || hasRelationshipBindings(prior)) return null;
  const children = Array.from(prior.children);
  if (!children.length || children.length > 13) return null;
  if (
    children.some((child) => !isSupportedSectionFormattingPriorChild(child))
  ) {
    return null;
  }
  const localNames = children.map((child) => child.localName);
  if (new Set(localNames).size !== localNames.length) return null;

  let orientation: 'portrait' | 'landscape' | undefined;
  let pageGeometry: WorkDocumentPageGeometry | undefined;
  let pageMargins: WorkDocumentPageMargins | undefined;
  let paperSource: WorkDocumentPaperSource | undefined;
  let columns: DocumentSectionColumnsSnapshot | undefined;
  let differentFirstPage: boolean | undefined;
  let rtlGutter: boolean | undefined;
  let documentGrid: WorkDocumentGrid | undefined;
  let lnNumType: WorkDocumentLnNumType | undefined;
  let pgNumType: WorkDocumentPgNumType | undefined;
  let formProt: boolean | undefined;
  let verticalAlign: WorkDocumentSectionVerticalAlign | undefined;
  let noEndnote: boolean | undefined;
  let textDirection: WorkDocumentSectionTextDirection | undefined;
  let bidi: boolean | undefined;
  for (const child of children) {
    if (child.localName === 'pgSz') {
      const value = importedPageSize(child);
      if (!value) return null;
      if (value.orientation) orientation = value.orientation;
      if (value.pageGeometry) pageGeometry = value.pageGeometry;
      continue;
    }
    if (child.localName === 'pgMar') {
      const value = importedPageMargins(child);
      if (!value) return null;
      pageMargins = value;
      continue;
    }
    if (child.localName === 'paperSrc') {
      const value = importedPaperSource(child);
      if (!value) return null;
      paperSource = value;
      continue;
    }
    if (child.localName === 'cols') {
      const value = importedSectionFormattingColumns(child);
      if (!value) return null;
      columns = value;
      continue;
    }
    if (child.localName === 'titlePg') {
      differentFirstPage = onOffValue(child);
      continue;
    }
    if (child.localName === 'rtlGutter') {
      rtlGutter = onOffValue(child);
      continue;
    }
    if (child.localName === 'docGrid') {
      const value = importedDocumentGrid(child);
      if (!value) return null;
      documentGrid = value;
      continue;
    }
    if (child.localName === 'lnNumType') {
      const value = importedLnNumType(child);
      if (!value) return null;
      lnNumType = value;
      continue;
    }
    if (child.localName === 'pgNumType') {
      const value = importedPgNumType(child);
      if (!value) return null;
      pgNumType = value;
      continue;
    }
    if (child.localName === 'formProt') {
      const value = importedFormProt(child);
      if (value === null) return null;
      formProt = value;
      continue;
    }
    if (child.localName === 'vAlign') {
      const value = importedVerticalAlign(child);
      if (value === null) return null;
      verticalAlign = value;
      continue;
    }
    if (child.localName === 'noEndnote') {
      const value = importedNoEndnote(child);
      if (value === null) return null;
      noEndnote = value;
      continue;
    }
    if (child.localName === 'textDirection') {
      const value = importedTextDirection(child);
      if (value === null) return null;
      textDirection = value;
      continue;
    }
    if (child.localName === 'bidi') {
      const value = importedBidi(child);
      if (value === null) return null;
      bidi = value;
    }
  }
  const before = serializeDocumentSectionFormatting({
    ...(orientation ? { orientation } : {}),
    ...(pageGeometry ? { pageGeometry } : {}),
    ...(pageMargins ? { pageMargins } : {}),
    ...(paperSource ? { paperSource } : {}),
    ...(columns ? { columns } : {}),
    ...(differentFirstPage !== undefined ? { differentFirstPage } : {}),
    ...(rtlGutter !== undefined ? { rtlGutter } : {}),
    ...(documentGrid ? { documentGrid } : {}),
    ...(lnNumType ? { lnNumType } : {}),
    ...(pgNumType ? { pgNumType } : {}),
    ...(formProt !== undefined ? { formProt } : {}),
    ...(verticalAlign !== undefined ? { verticalAlign } : {}),
    ...(noEndnote !== undefined ? { noEndnote } : {}),
    ...(textDirection !== undefined ? { textDirection } : {}),
    ...(bidi !== undefined ? { bidi } : {}),
  });
  return {
    id: `docx-section-format-change-${id}`,
    author,
    date: normalizeRevisionDate(rawDate),
    before,
  };
}

function importedPageSize(element: Element): {
  orientation?: 'portrait' | 'landscape';
  pageGeometry?: WorkDocumentPageGeometry;
} | null {
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  if (!attributes.length) return null;
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !PAGE_SIZE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;

  if (names.size === 1 && names.has('orient')) {
    const value = attributes[0]?.value.trim();
    if (value === 'landscape') return { orientation: 'landscape' };
    if (value === 'portrait') return { orientation: 'portrait' };
    return null;
  }

  if (!names.has('w') || !names.has('h')) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  const width = parseDocxTwipsMeasure(byName.get('w') ?? '', {
    minimum: 1,
    maximum: MAX_WORD_PAGE_TWIPS,
    signed: false,
    strict: false,
  });
  const height = parseDocxTwipsMeasure(byName.get('h') ?? '', {
    minimum: 1,
    maximum: MAX_WORD_PAGE_TWIPS,
    signed: false,
    strict: false,
  });
  const orientation = byName.get('orient');
  const codeSource = byName.get('code');
  const code =
    codeSource === undefined
      ? undefined
      : parseBoundedDocxInteger(codeSource, {
          minimum: 0,
          maximum: MAX_WORD_PAPER_CODE,
        });
  if (
    width === null ||
    height === null ||
    (orientation !== undefined &&
      orientation !== 'portrait' &&
      orientation !== 'landscape') ||
    (codeSource !== undefined && code === null)
  ) {
    return null;
  }
  const pageGeometry = normalizeDocumentPageGeometry({
    width,
    height,
    ...(orientation !== undefined ? { orientation } : {}),
    ...(code !== undefined ? { code } : {}),
  });
  return pageGeometry ? { pageGeometry } : null;
}

function isSupportedSectionFormattingPriorChild(child: Element): boolean {
  if (child.namespaceURI !== child.parentElement?.namespaceURI) return false;
  if (!SUPPORTED_PRIOR_CHILDREN.has(child.localName)) return false;
  if (child.localName === 'cols') {
    return Array.from(child.children).every(
      (column) =>
        column.localName === 'col' &&
        column.namespaceURI === child.namespaceURI &&
        isSupportedColumnChild(column),
    );
  }
  return child.children.length === 0;
}

function isSupportedColumnChild(column: Element): boolean {
  const attributes = Array.from(column.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(column, candidate) === column.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !COLUMN_CHILD_ATTRIBUTE_SET.has(name))) {
    return false;
  }
  if (names.size !== attributes.length || !names.has('w')) return false;
  const width = parseBoundedDocxInteger(
    attributes
      .find((candidate) => xmlAttributeLocalName(candidate) === 'w')
      ?.value.trim() ?? '',
    { minimum: 1, maximum: MAX_UNSIGNED_WORD_TWIPS },
  );
  if (width === null) return false;
  if (names.has('space')) {
    const space = parseDocxTwipsMeasure(
      attributes
        .find((candidate) => xmlAttributeLocalName(candidate) === 'space')
        ?.value.trim() ?? '',
      {
        minimum: 0,
        maximum: MAX_UNSIGNED_WORD_TWIPS,
        signed: false,
        strict: false,
      },
    );
    if (space === null) return false;
  }
  return column.children.length === 0;
}

function importedSectionFormattingColumns(
  element: Element,
): DocumentSectionColumnsSnapshot | null {
  const columnElements = Array.from(element.children).filter(
    (child) =>
      child.localName === 'col' && child.namespaceURI === element.namespaceURI,
  );
  if (columnElements.length > 0) {
    if (
      columnElements.length !== element.children.length ||
      columnElements.length < 2 ||
      columnElements.length > 6 ||
      !columnElements.every(isSupportedColumnChild)
    ) {
      return null;
    }
    const attributes = Array.from(element.attributes).filter(
      (candidate) =>
        xmlAttributeNamespace(element, candidate) === element.namespaceURI,
    );
    const names = new Set(
      attributes.map((candidate) => xmlAttributeLocalName(candidate)),
    );
    if ([...names].some((name) => !COLUMNS_ATTRIBUTE_SET.has(name))) {
      return null;
    }
    if (names.size !== attributes.length) return null;
    if (names.has('equalWidth')) {
      const equalWidth = attributes
        .find((candidate) => xmlAttributeLocalName(candidate) === 'equalWidth')
        ?.value.trim()
        .toLowerCase();
      if (equalWidth === '1' || equalWidth === 'true' || equalWidth === 'on') {
        return null;
      }
    }
    const imported = importDocxColumns(element, DEFAULT_DOCUMENT_COLUMNS);
    if (!imported.custom) return null;
    const spacing =
      attribute(element, 'space') !== null
        ? imported.spacing
        : (imported.custom[0]?.spacing ?? imported.spacing);
    return {
      count: imported.count,
      spacing,
      separator: imported.separator,
      custom: imported.custom.map((column) => ({
        widthPercent: column.widthPercent,
        spacing: column.spacing,
      })),
    };
  }

  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !COLUMNS_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  if (byName.has('equalWidth')) {
    const equalWidth = byName.get('equalWidth')?.toLowerCase();
    if (equalWidth !== '1' && equalWidth !== 'true' && equalWidth !== 'on') {
      return null;
    }
  }
  const countSource = byName.get('num');
  const count =
    countSource === undefined
      ? 1
      : parseBoundedDocxInteger(countSource, { minimum: 1, maximum: 6 });
  if (count === null) return null;
  const spaceSource = byName.get('space');
  let spacing = 12;
  if (spaceSource !== undefined) {
    const twips = parseDocxTwipsMeasure(spaceSource, {
      minimum: 0,
      maximum: MAX_UNSIGNED_WORD_TWIPS,
      signed: false,
      strict: false,
    });
    if (twips === null) return null;
    spacing = Math.round((twips / TWIPS_PER_MILLIMETER) * 10) / 10;
  }
  const sepSource = byName.get('sep')?.toLowerCase();
  const separator =
    sepSource === '1' || sepSource === 'true' || sepSource === 'on';
  return { count, spacing, separator };
}

function importedPaperSource(element: Element): WorkDocumentPaperSource | null {
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  if (!attributes.length) return null;
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !PAPER_SOURCE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const source: WorkDocumentPaperSource = {};
  for (const attribute of attributes) {
    const name = xmlAttributeLocalName(attribute);
    if (name !== 'first' && name !== 'other') return null;
    const parsed = parseBoundedDocxInteger(attribute.value.trim(), {
      minimum: 0,
      maximum: MAX_WORD_PAPER_SOURCE_CODE,
    });
    if (parsed === null) return null;
    source[name] = parsed;
  }
  const normalized = normalizeDocumentPaperSource(source);
  if (
    !normalized ||
    (normalized.first === undefined && normalized.other === undefined)
  ) {
    return null;
  }
  return normalized;
}

function importedPageMargins(element: Element): WorkDocumentPageMargins | null {
  const values: Partial<Record<WorkDocumentPageMarginKey, number>> = {};
  const seen = new Set<string>();
  for (const item of Array.from(element.attributes)) {
    const namespace =
      item.namespaceURI || xmlAttributeNamespace(element, item) || '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return null;
    if (namespace && namespace !== element.namespaceURI) return null;
    const name = xmlAttributeLocalName(item);
    if (!(PAGE_MARGIN_KEYS as readonly string[]).includes(name)) return null;
    if (seen.has(name)) return null;
    seen.add(name);
    const signed = name === 'top' || name === 'bottom';
    const value = parseDocxTwipsMeasure(item.value.trim(), {
      minimum: signed ? MIN_SIGNED_WORD_TWIPS : 0,
      maximum: signed ? MAX_SIGNED_WORD_TWIPS : MAX_UNSIGNED_WORD_TWIPS,
      signed,
      strict: false,
    });
    if (value === null) return null;
    values[name as WorkDocumentPageMarginKey] = value;
  }
  if (seen.size !== PAGE_MARGIN_KEYS.length) return null;
  return normalizeDocumentPageMargins(values);
}

function importedPgNumType(element: Element): WorkDocumentPgNumType | null {
  if (Array.from(element.children).length) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !PG_NUM_TYPE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  const next: WorkDocumentPgNumType = {};
  if (byName.has('fmt')) {
    const fmt = byName.get('fmt') ?? '';
    if (!PG_NUM_FMTS.has(fmt)) return null;
    next.fmt = fmt as WorkDocumentPgNumFmt;
  }
  if (byName.has('start')) {
    const start = parseBoundedDocxInteger(byName.get('start') ?? '', {
      minimum: 0,
      maximum: 32_767,
    });
    if (start === null) return null;
    next.start = start;
  }
  if (next.fmt === undefined && next.start === undefined) return null;
  return next;
}

function importedFormProt(element: Element): boolean | null {
  if (element.children.length > 0) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !FORM_PROT_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  return onOffValue(element);
}

function importedNoEndnote(element: Element): boolean | null {
  if (element.children.length > 0) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !NO_ENDNOTE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  return onOffValue(element);
}

function importedBidi(element: Element): boolean | null {
  if (element.children.length > 0) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !BIDI_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  return onOffValue(element);
}

function importedVerticalAlign(
  element: Element,
): WorkDocumentSectionVerticalAlign | null {
  if (element.children.length > 0) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !V_ALIGN_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length || !names.has('val')) return null;
  const value = attributes
    .find((candidate) => xmlAttributeLocalName(candidate) === 'val')
    ?.value.trim();
  if (!value || !V_ALIGN_VALUES.has(value)) return null;
  return value as WorkDocumentSectionVerticalAlign;
}

function importedTextDirection(
  element: Element,
): WorkDocumentSectionTextDirection | null {
  if (element.children.length > 0) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !TEXT_DIRECTION_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length || !names.has('val')) return null;
  const value = attributes
    .find((candidate) => xmlAttributeLocalName(candidate) === 'val')
    ?.value.trim();
  if (!value || !TEXT_DIRECTION_VALUES.has(value)) return null;
  return value as WorkDocumentSectionTextDirection;
}

function importedLnNumType(element: Element): WorkDocumentLnNumType | null {
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !LN_NUM_TYPE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  const next: WorkDocumentLnNumType = {};
  if (byName.has('countBy')) {
    const countBy = parseBoundedDocxInteger(byName.get('countBy') ?? '', {
      minimum: 1,
      maximum: 32_767,
    });
    if (countBy === null) return null;
    next.countBy = countBy;
  }
  if (byName.has('start')) {
    const start = parseBoundedDocxInteger(byName.get('start') ?? '', {
      minimum: 0,
      maximum: 32_767,
    });
    if (start === null) return null;
    next.start = start;
  }
  if (byName.has('distance')) {
    const distance = parseBoundedDocxInteger(byName.get('distance') ?? '', {
      minimum: 1,
      maximum: 31_680,
    });
    if (distance === null) return null;
    next.distance = distance;
  }
  if (byName.has('restart')) {
    const restart = byName.get('restart') ?? '';
    if (!LN_NUM_RESTARTS.has(restart)) return null;
    next.restart = restart as WorkDocumentLnNumRestart;
  }
  if (
    next.countBy === undefined &&
    next.start === undefined &&
    next.distance === undefined &&
    next.restart === undefined
  ) {
    next.countBy = 1;
  }
  return next;
}

function importedDocumentGrid(element: Element): WorkDocumentGrid | null {
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !DOC_GRID_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  const typeSource = byName.get('type');
  if (typeSource !== undefined && !DOC_GRID_TYPES.has(typeSource)) {
    return null;
  }
  const type: WorkDocumentGridType =
    typeSource === undefined ? 'default' : (typeSource as WorkDocumentGridType);
  const linePitchSource = byName.get('linePitch');
  const linePitchTwips =
    linePitchSource === undefined
      ? 360
      : parseBoundedDocxInteger(linePitchSource, {
          minimum: 1,
          maximum: MAX_DOC_GRID_LINE_PITCH_TWIPS,
        });
  if (linePitchTwips === null) return null;
  return {
    type,
    linePitch: Number((linePitchTwips / 20).toFixed(2)),
  };
}

function onOffValue(element: Element): boolean {
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === 'val' &&
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  if (matches.length > 1) return true;
  const value = matches[0]?.value ?? null;
  if (value === null || value === '') return true;
  const normalized = value.trim().toLowerCase();
  return (
    normalized === '1' ||
    normalized === 'true' ||
    normalized === 'on' ||
    normalized === 'yes'
  );
}

function wordAttribute(element: Element, localName: string): string | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === localName &&
      xmlAttributeNamespace(element, candidate) === namespace,
  );
  return matches.length === 1 ? (matches[0]?.value ?? null) : null;
}

function hasUnsupportedRevisionAttributes(element: Element): boolean {
  const namespace = element.namespaceURI;
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !REVISION_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
  );
}

function hasRelationshipBindings(element: Element): boolean {
  for (const attributeNode of Array.from(element.attributes)) {
    const namespace =
      attributeNode.namespaceURI ||
      xmlAttributeNamespace(element, attributeNode) ||
      '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return true;
  }
  for (const descendant of Array.from(element.querySelectorAll('*'))) {
    for (const attributeNode of Array.from(descendant.attributes)) {
      const namespace =
        attributeNode.namespaceURI ||
        xmlAttributeNamespace(descendant, attributeNode) ||
        '';
      if (RELATIONSHIP_NAMESPACES.has(namespace)) return true;
    }
  }
  return false;
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}
