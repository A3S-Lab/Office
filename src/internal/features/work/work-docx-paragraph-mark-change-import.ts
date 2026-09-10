import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { readDocxImageTransform } from './work-docx-image-transform';
import {
  EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  type DocxExternalHyperlinkTargets,
} from './work-docx-note-comment-hyperlink-relationships';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, directChildren } from './work-ooxml-package';

export type { DocxExternalHyperlinkTargets } from './work-docx-note-comment-hyperlink-relationships';
export {
  createDocxExternalHyperlinkTargets,
  EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
} from './work-docx-note-comment-hyperlink-relationships';

const TRANSITIONAL_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELATIONSHIP_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const IMAGE_RELATIONSHIP_TYPES = new Set([
  `${TRANSITIONAL_RELATIONSHIP_NAMESPACE}/image`,
  `${STRICT_RELATIONSHIP_NAMESPACE}/image`,
]);
const RELATIONSHIP_ID_PATTERN = /^[A-Za-z_][A-Za-z0-9_.-]{0,254}$/;

/** Relationship Ids that resolve to package image parts. */
export type DocxImageEmbedTargets = ReadonlySet<string>;

export const EMPTY_DOCX_IMAGE_EMBED_TARGETS: DocxImageEmbedTargets = new Set();

/**
 * Builds the admission lookup for relationship-bound inline pictures.
 * Only image relationship types without TargetMode (or with an empty one)
 * are retained; hyperlinks and external targets stay absent.
 */
export function createDocxImageEmbedTargets(
  relationships: Iterable<{
    id: string;
    type: string;
    targetMode?: string;
  }>,
): DocxImageEmbedTargets {
  const embeds = new Set<string>();
  for (const relationship of relationships) {
    if (
      !RELATIONSHIP_ID_PATTERN.test(relationship.id) ||
      !IMAGE_RELATIONSHIP_TYPES.has(relationship.type) ||
      (relationship.targetMode ?? '').trim() !== ''
    ) {
      continue;
    }
    embeds.add(relationship.id);
  }
  return embeds;
}

export type DocxParagraphMarkChangeKind = 'insertion' | 'deletion';

export interface ImportedDocxParagraphMarkChangeMarker {
  marker: string;
  id: string;
  kind: DocxParagraphMarkChangeKind;
  author: string;
  date: string;
}

export interface ImportedDocxParagraphMarkChangeMarkers {
  paragraphs: ImportedDocxParagraphMarkChangeMarker[];
}

interface SupportedDocxParagraphMarkChange
  extends Omit<ImportedDocxParagraphMarkChangeMarker, 'marker'> {
  element: Element;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const WORDPROCESSING_DRAWING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
]);
const DRAWINGML_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/main',
]);
const PICTURE_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://purl.oclc.org/ooxml/drawingml/picture',
]);
const MAX_PARAGRAPH_MARK_CHANGES = 65_536;
const MAX_REVISION_DATE_LENGTH = 64;
const PARAGRAPH_MARK_CHANGE_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_MARK_CHANGE_\d+__/g;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);

/**
 * Finds the bounded Word paragraph-mark insertion/deletion shape. Word and
 * WPS store the revision on the paragraph mark's run-properties node while
 * keeping the paragraph's text revisions in ordinary w:ins/w:del wrappers.
 */
export function markDocxParagraphMarkChanges(
  document: Document,
  externalHyperlinks: DocxExternalHyperlinkTargets = EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  imageEmbeds: DocxImageEmbedTargets = EMPTY_DOCX_IMAGE_EMBED_TARGETS,
): ImportedDocxParagraphMarkChangeMarkers {
  const paragraphs: ImportedDocxParagraphMarkChangeMarker[] = [];
  const changeIds = new Set<string>();
  for (const paragraph of descendants(document, 'p')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(paragraph.namespaceURI ?? '')) {
      continue;
    }
    const change = supportedParagraphMarkChange(
      paragraph,
      externalHyperlinks,
      imageEmbeds,
    );
    if (!change) continue;
    if (paragraphs.length >= MAX_PARAGRAPH_MARK_CHANGES) {
      throw new Error('Document exceeds the paragraph-mark revision limit.');
    }
    const marker = `__A3S_WORK_PARAGRAPH_MARK_CHANGE_${paragraphs.length + 1}__`;
    insertParagraphMarker(paragraph, marker);
    // Mammoth applies w:del on a paragraph mark by joining the following
    // paragraph.  The bounded Work model reviews the complete text-only
    // paragraph atomically, so remove the native boundary flag after capturing
    // it and let the marker retain the original block boundary through HTML.
    change.element.remove();
    paragraphs.push({
      marker,
      id: uniqueChangeId(change.id, changeIds),
      kind: change.kind,
      author: change.author,
      date: change.date,
    });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphMarkChangeMarkers(
  document: Document,
  markers: ImportedDocxParagraphMarkChangeMarkers,
): void {
  const changes = new Map(
    markers.paragraphs.map((change) => [change.marker, change]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_MARK_CHANGE_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_MARK_CHANGE_MARKER_PATTERN,
      (marker) => {
        const change = changes.get(marker);
        const block = change
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && change) {
          block.dataset.documentBlockChange = 'true';
          block.dataset.blockChangeKind = change.kind;
          block.dataset.blockChangeId = change.id;
          block.dataset.blockChangeAuthor = change.author;
          block.dataset.blockChangeDate = change.date;
        }
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphMarkChangeMarkers(
  markers: ImportedDocxParagraphMarkChangeMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

/** Returns true only for a native, unambiguous paragraph-mark revision. */
export function isSupportedDocxParagraphMarkChange(
  change: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets = EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  imageEmbeds: DocxImageEmbedTargets = EMPTY_DOCX_IMAGE_EMBED_TARGETS,
): boolean {
  const runProperties = change.parentElement;
  const properties = runProperties?.parentElement;
  const paragraph = properties?.parentElement;
  if (
    runProperties?.localName === 'rPr' &&
    properties?.localName === 'pPr' &&
    paragraph?.localName === 'p' &&
    runProperties.namespaceURI === change.namespaceURI &&
    properties.namespaceURI === change.namespaceURI &&
    paragraph.namespaceURI === change.namespaceURI
  ) {
    return (
      supportedParagraphMarkChange(paragraph, externalHyperlinks, imageEmbeds)
        ?.element === change
    );
  }
  return false;
}

export type DocxParagraphBreakMarkKind = 'merge' | 'split';

export interface InspectedDocxParagraphBreakMark {
  author: string;
  date: string;
  id: string;
  kind: DocxParagraphBreakMarkKind;
}

export interface ImportedDocxParagraphBreakChangeMarker {
  marker: string;
  id: string;
  kind: DocxParagraphBreakMarkKind;
  author: string;
  date: string;
}

export interface ImportedDocxParagraphBreakChangeMarkers {
  paragraphs: ImportedDocxParagraphBreakChangeMarker[];
}

const PARAGRAPH_BREAK_CHANGE_MARKER_PATTERN =
  /__A3S_WORK_PARAGRAPH_BREAK_CHANGE_\d+__/g;

/**
 * Mark-only paragraph revisions (no matching body wrap) are paragraph-break
 * merge/split candidates. Eligible neighbors promote them into reviewable
 * Work changes; others stay diagnostics-only.
 */
export function isIsolatedDocxParagraphBreakMarkChange(
  change: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets = EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  imageEmbeds: DocxImageEmbedTargets = EMPTY_DOCX_IMAGE_EMBED_TARGETS,
): boolean {
  return (
    isolatedParagraphBreakMarkChange(
      change,
      externalHyperlinks,
      imageEmbeds,
    ) !== null
  );
}

export function inspectDocxParagraphBreakMarkChanges(
  document: Document,
  externalHyperlinks: DocxExternalHyperlinkTargets = EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  imageEmbeds: DocxImageEmbedTargets = EMPTY_DOCX_IMAGE_EMBED_TARGETS,
): InspectedDocxParagraphBreakMark[] {
  const inspected: InspectedDocxParagraphBreakMark[] = [];
  for (const change of [
    ...descendants(document, 'ins'),
    ...descendants(document, 'del'),
  ]) {
    const breakMark = isolatedParagraphBreakMarkChange(
      change,
      externalHyperlinks,
      imageEmbeds,
    );
    if (breakMark) inspected.push(breakMark);
  }
  return inspected;
}

/**
 * Promotes eligible isolated paragraph-break marks into Work review markers.
 * Mammoth would join deleted marks; remove the native flag after capture so
 * both paragraphs survive for accept/reject.
 */
export function markDocxParagraphBreakChanges(
  document: Document,
  externalHyperlinks: DocxExternalHyperlinkTargets = EMPTY_DOCX_EXTERNAL_HYPERLINK_TARGETS,
  imageEmbeds: DocxImageEmbedTargets = EMPTY_DOCX_IMAGE_EMBED_TARGETS,
): ImportedDocxParagraphBreakChangeMarkers {
  const paragraphs: ImportedDocxParagraphBreakChangeMarker[] = [];
  const changeIds = new Set<string>();
  for (const paragraph of descendants(document, 'p')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(paragraph.namespaceURI ?? '')) {
      continue;
    }
    const change = reviewableParagraphBreakMarkChange(
      paragraph,
      externalHyperlinks,
      imageEmbeds,
    );
    if (!change) continue;
    if (paragraphs.length >= MAX_PARAGRAPH_MARK_CHANGES) {
      throw new Error('Document exceeds the paragraph-break revision limit.');
    }
    const marker = `__A3S_WORK_PARAGRAPH_BREAK_CHANGE_${paragraphs.length + 1}__`;
    insertParagraphMarker(paragraph, marker);
    change.element.remove();
    paragraphs.push({
      marker,
      id: uniqueChangeId(change.id, changeIds),
      kind: change.kind,
      author: change.author,
      date: change.date,
    });
  }
  return { paragraphs };
}

export function applyImportedDocxParagraphBreakChangeMarkers(
  document: Document,
  markers: ImportedDocxParagraphBreakChangeMarkers,
): void {
  const changes = new Map(
    markers.paragraphs.map((change) => [change.marker, change]),
  );
  for (const node of textNodes(document.body)) {
    if (!node.data.includes('__A3S_WORK_PARAGRAPH_BREAK_CHANGE_')) continue;
    node.data = node.data.replace(
      PARAGRAPH_BREAK_CHANGE_MARKER_PATTERN,
      (marker) => {
        const change = changes.get(marker);
        const block = change
          ? closestParagraphBlock(node.parentElement, node)
          : null;
        if (block && change) {
          block.dataset.paragraphBreakChange = 'true';
          block.dataset.paragraphBreakKind = change.kind;
          block.dataset.paragraphBreakId = change.id;
          block.dataset.paragraphBreakAuthor = change.author;
          block.dataset.paragraphBreakDate = change.date;
        }
        return '';
      },
    );
  }
  document.body.normalize();
}

export function hasImportedDocxParagraphBreakChangeMarkers(
  markers: ImportedDocxParagraphBreakChangeMarkers,
): boolean {
  return markers.paragraphs.length > 0;
}

function reviewableParagraphBreakMarkChange(
  paragraph: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): (InspectedDocxParagraphBreakMark & { element: Element }) | null {
  const properties = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (properties.length !== 1) return null;
  const runProperties = directChildren(properties[0] as Element, 'rPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (runProperties.length !== 1) return null;
  const revision = Array.from((runProperties[0] as Element).children).find(
    (element) =>
      (element.localName === 'ins' || element.localName === 'del') &&
      element.namespaceURI === paragraph.namespaceURI,
  );
  if (!revision) return null;
  const inspected = isolatedParagraphBreakMarkChange(
    revision,
    externalHyperlinks,
    imageEmbeds,
  );
  if (!inspected) return null;
  if (
    !paragraphBreakNeighborIsEligible(
      paragraph,
      inspected.kind,
      externalHyperlinks,
      imageEmbeds,
    )
  ) {
    return null;
  }
  return { ...inspected, element: revision };
}

function paragraphBreakNeighborIsEligible(
  paragraph: Element,
  kind: DocxParagraphBreakMarkKind,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): boolean {
  const sibling =
    kind === 'merge'
      ? nextWordParagraphSibling(paragraph)
      : previousWordParagraphSibling(paragraph);
  if (!sibling) return false;
  const properties = directChildren(sibling, 'pPr').find(
    (element) => element.namespaceURI === sibling.namespaceURI,
  );
  return paragraphBodyIsUntrackedTextOnly(
    sibling,
    properties,
    externalHyperlinks,
    imageEmbeds,
  );
}

function nextWordParagraphSibling(paragraph: Element): Element | null {
  let sibling = paragraph.nextElementSibling;
  while (sibling) {
    if (
      sibling.localName === 'p' &&
      sibling.namespaceURI === paragraph.namespaceURI
    ) {
      return sibling;
    }
    if (
      sibling.localName === 'tbl' ||
      sibling.localName === 'sectPr' ||
      sibling.localName === 'sdt'
    ) {
      return null;
    }
    sibling = sibling.nextElementSibling;
  }
  return null;
}

function previousWordParagraphSibling(paragraph: Element): Element | null {
  let sibling = paragraph.previousElementSibling;
  while (sibling) {
    if (
      sibling.localName === 'p' &&
      sibling.namespaceURI === paragraph.namespaceURI
    ) {
      return sibling;
    }
    if (
      sibling.localName === 'tbl' ||
      sibling.localName === 'sectPr' ||
      sibling.localName === 'sdt'
    ) {
      return null;
    }
    sibling = sibling.previousElementSibling;
  }
  return null;
}

function isolatedParagraphBreakMarkChange(
  change: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): InspectedDocxParagraphBreakMark | null {
  const runProperties = change.parentElement;
  const properties = runProperties?.parentElement;
  const paragraph = properties?.parentElement;
  if (
    runProperties?.localName !== 'rPr' ||
    properties?.localName !== 'pPr' ||
    paragraph?.localName !== 'p' ||
    runProperties.namespaceURI !== change.namespaceURI ||
    properties.namespaceURI !== change.namespaceURI ||
    paragraph.namespaceURI !== change.namespaceURI ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  if (
    supportedParagraphMarkChange(paragraph, externalHyperlinks, imageEmbeds)
      ?.element === change
  ) {
    return null;
  }
  const markChanges = Array.from(runProperties.children).filter(
    (element) =>
      (element.localName === 'ins' || element.localName === 'del') &&
      element.namespaceURI === paragraph.namespaceURI,
  );
  if (markChanges.length !== 1 || markChanges[0] !== change) return null;
  const parsed = paragraphMarkChangeFromElement(change);
  if (!parsed) return null;
  if (
    !paragraphBodyIsUntrackedTextOnly(
      paragraph,
      properties,
      externalHyperlinks,
      imageEmbeds,
    )
  ) {
    return null;
  }
  return {
    author: parsed.author,
    date: parsed.date,
    id: parsed.id,
    kind: parsed.kind === 'deletion' ? 'merge' : 'split',
  };
}

/**
 * Paragraph-break merge/split requires untracked text-only bodies on both the
 * marked paragraph and its eligible neighbor. Soft breaks, tabs, carriage
 * returns, last-rendered page breaks, page-number and date-field glyphs,
 * footnoteRef, endnoteRef, annotationRef, separator, and continuationSeparator
 * glyphs, non-breaking and soft hyphens, empty/`rPr`-only runs,
 * relationship-free internal hyperlinks, safe relationship-bound external
 * hyperlinks, relationship-free bookmarks, and supported inline DrawingML
 * pictures match the whole-paragraph mark admission set (including picture-only
 * bodies); floating anchors, tracked wrappers, and unsafe or unresolved links
 * stay fail-closed.
 */
function paragraphBodyIsUntrackedTextOnly(
  paragraph: Element,
  properties: Element | undefined,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): boolean {
  const body = directChildren(paragraph).filter(
    (element) => element !== properties,
  );
  if (!body.length) return true;
  for (const child of body) {
    if (child.namespaceURI !== paragraph.namespaceURI) return false;
    if (
      child.localName === 'bookmarkStart' ||
      child.localName === 'bookmarkEnd'
    ) {
      if (!isRelationshipFreeBookmarkMarker(child)) return false;
      continue;
    }
    if (child.localName === 'hyperlink') {
      if (!isAdmittedHyperlink(child, externalHyperlinks)) return false;
      const runs = directChildren(child);
      if (!runs.length) return false;
      let linkHasText = false;
      for (const run of runs) {
        if (
          run.localName !== 'r' ||
          run.namespaceURI !== paragraph.namespaceURI ||
          !runIsTextOnly(run, 'insertion') ||
          runHasTrackedMark(run)
        ) {
          return false;
        }
        linkHasText ||= runHasVisibleText(run, 'insertion');
      }
      if (!linkHasText) return false;
      continue;
    }
    if (child.localName !== 'r') return false;
    if (runHasTrackedMark(child)) return false;
    if (runIsSupportedInlinePicture(child, imageEmbeds)) {
      continue;
    }
    if (!runIsTextOnly(child, 'insertion')) return false;
  }
  return true;
}

function runHasTrackedMark(run: Element): boolean {
  const properties = directChildren(run).find(
    (child) =>
      child.localName === 'rPr' && child.namespaceURI === run.namespaceURI,
  );
  if (!properties) return false;
  return Array.from(properties.querySelectorAll('*')).some(
    (descendant) =>
      descendant.localName === 'ins' || descendant.localName === 'del',
  );
}

function supportedParagraphMarkChange(
  paragraph: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): SupportedDocxParagraphMarkChange | null {
  const properties = directChildren(paragraph, 'pPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (properties.length !== 1) return null;
  const runProperties = directChildren(properties[0] as Element, 'rPr').filter(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  if (runProperties.length !== 1) return null;
  const changes = (runProperties[0] as Element).children;
  const revision = Array.from(changes).find(
    (element) =>
      (element.localName === 'ins' || element.localName === 'del') &&
      element.namespaceURI === paragraph.namespaceURI,
  );
  if (
    !revision ||
    Array.from(changes).filter(
      (element) =>
        (element.localName === 'ins' || element.localName === 'del') &&
        element.namespaceURI === paragraph.namespaceURI,
    ).length !== 1
  ) {
    return null;
  }
  const change = paragraphMarkChangeFromElement(revision);
  return change &&
    paragraphBodyMatchesChange(
      paragraph,
      properties[0],
      change,
      externalHyperlinks,
      imageEmbeds,
    )
    ? { element: revision, ...change }
    : null;
}

function paragraphMarkChangeFromElement(
  change: Element,
): Omit<ImportedDocxParagraphMarkChangeMarker, 'marker'> | null {
  if (
    (change.localName !== 'ins' && change.localName !== 'del') ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '') ||
    change.children.length > 0 ||
    hasUnsupportedWordAttributes(change)
  ) {
    return null;
  }
  const id = wordAttribute(change, 'id')?.trim() ?? '';
  const author = wordAttribute(change, 'author')?.trim() ?? '';
  const date = wordAttribute(change, 'date');
  if (
    !/^\+?\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (date !== null && date.length > MAX_REVISION_DATE_LENGTH) ||
    (date !== null && !Number.isFinite(Date.parse(date)))
  ) {
    return null;
  }
  return {
    id: `docx-paragraph-mark-change-${id}`,
    kind: change.localName === 'del' ? 'deletion' : 'insertion',
    author,
    date: normalizeRevisionDate(date),
  };
}

function paragraphBodyMatchesChange(
  paragraph: Element,
  properties: Element | undefined,
  change: Omit<ImportedDocxParagraphMarkChangeMarker, 'marker'>,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): boolean {
  const body = directChildren(paragraph).filter(
    (element) => element !== properties,
  );
  if (!body.length) return false;
  const expectedName = change.kind === 'deletion' ? 'del' : 'ins';
  // Admit one or more consecutive matching body wrappers (Word/WPS often
  // split a whole-paragraph revision across run formatting siblings).
  // Unsafe relationship-bound hyperlinks and mismatched authors stay
  // excluded. Relationship-free bookmarkStart/End markers, untracked
  // text-only runs (including empty/rPr-only), and untracked supported
  // inline DrawingML picture runs may appear as siblings of the revision
  // wrappers without blocking admission. Floating anchors, empty or
  // malformed drawings, and unresolved embeds stay fail-closed.
  let sawRevision = false;
  for (const revision of body) {
    if (
      revision.localName === 'bookmarkStart' ||
      revision.localName === 'bookmarkEnd'
    ) {
      if (!isRelationshipFreeBookmarkMarker(revision)) return false;
      continue;
    }
    if (revision.localName === 'r') {
      if (
        revision.namespaceURI !== paragraph.namespaceURI ||
        !(
          isUntrackedTextOnlySiblingRun(revision) ||
          runIsSupportedInlinePicture(revision, imageEmbeds)
        )
      ) {
        return false;
      }
      continue;
    }
    if (
      revision.localName !== expectedName ||
      revision.namespaceURI !== paragraph.namespaceURI ||
      hasUnsupportedWordAttributes(revision) ||
      !revisionBodyIsTextOnly(
        revision,
        change.kind,
        externalHyperlinks,
        imageEmbeds,
      )
    ) {
      return false;
    }
    const id = wordAttribute(revision, 'id')?.trim() ?? '';
    const author = wordAttribute(revision, 'author')?.trim() ?? '';
    const rawDate = wordAttribute(revision, 'date');
    if (
      !/^\+?\d{1,10}$/.test(id) ||
      author !== change.author ||
      normalizeRevisionDate(rawDate) !== change.date ||
      (rawDate !== null && !Number.isFinite(Date.parse(rawDate)))
    ) {
      return false;
    }
    sawRevision = true;
  }
  return sawRevision;
}

function revisionBodyIsTextOnly(
  revision: Element,
  kind: DocxParagraphMarkChangeKind,
  externalHyperlinks: DocxExternalHyperlinkTargets,
  imageEmbeds: DocxImageEmbedTargets,
): boolean {
  const children = directChildren(revision);
  if (!children.length) return false;
  let hasText = false;
  let hasPicture = false;
  for (const child of children) {
    if (child.namespaceURI !== revision.namespaceURI) return false;
    if (child.localName === 'r') {
      if (runIsSupportedInlinePicture(child, imageEmbeds)) {
        hasPicture = true;
        continue;
      }
      if (!runIsTextOnly(child, kind)) return false;
      hasText ||= runHasVisibleText(child, kind);
      continue;
    }
    if (child.localName === 'hyperlink') {
      if (!isAdmittedHyperlink(child, externalHyperlinks)) return false;
      const runs = directChildren(child);
      if (!runs.length) return false;
      for (const run of runs) {
        if (
          run.localName !== 'r' ||
          run.namespaceURI !== revision.namespaceURI ||
          !runIsTextOnly(run, kind)
        ) {
          return false;
        }
        hasText ||= runHasVisibleText(run, kind);
      }
      continue;
    }
    if (
      child.localName === 'bookmarkStart' ||
      child.localName === 'bookmarkEnd'
    ) {
      if (!isRelationshipFreeBookmarkMarker(child)) return false;
      continue;
    }
    return false;
  }
  return hasText || hasPicture;
}

/**
 * Supported inline DrawingML picture run: optional rPr plus exactly one
 * w:drawing that resolves to a package image embed with a supported transform.
 */
function runIsSupportedInlinePicture(
  run: Element,
  imageEmbeds: DocxImageEmbedTargets,
): boolean {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(run.namespaceURI ?? '')) return false;
  const children = directChildren(run);
  const properties = children.filter(
    (child) =>
      child.localName === 'rPr' && child.namespaceURI === run.namespaceURI,
  );
  if (properties.length > 1) return false;
  const drawings = children.filter(
    (child) =>
      child.localName === 'drawing' && child.namespaceURI === run.namespaceURI,
  );
  if (drawings.length !== 1) return false;
  for (const child of children) {
    if (child.namespaceURI !== run.namespaceURI) return false;
    if (child.localName === 'rPr') {
      if (
        Array.from(child.querySelectorAll('*')).some(
          (descendant) => descendant.namespaceURI !== run.namespaceURI,
        )
      ) {
        return false;
      }
      continue;
    }
    if (child.localName !== 'drawing') return false;
  }
  return isSupportedInlinePictureDrawing(drawings[0] as Element, imageEmbeds);
}

function isSupportedInlinePictureDrawing(
  drawing: Element,
  imageEmbeds: DocxImageEmbedTargets,
): boolean {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(drawing.namespaceURI ?? '')) {
    return false;
  }
  const containers = directChildren(drawing).filter(
    (element) =>
      (element.localName === 'anchor' || element.localName === 'inline') &&
      WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (containers.length !== 1 || containers[0]?.localName !== 'inline') {
    return false;
  }
  const container = containers[0] as Element;
  const graphicData = descendants(container, 'graphicData').filter((element) =>
    DRAWINGML_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  const pictures = descendants(container, 'pic').filter((element) =>
    PICTURE_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  const blips = descendants(container, 'blip').filter((element) =>
    DRAWINGML_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (
    graphicData.length !== 1 ||
    !PICTURE_NAMESPACES.has(graphicData[0]?.getAttribute('uri') ?? '') ||
    pictures.length !== 1 ||
    blips.length !== 1
  ) {
    return false;
  }
  const embedId = relationshipEmbedId(blips[0] as Element);
  return (
    Boolean(embedId && imageEmbeds.has(embedId)) &&
    readDocxImageTransform(drawing).supported
  );
}

function relationshipEmbedId(element: Element): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === 'embed' &&
      RELATIONSHIP_NAMESPACES.has(xmlAttributeNamespace(element, item) ?? ''),
  );
  if (matches.length !== 1) return null;
  const value = matches[0]?.value.trim() ?? '';
  return value || null;
}

function runIsTextOnly(
  run: Element,
  kind: DocxParagraphMarkChangeKind,
): boolean {
  const textName = kind === 'deletion' ? 'delText' : 't';
  const children = directChildren(run);
  const properties = children.filter(
    (child) =>
      child.localName === 'rPr' && child.namespaceURI === run.namespaceURI,
  );
  if (properties.length > 1) return false;
  for (const child of children) {
    if (child.namespaceURI !== run.namespaceURI) return false;
    if (child.localName === 'rPr') {
      if (
        Array.from(child.querySelectorAll('*')).some(
          (descendant) => descendant.namespaceURI !== run.namespaceURI,
        )
      ) {
        return false;
      }
      continue;
    }
    if (child.localName === 'br') {
      if (!isTextWrappingBreak(child)) return false;
      continue;
    }
    if (isAdmittedEmptyRunGlyph(child)) continue;
    if (child.localName !== textName || child.children.length) return false;
  }
  return true;
}

function runHasVisibleText(
  run: Element,
  kind: DocxParagraphMarkChangeKind,
): boolean {
  const textName = kind === 'deletion' ? 'delText' : 't';
  return directChildren(run).some((child) => {
    if (child.namespaceURI !== run.namespaceURI) return false;
    if (isAdmittedEmptyRunGlyph(child)) return true;
    return (
      child.localName === textName &&
      Boolean(child.textContent) &&
      child.children.length === 0
    );
  });
}

/**
 * Untracked sibling runs beside a whole-paragraph mark revision may carry
 * visible text, soft breaks, and the same attribute-free empty glyphs as
 * tracked text-only runs. Tracked marks inside the run stay fail-closed.
 */
function isUntrackedTextOnlySiblingRun(run: Element): boolean {
  return runIsTextOnly(run, 'insertion') && !runHasTrackedMark(run);
}

const HYPERLINK_ATTRIBUTES = new Set([
  'anchor',
  'docLocation',
  'history',
  'tgtFrame',
  'tooltip',
]);
const BOOKMARK_START_ATTRIBUTES = new Set(['id', 'name']);
const BOOKMARK_END_ATTRIBUTES = new Set(['id']);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);

function isAdmittedHyperlink(
  element: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets,
): boolean {
  return (
    isRelationshipFreeInternalHyperlink(element) ||
    isSupportedExternalHyperlink(element, externalHyperlinks)
  );
}

function isRelationshipFreeInternalHyperlink(element: Element): boolean {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
    return false;
  }
  let hasAnchor = false;
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return false;
    if (namespace && namespace !== element.namespaceURI) return false;
    const localName = xmlAttributeLocalName(attribute);
    if (namespace === element.namespaceURI || !namespace) {
      if (!HYPERLINK_ATTRIBUTES.has(localName)) return false;
      if (localName === 'anchor') {
        const value = attribute.value.trim();
        if (
          !value ||
          value.length > 255 ||
          /[\u0000-\u001f\u007f]/.test(value)
        ) {
          return false;
        }
        hasAnchor = true;
      }
    }
  }
  return hasAnchor;
}

/**
 * Relationship-bound external hyperlinks are admitted only when the r:id
 * resolves to a safe http(s)/mailto destination from document relationships.
 * Anchors mixed with r:id, missing targets, and non-hyperlink relationships
 * stay fail-closed.
 */
function isSupportedExternalHyperlink(
  element: Element,
  externalHyperlinks: DocxExternalHyperlinkTargets,
): boolean {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
    return false;
  }
  let relationshipId: string | null = null;
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    const localName = xmlAttributeLocalName(attribute);
    if (RELATIONSHIP_NAMESPACES.has(namespace)) {
      if (localName !== 'id' || relationshipId !== null) return false;
      const value = attribute.value.trim();
      if (!value) return false;
      relationshipId = value;
      continue;
    }
    if (namespace && namespace !== element.namespaceURI) return false;
    if (namespace === element.namespaceURI || !namespace) {
      if (!HYPERLINK_ATTRIBUTES.has(localName)) return false;
      if (localName === 'anchor') return false;
    }
  }
  return Boolean(relationshipId && externalHyperlinks.has(relationshipId));
}

function isRelationshipFreeBookmarkMarker(element: Element): boolean {
  if (
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') ||
    element.children.length > 0
  ) {
    return false;
  }
  const allowed =
    element.localName === 'bookmarkStart'
      ? BOOKMARK_START_ATTRIBUTES
      : element.localName === 'bookmarkEnd'
        ? BOOKMARK_END_ATTRIBUTES
        : null;
  if (!allowed) return false;
  let hasId = false;
  let hasName = element.localName !== 'bookmarkStart';
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return false;
    if (namespace && namespace !== element.namespaceURI) return false;
    const localName = xmlAttributeLocalName(attribute);
    if (!(namespace === element.namespaceURI || !namespace)) return false;
    if (!allowed.has(localName)) return false;
    const value = attribute.value.trim();
    if (!value || value.length > 255 || /[\u0000-\u001f\u007f]/.test(value)) {
      return false;
    }
    if (localName === 'id') {
      if (!/^\+?\d{1,10}$/.test(value)) return false;
      hasId = true;
    }
    if (localName === 'name') hasName = true;
  }
  return hasId && hasName;
}

function isTextWrappingBreak(element: Element): boolean {
  if (element.children.length) return false;
  const type = wordAttribute(element, 'type');
  return type === null || type === 'textWrapping';
}

/** Empty CT_Empty run glyphs already projected elsewhere in Writer import. */
const ADMITTED_EMPTY_RUN_GLYPHS = new Set([
  'tab',
  'noBreakHyphen',
  'softHyphen',
  'cr',
  'lastRenderedPageBreak',
  'pgNum',
  'dayLong',
  'monthLong',
  'yearLong',
  'dayShort',
  'monthShort',
  'yearShort',
  'footnoteRef',
  'endnoteRef',
  'annotationRef',
  'separator',
  'continuationSeparator',
]);

function isAdmittedEmptyRunGlyph(element: Element): boolean {
  return (
    ADMITTED_EMPTY_RUN_GLYPHS.has(element.localName) &&
    element.children.length === 0 &&
    element.attributes.length === 0
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

function hasUnsupportedWordAttributes(element: Element): boolean {
  const namespace = element.namespaceURI;
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !REVISION_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
  );
}

function uniqueChangeId(base: string, ids: Set<string>): string {
  if (!ids.has(base)) {
    ids.add(base);
    return base;
  }
  let suffix = 2;
  while (ids.has(`${base}-${suffix}`)) suffix += 1;
  const id = `${base}-${suffix}`;
  ids.add(id);
  return id;
}

function insertParagraphMarker(paragraph: Element, marker: string): void {
  const document = paragraph.ownerDocument;
  const namespace = paragraph.namespaceURI ?? WORD_NAMESPACE;
  const prefix = paragraph.prefix ? `${paragraph.prefix}:` : '';
  const run = document.createElementNS(namespace, `${prefix}r`);
  const text = document.createElementNS(namespace, `${prefix}t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  const properties = directChildren(paragraph, 'pPr').find(
    (element) => element.namespaceURI === paragraph.namespaceURI,
  );
  paragraph.insertBefore(run, properties?.nextSibling ?? paragraph.firstChild);
}

function closestParagraphBlock(
  element: Element | null,
  markerNode: Node,
): HTMLElement | null {
  const explicit = element?.closest('p, h1, h2, h3, h4, h5, h6, figcaption');
  if (explicit instanceof HTMLElement) return explicit;
  const container = element?.closest('li, blockquote, td, th, div');
  return container instanceof HTMLElement
    ? wrapParagraphSegment(container, markerNode)
    : null;
}

function wrapParagraphSegment(
  container: HTMLElement,
  markerNode: Node,
): HTMLElement {
  let anchor: Node = markerNode;
  while (anchor.parentNode && anchor.parentNode !== container) {
    anchor = anchor.parentNode;
  }
  const children = Array.from(container.childNodes);
  const anchorIndex = children.indexOf(anchor as ChildNode);
  if (anchorIndex < 0) return container;
  let start = anchorIndex;
  let end = anchorIndex;
  while (start > 0 && !isParagraphBoundary(children[start - 1])) start -= 1;
  while (end + 1 < children.length && !isParagraphBoundary(children[end + 1])) {
    end += 1;
  }
  const paragraph = container.ownerDocument.createElement('p');
  const grouped = children.slice(start, end + 1);
  container.insertBefore(paragraph, grouped[0] ?? null);
  paragraph.append(...grouped);
  return paragraph;
}

function isParagraphBoundary(node: Node | undefined): boolean {
  if (!(node instanceof HTMLElement)) return false;
  return [
    'blockquote',
    'div',
    'figcaption',
    'h1',
    'h2',
    'h3',
    'h4',
    'h5',
    'h6',
    'hr',
    'ol',
    'p',
    'pre',
    'table',
    'ul',
  ].includes(node.tagName.toLowerCase());
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
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
