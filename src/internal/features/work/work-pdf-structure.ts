import type { jsPDF as JsPdf, OutlineItem } from 'jspdf';

export interface WorkPdfOutlineEntry {
  /** 1-based outline level for nesting (h1 / outlineLvl 0 → 1 … outlineLvl 8 → 9). */
  level?: number;
  pageNumber: number;
  title: string;
}

export interface WorkPdfDocumentStructure {
  language?: string;
  outline?: readonly WorkPdfOutlineEntry[];
  title?: string;
}

interface WorkPdfStructTreePlan {
  language: string | null;
  outline: readonly WorkPdfOutlineEntry[];
}

interface WorkPdfContentLink {
  mcid: number;
  pageNumber: number;
  text: string;
}

interface WorkPdfPageInfo {
  objId: number;
  pageNumber: number;
}

interface WorkPdfJsInternal {
  events?: {
    subscribe?: (name: string, handler: (...args: unknown[]) => void) => void;
  };
  getCurrentPageInfo?: () => WorkPdfPageInfo | undefined;
  getPageInfo?: (pageNumber: number) => WorkPdfPageInfo | undefined;
  newObjectDeferred?: () => number;
  newObjectDeferredBegin?: (objectId: number, doOutput?: boolean) => number;
  out?: (content: string) => void;
  write?: (...parts: string[]) => void;
  workPdfContentLinks?: WorkPdfContentLink[];
  workPdfMarkInfoSubscribed?: boolean;
  workPdfMcidCounters?: Record<number, number>;
  workPdfPageStructParents?: Record<number, number>;
  workPdfStructTreePlan?: WorkPdfStructTreePlan;
  workPdfStructTreeRootObjectId?: number;
  workPdfStructTreeSubscribed?: boolean;
  workPdfViewerPreferencesSubscribed?: boolean;
}

const MAX_OUTLINE_ENTRIES = 512;
const MAX_OUTLINE_TITLE_LENGTH = 200;
const MAX_ACTUAL_TEXT_LENGTH = 2048;
const MAX_CONTENT_LINKS = 2048;
const OUTLINE_SELECTOR = 'h1, h2, h3, h4, h5, h6, p[data-office-outline-level]';

/**
 * Applies bounded PDF document metadata and outline bookmarks. This is the
 * tagged/accessibility bootstrap: language + title + heading outline + MarkInfo
 * + StructTreeRoot (Document / H1–H6 / P) with ParentTree / MCID links for
 * vector-run Span content, without claiming full PDF/UA certification.
 */
export function applyWorkPdfDocumentStructure(
  pdf: JsPdf,
  structure: WorkPdfDocumentStructure,
): void {
  const title = structure.title?.trim();
  if (title) {
    pdf.setProperties({
      title,
      author: 'A3S Work',
      creator: 'A3S Work',
    });
  }
  const language = normalizePdfLanguage(structure.language);
  if (language) {
    try {
      pdf.setLanguage(language as Parameters<JsPdf['setLanguage']>[0]);
    } catch {
      // jsPDF only accepts a fixed language enum; ignore unknown tags.
    }
  }
  ensureWorkPdfMarkInfo(pdf);
  ensureWorkPdfViewerPreferences(pdf);
  const outline = structure.outline ?? [];
  ensureWorkPdfStructTreeRoot(pdf, { language, outline });
  const stack: Array<{ item: OutlineItem; level: number }> = [];
  let count = 0;
  for (const entry of outline) {
    if (count >= MAX_OUTLINE_ENTRIES) break;
    const entryTitle = entry.title.trim().slice(0, MAX_OUTLINE_TITLE_LENGTH);
    if (
      !entryTitle ||
      !Number.isSafeInteger(entry.pageNumber) ||
      entry.pageNumber < 1
    ) {
      continue;
    }
    const level =
      typeof entry.level === 'number' &&
      Number.isSafeInteger(entry.level) &&
      entry.level >= 1
        ? entry.level
        : 1;
    while ((stack.at(-1)?.level ?? 0) >= level) stack.pop();
    const parent = stack.at(-1)?.item ?? null;
    try {
      const item = pdf.outline.add(parent, entryTitle, {
        pageNumber: entry.pageNumber,
      });
      stack.push({ item, level });
      count += 1;
    } catch {
      // Outline plugin may be unavailable in some builds.
    }
  }
}

/**
 * Declares `/MarkInfo << /Marked true >>` once via jsPDF's putCatalog hook so
 * ActualText Span marked content is catalog-visible.
 */
export function ensureWorkPdfMarkInfo(pdf: JsPdf): void {
  const internal = workPdfJsInternal(pdf);
  if (!internal?.events?.subscribe || !internal.write) return;
  if (internal.workPdfMarkInfoSubscribed) return;
  internal.workPdfMarkInfoSubscribed = true;
  try {
    internal.events.subscribe('putCatalog', () => {
      internal.write?.('/MarkInfo << /Marked true >>');
    });
  } catch {
    internal.workPdfMarkInfoSubscribed = false;
  }
}

/**
 * Declares `/ViewerPreferences << /DisplayDocTitle true >>` so tagged-PDF
 * consumers prefer the document title. Not a PDF/UA certification claim.
 */
export function ensureWorkPdfViewerPreferences(pdf: JsPdf): void {
  const internal = workPdfJsInternal(pdf);
  if (!internal?.events?.subscribe || !internal.write) return;
  if (internal.workPdfViewerPreferencesSubscribed) return;
  internal.workPdfViewerPreferencesSubscribed = true;
  try {
    internal.events.subscribe('putCatalog', () => {
      internal.write?.('/ViewerPreferences << /DisplayDocTitle true >>');
    });
  } catch {
    internal.workPdfViewerPreferencesSubscribed = false;
  }
}

/**
 * Emits StructTreeRoot (Document + outline-derived H1–H6/P + Span content
 * links) with ParentTree / page StructParents via jsPDF hooks.
 * Page-matched vector-run Spans nest under the last outline role on that
 * page via `/K`; unmatched Spans stay Document kids. Not a full PDF/UA
 * certification claim.
 */
export function ensureWorkPdfStructTreeRoot(
  pdf: JsPdf,
  plan: {
    language?: string | null;
    outline?: readonly WorkPdfOutlineEntry[];
  },
): void {
  const internal = workPdfJsInternal(pdf);
  if (
    !internal?.events?.subscribe ||
    !internal.out ||
    !internal.write ||
    typeof internal.newObjectDeferred !== 'function' ||
    typeof internal.newObjectDeferredBegin !== 'function'
  ) {
    return;
  }
  const previous = internal.workPdfStructTreePlan;
  internal.workPdfStructTreePlan = {
    language: plan.language ?? previous?.language ?? null,
    outline: plan.outline ?? previous?.outline ?? [],
  };
  if (internal.workPdfStructTreeSubscribed) return;
  internal.workPdfStructTreeSubscribed = true;
  if (!internal.workPdfContentLinks) internal.workPdfContentLinks = [];
  if (!internal.workPdfMcidCounters) internal.workPdfMcidCounters = {};
  if (!internal.workPdfPageStructParents) {
    internal.workPdfPageStructParents = {};
  }
  try {
    internal.events.subscribe('putPage', (...args: unknown[]) => {
      const page = args[0] as
        | { pageNumber?: number; pageContext?: unknown }
        | undefined;
      const pageNumber =
        typeof page?.pageNumber === 'number' ? page.pageNumber : undefined;
      if (pageNumber === undefined) return;
      const parents = internal.workPdfPageStructParents;
      if (!parents || parents[pageNumber] === undefined) return;
      internal.write?.(`/StructParents ${parents[pageNumber]}`);
    });
    internal.events.subscribe('postPutResources', () => {
      internal.workPdfStructTreeRootObjectId =
        writeWorkPdfStructTreeObjects(internal);
    });
    internal.events.subscribe('putCatalog', () => {
      const rootId = internal.workPdfStructTreeRootObjectId;
      if (rootId !== undefined) {
        internal.write?.(`/StructTreeRoot ${rootId} 0 R`);
      }
    });
  } catch {
    internal.workPdfStructTreeSubscribed = false;
  }
}

/**
 * Maps outline levels to PDF structure roles: 1–6 → H1–H6, else P.
 * Exported for unit coverage of the role boundary.
 */
export function workPdfStructRoleFromOutlineLevel(level: number): string {
  if (Number.isSafeInteger(level) && level >= 1 && level <= 6) {
    return `H${level}`;
  }
  return 'P';
}

/**
 * Opens a `/Span` BDC with `/ActualText` and a page-local `/MCID` for a vector
 * text run. Pairs with {@link endWorkPdfActualTextSpan}. Fail-soft when jsPDF
 * internals are absent; falls back to ActualText-only when the MCID budget is
 * exhausted.
 */
export function beginWorkPdfActualTextSpan(pdf: JsPdf, text: string): boolean {
  const internal = workPdfJsInternal(pdf);
  if (!internal?.out) return false;
  const clipped = text.slice(0, MAX_ACTUAL_TEXT_LENGTH);
  if (!clipped) return false;
  ensureWorkPdfMarkInfo(pdf);
  ensureWorkPdfStructTreeRoot(pdf, {});
  const mcid = allocateWorkPdfMcid(internal);
  try {
    if (mcid === null) {
      internal.out(
        `/Span << /ActualText ${encodePdfActualTextOperand(clipped)} >> BDC`,
      );
    } else {
      internal.workPdfContentLinks?.push({
        mcid: mcid.mcid,
        pageNumber: mcid.pageNumber,
        text: clipped,
      });
      internal.out(
        `/Span << /ActualText ${encodePdfActualTextOperand(clipped)} /MCID ${mcid.mcid} >> BDC`,
      );
    }
    return true;
  } catch {
    return false;
  }
}

/** Closes a Span marked-content sequence opened by {@link beginWorkPdfActualTextSpan}. */
export function endWorkPdfActualTextSpan(pdf: JsPdf): void {
  const internal = workPdfJsInternal(pdf);
  if (!internal?.out) return;
  try {
    internal.out('EMC');
  } catch {
    // Ignore when the page stream is unavailable.
  }
}

/**
 * Encodes an ActualText operand: PDF literal for Latin-1, UTF-16BE hex (BOM)
 * otherwise. Exported for unit coverage of the escape boundary.
 */
export function encodePdfActualTextOperand(text: string): string {
  let latin1 = true;
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) > 0xff) {
      latin1 = false;
      break;
    }
  }
  if (latin1) {
    return `(${escapePdfLiteralString(text)})`;
  }
  let hex = 'FEFF';
  for (const char of text) {
    const codePoint = char.codePointAt(0);
    if (codePoint === undefined) continue;
    if (codePoint > 0xffff) {
      const offset = codePoint - 0x10000;
      const high = 0xd800 + (offset >> 10);
      const low = 0xdc00 + (offset & 0x3ff);
      hex += high.toString(16).toUpperCase().padStart(4, '0');
      hex += low.toString(16).toUpperCase().padStart(4, '0');
    } else {
      hex += codePoint.toString(16).toUpperCase().padStart(4, '0');
    }
  }
  return `<${hex}>`;
}

export function collectWorkPdfOutlineEntriesFromRoot(
  root: HTMLElement,
  pageNumber: number,
  pageBounds?: {
    height: number;
    left: number;
    top: number;
    width: number;
  },
): WorkPdfOutlineEntry[] {
  if (!Number.isSafeInteger(pageNumber) || pageNumber < 1) return [];
  const entries: WorkPdfOutlineEntry[] = [];
  for (const heading of Array.from(
    root.querySelectorAll<HTMLElement>(OUTLINE_SELECTOR),
  )) {
    if (entries.length >= MAX_OUTLINE_ENTRIES) break;
    const level = workPdfOutlineLevelFromElement(heading);
    if (level === null) continue;
    if (heading.closest('[aria-hidden="true"]')) {
      const hidden = heading.closest('[aria-hidden="true"]');
      if (
        hidden &&
        !hidden.classList.contains('work-document-live-pdf-snapshot') &&
        !hidden.classList.contains('work-pdf-export-page')
      ) {
        continue;
      }
    }
    if (pageBounds) {
      const rect = heading.getBoundingClientRect();
      const x = rect.left - pageBounds.left;
      const y = rect.top - pageBounds.top;
      if (
        x + rect.width <= 0 ||
        y + rect.height <= 0 ||
        x >= pageBounds.width ||
        y >= pageBounds.height
      ) {
        continue;
      }
    }
    const title = heading.textContent?.replace(/\s+/g, ' ').trim() ?? '';
    if (!title) continue;
    entries.push({
      level,
      pageNumber,
      title: title.slice(0, MAX_OUTLINE_TITLE_LENGTH),
    });
  }
  return entries;
}

export function normalizePdfLanguage(
  language: string | undefined,
): string | null {
  const trimmed = language?.trim();
  if (!trimmed || trimmed.length > 16) return null;
  if (!/^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$/.test(trimmed)) return null;
  return trimmed;
}

/** Matches Writer outline: h1–h6 and 0-based `data-office-outline-level` 0–8 → 1–9. */
function workPdfOutlineLevelFromElement(element: HTMLElement): number | null {
  const headingMatch = /^H([1-6])$/.exec(element.tagName);
  if (headingMatch) return Number(headingMatch[1]);
  if (element.tagName !== 'P') return null;
  const outlineLevel = Number(element.dataset.officeOutlineLevel);
  return Number.isInteger(outlineLevel) &&
    outlineLevel >= 0 &&
    outlineLevel <= 8
    ? outlineLevel + 1
    : null;
}

/**
 * Allocates a page-local MCID and StructParents key. Returns null when the
 * document content-link budget is exhausted.
 */
function allocateWorkPdfMcid(
  internal: WorkPdfJsInternal,
): { mcid: number; pageNumber: number } | null {
  const pageInfo = internal.getCurrentPageInfo?.();
  if (!pageInfo || !Number.isSafeInteger(pageInfo.pageNumber)) return null;
  const pageNumber = pageInfo.pageNumber;
  if (!internal.workPdfContentLinks) internal.workPdfContentLinks = [];
  if (internal.workPdfContentLinks.length >= MAX_CONTENT_LINKS) return null;
  if (!internal.workPdfMcidCounters) internal.workPdfMcidCounters = {};
  if (!internal.workPdfPageStructParents) {
    internal.workPdfPageStructParents = {};
  }
  if (internal.workPdfPageStructParents[pageNumber] === undefined) {
    internal.workPdfPageStructParents[pageNumber] = Object.keys(
      internal.workPdfPageStructParents,
    ).length;
  }
  const mcid = internal.workPdfMcidCounters[pageNumber] ?? 0;
  internal.workPdfMcidCounters[pageNumber] = mcid + 1;
  return { mcid, pageNumber };
}

/**
 * Writes StructTreeRoot → Document → outline roles (with page-matched Span
 * kids) + unmatched Span MCID kids, plus the ParentTree number tree.
 * Returns the root object id for the catalog.
 */
function writeWorkPdfStructTreeObjects(
  internal: WorkPdfJsInternal,
): number | undefined {
  const deferred = internal.newObjectDeferred;
  const begin = internal.newObjectDeferredBegin;
  const out = internal.out;
  if (!deferred || !begin || !out) return undefined;
  const plan = internal.workPdfStructTreePlan ?? {
    language: null,
    outline: [],
  };
  const outlineKids: Array<{
    alt: string;
    objectId: number;
    pageNumber: number;
    role: string;
  }> = [];
  for (const entry of plan.outline) {
    if (outlineKids.length >= MAX_OUTLINE_ENTRIES) break;
    const alt = entry.title.trim().slice(0, MAX_OUTLINE_TITLE_LENGTH);
    if (!alt) continue;
    if (!Number.isSafeInteger(entry.pageNumber) || entry.pageNumber < 1) {
      continue;
    }
    const level =
      typeof entry.level === 'number' &&
      Number.isSafeInteger(entry.level) &&
      entry.level >= 1
        ? entry.level
        : 1;
    outlineKids.push({
      alt,
      objectId: deferred(),
      pageNumber: entry.pageNumber,
      role: workPdfStructRoleFromOutlineLevel(level),
    });
  }

  const contentLinks = internal.workPdfContentLinks ?? [];
  const spanKids: Array<{
    alt: string;
    mcid: number;
    objectId: number;
    pageObjId: number;
    pageNumber: number;
    parentObjectId: number;
  }> = [];
  const rootObjectId = deferred();
  const documentObjectId = deferred();
  const spansByOutline = new Map<number, number[]>();
  const documentSpanIds: number[] = [];

  for (const link of contentLinks) {
    const pageInfo = internal.getPageInfo?.(link.pageNumber);
    if (!pageInfo || !Number.isSafeInteger(pageInfo.objId)) continue;
    const objectId = deferred();
    const outlineParent = findWorkPdfOutlineParentForPage(
      outlineKids,
      link.pageNumber,
    );
    const parentObjectId = outlineParent?.objectId ?? documentObjectId;
    spanKids.push({
      alt: link.text.slice(0, MAX_ACTUAL_TEXT_LENGTH),
      mcid: link.mcid,
      objectId,
      pageObjId: pageInfo.objId,
      pageNumber: link.pageNumber,
      parentObjectId,
    });
    if (outlineParent) {
      const kids = spansByOutline.get(outlineParent.objectId) ?? [];
      kids.push(objectId);
      spansByOutline.set(outlineParent.objectId, kids);
    } else {
      documentSpanIds.push(objectId);
    }
  }

  const parentTreeObjectId =
    spanKids.length > 0 ||
    Object.keys(internal.workPdfPageStructParents ?? {}).length > 0
      ? deferred()
      : undefined;

  for (const kid of outlineKids) {
    begin(kid.objectId, true);
    out('<<');
    out('/Type /StructElem');
    out(`/S /${kid.role}`);
    out(`/P ${documentObjectId} 0 R`);
    out(`/Alt ${encodePdfActualTextOperand(kid.alt)}`);
    const nested = spansByOutline.get(kid.objectId) ?? [];
    if (nested.length === 0) {
      out('/K []');
    } else {
      out(`/K [${nested.map((id) => `${id} 0 R`).join(' ')}]`);
    }
    out('>>');
    out('endobj');
  }

  for (const kid of spanKids) {
    begin(kid.objectId, true);
    out('<<');
    out('/Type /StructElem');
    out('/S /Span');
    out(`/P ${kid.parentObjectId} 0 R`);
    out(`/Pg ${kid.pageObjId} 0 R`);
    out(`/K ${kid.mcid}`);
    out(`/Alt ${encodePdfActualTextOperand(kid.alt)}`);
    out('>>');
    out('endobj');
  }

  begin(documentObjectId, true);
  out('<<');
  out('/Type /StructElem');
  out('/S /Document');
  out(`/P ${rootObjectId} 0 R`);
  if (plan.language) {
    out(`/Lang (${escapePdfLiteralString(plan.language)})`);
  }
  const documentKids = [
    ...outlineKids.map((kid) => `${kid.objectId} 0 R`),
    ...documentSpanIds.map((id) => `${id} 0 R`),
  ];
  if (documentKids.length === 0) {
    out('/K []');
  } else {
    out(`/K [${documentKids.join(' ')}]`);
  }
  out('>>');
  out('endobj');

  if (parentTreeObjectId !== undefined) {
    begin(parentTreeObjectId, true);
    out('<<');
    out(`/Nums [${writeWorkPdfParentTreeNums(internal, spanKids)}]`);
    out('>>');
    out('endobj');
  }

  const parentTreeNextKey = Object.keys(
    internal.workPdfPageStructParents ?? {},
  ).length;

  begin(rootObjectId, true);
  out('<<');
  out('/Type /StructTreeRoot');
  out(`/K [${documentObjectId} 0 R]`);
  if (parentTreeObjectId !== undefined) {
    out(`/ParentTree ${parentTreeObjectId} 0 R`);
    out(`/ParentTreeNextKey ${parentTreeNextKey}`);
  }
  out('>>');
  out('endobj');
  return rootObjectId;
}

/**
 * Picks the last outline role on `pageNumber` as the Span parent. Spans on
 * pages without an outline role stay under Document.
 */
function findWorkPdfOutlineParentForPage(
  outlineKids: readonly { objectId: number; pageNumber: number }[],
  pageNumber: number,
): { objectId: number; pageNumber: number } | undefined {
  for (let index = outlineKids.length - 1; index >= 0; index -= 1) {
    const kid = outlineKids[index];
    if (kid?.pageNumber === pageNumber) return kid;
  }
  return undefined;
}

/**
 * Builds ParentTree `/Nums` pairs: StructParents key → MCID-indexed array of
 * Span StructElem refs (null holes for unused MCID slots).
 */
function writeWorkPdfParentTreeNums(
  internal: WorkPdfJsInternal,
  spanKids: readonly {
    mcid: number;
    objectId: number;
    pageNumber: number;
  }[],
): string {
  const parents = internal.workPdfPageStructParents ?? {};
  const byPage = new Map<number, Array<{ mcid: number; objectId: number }>>();
  for (const kid of spanKids) {
    const list = byPage.get(kid.pageNumber) ?? [];
    list.push({ mcid: kid.mcid, objectId: kid.objectId });
    byPage.set(kid.pageNumber, list);
  }
  const parts: string[] = [];
  const keys = Object.entries(parents)
    .map(([pageNumber, key]) => ({
      key,
      pageNumber: Number(pageNumber),
    }))
    .sort((a, b) => a.key - b.key);
  for (const { key, pageNumber } of keys) {
    const links = byPage.get(pageNumber) ?? [];
    const maxMcid = links.reduce((max, link) => Math.max(max, link.mcid), -1);
    const slots: string[] = [];
    for (let mcid = 0; mcid <= maxMcid; mcid += 1) {
      const match = links.find((link) => link.mcid === mcid);
      slots.push(match ? `${match.objectId} 0 R` : 'null');
    }
    parts.push(`${key} [${slots.join(' ')}]`);
  }
  return parts.join(' ');
}

function workPdfJsInternal(pdf: JsPdf): WorkPdfJsInternal | null {
  const internal = (pdf as unknown as { internal?: WorkPdfJsInternal })
    .internal;
  return internal ?? null;
}

function escapePdfLiteralString(text: string): string {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\(/g, '\\(')
    .replace(/\)/g, '\\)')
    .replace(/\r/g, '\\r')
    .replace(/\n/g, '\\n');
}
