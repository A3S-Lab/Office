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

interface WorkPdfJsInternal {
  events?: {
    subscribe?: (name: string, handler: () => void) => void;
  };
  newObjectDeferred?: () => number;
  newObjectDeferredBegin?: (objectId: number, doOutput?: boolean) => number;
  out?: (content: string) => void;
  write?: (...parts: string[]) => void;
  workPdfMarkInfoSubscribed?: boolean;
  workPdfStructTreePlan?: WorkPdfStructTreePlan;
  workPdfStructTreeRootObjectId?: number;
  workPdfStructTreeSubscribed?: boolean;
}

const MAX_OUTLINE_ENTRIES = 512;
const MAX_OUTLINE_TITLE_LENGTH = 200;
const MAX_ACTUAL_TEXT_LENGTH = 2048;
const OUTLINE_SELECTOR = 'h1, h2, h3, h4, h5, h6, p[data-office-outline-level]';

/**
 * Applies bounded PDF document metadata and outline bookmarks. This is the
 * tagged/accessibility bootstrap: language + title + heading outline + MarkInfo
 * + a stub StructTreeRoot (Document / H1–H6 / P), without inventing a second
 * layout model or claiming full PDF/UA certification / MCID parent trees.
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
 * Emits a bounded StructTreeRoot stub (Document + outline-derived H1–H6/P
 * StructElems with Alt, empty K) via jsPDF postPutResources / putCatalog.
 * Not a full PDF/UA parent tree: no MCIDs, no ParentTree, no page /K links.
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
  internal.workPdfStructTreePlan = {
    language: plan.language ?? null,
    outline: plan.outline ?? [],
  };
  if (internal.workPdfStructTreeSubscribed) return;
  internal.workPdfStructTreeSubscribed = true;
  try {
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
 * Opens a `/Span` BDC with `/ActualText` for a vector text run. Pairs with
 * {@link endWorkPdfActualTextSpan}. Fail-soft when jsPDF internals are absent.
 */
export function beginWorkPdfActualTextSpan(pdf: JsPdf, text: string): boolean {
  const internal = workPdfJsInternal(pdf);
  if (!internal?.out) return false;
  const clipped = text.slice(0, MAX_ACTUAL_TEXT_LENGTH);
  if (!clipped) return false;
  ensureWorkPdfMarkInfo(pdf);
  try {
    internal.out(
      `/Span << /ActualText ${encodePdfActualTextOperand(clipped)} >> BDC`,
    );
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
 * Writes StructTreeRoot → Document → outline-derived H1–H6/P stubs. Returns the
 * root object id for the catalog `/StructTreeRoot` reference.
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
  const kids: Array<{ alt: string; objectId: number; role: string }> = [];
  for (const entry of plan.outline) {
    if (kids.length >= MAX_OUTLINE_ENTRIES) break;
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
    kids.push({
      alt,
      objectId: deferred(),
      role: workPdfStructRoleFromOutlineLevel(level),
    });
  }
  const rootObjectId = deferred();
  const documentObjectId = deferred();
  for (const kid of kids) {
    begin(kid.objectId, true);
    out('<<');
    out('/Type /StructElem');
    out(`/S /${kid.role}`);
    out(`/P ${documentObjectId} 0 R`);
    out(`/Alt ${encodePdfActualTextOperand(kid.alt)}`);
    out('/K []');
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
  if (kids.length === 0) {
    out('/K []');
  } else {
    out(`/K [${kids.map((kid) => `${kid.objectId} 0 R`).join(' ')}]`);
  }
  out('>>');
  out('endobj');
  begin(rootObjectId, true);
  out('<<');
  out('/Type /StructTreeRoot');
  out(`/K [${documentObjectId} 0 R]`);
  out('>>');
  out('endobj');
  return rootObjectId;
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
