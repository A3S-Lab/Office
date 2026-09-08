import type { jsPDF as JsPdf } from 'jspdf';

export interface WorkPdfOutlineEntry {
  pageNumber: number;
  title: string;
}

export interface WorkPdfDocumentStructure {
  language?: string;
  outline?: readonly WorkPdfOutlineEntry[];
  title?: string;
}

const MAX_OUTLINE_ENTRIES = 512;
const MAX_OUTLINE_TITLE_LENGTH = 200;

/**
 * Applies bounded PDF document metadata and outline bookmarks. This is the
 * first tagged/accessibility bootstrap: language + title + heading outline,
 * without inventing a second layout model or a full PDF/UA structure tree.
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
  const outline = structure.outline ?? [];
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
    try {
      pdf.outline.add(null, entryTitle, { pageNumber: entry.pageNumber });
      count += 1;
    } catch {
      // Outline plugin may be unavailable in some builds.
    }
  }
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
    root.querySelectorAll<HTMLElement>('h1, h2, h3, h4, h5, h6'),
  )) {
    if (entries.length >= MAX_OUTLINE_ENTRIES) break;
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
