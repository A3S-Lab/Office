import type { jsPDF as JsPdf } from 'jspdf';
import { registerWorkPdfTableStructEntry } from './work-pdf-structure';
import type { WorkPdfPageBounds } from './work-pdf-text-layer';

export interface WorkPdfTableCellStruct {
  alt: string;
  header: boolean;
}

export interface WorkPdfTableRowStruct {
  cells: WorkPdfTableCellStruct[];
}

export interface WorkPdfTableStruct {
  rows: WorkPdfTableRowStruct[];
}

const MAX_TABLES = 256;
const MAX_ROWS = 256;
const MAX_CELLS = 64;
const MAX_ALT_LENGTH = 512;

/**
 * Collects simple HTML tables that intersect the page bounds. Cell `/Alt`
 * uses trimmed text content; empty cells keep an empty alt. Not a PDF/UA
 * certification claim.
 */
export function collectWorkPdfTableStructs(
  root: ParentNode,
  page: WorkPdfPageBounds,
): WorkPdfTableStruct[] {
  const tables: WorkPdfTableStruct[] = [];
  for (const node of Array.from(root.querySelectorAll('table'))) {
    if (tables.length >= MAX_TABLES) break;
    if (!(node instanceof HTMLTableElement)) continue;
    let rect: DOMRect;
    try {
      rect = node.getBoundingClientRect();
    } catch {
      continue;
    }
    if (
      !Number.isFinite(rect.width) ||
      !Number.isFinite(rect.height) ||
      rect.width <= 0 ||
      rect.height <= 0
    ) {
      continue;
    }
    const x = rect.left - page.left;
    const y = rect.top - page.top;
    if (
      x + rect.width <= 0 ||
      y + rect.height <= 0 ||
      x >= page.width ||
      y >= page.height
    ) {
      continue;
    }
    const rows: WorkPdfTableRowStruct[] = [];
    for (const rowNode of Array.from(node.rows)) {
      if (rows.length >= MAX_ROWS) break;
      const cells: WorkPdfTableCellStruct[] = [];
      for (const cellNode of Array.from(rowNode.cells)) {
        if (cells.length >= MAX_CELLS) break;
        const alt = (cellNode.textContent ?? '')
          .replace(/\s+/g, ' ')
          .trim()
          .slice(0, MAX_ALT_LENGTH);
        cells.push({
          alt,
          header: cellNode.tagName.toLowerCase() === 'th',
        });
      }
      if (cells.length) rows.push({ cells });
    }
    if (rows.length) tables.push({ rows });
  }
  return tables;
}

/**
 * Registers `/Table` → `/TR` → `/TH`|`/TD` StructElems for collected tables on
 * the current PDF page.
 */
export function appendWorkPdfTableStructEntries(
  pdf: JsPdf,
  tables: readonly WorkPdfTableStruct[],
): void {
  for (const table of tables) {
    registerWorkPdfTableStructEntry(pdf, table);
  }
}
