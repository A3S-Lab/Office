import type { CellObject } from 'xlsx';

export interface SpreadsheetWorksheetCompatibilitySummary {
  hasArrayFormulas: boolean;
  hasComments: boolean;
  hasCommentThreads: boolean;
  hasLinks: boolean;
  hasRichText: boolean;
}

export function emptySpreadsheetWorksheetCompatibilitySummary(): SpreadsheetWorksheetCompatibilitySummary {
  return {
    hasArrayFormulas: false,
    hasComments: false,
    hasCommentThreads: false,
    hasLinks: false,
    hasRichText: false,
  };
}

export function updateSpreadsheetWorksheetCompatibilitySummary(
  summary: SpreadsheetWorksheetCompatibilitySummary,
  cell: CellObject,
): void {
  const source = cell as CellObject & {
    c?: Array<{ T?: boolean }>;
    l?: { Target?: string };
    F?: string;
    r?: string;
  };
  if (source.c?.length) {
    summary.hasComments = true;
    if (source.c.length > 1 || source.c.some((comment) => comment.T)) {
      summary.hasCommentThreads = true;
    }
  }
  if (source.l?.Target) summary.hasLinks = true;
  if (source.F) summary.hasArrayFormulas = true;
  if (source.r) summary.hasRichText = true;
}
