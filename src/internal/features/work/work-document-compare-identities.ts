import type {
  WorkDocumentChange,
  WorkDocumentChangeIdentity,
} from './work-document-changes';
import {
  boundedMetadata,
  normalizedDate,
  stableHash,
  wordParagraphId,
} from './work-document-compare-stability';
import type { WorkDocumentChangeKind } from './work-types';

export type DocumentComparisonMode = 'compare' | 'combine';

export interface DocumentComparisonSummary {
  insertions: number;
  deletions: number;
  formatting: number;
  paragraphFormatting: number;
  /** Number of paired text moves inferred by Compare (when non-zero). */
  moves?: number;
}

export interface DocumentComparisonOptions {
  mode: DocumentComparisonMode;
  author: string;
  date: string;
  sourceName: string;
}

export interface ComparisonIdentityFactory {
  create(
    kind: WorkDocumentChangeKind,
    before?: string,
  ): WorkDocumentChangeIdentity;
  paragraphIdentity(channel: string): { paragraphId: string; textId: string };
  summary: DocumentComparisonSummary;
}

export function createComparisonIdentityFactory(
  currentSignature: string,
  revisedSignature: string,
  options: DocumentComparisonOptions,
): ComparisonIdentityFactory {
  const seed = stableHash(
    `${options.mode}\u0000${currentSignature}\u0000${revisedSignature}`,
  );
  const author = boundedMetadata(options.author, 256) || 'A3S Work user';
  const date = normalizedDate(options.date);
  const summary = emptyComparisonSummary();
  let sequence = 0;
  return {
    summary,
    create(kind) {
      sequence += 1;
      if (kind === 'insertion') summary.insertions += 1;
      else if (kind === 'deletion') summary.deletions += 1;
      else if (kind === 'formatting') summary.formatting += 1;
      else if (kind === 'move') summary.moves = (summary.moves ?? 0) + 1;
      else summary.paragraphFormatting += 1;
      return {
        id: `compare-${seed}-${kind}-${sequence.toString(36)}`,
        author,
        date,
      };
    },
    paragraphIdentity(channel) {
      return {
        paragraphId: wordParagraphId(`${seed}:${channel}:paragraph`),
        textId: wordParagraphId(`${seed}:${channel}:text`),
      };
    },
  };
}

export function summarizeComparisonChanges(
  changes: readonly WorkDocumentChange[],
): DocumentComparisonSummary {
  const summary = emptyComparisonSummary();
  for (const change of changes) {
    if (change.kind === 'insertion') summary.insertions += 1;
    else if (change.kind === 'deletion') summary.deletions += 1;
    else if (change.kind === 'formatting') summary.formatting += 1;
    else if (change.kind === 'move') summary.moves = (summary.moves ?? 0) + 1;
    else summary.paragraphFormatting += 1;
  }
  return summary;
}

export function emptyComparisonSummary(): DocumentComparisonSummary {
  return {
    deletions: 0,
    formatting: 0,
    insertions: 0,
    paragraphFormatting: 0,
  };
}
