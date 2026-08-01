import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

export interface DocumentTextMatch {
  from: number;
  to: number;
  before: string;
  matchedText: string;
  after: string;
  truncatedBefore: boolean;
  truncatedAfter: boolean;
}

interface DocumentTextRun {
  from: number;
  text: string;
}

const SEARCH_CONTEXT_LENGTH = 34;

export function documentTextMatches(
  document: ProseMirrorNode,
  rawQuery: string,
): DocumentTextMatch[] {
  const query = rawQuery.toLocaleLowerCase();
  if (!query) return [];

  return collectDocumentTextRuns(document).flatMap((run) => {
    const matches: DocumentTextMatch[] = [];
    const searchableText = run.text.toLocaleLowerCase();
    let offset = 0;
    while (offset <= searchableText.length - query.length) {
      const index = searchableText.indexOf(query, offset);
      if (index < 0) break;
      const end = index + query.length;
      const contextStart = Math.max(0, index - SEARCH_CONTEXT_LENGTH);
      const contextEnd = Math.min(run.text.length, end + SEARCH_CONTEXT_LENGTH);
      matches.push({
        from: run.from + index,
        to: run.from + end,
        before: normalizedSearchContext(
          run.text.slice(contextStart, index),
          true,
        ),
        matchedText: run.text.slice(index, end),
        after: normalizedSearchContext(run.text.slice(end, contextEnd), false),
        truncatedBefore: contextStart > 0,
        truncatedAfter: contextEnd < run.text.length,
      });
      offset = index + Math.max(1, query.length);
    }
    return matches;
  });
}

function collectDocumentTextRuns(document: ProseMirrorNode): DocumentTextRun[] {
  const runs: DocumentTextRun[] = [];
  document.descendants((node, position) => {
    if (!node.isText || !node.text) return;
    const previous = runs.at(-1);
    if (previous && previous.from + previous.text.length === position) {
      previous.text += node.text;
    } else {
      runs.push({ from: position, text: node.text });
    }
  });
  return runs;
}

function normalizedSearchContext(value: string, trailing: boolean): string {
  const normalized = value.replace(/\s+/g, ' ');
  return trailing ? normalized.trimStart() : normalized.trimEnd();
}
