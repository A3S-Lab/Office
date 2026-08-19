import type { WorkDocumentNode } from './work-types';

const SECTION_CLOSE = '</section>';
const TABLE_BODY_OPEN = '<tbody>';
const TABLE_CLOSE = '</tbody></table>';

export type DocumentLazyHtmlTablePart =
  | 'complete'
  | 'first'
  | 'middle'
  | 'last';

export interface DocumentLazyHtmlRange {
  from: number;
  id: string;
  tablePart: DocumentLazyHtmlTablePart | null;
  to: number;
}

export interface DocumentLazyHtmlProjection {
  html: string;
  orderedRanges: DocumentLazyHtmlRange[];
  ranges: ReadonlyMap<string, DocumentLazyHtmlRange>;
}

/**
 * Indexes canonical HTML without copying its chunk strings. This projection is
 * available only for the parser-authenticated paragraph/table subset.
 */
export function createDocumentLazyHtmlProjection(
  html: string,
  root: WorkDocumentNode,
): DocumentLazyHtmlProjection | null {
  if (!html.startsWith('<section') || !html.endsWith(SECTION_CLOSE)) {
    return null;
  }
  const sectionOpenEnd = html.indexOf('>') + 1;
  if (sectionOpenEnd <= 0) return null;
  let cursor = sectionOpenEnd;
  const orderedRanges: DocumentLazyHtmlRange[] = [];
  const ranges = new Map<string, DocumentLazyHtmlRange>();

  for (const chunk of lazyDocumentLeafChunks(root)) {
    const id = typeof chunk.attrs?.id === 'string' ? chunk.attrs.id : '';
    if (!id || ranges.has(id)) return null;
    const from = cursor;
    let tablePart: DocumentLazyHtmlTablePart | null = null;
    const content = chunk.content ?? [];
    for (const node of content) {
      const scanned = scanSimpleDocumentNodeHtml(html, cursor, node);
      if (!scanned) return null;
      cursor = scanned.to;
      if (scanned.tablePart) {
        if (content.length !== 1 || tablePart) return null;
        tablePart = scanned.tablePart;
      }
    }
    const range = { from, id, tablePart, to: cursor };
    orderedRanges.push(range);
    ranges.set(id, range);
  }

  if (cursor !== html.length - SECTION_CLOSE.length) return null;
  return { html, orderedRanges, ranges };
}

export function patchDocumentLazyHtmlProjection(
  projection: DocumentLazyHtmlProjection,
  replacements: ReadonlyMap<string, string>,
): string | null {
  if (!replacements.size) return projection.html;
  for (const id of replacements.keys()) {
    if (!projection.ranges.has(id)) return null;
  }

  const parts: string[] = [];
  let cursor = 0;
  for (const range of projection.orderedRanges) {
    const replacement = replacements.get(range.id);
    if (replacement === undefined) continue;
    parts.push(projection.html.slice(cursor, range.from), replacement);
    cursor = range.to;
  }
  parts.push(projection.html.slice(cursor));
  const html = parts.join('');

  let offset = 0;
  for (const range of projection.orderedRanges) {
    const previousFrom = range.from;
    const previousLength = range.to - previousFrom;
    const replacement = replacements.get(range.id);
    const nextLength = replacement?.length ?? previousLength;
    range.from = previousFrom + offset;
    range.to = range.from + nextLength;
    offset += nextLength - previousLength;
  }
  projection.html = html;
  return html;
}

export function documentLazyHtmlChunkFragment(
  html: string,
  tablePart: DocumentLazyHtmlTablePart | null,
): string | null {
  if (!tablePart || tablePart === 'complete') return html;
  const bodyOpen = html.indexOf(TABLE_BODY_OPEN);
  const bodyClose = html.lastIndexOf('</tbody>');
  if (bodyOpen < 0 || bodyClose < bodyOpen) return null;
  const rowsFrom = bodyOpen + TABLE_BODY_OPEN.length;
  if (tablePart === 'first') return html.slice(0, bodyClose);
  if (tablePart === 'middle') return html.slice(rowsFrom, bodyClose);
  return html.slice(rowsFrom);
}

function lazyDocumentLeafChunks(root: WorkDocumentNode): WorkDocumentNode[] {
  const chunks: WorkDocumentNode[] = [];
  const pending = [root];
  while (pending.length) {
    const node = pending.pop();
    if (!node) continue;
    if (node.type === 'documentChunk' && node.attrs?.windowContainer !== true) {
      chunks.push(node);
      continue;
    }
    for (let index = (node.content?.length ?? 0) - 1; index >= 0; index -= 1) {
      const child = node.content?.[index];
      if (child) pending.push(child);
    }
  }
  return chunks;
}

function scanSimpleDocumentNodeHtml(
  html: string,
  from: number,
  node: WorkDocumentNode,
): { tablePart: DocumentLazyHtmlTablePart | null; to: number } | null {
  if (node.type === 'paragraph') {
    if (!html.startsWith('<p>', from)) return null;
    const close = html.indexOf('</p>', from + 3);
    return close < 0 ? null : { tablePart: null, to: close + 4 };
  }
  if (node.type !== 'table') return null;

  const virtualIndex = Number(node.attrs?.virtualTableIndex);
  const virtualCount = Number(node.attrs?.virtualTableCount);
  const virtual = Boolean(
    typeof node.attrs?.virtualTableId === 'string' &&
      node.attrs.virtualTableId &&
      Number.isSafeInteger(virtualIndex) &&
      Number.isSafeInteger(virtualCount) &&
      virtualIndex >= 0 &&
      virtualCount > 0 &&
      virtualIndex < virtualCount,
  );
  let cursor = from;
  if (!virtual || virtualIndex === 0) {
    if (!html.startsWith('<table', cursor)) return null;
    const body = html.indexOf(TABLE_BODY_OPEN, cursor);
    if (body < 0) return null;
    cursor = body + TABLE_BODY_OPEN.length;
  }
  for (const row of node.content ?? []) {
    if (row.type !== 'tableRow' || !html.startsWith('<tr>', cursor)) {
      return null;
    }
    const close = html.indexOf('</tr>', cursor + 4);
    if (close < 0) return null;
    cursor = close + 5;
  }
  if (!virtual || virtualIndex === virtualCount - 1) {
    if (!html.startsWith(TABLE_CLOSE, cursor)) return null;
    cursor += TABLE_CLOSE.length;
  }
  const tablePart = !virtual
    ? 'complete'
    : virtualCount === 1
      ? 'complete'
      : virtualIndex === 0
        ? 'first'
        : virtualIndex === virtualCount - 1
          ? 'last'
          : 'middle';
  return { tablePart, to: cursor };
}
