import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { Transaction } from '@tiptap/pm/state';

export type DocumentParagraphBreakKind = 'merge' | 'split';

export interface DocumentParagraphBreakChangeSegment {
  id: string;
  kind: DocumentParagraphBreakKind;
  position: number;
  from: number;
  to: number;
}

export function clearDocumentParagraphBreakChangeAttributes(
  attributes: Record<string, unknown>,
): Record<string, unknown> {
  return {
    ...attributes,
    paragraphBreakChangeKind: null,
    paragraphBreakChangeId: '',
    paragraphBreakChangeAuthor: '',
    paragraphBreakChangeDate: '',
  };
}

export function documentParagraphBreakChangeSegments(
  document: ProseMirrorNode,
): DocumentParagraphBreakChangeSegment[] {
  const segments: DocumentParagraphBreakChangeSegment[] = [];
  document.descendants((node, position) => {
    if (
      (node.type.name !== 'paragraph' && node.type.name !== 'heading') ||
      (node.attrs.paragraphBreakChangeKind !== 'merge' &&
        node.attrs.paragraphBreakChangeKind !== 'split')
    ) {
      return;
    }
    segments.push({
      id:
        typeof node.attrs.paragraphBreakChangeId === 'string' &&
        node.attrs.paragraphBreakChangeId
          ? node.attrs.paragraphBreakChangeId
          : `paragraph-break-at-${position}`,
      kind: node.attrs.paragraphBreakChangeKind,
      position,
      from: position + 1,
      to: position + 1 + node.content.size,
    });
  });
  return segments;
}

/**
 * Accept merge / reject split joins this paragraph with its neighbor.
 * Accept split / reject merge only clears the break attrs.
 */
export function resolveDocumentParagraphBreakChange(
  tr: Transaction,
  segment: DocumentParagraphBreakChangeSegment,
  decision: 'accept' | 'reject',
): boolean {
  const position = tr.mapping.map(segment.position);
  const node = tr.doc.nodeAt(position);
  if (
    !node ||
    (node.type.name !== 'paragraph' && node.type.name !== 'heading')
  ) {
    return false;
  }
  const shouldJoin =
    (decision === 'accept' && segment.kind === 'merge') ||
    (decision === 'reject' && segment.kind === 'split');
  if (!shouldJoin) {
    tr.setNodeMarkup(
      position,
      undefined,
      clearDocumentParagraphBreakChangeAttributes(node.attrs),
    );
    return true;
  }
  if (segment.kind === 'merge') {
    return joinParagraphWithNext(tr, position, node);
  }
  return joinParagraphWithPrevious(tr, position, node);
}

function joinParagraphWithNext(
  tr: Transaction,
  position: number,
  node: ProseMirrorNode,
): boolean {
  const after = position + node.nodeSize;
  const next = tr.doc.nodeAt(after);
  if (
    !next ||
    (next.type.name !== 'paragraph' && next.type.name !== 'heading')
  ) {
    return false;
  }
  const joined = node.copy(node.content.append(next.content));
  tr.replaceWith(
    position,
    after + next.nodeSize,
    joined.type.create(
      clearDocumentParagraphBreakChangeAttributes(joined.attrs),
      joined.content,
      joined.marks,
    ),
  );
  return true;
}

function joinParagraphWithPrevious(
  tr: Transaction,
  position: number,
  node: ProseMirrorNode,
): boolean {
  const resolved = tr.doc.resolve(position);
  const previous = resolved.nodeBefore;
  if (
    !previous ||
    (previous.type.name !== 'paragraph' && previous.type.name !== 'heading')
  ) {
    return false;
  }
  const previousPos = position - previous.nodeSize;
  const joined = previous.copy(previous.content.append(node.content));
  tr.replaceWith(
    previousPos,
    position + node.nodeSize,
    joined.type.create(joined.attrs, joined.content, joined.marks),
  );
  return true;
}

export function paragraphBreakChangeAttribute(
  field: 'kind' | 'id' | 'author' | 'date',
) {
  const modelName = `paragraphBreakChange${field[0]?.toUpperCase() ?? ''}${field.slice(
    1,
  )}`;
  const htmlName =
    field === 'kind'
      ? 'data-paragraph-break-kind'
      : `data-paragraph-break-${field}`;
  return {
    default: field === 'kind' ? null : '',
    parseHTML: (element: HTMLElement) => {
      if (element.getAttribute('data-paragraph-break-change') !== 'true') {
        return field === 'kind' ? null : '';
      }
      const value = element.getAttribute(htmlName) ?? '';
      if (field === 'kind') {
        return value === 'merge' || value === 'split' ? value : null;
      }
      return value;
    },
    renderHTML: (attributes: Record<string, unknown>) => {
      const kind = attributes.paragraphBreakChangeKind;
      if (kind !== 'merge' && kind !== 'split') return {};
      if (field === 'kind') {
        return {
          'data-paragraph-break-change': 'true',
          'data-paragraph-break-kind': kind,
        };
      }
      const value =
        typeof attributes[modelName] === 'string' ? attributes[modelName] : '';
      return value ? { [htmlName]: value } : {};
    },
  };
}
