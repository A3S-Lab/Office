import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { OfficeKernelLayoutBlock } from '../../kernel/office-kernel-protocol';
import type { MeasuredDocumentLayoutBlock } from './work-document-pagination-types';
import type { WorkDocumentSectionLayout } from './work-types';

export function reusableDocumentLayoutBlocks(
  previous: readonly MeasuredDocumentLayoutBlock[],
  blockId: string,
  element: HTMLElement,
  nodeTo: number,
  dirtyFrom: number,
): MeasuredDocumentLayoutBlock[] {
  if (nodeTo > dirtyFrom) return [];
  return previous
    .filter(
      (candidate) =>
        candidate.block.id === blockId || candidate.block.flowId === blockId,
    )
    .map((candidate) => ({
      ...candidate,
      block: { ...candidate.block },
      element,
    }));
}

export function isDocumentListNode(node: ProseMirrorNode): boolean {
  return node.type.name === 'bulletList' || node.type.name === 'orderedList';
}

export function documentListChildId(
  listId: string,
  itemIndex: number,
  node: ProseMirrorNode,
  childIndex: number,
  nestedList = isDocumentListNode(node),
): string {
  const kind = nestedList ? 'list' : node.type.name;
  return `${listId}-item-${itemIndex}-${kind}-${childIndex}`;
}

export function verticalBlockStart(style: CSSStyleDeclaration): number {
  return (
    nonNegativePixels(style.marginTop) +
    nonNegativePixels(style.paddingTop) +
    nonNegativePixels(style.borderTopWidth)
  );
}

export function verticalBlockEnd(style: CSSStyleDeclaration): number {
  return (
    nonNegativePixels(style.marginBottom) +
    nonNegativePixels(style.paddingBottom) +
    nonNegativePixels(style.borderBottomWidth)
  );
}

export function documentInlineOffsets(
  editor: Editor,
  element: HTMLElement,
): { inlineOffsetLeft: number; inlineOffsetRight: number } {
  const elementRect = element.getBoundingClientRect();
  const editorRect = editor.view.dom.getBoundingClientRect();
  const scaleX =
    editor.view.dom.offsetWidth > 0
      ? editorRect.width / editor.view.dom.offsetWidth
      : 1;
  return {
    inlineOffsetLeft: Math.max(
      0,
      (elementRect.left - editorRect.left) / Math.max(scaleX, 0.01),
    ),
    inlineOffsetRight: Math.max(
      0,
      (editorRect.right - elementRect.right) / Math.max(scaleX, 0.01),
    ),
  };
}

export interface DocumentNodeParagraphPagination {
  keepLines: boolean;
  keepWithNext: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
}

export function documentNodeParagraphPagination(
  node: ProseMirrorNode,
): DocumentNodeParagraphPagination {
  return {
    keepLines: directNodeBoolean(node, 'keepLines') ?? false,
    keepWithNext:
      directNodeBoolean(node, 'keepWithNext') ?? node.type.name === 'heading',
    pageBreakBefore: directNodeBoolean(node, 'pageBreakBefore') ?? false,
    widowControl: directNodeBoolean(node, 'widowControl') ?? true,
  };
}

function directNodeBoolean(
  node: ProseMirrorNode,
  attribute: 'keepLines' | 'keepWithNext' | 'pageBreakBefore' | 'widowControl',
): boolean | null {
  const value: unknown = node.attrs[attribute];
  return typeof value === 'boolean' ? value : null;
}

export function measuredDocumentBlock({
  block,
  element,
  from,
  to,
}: {
  block: OfficeKernelLayoutBlock;
  element: HTMLElement;
  from: number;
  to: number;
}): MeasuredDocumentLayoutBlock {
  return {
    block,
    element,
    from,
    to,
    inlineOffsetLeft: 0,
    inlineOffsetRight: 0,
    observeResize: true,
  };
}

export function shouldKeepDocumentBlockTogether(
  node: ProseMirrorNode,
): boolean {
  return (
    node.type.name === 'table' ||
    node.type.name === 'blockquote' ||
    node.type.name === 'codeBlock' ||
    node.type.name === 'image' ||
    node.type.name === 'documentNote'
  );
}

export function documentPaginationLayoutKey(
  layout: WorkDocumentSectionLayout,
): string {
  return JSON.stringify({
    margins: layout.margins,
    orientation: layout.orientation,
    pageSize: layout.pageSize,
  });
}

export function documentBlockId(
  sectionPosition: number,
  index: number,
  position: number,
): string {
  return `section-${sectionPosition}-block-${index}-${position}`;
}

export function outerHeight(element: HTMLElement | null): number {
  if (!element) return 0;
  const style = getComputedStyle(element);
  return (
    element.offsetHeight +
    nonNegativePixels(style.marginTop) +
    nonNegativePixels(style.marginBottom)
  );
}

export function nonNegativePixels(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? Math.max(0, parsed) : 0;
}

export function positivePixels(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

export function finitePixels(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function elementForNode(
  editor: Editor,
  position: number,
): HTMLElement | null {
  const node = editor.view.nodeDOM(position);
  if (!node || node.nodeType !== Node.ELEMENT_NODE) return null;
  return node as HTMLElement;
}
