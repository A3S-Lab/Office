import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type {
  OfficeKernelPageMetrics,
  OfficeKernelTextLayoutParagraphResult,
} from '../../kernel/office-kernel-protocol';
import { millimetersToPixels } from './work-document-layout';
import { measureParagraphLineFragments } from './work-document-line-measurement';
import {
  documentBlockId,
  documentInlineOffsets,
  documentListChildId,
  documentNodeParagraphPagination,
  documentPaginationLayoutKey,
  elementForNode,
  isDocumentListNode,
  measuredDocumentBlock,
  outerHeight,
  reusableDocumentLayoutBlocks,
  shouldKeepDocumentBlockTogether,
  verticalBlockEnd,
  verticalBlockStart,
} from './work-document-pagination-dom';
import type {
  DocumentPaginationSection,
  DocumentPaginationSnapshot,
  MeasuredDocumentLayoutBlock,
} from './work-document-pagination-types';
import {
  documentSectionLayoutFromNodeAttributes,
  type DocumentSectionNodeAttributes,
} from './work-document-section';
import { measureDocumentTableRows } from './work-document-table-pagination';
import type { WorkDocumentSectionLayout } from './work-types';

export function measureDocumentLayoutBlocks(
  editor: Editor,
  previous: DocumentPaginationSnapshot | null = null,
  dirtyFrom = 0,
  textLayouts: ReadonlyMap<
    string,
    OfficeKernelTextLayoutParagraphResult
  > = new Map(),
  maximumFragmentedTableRowHeight = 1_000_000,
): DocumentPaginationSnapshot {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  const counts = { measured: 0, reused: 0 };
  let unsupportedLayout = false;
  let firstLayoutKey: string | null = null;
  editor.state.doc.forEach((section, sectionPosition, sectionIndex) => {
    if (section.type.name !== 'documentSection') return;
    const layout = documentSectionLayoutFromNodeAttributes(
      section.attrs as Partial<DocumentSectionNodeAttributes>,
    );
    const layoutKey = documentPaginationLayoutKey(layout);
    firstLayoutKey ??= layoutKey;
    if (layout.columns.count > 1 || layoutKey !== firstLayoutKey) {
      unsupportedLayout = true;
    }
    measureSectionBlocks(
      editor,
      section,
      sectionPosition,
      sectionIndex,
      layout,
      previous?.blocks ?? [],
      dirtyFrom,
      textLayouts,
      maximumFragmentedTableRowHeight,
      counts,
      blocks,
    );
  });
  return {
    blocks,
    measuredBlockCount: counts.measured,
    reusedBlockCount: counts.reused,
    unsupportedLayout,
  };
}

export function documentPageMetrics(
  layout: WorkDocumentSectionLayout,
): OfficeKernelPageMetrics {
  const portrait =
    layout.pageSize === 'letter'
      ? { width: 816, height: 1056 }
      : {
          width: millimetersToPixels(210),
          height: millimetersToPixels(297),
        };
  const landscape = layout.orientation === 'landscape';
  return {
    width: landscape ? portrait.height : portrait.width,
    height: landscape ? portrait.width : portrait.height,
    marginTop: millimetersToPixels(layout.margins.top),
    marginRight: millimetersToPixels(layout.margins.right),
    marginBottom: millimetersToPixels(layout.margins.bottom),
    marginLeft: millimetersToPixels(layout.margins.left),
    headerHeight: 0,
    footerHeight: 0,
    pageGap: 28,
  };
}

export function documentPageChromeHeights(
  editor: Editor,
): Pick<OfficeKernelPageMetrics, 'headerHeight' | 'footerHeight'> {
  const page = editor.view.dom.closest<HTMLElement>('.work-document-page');
  return {
    headerHeight: outerHeight(
      page?.querySelector<HTMLElement>('.work-document-page-header') ?? null,
    ),
    footerHeight: outerHeight(
      page?.querySelector<HTMLElement>('.work-document-page-footer') ?? null,
    ),
  };
}

function measureSectionBlocks(
  editor: Editor,
  section: ProseMirrorNode,
  sectionPosition: number,
  sectionIndex: number,
  layout: WorkDocumentSectionLayout,
  previous: readonly MeasuredDocumentLayoutBlock[],
  dirtyFrom: number,
  textLayouts: ReadonlyMap<string, OfficeKernelTextLayoutParagraphResult>,
  maximumFragmentedTableRowHeight: number,
  counts: { measured: number; reused: number },
  result: MeasuredDocumentLayoutBlock[],
): void {
  const sectionBlocks: MeasuredDocumentLayoutBlock[] = [];
  section.forEach((node, offset, index) => {
    const position = sectionPosition + offset + 1;
    const element = elementForNode(editor, position);
    if (!element) return;
    const id = documentBlockId(sectionPosition, index, position);
    const paragraphPagination = documentNodeParagraphPagination(node);
    if (isDocumentListNode(node)) {
      const reused = reusableDocumentListLayoutBlocks(
        previous,
        id,
        position + node.nodeSize,
        dirtyFrom,
      );
      if (reused.length) {
        counts.reused += reused.length;
        sectionBlocks.push(...reused);
        return;
      }
      const listBlocks = measureDocumentListBlocks(
        editor,
        node,
        element,
        id,
        position,
        textLayouts,
        maximumFragmentedTableRowHeight,
      );
      if (listBlocks.length) {
        counts.measured += listBlocks.length;
        sectionBlocks.push(...listBlocks);
        return;
      }
    }
    const reused = reusableDocumentLayoutBlocks(
      previous,
      id,
      element,
      position + node.nodeSize,
      dirtyFrom,
    );
    if (reused.length) {
      counts.reused += reused.length;
      sectionBlocks.push(...reused);
      return;
    }
    if (node.type.name === 'table') {
      const tableRows = measureDocumentTableRows(
        editor,
        node,
        element,
        id,
        position,
        maximumFragmentedTableRowHeight,
      );
      if (tableRows.length) {
        counts.measured += tableRows.length;
        sectionBlocks.push(...tableRows);
        return;
      }
    }
    const lineFragments = measureParagraphLineFragments(
      editor,
      node,
      element,
      id,
      position,
      position + node.nodeSize,
      paragraphPagination,
      textLayouts.get(id),
    );
    if (lineFragments.length > 1) {
      counts.measured += lineFragments.length;
      sectionBlocks.push(...lineFragments);
      return;
    }
    counts.measured += 1;
    sectionBlocks.push(
      measuredDocumentBlock({
        block: {
          id,
          height: Math.max(1, outerHeight(element)),
          breakBefore: paragraphPagination.pageBreakBefore,
          breakAfter: node.type.name === 'pageBreak',
          keepTogether:
            paragraphPagination.keepLines ||
            shouldKeepDocumentBlockTogether(node),
          keepWithNext: paragraphPagination.keepWithNext,
        },
        element,
        from: position,
        to: position + node.nodeSize,
      }),
    );
  });

  const last = sectionBlocks.at(-1);
  if (
    last &&
    layout.breakAfter !== 'continuous' &&
    layout.breakAfter !== 'nextColumn'
  ) {
    last.block.breakAfter = true;
  }
  const sectionMetadata: DocumentPaginationSection = {
    id:
      typeof section.attrs.id === 'string' && section.attrs.id
        ? section.attrs.id
        : `document-section-${sectionIndex + 1}`,
    index: sectionIndex,
    position: sectionPosition,
    layout,
  };
  for (const block of sectionBlocks) block.section = sectionMetadata;
  result.push(...sectionBlocks);
}

function measureDocumentListBlocks(
  editor: Editor,
  list: ProseMirrorNode,
  listElement: HTMLElement,
  listId: string,
  listPosition: number,
  textLayouts: ReadonlyMap<string, OfficeKernelTextLayoutParagraphResult>,
  maximumFragmentedTableRowHeight: number,
): MeasuredDocumentLayoutBlock[] {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  let complete = true;
  list.forEach((item, itemOffset, itemIndex) => {
    if (item.type.name !== 'listItem') {
      complete = false;
      return;
    }
    const itemPosition = listPosition + itemOffset + 1;
    const itemElement = elementForNode(editor, itemPosition);
    if (!itemElement || itemElement.tagName.toLowerCase() !== 'li') {
      complete = false;
      return;
    }
    const itemBlocks: MeasuredDocumentLayoutBlock[] = [];
    item.forEach((node, offset, index) => {
      const position = itemPosition + offset + 1;
      const element = elementForNode(editor, position);
      if (!element) {
        complete = false;
        return;
      }
      const nestedList = isDocumentListNode(node);
      const id = documentListChildId(
        listId,
        itemIndex,
        node,
        index,
        nestedList,
      );
      if (nestedList) {
        const nestedBlocks = measureDocumentListBlocks(
          editor,
          node,
          element,
          id,
          position,
          textLayouts,
          maximumFragmentedTableRowHeight,
        );
        if (!nestedBlocks.length) {
          complete = false;
          return;
        }
        itemBlocks.push(...nestedBlocks);
        return;
      }
      itemBlocks.push(
        ...measureDocumentListItemBlock(
          editor,
          node,
          element,
          id,
          position,
          textLayouts.get(id),
          maximumFragmentedTableRowHeight,
        ),
      );
    });
    if (!itemBlocks.length) {
      complete = false;
      return;
    }
    fitDocumentBlocksToContainer(itemBlocks, itemElement);
    blocks.push(...itemBlocks);
  });
  if (!complete || !blocks.length) return [];
  fitDocumentBlocksToContainer(blocks, listElement);
  return blocks;
}

function measureDocumentListItemBlock(
  editor: Editor,
  node: ProseMirrorNode,
  element: HTMLElement,
  id: string,
  position: number,
  shapedLayout: OfficeKernelTextLayoutParagraphResult | undefined,
  maximumFragmentedTableRowHeight: number,
): MeasuredDocumentLayoutBlock[] {
  if (node.type.name === 'table') {
    const rows = measureDocumentTableRows(
      editor,
      node,
      element,
      id,
      position,
      maximumFragmentedTableRowHeight,
    );
    if (rows.length) return rows;
  }
  const pagination = documentNodeParagraphPagination(node);
  const lineFragments = measureParagraphLineFragments(
    editor,
    node,
    element,
    id,
    position,
    position + node.nodeSize,
    pagination,
    shapedLayout,
  );
  const { inlineOffsetLeft, inlineOffsetRight } = documentInlineOffsets(
    editor,
    element,
  );
  if (lineFragments.length > 1) {
    return lineFragments.map((fragment) => ({
      ...fragment,
      inlineOffsetLeft,
      inlineOffsetRight,
    }));
  }
  return [
    {
      ...measuredDocumentBlock({
        block: {
          id,
          height: Math.max(1, outerHeight(element)),
          breakBefore: pagination.pageBreakBefore,
          breakAfter: node.type.name === 'pageBreak',
          keepTogether:
            pagination.keepLines || shouldKeepDocumentBlockTogether(node),
          keepWithNext: pagination.keepWithNext,
        },
        element,
        from: position,
        to: position + node.nodeSize,
      }),
      inlineOffsetLeft,
      inlineOffsetRight,
    },
  ];
}

function fitDocumentBlocksToContainer(
  blocks: MeasuredDocumentLayoutBlock[],
  container: HTMLElement,
): void {
  const target = outerHeight(container);
  if (!(target > 0) || blocks.length === 0) return;
  const measured = blocks.reduce(
    (height, candidate) => height + candidate.block.height,
    0,
  );
  let difference = target - measured;
  if (Math.abs(difference) < 0.01) return;

  const first = blocks[0];
  const last = blocks.at(-1) as MeasuredDocumentLayoutBlock;
  if (difference > 0) {
    const style = getComputedStyle(container);
    const start = Math.min(difference, verticalBlockStart(style));
    first.block.height += start;
    difference -= start;
    const end = Math.min(difference, verticalBlockEnd(style));
    last.block.height += end;
    difference -= end;
    last.block.height += difference;
    return;
  }

  let remaining = -difference;
  const candidates = Array.from(
    new Set([last, first, ...blocks.slice(1, -1).reverse()]),
  );
  for (const candidate of candidates) {
    if (remaining <= 0) break;
    const removable = Math.max(0, candidate.block.height - 1);
    const amount = Math.min(removable, remaining);
    candidate.block.height -= amount;
    remaining -= amount;
  }
}

function reusableDocumentListLayoutBlocks(
  previous: readonly MeasuredDocumentLayoutBlock[],
  listId: string,
  nodeTo: number,
  dirtyFrom: number,
): MeasuredDocumentLayoutBlock[] {
  if (nodeTo > dirtyFrom) return [];
  const prefix = `${listId}-item-`;
  return previous
    .filter(
      ({ block }) =>
        block.id.startsWith(prefix) || block.flowId?.startsWith(prefix),
    )
    .map((candidate) => ({
      ...candidate,
      block: { ...candidate.block },
    }));
}
