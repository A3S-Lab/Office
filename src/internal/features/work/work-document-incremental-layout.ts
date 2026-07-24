import type {
  OfficeKernelLayoutBlock,
  OfficeKernelLayoutBreak,
  OfficeKernelLayoutPage,
  OfficeKernelLayoutResult,
  OfficeKernelPageMetrics,
} from '../../kernel/office-kernel-protocol';

export interface PositionedOfficeKernelLayoutBlock {
  block: OfficeKernelLayoutBlock;
  from: number;
  to: number;
}

export interface IncrementalDocumentLayoutPlan {
  blocks: OfficeKernelLayoutBlock[];
  reusedPageCount: number;
  startBlockIndex: number;
  startPageIndex: number;
}

export function planIncrementalDocumentLayout(
  previous: {
    blocks: readonly PositionedOfficeKernelLayoutBlock[];
    layout: OfficeKernelLayoutResult;
  } | null,
  next: readonly PositionedOfficeKernelLayoutBlock[],
  dirtyFrom: number,
): IncrementalDocumentLayoutPlan {
  const fullLayout = (): IncrementalDocumentLayoutPlan => ({
    blocks: next.map(({ block }) => block),
    reusedPageCount: 0,
    startBlockIndex: 0,
    startPageIndex: 0,
  });
  if (
    !previous ||
    !Number.isFinite(dirtyFrom) ||
    dirtyFrom <= 0 ||
    previous.layout.startPageIndex !== 0 ||
    previous.layout.pages.length < 3
  ) {
    return fullLayout();
  }

  const affectedBlockIndex = earliestAffectedBlockIndex(
    previous.blocks,
    next,
    dirtyFrom,
  );
  const pageByBlockId = new Map(
    previous.layout.pages.flatMap((page) =>
      page.placements.map(
        (placement) => [placement.blockId, page.index] as const,
      ),
    ),
  );
  const affectedPageIndex = [
    next[affectedBlockIndex]?.block.id,
    previous.blocks[affectedBlockIndex]?.block.id,
  ].reduce<number | null>((earliest, blockId) => {
    const pageIndex = blockId ? pageByBlockId.get(blockId) : undefined;
    if (pageIndex === undefined) return earliest;
    return earliest === null ? pageIndex : Math.min(earliest, pageIndex);
  }, null);
  if (affectedPageIndex === null || affectedPageIndex < 2) {
    return fullLayout();
  }

  const nextBlockIndex = new Map(
    next.map(({ block }, index) => [block.id, index] as const),
  );
  const previousBlockById = new Map(
    previous.blocks.map(({ block }) => [block.id, block] as const),
  );
  let startPageIndex = affectedPageIndex - 1;

  while (startPageIndex > 0) {
    const startBlockId: string | undefined =
      previous.layout.pages[startPageIndex]?.placements[0]?.blockId;
    const startBlockIndex =
      startBlockId === undefined ? undefined : nextBlockIndex.get(startBlockId);
    if (startBlockIndex === undefined || startBlockIndex <= 0) {
      startPageIndex -= 1;
      continue;
    }

    const startBlock = next[startBlockIndex]?.block;
    if (startBlock?.flowId && startBlock.flowIndex !== 0) {
      const flowStart = next.find(
        ({ block }) =>
          block.flowId === startBlock.flowId && block.flowIndex === 0,
      );
      const flowStartPage = flowStart
        ? pageByBlockId.get(flowStart.block.id)
        : undefined;
      startPageIndex =
        flowStartPage === undefined
          ? startPageIndex - 1
          : Math.max(0, flowStartPage - 1);
      continue;
    }

    const previousPageLastId =
      previous.layout.pages[startPageIndex - 1]?.placements.at(-1)?.blockId;
    const previousPageLast =
      previousPageLastId === undefined
        ? undefined
        : next[nextBlockIndex.get(previousPageLastId) ?? -1]?.block;
    if (
      !previousPageLast ||
      previousPageLast.keepWithNext ||
      (previousPageLast.flowId &&
        previousPageLast.flowId === startBlock?.flowId)
    ) {
      startPageIndex -= 1;
      continue;
    }

    if (
      !prefixMatches(
        previous.layout.pages,
        startPageIndex,
        previousBlockById,
        next,
        startBlockIndex,
      )
    ) {
      startPageIndex -= 1;
      continue;
    }

    return {
      blocks: next.slice(startBlockIndex).map(({ block }) => block),
      reusedPageCount: startPageIndex,
      startBlockIndex,
      startPageIndex,
    };
  }

  return fullLayout();
}

export function mergeIncrementalDocumentLayout(
  previous: OfficeKernelLayoutResult,
  partial: OfficeKernelLayoutResult,
  page: OfficeKernelPageMetrics,
): OfficeKernelLayoutResult {
  if (partial.startPageIndex === 0) return partial;
  const startPageIndex = partial.startPageIndex;
  const prefixPages = previous.pages.slice(0, startPageIndex);
  if (
    prefixPages.length !== startPageIndex ||
    partial.pages[0]?.index !== startPageIndex
  ) {
    throw new Error('Incremental document layout has an invalid page prefix.');
  }

  const prefixBreaks = previous.breaks.filter(
    (pageBreak) => pageBreak.pageIndex < startPageIndex,
  );
  const boundaryBreak = incrementalBoundaryBreak(
    prefixPages.at(-1),
    partial.pages[0],
    page,
  );
  return {
    ...partial,
    startPageIndex: 0,
    pages: [...prefixPages, ...partial.pages],
    breaks: [
      ...prefixBreaks,
      ...(boundaryBreak ? [boundaryBreak] : []),
      ...partial.breaks,
    ],
  };
}

function earliestAffectedBlockIndex(
  previous: readonly PositionedOfficeKernelLayoutBlock[],
  next: readonly PositionedOfficeKernelLayoutBlock[],
  dirtyFrom: number,
): number {
  const positionIndex = next.findIndex((candidate) => candidate.to > dirtyFrom);
  const sharedLength = Math.min(previous.length, next.length);
  let divergentIndex = -1;
  for (let index = 0; index < sharedLength; index += 1) {
    if (!sameLayoutBlock(previous[index].block, next[index].block)) {
      divergentIndex = index;
      break;
    }
  }
  if (divergentIndex < 0 && previous.length !== next.length) {
    divergentIndex = sharedLength;
  }
  const candidates = [positionIndex, divergentIndex].filter(
    (index) => index >= 0,
  );
  return Math.min(
    ...(candidates.length ? candidates : [Math.max(0, next.length - 1)]),
  );
}

function prefixMatches(
  pages: readonly OfficeKernelLayoutPage[],
  startPageIndex: number,
  previousBlockById: ReadonlyMap<string, OfficeKernelLayoutBlock>,
  next: readonly PositionedOfficeKernelLayoutBlock[],
  startBlockIndex: number,
): boolean {
  const prefixIds = pages
    .slice(0, startPageIndex)
    .flatMap((page) => page.placements.map((placement) => placement.blockId));
  if (prefixIds.length !== startBlockIndex) return false;
  return prefixIds.every((blockId, index) => {
    const previous = previousBlockById.get(blockId);
    const candidate = next[index]?.block;
    return (
      previous !== undefined &&
      candidate?.id === blockId &&
      sameLayoutBlock(previous, candidate)
    );
  });
}

function sameLayoutBlock(
  left: OfficeKernelLayoutBlock,
  right: OfficeKernelLayoutBlock,
): boolean {
  return (
    left.id === right.id &&
    left.height === right.height &&
    left.breakBefore === right.breakBefore &&
    left.breakAfter === right.breakAfter &&
    left.keepTogether === right.keepTogether &&
    left.keepWithNext === right.keepWithNext &&
    left.flowId === right.flowId &&
    left.flowIndex === right.flowIndex &&
    left.flowCount === right.flowCount &&
    left.minimumFragmentsPerPage === right.minimumFragmentsPerPage &&
    left.repeatHeaderCount === right.repeatHeaderCount &&
    left.repeatHeaderHeight === right.repeatHeaderHeight
  );
}

function incrementalBoundaryBreak(
  previous: OfficeKernelLayoutPage | undefined,
  next: OfficeKernelLayoutPage | undefined,
  page: OfficeKernelPageMetrics,
): OfficeKernelLayoutBreak | null {
  const beforeBlockId = next?.placements[0]?.blockId;
  if (!previous || !next || !beforeBlockId) return null;
  const remainingBodyHeight = Math.max(
    0,
    previous.availableHeight - previous.usedHeight,
  );
  return {
    beforeBlockId,
    pageIndex: next.index,
    remainingBodyHeight,
    spacerHeight:
      remainingBodyHeight +
      page.marginBottom +
      page.footerHeight +
      page.pageGap +
      page.marginTop +
      page.headerHeight,
  };
}
