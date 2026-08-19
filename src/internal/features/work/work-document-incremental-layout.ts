import type {
  OfficeKernelLayoutBlock,
  OfficeKernelLayoutBreak,
  OfficeKernelLayoutPage,
  OfficeKernelLayoutResult,
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
  reusedPrefixBlockCount = 0,
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

  const stablePrefixBlockCount = Math.max(
    0,
    Math.min(
      Math.trunc(reusedPrefixBlockCount),
      previous.blocks.length,
      next.length,
    ),
  );
  const affectedBlockIndex = earliestAffectedBlockIndex(
    previous.blocks,
    next,
    dirtyFrom,
    stablePrefixBlockCount,
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
        previous.blocks,
        next,
        startBlockIndex,
        stablePrefixBlockCount,
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
  stablePrefixBlockCount: number,
): number {
  let lower = 0;
  let upper = next.length;
  while (lower < upper) {
    const middle = Math.floor((lower + upper) / 2);
    if (next[middle].to <= dirtyFrom) lower = middle + 1;
    else upper = middle;
  }
  const positionIndex = lower < next.length ? lower : -1;
  const sharedLength = Math.min(previous.length, next.length);
  let divergentIndex =
    stablePrefixBlockCount > 0 && stablePrefixBlockCount < sharedLength
      ? stablePrefixBlockCount
      : -1;
  if (stablePrefixBlockCount === 0) {
    for (let index = 0; index < sharedLength; index += 1) {
      if (!sameLayoutBlock(previous[index].block, next[index].block)) {
        divergentIndex = index;
        break;
      }
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
  previous: readonly PositionedOfficeKernelLayoutBlock[],
  next: readonly PositionedOfficeKernelLayoutBlock[],
  startBlockIndex: number,
  stablePrefixBlockCount: number,
): boolean {
  let blockIndex = 0;
  for (let pageIndex = 0; pageIndex < startPageIndex; pageIndex += 1) {
    for (const placement of pages[pageIndex]?.placements ?? []) {
      if (blockIndex >= startBlockIndex) return false;
      if (blockIndex >= stablePrefixBlockCount) {
        const previousBlock = previous[blockIndex]?.block;
        const nextBlock = next[blockIndex]?.block;
        if (
          previousBlock?.id !== placement.blockId ||
          nextBlock?.id !== placement.blockId ||
          !sameLayoutBlock(previousBlock, nextBlock)
        ) {
          return false;
        }
      }
      blockIndex += 1;
    }
  }
  return blockIndex === startBlockIndex;
}

function sameLayoutBlock(
  left: OfficeKernelLayoutBlock,
  right: OfficeKernelLayoutBlock,
): boolean {
  return (
    left.id === right.id &&
    left.height === right.height &&
    left.pageStyleId === right.pageStyleId &&
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
      previous.page.marginBottom +
      previous.page.pageGap +
      next.page.marginTop,
  };
}
