import { describe, expect, test } from '@rstest/core';
import {
  mergeIncrementalDocumentLayout,
  planIncrementalDocumentLayout,
  type PositionedOfficeKernelLayoutBlock,
} from '../src/internal/features/work/work-document-incremental-layout';
import { layoutOfficeDocumentInJavaScript } from '../src/internal/kernel/office-kernel-fallback';
import {
  type OfficeKernelLayoutBlock,
  type OfficeKernelLayoutRequest,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from '../src/internal/kernel/office-kernel-protocol';

describe('incremental document layout', () => {
  test('reuses complete stable pages before a tail edit', () => {
    const previousBlocks = positionedBlocks([
      block('a', 900),
      block('b', 900),
      block('c', 900),
      block('d', 900),
      block('e', 900),
    ]);
    const previousLayout = layoutOfficeDocumentInJavaScript(
      request(previousBlocks.map(({ block }) => block)),
    );
    const nextBlocks = previousBlocks.map((candidate) => ({
      ...candidate,
      block:
        candidate.block.id === 'e'
          ? { ...candidate.block, height: 880 }
          : candidate.block,
    }));

    const plan = planIncrementalDocumentLayout(
      { blocks: previousBlocks, layout: previousLayout },
      nextBlocks,
      nextBlocks[4].from,
    );

    expect(plan).toMatchObject({
      reusedPageCount: 3,
      startBlockIndex: 3,
      startPageIndex: 3,
    });
    expect(plan.blocks.map((candidate) => candidate.id)).toEqual(['d', 'e']);
  });

  test('rewinds another page when keep-with-next crosses the prefix boundary', () => {
    const previousBlocks = positionedBlocks([
      block('a', 900),
      block('b', 900),
      { ...block('c', 900), keepWithNext: true },
      block('d', 900),
      block('e', 900),
    ]);
    const previousLayout = layoutOfficeDocumentInJavaScript(
      request(previousBlocks.map(({ block }) => block)),
    );

    const plan = planIncrementalDocumentLayout(
      { blocks: previousBlocks, layout: previousLayout },
      previousBlocks,
      previousBlocks[4].from,
    );

    expect(plan.startPageIndex).toBe(2);
    expect(plan.blocks[0].id).toBe('c');
  });

  test('merges reused pages with absolute page indices and a fresh boundary', () => {
    const previousBlocks = positionedBlocks([
      block('a', 900),
      block('b', 900),
      block('c', 900),
      block('d', 900),
      block('e', 900),
    ]);
    const previousLayout = layoutOfficeDocumentInJavaScript(
      request(previousBlocks.map(({ block }) => block)),
    );
    const partial = layoutOfficeDocumentInJavaScript({
      ...request([block('d', 900), block('e', 880)]),
      documentRevision: 5,
      startPageIndex: 3,
    });

    const merged = mergeIncrementalDocumentLayout(
      previousLayout,
      partial,
      partialPageMetrics(),
    );

    expect(merged.startPageIndex).toBe(0);
    expect(merged.documentRevision).toBe(5);
    expect(merged.pages.map((page) => page.index)).toEqual([0, 1, 2, 3, 4]);
    expect(merged.pages.map((page) => page.placements[0]?.blockId)).toEqual([
      'a',
      'b',
      'c',
      'd',
      'e',
    ]);
    expect(merged.breaks.map((pageBreak) => pageBreak.pageIndex)).toEqual([
      1, 2, 3, 4,
    ]);
    expect(merged.breaks[2].beforeBlockId).toBe('d');
  });
});

function positionedBlocks(
  blocks: OfficeKernelLayoutBlock[],
): PositionedOfficeKernelLayoutBlock[] {
  return blocks.map((block, index) => ({
    block,
    from: index * 10 + 1,
    to: index * 10 + 10,
  }));
}

function request(blocks: OfficeKernelLayoutBlock[]): OfficeKernelLayoutRequest {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'layout',
    requestId: 7,
    revision: 12,
    documentRevision: 4,
    startPageIndex: 0,
    page: partialPageMetrics(),
    blocks,
  };
}

function partialPageMetrics(): OfficeKernelLayoutRequest['page'] {
  return {
    width: 794,
    height: 1123,
    marginTop: 80,
    marginRight: 80,
    marginBottom: 80,
    marginLeft: 80,
    headerHeight: 0,
    footerHeight: 0,
    pageGap: 24,
  };
}

function block(id: string, height: number): OfficeKernelLayoutBlock {
  return { id, height };
}
