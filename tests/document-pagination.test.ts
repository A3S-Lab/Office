import { expect, test } from '@rstest/core';
import {
  createDocumentTableRowFragmentPlan,
  createDocumentLineFragments,
  createShapedDocumentLineFragments,
  findDocumentLineStartOffset,
  reusableDocumentLayoutBlocks,
} from '../src/internal/features/work/work-document-pagination';
import type { MeasuredDocumentLayoutBlock } from '../src/internal/features/work/work-document-pagination';
import {
  documentPaginationPageDescriptors,
  pageForPosition,
  type DocumentPaginationResult,
} from '../src/internal/features/work/editors/use-document-pagination';
import {
  type OfficeKernelLayoutResult,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from '../src/internal/kernel/office-kernel-protocol';
import type { WorkDocumentSectionLayout } from '../src/internal/features/work/work-types';

test('converts measured browser lines into contiguous layout fragments', () => {
  expect(
    createDocumentLineFragments(
      [
        { top: 100, position: 10 },
        { top: 120, position: 20 },
        { top: 140, position: 30 },
      ],
      90,
      165,
    ),
  ).toEqual([
    { from: 10, to: 20, height: 30 },
    { from: 20, to: 30, height: 20 },
    { from: 30, to: 31, height: 25 },
  ]);
});

test('converts shaped UTF-16 line offsets into document fragments', () => {
  expect(
    createShapedDocumentLineFragments(
      [
        { startUtf16: 0, endUtf16: 4, height: 21 },
        { startUtf16: 4, endUtf16: 8, height: 21 },
        { startUtf16: 8, endUtf16: 10, height: 21 },
      ],
      12,
      24,
      6,
      9,
    ),
  ).toEqual([
    { from: 12, to: 17, height: 27 },
    { from: 17, to: 21, height: 21 },
    { from: 21, to: 24, height: 30 },
  ]);
});

test('rejects overlapping shaped line offsets', () => {
  expect(
    createShapedDocumentLineFragments(
      [
        { startUtf16: 0, endUtf16: 5, height: 20 },
        { startUtf16: 4, endUtf16: 8, height: 20 },
      ],
      2,
      12,
      0,
      0,
    ),
  ).toEqual([]);
});

test('normalizes visual measurements produced under editor zoom', () => {
  expect(
    createDocumentLineFragments(
      [
        { top: 50, position: 2 },
        { top: 60, position: 8 },
      ],
      45,
      72.5,
      0.5,
    ),
  ).toEqual([
    { from: 2, to: 8, height: 30 },
    { from: 8, to: 9, height: 25 },
  ]);
});

test('finds a visual line start without viewport hit testing', () => {
  const lineHeight = 20;
  const charactersPerLine = 10;

  expect(
    findDocumentLineStartOffset(100, { top: -160, bottom: -140 }, (offset) => {
      const line = Math.floor(offset / charactersPerLine);
      const top = -200 + line * lineHeight;
      return { top, bottom: top + 14 };
    }),
  ).toBe(20);
});

test('reuses stable measured blocks before the earliest document change', () => {
  const previousElement = document.createElement('p');
  const currentElement = document.createElement('p');
  const previous: MeasuredDocumentLayoutBlock[] = [
    {
      block: {
        id: 'section-0-block-0-1-line-0',
        flowId: 'section-0-block-0-1',
        flowIndex: 0,
        flowCount: 2,
        height: 20,
      },
      element: previousElement,
      from: 1,
      to: 8,
      inlineOffsetLeft: 0,
      inlineOffsetRight: 0,
      observeResize: false,
    },
    {
      block: {
        id: 'section-0-block-0-1-line-1',
        flowId: 'section-0-block-0-1',
        flowIndex: 1,
        flowCount: 2,
        height: 20,
      },
      element: previousElement,
      from: 8,
      to: 16,
      inlineOffsetLeft: 10,
      inlineOffsetRight: 5,
      observeResize: false,
    },
  ];

  const reused = reusableDocumentLayoutBlocks(
    previous,
    'section-0-block-0-1',
    currentElement,
    16,
    24,
  );

  expect(reused).toHaveLength(2);
  expect(reused[0]).not.toBe(previous[0]);
  expect(reused[0].block).not.toBe(previous[0].block);
  expect(reused.every((block) => block.element === currentElement)).toBe(true);
  expect(
    reusableDocumentLayoutBlocks(
      previous,
      'section-0-block-0-1',
      currentElement,
      16,
      8,
    ),
  ).toEqual([]);
});

test('aligns safe table-cell boundaries into row fragments', () => {
  const fragments = createDocumentTableRowFragmentPlan(
    [
      {
        cellIndex: 0,
        from: 10,
        to: 30,
        boundaries: [
          { position: 10, y: 5 },
          { position: 20, y: 40 },
          { position: 30, y: 80 },
        ],
      },
      {
        cellIndex: 1,
        from: 40,
        to: 60,
        boundaries: [
          { position: 40, y: 5 },
          { position: 50, y: 55 },
          { position: 60, y: 80 },
        ],
      },
    ],
    80,
    100,
  );

  expect(fragments.map((fragment) => fragment.height)).toEqual([40, 15, 25]);
  expect(fragments[1]?.cellBreaks).toEqual([
    { cellIndex: 0, position: 20, alignmentOffset: 0 },
    { cellIndex: 1, position: 40, alignmentOffset: 35 },
  ]);
  expect(fragments[2]?.cellBreaks).toEqual([
    { cellIndex: 0, position: 20, alignmentOffset: 15 },
    { cellIndex: 1, position: 50, alignmentOffset: 0 },
  ]);
  expect(fragments[2]?.cellRanges).toEqual([
    { from: 20, to: 30 },
    { from: 50, to: 60 },
  ]);
});

test('keeps rows larger than the continuation page atomic', () => {
  expect(
    createDocumentTableRowFragmentPlan(
      [
        {
          cellIndex: 0,
          from: 10,
          to: 30,
          boundaries: [
            { position: 10, y: 0 },
            { position: 20, y: 60 },
            { position: 30, y: 120 },
          ],
        },
      ],
      120,
      100,
    ),
  ).toEqual([
    {
      height: 120,
      cellRanges: [{ from: 10, to: 30 }],
    },
  ]);
});

test('maps a selection in a fragmented table row to its visual page', () => {
  const element = document.createElement('table');
  const blocks: MeasuredDocumentLayoutBlock[] = [
    {
      block: { id: 'row-fragment-0', height: 40 },
      element,
      from: 10,
      to: 60,
      inlineOffsetLeft: 0,
      inlineOffsetRight: 0,
      observeResize: false,
      selectionRanges: [
        { from: 10, to: 20 },
        { from: 40, to: 50 },
      ],
    },
    {
      block: { id: 'row-fragment-1', height: 40 },
      element,
      from: 10,
      to: 60,
      inlineOffsetLeft: 0,
      inlineOffsetRight: 0,
      observeResize: false,
      selectionRanges: [
        { from: 20, to: 30 },
        { from: 50, to: 60 },
      ],
    },
  ];
  const pagination: DocumentPaginationResult = {
    layout: {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'layoutResult',
      requestId: 1,
      revision: 1,
      documentRevision: 1,
      startPageIndex: 0,
      engine: 'javascript',
      pages: [],
      breaks: [],
    },
    blocks,
    pages: [],
    pageByBlockId: new Map([
      ['row-fragment-0', 1],
      ['row-fragment-1', 2],
    ]),
  };

  expect(pageForPosition(pagination, 15)).toBe(1);
  expect(pageForPosition(pagination, 20)).toBe(2);
  expect(pageForPosition(pagination, 55)).toBe(2);
});

test('resolves page chrome from physical and section-aware pagination', () => {
  const firstSection = sectionLayout({
    pageNumberStart: 3,
    pageChrome: {
      differentFirstPage: true,
      differentOddEvenPages: true,
      default: chrome('Default'),
      first: chrome('First'),
      even: chrome('Even'),
    },
  });
  const secondSection = sectionLayout({
    pageNumberStart: 10,
    pageChrome: {
      differentFirstPage: true,
      differentOddEvenPages: false,
      default: chrome('Second default'),
      first: chrome('Second first'),
      even: chrome('Second even'),
    },
  });
  const element = document.createElement('p');
  const blocks: MeasuredDocumentLayoutBlock[] = [
    measuredBlock('a', 'section-a', 0, firstSection, element),
    measuredBlock('b', 'section-a', 0, firstSection, element),
    measuredBlock('c', 'section-a', 0, firstSection, element),
    measuredBlock('d', 'section-b', 1, secondSection, element),
  ];
  const layout = paginationLayout(['a', 'b', 'c', 'd']);

  expect(
    documentPaginationPageDescriptors(layout, blocks).map((page) => ({
      pageNumber: page.pageNumber,
      physicalPage: page.physicalPage,
      sectionPage: page.sectionPage,
      sectionId: page.sectionId,
      variant: page.pageChrome.variant,
      headerHtml: page.pageChrome.headerHtml,
    })),
  ).toEqual([
    {
      pageNumber: 3,
      physicalPage: 1,
      sectionPage: 1,
      sectionId: 'section-a',
      variant: 'first',
      headerHtml: '<p>First</p>',
    },
    {
      pageNumber: 4,
      physicalPage: 2,
      sectionPage: 2,
      sectionId: 'section-a',
      variant: 'even',
      headerHtml: '<p>Even</p>',
    },
    {
      pageNumber: 5,
      physicalPage: 3,
      sectionPage: 3,
      sectionId: 'section-a',
      variant: 'default',
      headerHtml: '<p>Default</p>',
    },
    {
      pageNumber: 10,
      physicalPage: 4,
      sectionPage: 1,
      sectionId: 'section-b',
      variant: 'first',
      headerHtml: '<p>Second first</p>',
    },
  ]);
});

test('counts a continuous section on every physical page it occupies', () => {
  const firstSection = sectionLayout();
  const secondSection = sectionLayout();
  const element = document.createElement('p');
  const blocks = [
    measuredBlock('a', 'section-a', 0, firstSection, element),
    measuredBlock('b', 'section-b', 1, secondSection, element),
    measuredBlock('c', 'section-b', 1, secondSection, element),
  ];
  const layout = paginationLayout([['a', 'b'], ['c']]);

  const pages = documentPaginationPageDescriptors(layout, blocks);

  expect(pages[0]).toMatchObject({
    sectionId: 'section-a',
    sectionPage: 1,
  });
  expect(pages[1]).toMatchObject({
    sectionId: 'section-b',
    sectionPage: 2,
  });
});

function sectionLayout(
  patch: Partial<WorkDocumentSectionLayout> = {},
): WorkDocumentSectionLayout {
  return {
    pageSize: 'a4',
    orientation: 'portrait',
    margins: { top: 25, right: 23, bottom: 25, left: 23 },
    columns: { count: 1, spacing: 12, separator: false },
    breakAfter: 'nextPage',
    ...patch,
  };
}

function chrome(label: string) {
  return {
    headerHtml: `<p>${label}</p>`,
    footerHtml: '',
    showPageNumber: true,
  };
}

function measuredBlock(
  id: string,
  sectionId: string,
  sectionIndex: number,
  layout: WorkDocumentSectionLayout,
  element: HTMLElement,
): MeasuredDocumentLayoutBlock {
  return {
    block: { id, height: 20 },
    element,
    from: sectionIndex * 100 + 1,
    to: sectionIndex * 100 + 20,
    inlineOffsetLeft: 0,
    inlineOffsetRight: 0,
    observeResize: false,
    section: {
      id: sectionId,
      index: sectionIndex,
      position: sectionIndex * 100,
      layout,
    },
  };
}

function paginationLayout(
  pageBlocks: Array<string | string[]>,
): OfficeKernelLayoutResult {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'layoutResult',
    requestId: 1,
    revision: 1,
    documentRevision: 1,
    startPageIndex: 0,
    engine: 'javascript',
    pages: pageBlocks.map((source, index) => ({
      index,
      usedHeight: 20,
      availableHeight: 900,
      placements: (Array.isArray(source) ? source : [source]).map(
        (blockId, placementIndex) => ({
          blockId,
          y: placementIndex * 20,
          height: 20,
          overflow: false,
        }),
      ),
    })),
    breaks: [],
  };
}
