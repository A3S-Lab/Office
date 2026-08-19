import {
  type OfficeKernelLayoutBreak,
  type OfficeKernelLayoutPage,
  type OfficeKernelLayoutRequest,
  type OfficeKernelLayoutResult,
  type OfficeKernelPageMetrics,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

const MAX_LAYOUT_BLOCKS = 200_000;
const MAX_LAYOUT_PAGE_STYLES = 10_000;
const MAX_LAYOUT_EXTENT = 1_000_000;
const MAX_LAYOUT_PAGE_INDEX = 1_000_000;

export function layoutOfficeDocumentInJavaScript(
  request: OfficeKernelLayoutRequest,
): OfficeKernelLayoutResult {
  validateLayoutRequest(request);
  const pageStyles = new Map(
    (request.pageStyles ?? []).map(({ id, page }) => [id, page] as const),
  );
  const firstPage = pageMetricsForBlock(
    request.blocks[0],
    request.page,
    pageStyles,
  );
  const pages: OfficeKernelLayoutPage[] = [
    emptyPage(request.startPageIndex, firstPage),
  ];

  let index = 0;
  while (index < request.blocks.length) {
    const block = request.blocks[index];
    const page = pageMetricsForBlock(block, request.page, pageStyles);
    ensurePageMetrics(pages, page);
    if (block.flowCount !== undefined) {
      const end = index + block.flowCount;
      layoutFlow(
        request.blocks,
        index,
        end,
        pages,
        page,
        request.page,
        pageStyles,
        end < request.blocks.length,
      );
      index = end;
    } else {
      layoutSingleBlock(
        block,
        request.blocks,
        index + 1,
        pages,
        page,
        request.page,
        pageStyles,
        index + 1 < request.blocks.length,
      );
      index += 1;
    }
  }

  if (pages.length > 1 && pages.at(-1)?.placements.length === 0) pages.pop();

  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'layoutResult',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    startPageIndex: request.startPageIndex,
    engine: 'javascript',
    pages,
    breaks: layoutBreaks(pages),
  };
}

function validateLayoutRequest(request: OfficeKernelLayoutRequest): void {
  if (request.protocol !== OFFICE_KERNEL_PROTOCOL_VERSION) {
    throw kernelError(
      'office.kernel.protocol_unsupported',
      `Office kernel protocol ${request.protocol} is unsupported.`,
    );
  }
  if (request.kind !== 'layout') {
    throw kernelError(
      'office.kernel.request_kind_invalid',
      'The Office kernel only accepts layout requests at this boundary.',
    );
  }
  if (
    !Number.isSafeInteger(request.requestId) ||
    request.requestId < 0 ||
    !Number.isSafeInteger(request.revision) ||
    request.revision < 0 ||
    !Number.isSafeInteger(request.documentRevision) ||
    request.documentRevision < 0 ||
    !Number.isSafeInteger(request.startPageIndex) ||
    request.startPageIndex < 0
  ) {
    throw kernelError(
      'office.kernel.revision_invalid',
      'Request, layout, and document revisions must be non-negative safe integers.',
    );
  }
  if (request.startPageIndex > MAX_LAYOUT_PAGE_INDEX) {
    throw kernelError(
      'office.kernel.page_index_invalid',
      `startPageIndex may not exceed ${MAX_LAYOUT_PAGE_INDEX}.`,
    );
  }
  if (request.blocks.length > MAX_LAYOUT_BLOCKS) {
    throw kernelError(
      'office.kernel.block_limit_exceeded',
      `A layout request may contain at most ${MAX_LAYOUT_BLOCKS} blocks.`,
    );
  }
  validatePageMetrics('page', request.page);
  const pageStyles = request.pageStyles ?? [];
  if (pageStyles.length > MAX_LAYOUT_PAGE_STYLES) {
    throw kernelError(
      'office.kernel.page_style_limit_exceeded',
      `A layout request may contain at most ${MAX_LAYOUT_PAGE_STYLES} page styles.`,
    );
  }
  const pageStyleIds = new Set<string>();
  for (const style of pageStyles) {
    if (!style.id.trim() || style.id.length > 256) {
      throw kernelError(
        'office.kernel.page_style_id_invalid',
        'Every page style requires a non-empty ID of at most 256 bytes.',
      );
    }
    if (pageStyleIds.has(style.id)) {
      throw kernelError(
        'office.kernel.page_style_id_duplicate',
        `Page style ID '${style.id}' is duplicated.`,
      );
    }
    pageStyleIds.add(style.id);
    validatePageMetrics(`pageStyles.${style.id}.page`, style.page);
  }
  const blockIds = new Set<string>();
  for (const block of request.blocks) {
    if (!block.id.trim() || block.id.length > 256) {
      throw kernelError(
        'office.kernel.block_id_invalid',
        'Every layout block requires a non-empty ID of at most 256 bytes.',
      );
    }
    if (blockIds.has(block.id)) {
      throw kernelError(
        'office.kernel.block_id_duplicate',
        `Layout block ID '${block.id}' is duplicated.`,
      );
    }
    blockIds.add(block.id);
    if (
      block.pageStyleId !== undefined &&
      (!block.pageStyleId.trim() ||
        block.pageStyleId.length > 256 ||
        !pageStyleIds.has(block.pageStyleId))
    ) {
      throw kernelError(
        'office.kernel.page_style_reference_invalid',
        `Layout block '${block.id}' references an unknown page style.`,
      );
    }
    validateExtent('block.height', block.height);
    validateFlowMetadata(block);
  }
  validateFlowSequences(request.blocks);
}

function layoutSingleBlock(
  block: OfficeKernelLayoutRequest['blocks'][number],
  blocks: OfficeKernelLayoutRequest['blocks'],
  nextIndex: number,
  pages: OfficeKernelLayoutPage[],
  page: OfficeKernelPageMetrics,
  defaultPage: OfficeKernelPageMetrics,
  pageStyles: ReadonlyMap<string, OfficeKernelPageMetrics>,
  hasMoreBlocks: boolean,
): void {
  const availableHeight = pageBodyHeight(page);
  let current = pages.at(-1) as OfficeKernelLayoutPage;
  let currentHasContent = current.placements.length > 0;
  if (block.breakBefore && currentHasContent) {
    current = emptyPage(nextPageIndex(pages), page);
    pages.push(current);
    currentHasContent = false;
  }

  const remaining = Math.max(0, availableHeight - current.usedHeight);
  const nextHeight = nextBlockPreviewHeight(
    blocks,
    nextIndex,
    page,
    defaultPage,
    pageStyles,
  );
  const groupedHeight = block.height + (block.keepWithNext ? nextHeight : 0);
  const shouldAdvance =
    currentHasContent &&
    (block.height > remaining ||
      ((block.keepTogether || block.keepWithNext) &&
        groupedHeight <= availableHeight &&
        groupedHeight > remaining));
  if (shouldAdvance) {
    current = emptyPage(nextPageIndex(pages), page);
    pages.push(current);
  }

  placeFragment(block, current, availableHeight);
  if (block.breakAfter && hasMoreBlocks) {
    pages.push(emptyPage(nextPageIndex(pages), page));
  }
}

function layoutFlow(
  blocks: OfficeKernelLayoutRequest['blocks'],
  start: number,
  end: number,
  pages: OfficeKernelLayoutPage[],
  page: OfficeKernelPageMetrics,
  defaultPage: OfficeKernelPageMetrics,
  pageStyles: ReadonlyMap<string, OfficeKernelPageMetrics>,
  hasMoreBlocks: boolean,
): void {
  const availableHeight = pageBodyHeight(page);
  const first = blocks[start];
  const last = blocks[end - 1] as OfficeKernelLayoutRequest['blocks'][number];
  if (first.breakBefore && pages.at(-1)?.placements.length) {
    pages.push(emptyPage(nextPageIndex(pages), page));
  }

  const totalHeight = fragmentHeight(blocks, start, end);
  const nextHeight = last.keepWithNext
    ? nextBlockPreviewHeight(blocks, end, page, defaultPage, pageStyles)
    : 0;
  let current = pages.at(-1) as OfficeKernelLayoutPage;
  const repeatHeaderCount = first.repeatHeaderCount ?? 0;
  const repeatHeaderHeight = first.repeatHeaderHeight ?? 0;
  const groupedHeight = totalHeight + nextHeight;
  if (
    first.keepTogether &&
    current.placements.length &&
    groupedHeight <= availableHeight &&
    groupedHeight > Math.max(0, availableHeight - current.usedHeight)
  ) {
    current = emptyPage(nextPageIndex(pages), page);
    pages.push(current);
  }
  if (repeatHeaderCount > 0 && repeatHeaderCount < end - start) {
    const leadingHeight = fragmentHeight(
      blocks,
      start,
      start + repeatHeaderCount + 1,
    );
    const remainingHeight = Math.max(0, availableHeight - current.usedHeight);
    if (
      current.placements.length > 0 &&
      leadingHeight <= availableHeight &&
      leadingHeight > remainingHeight
    ) {
      current = emptyPage(nextPageIndex(pages), page);
      pages.push(current);
    }
  }

  const minimum = Math.max(1, first.minimumFragmentsPerPage ?? 1);
  let cursor = start;
  let remainingFlowHeight = totalHeight;
  while (cursor < end) {
    current = pages.at(-1) as OfficeKernelLayoutPage;
    if (
      cursor - start >= repeatHeaderCount &&
      cursor > start &&
      repeatHeaderHeight > 0 &&
      current.placements.length === 0 &&
      current.usedHeight === 0
    ) {
      current.usedHeight = repeatHeaderHeight;
    }
    const currentHasContent = current.placements.length > 0;
    const remainingHeight = Math.max(0, availableHeight - current.usedHeight);
    const remainingFragments = end - cursor;
    if (
      nextHeight > 0 &&
      currentHasContent &&
      remainingFlowHeight <= remainingHeight &&
      remainingFlowHeight + nextHeight > remainingHeight &&
      remainingFlowHeight + nextHeight <= availableHeight
    ) {
      pages.push(emptyPage(nextPageIndex(pages), page));
      continue;
    }
    let fitting = fragmentsFitting(blocks, cursor, end, remainingHeight);

    if (fitting === remainingFragments) {
      placeFragments(blocks, cursor, end, current, availableHeight);
      cursor = end;
      remainingFlowHeight = 0;
      continue;
    }
    if (fitting === 0) {
      if (currentHasContent) {
        pages.push(emptyPage(nextPageIndex(pages), page));
        continue;
      }
      fitting = 1;
    }
    if (remainingFragments > minimum) {
      fitting = Math.min(fitting, remainingFragments - minimum);
    }
    const minimumHere = Math.min(minimum, remainingFragments);
    if (fitting < minimumHere && currentHasContent) {
      pages.push(emptyPage(nextPageIndex(pages), page));
      continue;
    }
    fitting = Math.max(1, fitting);
    remainingFlowHeight -= fragmentHeight(blocks, cursor, cursor + fitting);
    placeFragments(blocks, cursor, cursor + fitting, current, availableHeight);
    cursor += fitting;
    if (cursor < end) {
      pages.push(emptyPage(nextPageIndex(pages), page));
    }
  }

  if (last.breakAfter && hasMoreBlocks) {
    pages.push(emptyPage(nextPageIndex(pages), page));
  }
}

function nextBlockPreviewHeight(
  blocks: OfficeKernelLayoutRequest['blocks'],
  start: number,
  currentPage: OfficeKernelPageMetrics,
  defaultPage: OfficeKernelPageMetrics,
  pageStyles: ReadonlyMap<string, OfficeKernelPageMetrics>,
): number {
  const first = blocks[start];
  if (!first || first.breakBefore) return 0;
  if (
    !pageMetricsEqual(
      currentPage,
      pageMetricsForBlock(first, defaultPage, pageStyles),
    )
  ) {
    return 0;
  }
  const count = Math.min(2, first.flowCount ?? 1);
  return fragmentHeight(blocks, start, Math.min(blocks.length, start + count));
}

function fragmentsFitting(
  blocks: OfficeKernelLayoutRequest['blocks'],
  start: number,
  end: number,
  availableHeight: number,
): number {
  let used = 0;
  let count = 0;
  for (let index = start; index < end; index += 1) {
    const block = blocks[index];
    if (used + block.height > availableHeight) break;
    used += block.height;
    count += 1;
  }
  return count;
}

function placeFragments(
  blocks: OfficeKernelLayoutRequest['blocks'],
  start: number,
  end: number,
  page: OfficeKernelLayoutPage,
  availableHeight: number,
): void {
  for (let index = start; index < end; index += 1) {
    placeFragment(blocks[index], page, availableHeight);
  }
}

function fragmentHeight(
  blocks: OfficeKernelLayoutRequest['blocks'],
  start: number,
  end: number,
): number {
  let height = 0;
  for (let index = start; index < end; index += 1) {
    height += blocks[index].height;
  }
  return height;
}

function placeFragment(
  block: OfficeKernelLayoutRequest['blocks'][number],
  page: OfficeKernelLayoutPage,
  availableHeight: number,
): void {
  const y = page.usedHeight;
  page.placements.push({
    blockId: block.id,
    y,
    height: block.height,
    overflow:
      block.height > availableHeight || y + block.height > availableHeight,
  });
  page.usedHeight += block.height;
}

function validateFlowMetadata(
  block: OfficeKernelLayoutRequest['blocks'][number],
): void {
  const values = [
    block.flowId,
    block.flowIndex,
    block.flowCount,
    block.minimumFragmentsPerPage,
  ];
  const repeatValues = [block.repeatHeaderCount, block.repeatHeaderHeight];
  if (
    values.every((value) => value === undefined) &&
    repeatValues.every((value) => value === undefined)
  ) {
    return;
  }
  if (values.some((value) => value === undefined)) {
    throw kernelError(
      'office.kernel.flow_metadata_incomplete',
      'Flow metadata requires flowId, flowIndex, flowCount, and minimumFragmentsPerPage.',
    );
  }
  if (
    typeof block.flowId !== 'string' ||
    !block.flowId.trim() ||
    block.flowId.length > 256
  ) {
    throw kernelError(
      'office.kernel.flow_id_invalid',
      'Every layout flow requires a non-empty ID of at most 256 bytes.',
    );
  }
  if (
    !Number.isSafeInteger(block.flowIndex) ||
    !Number.isSafeInteger(block.flowCount) ||
    (block.flowCount as number) <= 0 ||
    (block.flowIndex as number) < 0 ||
    (block.flowIndex as number) >= (block.flowCount as number)
  ) {
    throw kernelError(
      'office.kernel.flow_index_invalid',
      'Flow indices must be within a non-empty flow.',
    );
  }
  if (
    !Number.isSafeInteger(block.minimumFragmentsPerPage) ||
    (block.minimumFragmentsPerPage as number) <= 0 ||
    (block.minimumFragmentsPerPage as number) > (block.flowCount as number)
  ) {
    throw kernelError(
      'office.kernel.flow_minimum_invalid',
      'minimumFragmentsPerPage must be within the flow fragment count.',
    );
  }
  if (repeatValues.every((value) => value === undefined)) return;
  if (repeatValues.some((value) => value === undefined)) {
    throw kernelError(
      'office.kernel.repeat_header_metadata_incomplete',
      'Repeated table headers require repeatHeaderCount and repeatHeaderHeight.',
    );
  }
  if (
    !Number.isSafeInteger(block.repeatHeaderCount) ||
    (block.repeatHeaderCount as number) <= 0 ||
    (block.repeatHeaderCount as number) >= (block.flowCount as number)
  ) {
    throw kernelError(
      'office.kernel.repeat_header_count_invalid',
      'repeatHeaderCount must identify a non-empty header before the body rows.',
    );
  }
  validateExtent(
    'block.repeatHeaderHeight',
    block.repeatHeaderHeight as number,
  );
  if ((block.repeatHeaderHeight as number) <= 0) {
    throw kernelError(
      'office.kernel.repeat_header_height_invalid',
      'repeatHeaderHeight must be greater than zero.',
    );
  }
}

function validateFlowSequences(
  blocks: OfficeKernelLayoutRequest['blocks'],
): void {
  const flowIds = new Set<string>();
  let index = 0;
  while (index < blocks.length) {
    const block = blocks[index];
    if (block.flowId === undefined) {
      index += 1;
      continue;
    }
    if (block.flowIndex !== 0) {
      throw kernelError(
        'office.kernel.flow_sequence_invalid',
        'A layout flow must begin with fragment index zero.',
      );
    }
    if (flowIds.has(block.flowId)) {
      throw kernelError(
        'office.kernel.flow_id_duplicate',
        `Layout flow ID '${block.flowId}' is duplicated.`,
      );
    }
    flowIds.add(block.flowId);
    const count = block.flowCount as number;
    const repeatHeaderCount = block.repeatHeaderCount;
    const repeatHeaderHeight = block.repeatHeaderHeight;
    if (index + count > blocks.length) {
      throw kernelError(
        'office.kernel.flow_sequence_invalid',
        'A layout flow must contain its declared number of fragments.',
      );
    }
    for (let flowIndex = 0; flowIndex < count; flowIndex += 1) {
      const fragment = blocks[index + flowIndex];
      if (
        fragment.flowId !== block.flowId ||
        fragment.flowIndex !== flowIndex ||
        fragment.flowCount !== count ||
        fragment.minimumFragmentsPerPage !== block.minimumFragmentsPerPage ||
        fragment.repeatHeaderCount !== repeatHeaderCount ||
        fragment.repeatHeaderHeight !== repeatHeaderHeight ||
        fragment.pageStyleId !== block.pageStyleId
      ) {
        throw kernelError(
          'office.kernel.flow_sequence_invalid',
          'Layout flow fragments must be consecutive and consistently indexed.',
        );
      }
      if (flowIndex > 0 && fragment.breakBefore) {
        throw kernelError(
          'office.kernel.flow_break_invalid',
          'Only the first flow fragment may request breakBefore.',
        );
      }
      if (flowIndex + 1 < count && fragment.breakAfter) {
        throw kernelError(
          'office.kernel.flow_break_invalid',
          'Only the last flow fragment may request breakAfter.',
        );
      }
    }
    index += count;
  }
}

function emptyPage(
  index: number,
  page: OfficeKernelPageMetrics,
): OfficeKernelLayoutPage {
  return {
    index,
    page: { ...page },
    usedHeight: 0,
    availableHeight: pageBodyHeight(page),
    placements: [],
  };
}

function nextPageIndex(pages: readonly OfficeKernelLayoutPage[]): number {
  return (pages.at(-1)?.index ?? -1) + 1;
}

function layoutBreaks(
  pages: OfficeKernelLayoutPage[],
): OfficeKernelLayoutBreak[] {
  return pages.slice(1).flatMap((page, localIndex) => {
    const beforeBlockId = page.placements[0]?.blockId;
    const previous = pages[localIndex];
    if (!beforeBlockId || !previous) return [];
    const remainingBodyHeight = Math.max(
      0,
      previous.availableHeight - previous.usedHeight,
    );
    return [
      {
        beforeBlockId,
        pageIndex: page.index,
        remainingBodyHeight,
        spacerHeight:
          remainingBodyHeight +
          previous.page.marginBottom +
          previous.page.pageGap +
          page.page.marginTop,
      },
    ];
  });
}

function ensurePageMetrics(
  pages: OfficeKernelLayoutPage[],
  page: OfficeKernelPageMetrics,
): void {
  const current = pages.at(-1) as OfficeKernelLayoutPage;
  if (pageMetricsEqual(current.page, page)) return;
  if (current.placements.length === 0 && current.usedHeight === 0) {
    current.page = { ...page };
    current.availableHeight = pageBodyHeight(page);
    return;
  }
  pages.push(emptyPage(nextPageIndex(pages), page));
}

function pageMetricsForBlock(
  block: OfficeKernelLayoutRequest['blocks'][number] | undefined,
  defaultPage: OfficeKernelPageMetrics,
  pageStyles: ReadonlyMap<string, OfficeKernelPageMetrics>,
): OfficeKernelPageMetrics {
  return block?.pageStyleId
    ? (pageStyles.get(block.pageStyleId) ?? defaultPage)
    : defaultPage;
}

function pageBodyHeight(page: OfficeKernelPageMetrics): number {
  return Math.max(1, page.height - page.marginTop - page.marginBottom);
}

function pageMetricsEqual(
  left: OfficeKernelPageMetrics,
  right: OfficeKernelPageMetrics,
): boolean {
  return (
    left.width === right.width &&
    left.height === right.height &&
    left.marginTop === right.marginTop &&
    left.marginRight === right.marginRight &&
    left.marginBottom === right.marginBottom &&
    left.marginLeft === right.marginLeft &&
    left.headerHeight === right.headerHeight &&
    left.footerHeight === right.footerHeight &&
    left.pageGap === right.pageGap
  );
}

function validatePageMetrics(
  name: string,
  page: OfficeKernelPageMetrics,
): void {
  for (const [property, value] of Object.entries(page)) {
    validateExtent(`${name}.${property}`, value);
  }
  if (page.width <= page.marginLeft + page.marginRight) {
    throw kernelError(
      'office.kernel.page_width_invalid',
      'Page width must be greater than its horizontal margins.',
    );
  }
  if (page.height <= page.marginTop + page.marginBottom) {
    throw kernelError(
      'office.kernel.page_height_invalid',
      'Page height must leave a positive body area.',
    );
  }
}

function validateExtent(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0 || value > MAX_LAYOUT_EXTENT) {
    throw kernelError(
      'office.kernel.extent_invalid',
      `${name} must be a finite non-negative number.`,
    );
  }
}

function kernelError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}
