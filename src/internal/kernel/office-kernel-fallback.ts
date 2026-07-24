import {
  type OfficeKernelLayoutBreak,
  type OfficeKernelLayoutPage,
  type OfficeKernelLayoutRequest,
  type OfficeKernelLayoutResult,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

const MAX_LAYOUT_BLOCKS = 10_000;
const MAX_LAYOUT_EXTENT = 1_000_000;
const MAX_LAYOUT_PAGE_INDEX = 1_000_000;

export function layoutOfficeDocumentInJavaScript(
  request: OfficeKernelLayoutRequest,
): OfficeKernelLayoutResult {
  validateLayoutRequest(request);
  const availableHeight = Math.max(
    1,
    request.page.height -
      request.page.marginTop -
      request.page.marginBottom -
      request.page.headerHeight -
      request.page.footerHeight,
  );
  const pages: OfficeKernelLayoutPage[] = [
    emptyPage(request.startPageIndex, availableHeight),
  ];

  let index = 0;
  while (index < request.blocks.length) {
    const block = request.blocks[index];
    if (block.flowCount !== undefined) {
      const end = index + block.flowCount;
      layoutFlow(
        request.blocks.slice(index, end),
        request.blocks.slice(end),
        pages,
        availableHeight,
        end < request.blocks.length,
      );
      index = end;
    } else {
      layoutSingleBlock(
        block,
        request.blocks.slice(index + 1),
        pages,
        availableHeight,
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
    breaks: layoutBreaks(pages, request),
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
  for (const [name, value] of Object.entries(request.page)) {
    validateExtent(`page.${name}`, value);
  }
  if (
    request.page.width <=
    request.page.marginLeft + request.page.marginRight
  ) {
    throw kernelError(
      'office.kernel.page_width_invalid',
      'Page width must be greater than its horizontal margins.',
    );
  }
  if (
    request.page.height <=
    request.page.marginTop +
      request.page.marginBottom +
      request.page.headerHeight +
      request.page.footerHeight
  ) {
    throw kernelError(
      'office.kernel.page_height_invalid',
      'Page height must leave a positive body area.',
    );
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
    validateExtent('block.height', block.height);
    validateFlowMetadata(block);
  }
  validateFlowSequences(request.blocks);
}

function layoutSingleBlock(
  block: OfficeKernelLayoutRequest['blocks'][number],
  next: OfficeKernelLayoutRequest['blocks'],
  pages: OfficeKernelLayoutPage[],
  availableHeight: number,
  hasMoreBlocks: boolean,
): void {
  let current = pages.at(-1) as OfficeKernelLayoutPage;
  let currentHasContent = current.placements.length > 0;
  if (block.breakBefore && currentHasContent) {
    current = emptyPage(nextPageIndex(pages), availableHeight);
    pages.push(current);
    currentHasContent = false;
  }

  const remaining = Math.max(0, availableHeight - current.usedHeight);
  const nextHeight = nextBlockPreviewHeight(next);
  const groupedHeight = block.height + (block.keepWithNext ? nextHeight : 0);
  const shouldAdvance =
    currentHasContent &&
    (block.height > remaining ||
      ((block.keepTogether || block.keepWithNext) &&
        groupedHeight <= availableHeight &&
        groupedHeight > remaining));
  if (shouldAdvance) {
    current = emptyPage(nextPageIndex(pages), availableHeight);
    pages.push(current);
  }

  placeFragment(block, current, availableHeight);
  if (block.breakAfter && hasMoreBlocks) {
    pages.push(emptyPage(nextPageIndex(pages), availableHeight));
  }
}

function layoutFlow(
  blocks: OfficeKernelLayoutRequest['blocks'],
  next: OfficeKernelLayoutRequest['blocks'],
  pages: OfficeKernelLayoutPage[],
  availableHeight: number,
  hasMoreBlocks: boolean,
): void {
  const first = blocks[0];
  const last = blocks.at(-1) as OfficeKernelLayoutRequest['blocks'][number];
  if (first.breakBefore && pages.at(-1)?.placements.length) {
    pages.push(emptyPage(nextPageIndex(pages), availableHeight));
  }

  const totalHeight = fragmentHeight(blocks, 0, blocks.length);
  const nextHeight = last.keepWithNext ? nextBlockPreviewHeight(next) : 0;
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
    current = emptyPage(nextPageIndex(pages), availableHeight);
    pages.push(current);
  }
  if (repeatHeaderCount > 0 && repeatHeaderCount < blocks.length) {
    const leadingHeight = fragmentHeight(blocks, 0, repeatHeaderCount + 1);
    const remainingHeight = Math.max(0, availableHeight - current.usedHeight);
    if (
      current.placements.length > 0 &&
      leadingHeight <= availableHeight &&
      leadingHeight > remainingHeight
    ) {
      current = emptyPage(nextPageIndex(pages), availableHeight);
      pages.push(current);
    }
  }

  const minimum = Math.max(1, first.minimumFragmentsPerPage ?? 1);
  let cursor = 0;
  while (cursor < blocks.length) {
    current = pages.at(-1) as OfficeKernelLayoutPage;
    if (
      cursor >= repeatHeaderCount &&
      cursor > 0 &&
      repeatHeaderHeight > 0 &&
      current.placements.length === 0 &&
      current.usedHeight === 0
    ) {
      current.usedHeight = repeatHeaderHeight;
    }
    const currentHasContent = current.placements.length > 0;
    const remainingHeight = Math.max(0, availableHeight - current.usedHeight);
    const remainingFragments = blocks.length - cursor;
    const remainingFlowHeight = fragmentHeight(blocks, cursor, blocks.length);
    if (
      nextHeight > 0 &&
      currentHasContent &&
      remainingFlowHeight <= remainingHeight &&
      remainingFlowHeight + nextHeight > remainingHeight &&
      remainingFlowHeight + nextHeight <= availableHeight
    ) {
      pages.push(emptyPage(nextPageIndex(pages), availableHeight));
      continue;
    }
    let fitting = fragmentsFitting(
      blocks,
      cursor,
      blocks.length,
      remainingHeight,
    );

    if (fitting === remainingFragments) {
      placeFragments(blocks, cursor, blocks.length, current, availableHeight);
      cursor = blocks.length;
      continue;
    }
    if (fitting === 0) {
      if (currentHasContent) {
        pages.push(emptyPage(nextPageIndex(pages), availableHeight));
        continue;
      }
      fitting = 1;
    }
    if (remainingFragments > minimum) {
      fitting = Math.min(fitting, remainingFragments - minimum);
    }
    const minimumHere = Math.min(minimum, remainingFragments);
    if (fitting < minimumHere && currentHasContent) {
      pages.push(emptyPage(nextPageIndex(pages), availableHeight));
      continue;
    }
    fitting = Math.max(1, fitting);
    placeFragments(blocks, cursor, cursor + fitting, current, availableHeight);
    cursor += fitting;
    if (cursor < blocks.length) {
      pages.push(emptyPage(nextPageIndex(pages), availableHeight));
    }
  }

  if (last.breakAfter && hasMoreBlocks) {
    pages.push(emptyPage(nextPageIndex(pages), availableHeight));
  }
}

function nextBlockPreviewHeight(
  blocks: OfficeKernelLayoutRequest['blocks'],
): number {
  const first = blocks[0];
  if (!first || first.breakBefore) return 0;
  const count = Math.min(2, first.flowCount ?? 1);
  return fragmentHeight(blocks, 0, count);
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
    const flow = blocks.slice(index, index + count);
    if (flow.length !== count) {
      throw kernelError(
        'office.kernel.flow_sequence_invalid',
        'A layout flow must contain its declared number of fragments.',
      );
    }
    flow.forEach((fragment, flowIndex) => {
      if (
        fragment.flowId !== block.flowId ||
        fragment.flowIndex !== flowIndex ||
        fragment.flowCount !== count ||
        fragment.minimumFragmentsPerPage !== block.minimumFragmentsPerPage ||
        fragment.repeatHeaderCount !== repeatHeaderCount ||
        fragment.repeatHeaderHeight !== repeatHeaderHeight
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
    });
    index += count;
  }
}

function emptyPage(
  index: number,
  availableHeight: number,
): OfficeKernelLayoutPage {
  return {
    index,
    usedHeight: 0,
    availableHeight,
    placements: [],
  };
}

function nextPageIndex(pages: readonly OfficeKernelLayoutPage[]): number {
  return (pages.at(-1)?.index ?? -1) + 1;
}

function layoutBreaks(
  pages: OfficeKernelLayoutPage[],
  request: OfficeKernelLayoutRequest,
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
          request.page.marginBottom +
          request.page.footerHeight +
          request.page.pageGap +
          request.page.marginTop +
          request.page.headerHeight,
      },
    ];
  });
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
