import {
  type OfficeKernelLayoutBreak,
  type OfficeKernelLayoutPage,
  type OfficeKernelLayoutRequest,
  type OfficeKernelLayoutResult,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

const MAX_LAYOUT_BLOCKS = 10_000;
const MAX_LAYOUT_EXTENT = 1_000_000;

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
  const pages: OfficeKernelLayoutPage[] = [emptyPage(0, availableHeight)];

  request.blocks.forEach((block, index) => {
    let current = pages.at(-1) as OfficeKernelLayoutPage;
    let currentHasContent = current.placements.length > 0;
    if (block.breakBefore && currentHasContent) {
      current = emptyPage(pages.length, availableHeight);
      pages.push(current);
      currentHasContent = false;
    }

    const remaining = Math.max(0, availableHeight - current.usedHeight);
    const next = request.blocks[index + 1];
    const nextHeight = next && !next.breakBefore ? next.height : 0;
    const groupedHeight = block.height + (block.keepWithNext ? nextHeight : 0);
    const shouldAdvance =
      currentHasContent &&
      (block.height > remaining ||
        ((block.keepTogether || block.keepWithNext) &&
          groupedHeight <= availableHeight &&
          groupedHeight > remaining));
    if (shouldAdvance) {
      current = emptyPage(pages.length, availableHeight);
      pages.push(current);
    }

    const y = current.usedHeight;
    current.placements.push({
      blockId: block.id,
      y,
      height: block.height,
      overflow:
        block.height > availableHeight || y + block.height > availableHeight,
    });
    current.usedHeight += block.height;

    if (block.breakAfter && index + 1 < request.blocks.length) {
      pages.push(emptyPage(pages.length, availableHeight));
    }
  });

  if (pages.length > 1 && pages.at(-1)?.placements.length === 0) pages.pop();

  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'layoutResult',
    requestId: request.requestId,
    revision: request.revision,
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

function layoutBreaks(
  pages: OfficeKernelLayoutPage[],
  request: OfficeKernelLayoutRequest,
): OfficeKernelLayoutBreak[] {
  return pages.slice(1).flatMap((page) => {
    const beforeBlockId = page.placements[0]?.blockId;
    const previous = pages[page.index - 1];
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
