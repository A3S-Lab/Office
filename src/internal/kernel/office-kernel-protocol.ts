export const OFFICE_KERNEL_PROTOCOL_VERSION = 1 as const;

export type OfficeKernelEngine = 'wasm' | 'javascript';

export interface OfficeKernelPageMetrics {
  width: number;
  height: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  headerHeight: number;
  footerHeight: number;
  pageGap: number;
}

export interface OfficeKernelLayoutBlock {
  id: string;
  height: number;
  breakBefore?: boolean;
  breakAfter?: boolean;
  keepTogether?: boolean;
  keepWithNext?: boolean;
}

export interface OfficeKernelLayoutRequest {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'layout';
  requestId: number;
  revision: number;
  page: OfficeKernelPageMetrics;
  blocks: OfficeKernelLayoutBlock[];
}

export interface OfficeKernelLayoutPlacement {
  blockId: string;
  y: number;
  height: number;
  overflow: boolean;
}

export interface OfficeKernelLayoutPage {
  index: number;
  usedHeight: number;
  availableHeight: number;
  placements: OfficeKernelLayoutPlacement[];
}

export interface OfficeKernelLayoutBreak {
  beforeBlockId: string;
  pageIndex: number;
  spacerHeight: number;
  remainingBodyHeight: number;
}

export interface OfficeKernelLayoutResult {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'layoutResult';
  requestId: number;
  revision: number;
  engine: OfficeKernelEngine;
  pages: OfficeKernelLayoutPage[];
  breaks: OfficeKernelLayoutBreak[];
}

export interface OfficeKernelErrorDetail {
  code: string;
  message: string;
}

export interface OfficeKernelErrorResponse {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'error';
  requestId: number;
  revision: number;
  engine: OfficeKernelEngine;
  error: OfficeKernelErrorDetail;
}

export type OfficeKernelResponse =
  | OfficeKernelLayoutResult
  | OfficeKernelErrorResponse;

export type OfficeKernelWorkerRequest =
  | {
      kind: 'initialize';
      wasmUrl?: string;
    }
  | {
      kind: 'layout';
      request: OfficeKernelLayoutRequest;
    }
  | {
      kind: 'cancel';
      requestId: number;
    };

export interface OfficeKernelWorkerResponse {
  kind: 'response';
  response: OfficeKernelResponse;
}

export function isOfficeKernelResponse(
  value: unknown,
): value is OfficeKernelResponse {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Record<string, unknown>;
  if (
    candidate.protocol !== OFFICE_KERNEL_PROTOCOL_VERSION ||
    (candidate.kind !== 'layoutResult' && candidate.kind !== 'error') ||
    !isNonNegativeInteger(candidate.requestId) ||
    !isNonNegativeInteger(candidate.revision) ||
    (candidate.engine !== 'wasm' && candidate.engine !== 'javascript')
  ) {
    return false;
  }
  if (candidate.kind === 'error') {
    const error = candidate.error as Record<string, unknown> | undefined;
    return (
      Boolean(error) &&
      typeof error?.code === 'string' &&
      typeof error.message === 'string'
    );
  }
  return (
    Array.isArray(candidate.pages) &&
    candidate.pages.every(isLayoutPage) &&
    Array.isArray(candidate.breaks) &&
    candidate.breaks.every(isLayoutBreak)
  );
}

function isLayoutPage(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const page = value as Record<string, unknown>;
  return (
    isNonNegativeInteger(page.index) &&
    isNonNegativeNumber(page.usedHeight) &&
    isNonNegativeNumber(page.availableHeight) &&
    Array.isArray(page.placements) &&
    page.placements.every(isLayoutPlacement)
  );
}

function isLayoutPlacement(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const placement = value as Record<string, unknown>;
  return (
    typeof placement.blockId === 'string' &&
    isNonNegativeNumber(placement.y) &&
    isNonNegativeNumber(placement.height) &&
    typeof placement.overflow === 'boolean'
  );
}

function isLayoutBreak(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const pageBreak = value as Record<string, unknown>;
  return (
    typeof pageBreak.beforeBlockId === 'string' &&
    isNonNegativeInteger(pageBreak.pageIndex) &&
    isNonNegativeNumber(pageBreak.spacerHeight) &&
    isNonNegativeNumber(pageBreak.remainingBodyHeight)
  );
}

function isNonNegativeInteger(value: unknown): boolean {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function isNonNegativeNumber(value: unknown): boolean {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0;
}
