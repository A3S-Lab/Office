export const OFFICE_KERNEL_PROTOCOL_VERSION = 11 as const;

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
  flowId?: string;
  flowIndex?: number;
  flowCount?: number;
  minimumFragmentsPerPage?: number;
  repeatHeaderCount?: number;
  repeatHeaderHeight?: number;
}

export interface OfficeKernelLayoutRequest {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'layout';
  requestId: number;
  revision: number;
  documentRevision: number;
  startPageIndex: number;
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
  documentRevision: number;
  startPageIndex: number;
  engine: OfficeKernelEngine;
  pages: OfficeKernelLayoutPage[];
  breaks: OfficeKernelLayoutBreak[];
}

export type OfficeKernelPresentationAlignment =
  | 'bottom'
  | 'center'
  | 'left'
  | 'middle'
  | 'right'
  | 'top';

export interface OfficeKernelPresentationGeometryElement {
  id: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface OfficeKernelPresentationGeometryRequest {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'presentationGeometry';
  requestId: number;
  revision: number;
  documentRevision: number;
  operation: {
    type: 'alignToSlide';
    alignment: OfficeKernelPresentationAlignment;
  };
  elements: OfficeKernelPresentationGeometryElement[];
}

export interface OfficeKernelPresentationGeometryResult {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'presentationGeometryResult';
  requestId: number;
  revision: number;
  documentRevision: number;
  engine: OfficeKernelEngine;
  elements: OfficeKernelPresentationGeometryElement[];
}

export type OfficeKernelTextDirection = 'auto' | 'ltr' | 'rtl';
export type OfficeKernelTextWhiteSpace = 'breakSpaces' | 'normal';
export type OfficeKernelTextTabAlignment =
  | 'center'
  | 'decimal'
  | 'left'
  | 'right';

export interface OfficeKernelTextTabStop {
  position: number;
  alignment: OfficeKernelTextTabAlignment;
}

export interface OfficeKernelTextTabLayout {
  origin: number;
  firstLineIndent: number;
  defaultInterval: number;
  stops: OfficeKernelTextTabStop[];
}

export interface OfficeKernelTextLayoutRun {
  startUtf16: number;
  endUtf16: number;
  fontId: string;
  fallbackFontIds?: string[];
  fontSize: number;
  lineHeight: number;
  letterSpacing?: number;
  ligatures?: boolean;
  kerning?: boolean;
}

export interface OfficeKernelTextLayoutParagraph {
  id: string;
  text: string;
  runs: OfficeKernelTextLayoutRun[];
  maxWidth: number;
  firstLineMaxWidth?: number;
  direction?: OfficeKernelTextDirection;
  whiteSpace?: OfficeKernelTextWhiteSpace;
  tabLayout?: OfficeKernelTextTabLayout;
}

export interface OfficeKernelTextLayoutRequest {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'textLayout';
  requestId: number;
  revision: number;
  documentRevision: number;
  paragraphs: OfficeKernelTextLayoutParagraph[];
}

export interface OfficeKernelTextLayoutLine {
  startUtf16: number;
  endUtf16: number;
  width: number;
  ascent: number;
  descent: number;
  height: number;
  hardBreak: boolean;
}

export interface OfficeKernelTextLayoutParagraphResult {
  id: string;
  glyphCount: number;
  fallbackGlyphCount: number;
  missingGlyphCount: number;
  lines: OfficeKernelTextLayoutLine[];
}

export interface OfficeKernelTextLayoutResult {
  protocol: typeof OFFICE_KERNEL_PROTOCOL_VERSION;
  kind: 'textLayoutResult';
  requestId: number;
  revision: number;
  documentRevision: number;
  engine: OfficeKernelEngine;
  layouts: OfficeKernelTextLayoutParagraphResult[];
  unsupportedParagraphIds: string[];
}

export interface OfficeKernelFontSource {
  id: string;
  url: string;
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
  documentRevision: number;
  engine: OfficeKernelEngine;
  error: OfficeKernelErrorDetail;
}

export type OfficeKernelResponse =
  | OfficeKernelLayoutResult
  | OfficeKernelPresentationGeometryResult
  | OfficeKernelTextLayoutResult
  | OfficeKernelErrorResponse;

export type OfficeKernelWorkerRequest =
  | {
      kind: 'initialize';
      wasmUrl?: string;
      fonts?: OfficeKernelFontSource[];
    }
  | {
      kind: 'layout';
      request: OfficeKernelLayoutRequest;
    }
  | {
      kind: 'presentationGeometry';
      request: OfficeKernelPresentationGeometryRequest;
    }
  | {
      kind: 'textLayout';
      request: OfficeKernelTextLayoutRequest;
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
    (candidate.kind !== 'layoutResult' &&
      candidate.kind !== 'presentationGeometryResult' &&
      candidate.kind !== 'textLayoutResult' &&
      candidate.kind !== 'error')
  ) {
    return false;
  }
  if (
    !isNonNegativeInteger(candidate.requestId) ||
    !isNonNegativeInteger(candidate.revision) ||
    !isNonNegativeInteger(candidate.documentRevision) ||
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
  if (candidate.kind === 'presentationGeometryResult') {
    return (
      Array.isArray(candidate.elements) &&
      candidate.elements.every(isPresentationGeometryElement)
    );
  }
  if (candidate.kind === 'textLayoutResult') {
    return (
      Array.isArray(candidate.layouts) &&
      candidate.layouts.every(isTextLayoutParagraphResult) &&
      Array.isArray(candidate.unsupportedParagraphIds) &&
      candidate.unsupportedParagraphIds.every(
        (id) => typeof id === 'string' && id.length > 0,
      )
    );
  }
  return (
    isNonNegativeInteger(candidate.startPageIndex) &&
    Array.isArray(candidate.pages) &&
    candidate.pages.every(
      (page, index) =>
        isLayoutPage(page) &&
        page.index === Number(candidate.startPageIndex) + index,
    ) &&
    Array.isArray(candidate.breaks) &&
    candidate.breaks.every(isLayoutBreak)
  );
}

function isTextLayoutParagraphResult(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const layout = value as Record<string, unknown>;
  if (
    typeof layout.id !== 'string' ||
    layout.id.length === 0 ||
    !isNonNegativeInteger(layout.glyphCount) ||
    !isNonNegativeInteger(layout.fallbackGlyphCount) ||
    !isNonNegativeInteger(layout.missingGlyphCount) ||
    Number(layout.fallbackGlyphCount) > Number(layout.glyphCount) ||
    Number(layout.missingGlyphCount) > Number(layout.glyphCount) ||
    !Array.isArray(layout.lines)
  ) {
    return false;
  }
  let previousEnd = 0;
  return layout.lines.every((line, index) => {
    if (!isTextLayoutLine(line)) return false;
    const candidate = line as Record<string, unknown>;
    const start = Number(candidate.startUtf16);
    const end = Number(candidate.endUtf16);
    if (index === 0 && start !== 0) return false;
    if (start < previousEnd) return false;
    previousEnd = end;
    return true;
  });
}

function isTextLayoutLine(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const line = value as Record<string, unknown>;
  return (
    isNonNegativeInteger(line.startUtf16) &&
    isNonNegativeInteger(line.endUtf16) &&
    Number(line.endUtf16) >= Number(line.startUtf16) &&
    isNonNegativeNumber(line.width) &&
    isNonNegativeNumber(line.ascent) &&
    isNonNegativeNumber(line.descent) &&
    isPositiveNumber(line.height) &&
    typeof line.hardBreak === 'boolean'
  );
}

function isPresentationGeometryElement(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  const element = value as Record<string, unknown>;
  return (
    typeof element.id === 'string' &&
    isNonNegativeNumber(element.x) &&
    isNonNegativeNumber(element.y) &&
    isNonNegativeNumber(element.width) &&
    isNonNegativeNumber(element.height)
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

function isPositiveNumber(value: unknown): boolean {
  return typeof value === 'number' && Number.isFinite(value) && value > 0;
}
