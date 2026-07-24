import {
  type OfficeKernelPresentationAlignment,
  type OfficeKernelPresentationGeometryElement,
  type OfficeKernelPresentationGeometryRequest,
  type OfficeKernelPresentationGeometryResult,
  type OfficeKernelPresentationSnapGuide,
  type OfficeKernelPresentationTransformMode,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

const MAX_PRESENTATION_ELEMENTS = 10_000;
const MAX_PRESENTATION_EXTENT = 1_000_000;
const MAX_PRESENTATION_SNAP_THRESHOLD = 10;
const PRESENTATION_SLIDE_EXTENT = 100;
const PRESENTATION_SNAP_EPSILON = 0.000_001;

export function resolveOfficePresentationGeometryInJavaScript(
  request: OfficeKernelPresentationGeometryRequest,
): OfficeKernelPresentationGeometryResult {
  validatePresentationGeometryRequest(request);
  const operation = request.operation;
  const resolved =
    operation.type === 'alignToSlide'
      ? {
          elements: request.elements.map((element) =>
            alignElementToSlide(element, operation.alignment),
          ),
          guides: [],
        }
      : snapPresentationElement(
          request.elements,
          operation.movingElementId,
          operation.mode,
          operation.thresholdX,
          operation.thresholdY,
        );
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'presentationGeometryResult',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    engine: 'javascript',
    ...resolved,
  };
}

function alignElementToSlide(
  element: OfficeKernelPresentationGeometryElement,
  alignment: OfficeKernelPresentationAlignment,
): OfficeKernelPresentationGeometryElement {
  const maximumX = Math.max(0, PRESENTATION_SLIDE_EXTENT - element.width);
  const maximumY = Math.max(0, PRESENTATION_SLIDE_EXTENT - element.height);
  switch (alignment) {
    case 'left':
      return { ...element, x: 0 };
    case 'center':
      return { ...element, x: maximumX / 2 };
    case 'right':
      return { ...element, x: maximumX };
    case 'top':
      return { ...element, y: 0 };
    case 'middle':
      return { ...element, y: maximumY / 2 };
    case 'bottom':
      return { ...element, y: maximumY };
  }
}

function snapPresentationElement(
  elements: OfficeKernelPresentationGeometryElement[],
  movingElementId: string,
  mode: OfficeKernelPresentationTransformMode,
  thresholdX: number,
  thresholdY: number,
): {
  elements: OfficeKernelPresentationGeometryElement[];
  guides: OfficeKernelPresentationSnapGuide[];
} {
  const moving = elements.find((element) => element.id === movingElementId);
  if (!moving) {
    throw kernelError(
      'office.kernel.moving_element_missing',
      `Presentation element '${movingElementId}' does not exist.`,
    );
  }
  const xSnap = presentationSnapForAxis(
    moving,
    elements,
    'x',
    mode,
    thresholdX,
  );
  const ySnap = presentationSnapForAxis(
    moving,
    elements,
    'y',
    mode,
    thresholdY,
  );
  const snapped = applyPresentationSnap(moving, mode, xSnap, ySnap);
  const guides = [
    ...(xSnap &&
    presentationSnapApplied(moving, snapped, mode, 'x', xSnap.delta)
      ? [xSnap.guide]
      : []),
    ...(ySnap &&
    presentationSnapApplied(moving, snapped, mode, 'y', ySnap.delta)
      ? [ySnap.guide]
      : []),
  ];
  return {
    elements: elements.map((element) =>
      element.id === movingElementId ? snapped : element,
    ),
    guides,
  };
}

interface PresentationSnap {
  delta: number;
  guide: OfficeKernelPresentationSnapGuide;
}

function presentationSnapForAxis(
  moving: OfficeKernelPresentationGeometryElement,
  elements: OfficeKernelPresentationGeometryElement[],
  axis: 'x' | 'y',
  mode: OfficeKernelPresentationTransformMode,
  threshold: number,
): PresentationSnap | null {
  const targets = [
    ...[0, PRESENTATION_SLIDE_EXTENT / 2, PRESENTATION_SLIDE_EXTENT].map(
      (position) => ({
        position,
        guide: {
          axis,
          position,
          source: 'slide' as const,
        },
      }),
    ),
    ...elements.flatMap((element) =>
      element.id === moving.id
        ? []
        : presentationElementAnchors(element, axis).map((position) => ({
            position,
            guide: {
              axis,
              position,
              source: 'element' as const,
              targetId: element.id,
            },
          })),
    ),
  ];
  let best: (PresentationSnap & { distance: number }) | null = null;
  for (const movingPosition of presentationMovingAnchors(moving, axis, mode)) {
    for (const target of targets) {
      const delta = target.position - movingPosition;
      const distance = Math.abs(delta);
      if (
        distance <= threshold &&
        !presentationResizeSnapCollapsesElement(moving, axis, mode, delta) &&
        (!best || distance < best.distance - PRESENTATION_SNAP_EPSILON)
      ) {
        best = { delta, distance, guide: target.guide };
      }
    }
  }
  return best
    ? {
        delta: best.delta,
        guide: best.guide,
      }
    : null;
}

function presentationResizeSnapCollapsesElement(
  element: OfficeKernelPresentationGeometryElement,
  axis: 'x' | 'y',
  mode: OfficeKernelPresentationTransformMode,
  delta: number,
): boolean {
  if (mode !== 'resize') return false;
  const size = axis === 'x' ? element.width : element.height;
  return size + delta <= 0;
}

function presentationMovingAnchors(
  element: OfficeKernelPresentationGeometryElement,
  axis: 'x' | 'y',
  mode: OfficeKernelPresentationTransformMode,
): number[] {
  const start = axis === 'x' ? element.x : element.y;
  const size = axis === 'x' ? element.width : element.height;
  return mode === 'resize'
    ? [start + size]
    : [start, start + size / 2, start + size];
}

function presentationElementAnchors(
  element: OfficeKernelPresentationGeometryElement,
  axis: 'x' | 'y',
): number[] {
  const start = axis === 'x' ? element.x : element.y;
  const size = axis === 'x' ? element.width : element.height;
  return [start, start + size / 2, start + size];
}

function applyPresentationSnap(
  element: OfficeKernelPresentationGeometryElement,
  mode: OfficeKernelPresentationTransformMode,
  xSnap: PresentationSnap | null,
  ySnap: PresentationSnap | null,
): OfficeKernelPresentationGeometryElement {
  if (mode === 'move') {
    return {
      ...element,
      x: clampPresentationExtent(
        element.x + (xSnap?.delta ?? 0),
        PRESENTATION_SLIDE_EXTENT - element.width,
      ),
      y: clampPresentationExtent(
        element.y + (ySnap?.delta ?? 0),
        PRESENTATION_SLIDE_EXTENT - element.height,
      ),
    };
  }
  return {
    ...element,
    width: clampPresentationExtent(
      element.width + (xSnap?.delta ?? 0),
      PRESENTATION_SLIDE_EXTENT - element.x,
    ),
    height: clampPresentationExtent(
      element.height + (ySnap?.delta ?? 0),
      PRESENTATION_SLIDE_EXTENT - element.y,
    ),
  };
}

function presentationSnapApplied(
  before: OfficeKernelPresentationGeometryElement,
  after: OfficeKernelPresentationGeometryElement,
  mode: OfficeKernelPresentationTransformMode,
  axis: 'x' | 'y',
  delta: number,
): boolean {
  const applied =
    mode === 'move'
      ? axis === 'x'
        ? after.x - before.x
        : after.y - before.y
      : axis === 'x'
        ? after.width - before.width
        : after.height - before.height;
  return Math.abs(applied - delta) <= PRESENTATION_SNAP_EPSILON;
}

function clampPresentationExtent(value: number, maximum: number): number {
  return Math.min(Math.max(value, 0), Math.max(0, maximum));
}

function validatePresentationGeometryRequest(
  request: OfficeKernelPresentationGeometryRequest,
): void {
  if (request.protocol !== OFFICE_KERNEL_PROTOCOL_VERSION) {
    throw kernelError(
      'office.kernel.protocol_unsupported',
      `Office kernel protocol ${request.protocol} is unsupported.`,
    );
  }
  if (request.kind !== 'presentationGeometry') {
    throw kernelError(
      'office.kernel.request_kind_invalid',
      'The presentation geometry kernel received an invalid request kind.',
    );
  }
  for (const [name, value] of [
    ['requestId', request.requestId],
    ['revision', request.revision],
    ['documentRevision', request.documentRevision],
  ] as const) {
    if (!Number.isSafeInteger(value) || value < 0) {
      throw kernelError(
        'office.kernel.revision_invalid',
        `${name} must be a non-negative safe integer.`,
      );
    }
  }
  if (request.elements.length > MAX_PRESENTATION_ELEMENTS) {
    throw kernelError(
      'office.kernel.element_limit_exceeded',
      `A presentation geometry request may contain at most ${MAX_PRESENTATION_ELEMENTS} elements.`,
    );
  }
  if (request.operation.type === 'snapElement') {
    if (!request.operation.movingElementId.trim()) {
      throw kernelError(
        'office.kernel.moving_element_invalid',
        'A presentation snap request requires a moving element ID.',
      );
    }
    for (const [name, value] of [
      ['thresholdX', request.operation.thresholdX],
      ['thresholdY', request.operation.thresholdY],
    ] as const) {
      if (
        !Number.isFinite(value) ||
        value < 0 ||
        value > MAX_PRESENTATION_SNAP_THRESHOLD
      ) {
        throw kernelError(
          'office.kernel.snap_threshold_invalid',
          `${name} must be between 0 and ${MAX_PRESENTATION_SNAP_THRESHOLD}.`,
        );
      }
    }
  }
  const ids = new Set<string>();
  for (const element of request.elements) {
    if (!element.id.trim() || element.id.length > 256) {
      throw kernelError(
        'office.kernel.element_id_invalid',
        'Every presentation element requires a non-empty ID of at most 256 bytes.',
      );
    }
    if (ids.has(element.id)) {
      throw kernelError(
        'office.kernel.element_id_duplicate',
        `Presentation element ID '${element.id}' is duplicated.`,
      );
    }
    ids.add(element.id);
    for (const [name, value] of [
      ['x', element.x],
      ['y', element.y],
      ['width', element.width],
      ['height', element.height],
    ] as const) {
      if (
        !Number.isFinite(value) ||
        value < 0 ||
        value > MAX_PRESENTATION_EXTENT
      ) {
        throw kernelError(
          'office.kernel.extent_invalid',
          `element.${name} must be a finite non-negative number.`,
        );
      }
    }
    if (element.width <= 0 || element.height <= 0) {
      throw kernelError(
        'office.kernel.element_size_invalid',
        'Presentation element width and height must be positive.',
      );
    }
  }
}

function kernelError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}
