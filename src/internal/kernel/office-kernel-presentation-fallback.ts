import {
  type OfficeKernelPresentationAlignment,
  type OfficeKernelPresentationGeometryElement,
  type OfficeKernelPresentationGeometryRequest,
  type OfficeKernelPresentationGeometryResult,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

const MAX_PRESENTATION_ELEMENTS = 10_000;
const MAX_PRESENTATION_EXTENT = 1_000_000;
const PRESENTATION_SLIDE_EXTENT = 100;

export function alignOfficePresentationInJavaScript(
  request: OfficeKernelPresentationGeometryRequest,
): OfficeKernelPresentationGeometryResult {
  validatePresentationGeometryRequest(request);
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'presentationGeometryResult',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    engine: 'javascript',
    elements: request.elements.map((element) =>
      alignElementToSlide(element, request.operation.alignment),
    ),
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

function validatePresentationGeometryRequest(
  request: OfficeKernelPresentationGeometryRequest,
): void {
  if (request.protocol !== OFFICE_KERNEL_PROTOCOL_VERSION) {
    throw kernelError(
      'office.kernel.protocol_unsupported',
      `Office kernel protocol ${request.protocol} is unsupported.`,
    );
  }
  if (
    request.kind !== 'presentationGeometry' ||
    request.operation.type !== 'alignToSlide'
  ) {
    throw kernelError(
      'office.kernel.request_kind_invalid',
      'The presentation geometry kernel only accepts slide-alignment requests.',
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
