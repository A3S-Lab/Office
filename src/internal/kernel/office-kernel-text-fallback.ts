import type {
  OfficeKernelTextLayoutRequest,
  OfficeKernelTextLayoutResult,
} from './office-kernel-protocol';
import { OFFICE_KERNEL_PROTOCOL_VERSION } from './office-kernel-protocol';

/**
 * Browser font fallback is intentionally reported as unsupported.
 *
 * The caller can retain DOM line measurement when the WebAssembly shaper or
 * an exact registered font is unavailable. Returning estimated line boxes
 * here would make pagination appear deterministic while using different font
 * metrics from the rendered document.
 */
export function layoutOfficeTextInJavaScript(
  request: OfficeKernelTextLayoutRequest,
): OfficeKernelTextLayoutResult {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'textLayoutResult',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    engine: 'javascript',
    layouts: [],
    unsupportedParagraphIds: request.paragraphs.map(
      (paragraph) => paragraph.id,
    ),
  };
}
