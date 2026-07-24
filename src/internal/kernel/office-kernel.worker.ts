import { layoutOfficeDocumentInJavaScript } from './office-kernel-fallback';
import { alignOfficePresentationInJavaScript } from './office-kernel-presentation-fallback';
import { layoutOfficeTextInJavaScript } from './office-kernel-text-fallback';
import {
  isOfficeKernelResponse,
  type OfficeKernelFontSource,
  type OfficeKernelLayoutRequest,
  type OfficeKernelPresentationGeometryRequest,
  type OfficeKernelResponse,
  type OfficeKernelTextLayoutRequest,
  type OfficeKernelWorkerRequest,
  type OfficeKernelWorkerResponse,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from './office-kernel-protocol';

interface OfficeKernelWasmExports {
  memory: WebAssembly.Memory;
  office_kernel_abi_version: () => number;
  office_kernel_alloc: (length: number) => number;
  office_kernel_dealloc: (pointer: number, length: number) => void;
  office_kernel_layout: (pointer: number, length: number) => number;
  office_kernel_presentation_geometry: (
    pointer: number,
    length: number,
  ) => number;
  office_kernel_register_font: (
    idPointer: number,
    idLength: number,
    dataPointer: number,
    dataLength: number,
  ) => number;
  office_kernel_text_layout: (pointer: number, length: number) => number;
  office_kernel_result_pointer: () => number;
  office_kernel_result_length: () => number;
}

interface KernelWorkerScope {
  onmessage: ((event: MessageEvent<OfficeKernelWorkerRequest>) => void) | null;
  postMessage: (message: OfficeKernelWorkerResponse) => void;
}

const scope = globalThis as unknown as KernelWorkerScope;
const cancelledRequests = new Set<number>();
const MAX_FONT_BYTES = 32 * 1024 * 1024;
const MAX_REGISTERED_FONTS = 16;
let wasmInitialization: Promise<OfficeKernelWasmExports | null> =
  Promise.resolve(null);

scope.onmessage = (event) => {
  const message = event.data;
  if (message.kind === 'initialize') {
    wasmInitialization = message.wasmUrl
      ? initializeOfficeKernel(message.wasmUrl, message.fonts ?? []).catch(
          () => null,
        )
      : Promise.resolve(null);
    return;
  }
  if (message.kind === 'cancel') {
    cancelledRequests.add(message.requestId);
    return;
  }
  if (message.kind === 'layout') {
    void respondToLayout(message.request);
    return;
  }
  if (message.kind === 'presentationGeometry') {
    void respondToPresentationGeometry(message.request);
    return;
  }
  void respondToTextLayout(message.request);
};

async function respondToLayout(
  request: OfficeKernelLayoutRequest,
): Promise<void> {
  let response: OfficeKernelResponse;
  try {
    const wasm = await wasmInitialization;
    response = wasm
      ? layoutWithWasm(wasm, request)
      : layoutOfficeDocumentInJavaScript(request);
  } catch (error) {
    response = {
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'error',
      requestId: request.requestId,
      revision: request.revision,
      documentRevision: request.documentRevision,
      engine: 'javascript',
      error: {
        code:
          typeof error === 'object' &&
          error &&
          'code' in error &&
          typeof error.code === 'string'
            ? error.code
            : 'office.kernel.layout_failed',
        message:
          error instanceof Error ? error.message : 'Document layout failed.',
      },
    };
  }
  if (cancelledRequests.delete(request.requestId)) return;
  scope.postMessage({ kind: 'response', response });
}

async function respondToPresentationGeometry(
  request: OfficeKernelPresentationGeometryRequest,
): Promise<void> {
  let response: OfficeKernelResponse;
  try {
    const wasm = await wasmInitialization;
    response = wasm
      ? presentationGeometryWithWasm(wasm, request)
      : alignOfficePresentationInJavaScript(request);
  } catch (error) {
    response = errorResponse(request, error);
  }
  if (cancelledRequests.delete(request.requestId)) return;
  scope.postMessage({ kind: 'response', response });
}

async function respondToTextLayout(
  request: OfficeKernelTextLayoutRequest,
): Promise<void> {
  let response: OfficeKernelResponse;
  try {
    const wasm = await wasmInitialization;
    response = wasm
      ? textLayoutWithWasm(wasm, request)
      : layoutOfficeTextInJavaScript(request);
  } catch {
    response = layoutOfficeTextInJavaScript(request);
  }
  if (cancelledRequests.delete(request.requestId)) return;
  scope.postMessage({ kind: 'response', response });
}

async function initializeOfficeKernel(
  wasmUrl: string,
  fonts: readonly OfficeKernelFontSource[],
): Promise<OfficeKernelWasmExports> {
  const wasm = await loadOfficeKernelWasm(wasmUrl);
  await Promise.all(
    fonts.slice(0, MAX_REGISTERED_FONTS).map(async (font) => {
      try {
        const encodedId = new TextEncoder().encode(font.id);
        if (encodedId.byteLength === 0 || encodedId.byteLength > 128) return;
        const response = await fetch(font.url);
        if (!response.ok) return;
        const declaredSize = Number(response.headers.get('content-length'));
        if (Number.isFinite(declaredSize) && declaredSize > MAX_FONT_BYTES) {
          return;
        }
        const data = await response.arrayBuffer();
        if (data.byteLength === 0 || data.byteLength > MAX_FONT_BYTES) return;
        registerFontWithWasm(wasm, font.id, new Uint8Array(data));
      } catch {
        // Text layout reports this font as unsupported when registration fails.
      }
    }),
  );
  return wasm;
}

async function loadOfficeKernelWasm(
  wasmUrl: string,
): Promise<OfficeKernelWasmExports> {
  const response = await fetch(wasmUrl);
  if (!response.ok) {
    throw new Error(`Office kernel request failed with ${response.status}.`);
  }
  let source: WebAssembly.WebAssemblyInstantiatedSource;
  try {
    source = await WebAssembly.instantiateStreaming(response.clone(), {});
  } catch {
    source = await WebAssembly.instantiate(await response.arrayBuffer(), {});
  }
  const exports = source.instance.exports as unknown as OfficeKernelWasmExports;
  if (
    !(exports.memory instanceof WebAssembly.Memory) ||
    exports.office_kernel_abi_version() !== OFFICE_KERNEL_PROTOCOL_VERSION ||
    typeof exports.office_kernel_presentation_geometry !== 'function' ||
    typeof exports.office_kernel_register_font !== 'function' ||
    typeof exports.office_kernel_text_layout !== 'function'
  ) {
    throw new Error('Office kernel WebAssembly ABI is incompatible.');
  }
  return exports;
}

function registerFontWithWasm(
  wasm: OfficeKernelWasmExports,
  id: string,
  data: Uint8Array,
): void {
  const encodedId = new TextEncoder().encode(id);
  const idPointer = wasm.office_kernel_alloc(encodedId.byteLength);
  const dataPointer = wasm.office_kernel_alloc(data.byteLength);
  try {
    new Uint8Array(wasm.memory.buffer, idPointer, encodedId.byteLength).set(
      encodedId,
    );
    new Uint8Array(wasm.memory.buffer, dataPointer, data.byteLength).set(data);
    if (
      wasm.office_kernel_register_font(
        idPointer,
        encodedId.byteLength,
        dataPointer,
        data.byteLength,
      ) !== 0
    ) {
      throw new Error(`Office kernel rejected font '${id}'.`);
    }
  } finally {
    wasm.office_kernel_dealloc(idPointer, encodedId.byteLength);
    wasm.office_kernel_dealloc(dataPointer, data.byteLength);
  }
}

function layoutWithWasm(
  wasm: OfficeKernelWasmExports,
  request: OfficeKernelLayoutRequest,
): OfficeKernelResponse {
  const response = requestWithWasm(wasm, request, wasm.office_kernel_layout);
  if (response.kind !== 'layoutResult' && response.kind !== 'error') {
    throw new Error('Office layout kernel returned an unexpected response.');
  }
  return response;
}

function presentationGeometryWithWasm(
  wasm: OfficeKernelWasmExports,
  request: OfficeKernelPresentationGeometryRequest,
): OfficeKernelResponse {
  const response = requestWithWasm(
    wasm,
    request,
    wasm.office_kernel_presentation_geometry,
  );
  if (
    response.kind !== 'presentationGeometryResult' &&
    response.kind !== 'error'
  ) {
    throw new Error(
      'Office presentation kernel returned an unexpected response.',
    );
  }
  return response;
}

function textLayoutWithWasm(
  wasm: OfficeKernelWasmExports,
  request: OfficeKernelTextLayoutRequest,
): OfficeKernelResponse {
  const response = requestWithWasm(
    wasm,
    request,
    wasm.office_kernel_text_layout,
  );
  if (response.kind !== 'textLayoutResult' && response.kind !== 'error') {
    throw new Error('Office text kernel returned an unexpected response.');
  }
  return response.kind === 'error'
    ? layoutOfficeTextInJavaScript(request)
    : response;
}

function requestWithWasm(
  wasm: OfficeKernelWasmExports,
  request:
    | OfficeKernelLayoutRequest
    | OfficeKernelPresentationGeometryRequest
    | OfficeKernelTextLayoutRequest,
  execute: (pointer: number, length: number) => number,
): OfficeKernelResponse {
  const input = new TextEncoder().encode(JSON.stringify(request));
  const pointer = wasm.office_kernel_alloc(input.byteLength);
  try {
    new Uint8Array(wasm.memory.buffer, pointer, input.byteLength).set(input);
    execute(pointer, input.byteLength);
  } finally {
    wasm.office_kernel_dealloc(pointer, input.byteLength);
  }
  const resultPointer = wasm.office_kernel_result_pointer();
  const resultLength = wasm.office_kernel_result_length();
  const output = new Uint8Array(
    wasm.memory.buffer,
    resultPointer,
    resultLength,
  ).slice();
  const parsed: unknown = JSON.parse(new TextDecoder().decode(output));
  if (!isOfficeKernelResponse(parsed)) {
    throw new Error('Office kernel returned an invalid response.');
  }
  return parsed;
}

function errorResponse(
  request:
    | OfficeKernelLayoutRequest
    | OfficeKernelPresentationGeometryRequest
    | OfficeKernelTextLayoutRequest,
  error: unknown,
): OfficeKernelResponse {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'error',
    requestId: request.requestId,
    revision: request.revision,
    documentRevision: request.documentRevision,
    engine: 'javascript',
    error: {
      code:
        typeof error === 'object' &&
        error &&
        'code' in error &&
        typeof error.code === 'string'
          ? error.code
          : 'office.kernel.request_failed',
      message:
        error instanceof Error
          ? error.message
          : 'Office kernel request failed.',
    },
  };
}
