import { layoutOfficeDocumentInJavaScript } from './office-kernel-fallback';
import {
  isOfficeKernelResponse,
  type OfficeKernelLayoutRequest,
  type OfficeKernelResponse,
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
  office_kernel_result_pointer: () => number;
  office_kernel_result_length: () => number;
}

interface KernelWorkerScope {
  onmessage: ((event: MessageEvent<OfficeKernelWorkerRequest>) => void) | null;
  postMessage: (message: OfficeKernelWorkerResponse) => void;
}

const scope = globalThis as unknown as KernelWorkerScope;
const cancelledRequests = new Set<number>();
let wasmInitialization: Promise<OfficeKernelWasmExports | null> =
  Promise.resolve(null);

scope.onmessage = (event) => {
  const message = event.data;
  if (message.kind === 'initialize') {
    wasmInitialization = message.wasmUrl
      ? loadOfficeKernelWasm(message.wasmUrl).catch(() => null)
      : Promise.resolve(null);
    return;
  }
  if (message.kind === 'cancel') {
    cancelledRequests.add(message.requestId);
    return;
  }
  void respondToLayout(message.request);
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
    exports.office_kernel_abi_version() !== OFFICE_KERNEL_PROTOCOL_VERSION
  ) {
    throw new Error('Office kernel WebAssembly ABI is incompatible.');
  }
  return exports;
}

function layoutWithWasm(
  wasm: OfficeKernelWasmExports,
  request: OfficeKernelLayoutRequest,
): OfficeKernelResponse {
  const input = new TextEncoder().encode(JSON.stringify(request));
  const pointer = wasm.office_kernel_alloc(input.byteLength);
  try {
    new Uint8Array(wasm.memory.buffer, pointer, input.byteLength).set(input);
    wasm.office_kernel_layout(pointer, input.byteLength);
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
