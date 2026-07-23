interface OfficeKernelWasmExports {
  memory: WebAssembly.Memory;
  office_kernel_abi_version: () => number;
  office_kernel_alloc: (length: number) => number;
  office_kernel_dealloc: (pointer: number, length: number) => void;
  office_kernel_layout: (pointer: number, length: number) => number;
  office_kernel_result_pointer: () => number;
  office_kernel_result_length: () => number;
}

const wasm = await WebAssembly.instantiate(
  await Bun.file(
    new URL('../generated/office-kernel.wasm', import.meta.url),
  ).arrayBuffer(),
  {},
);
const exports = wasm.instance.exports as unknown as OfficeKernelWasmExports;

assert(
  exports.office_kernel_abi_version() === 1,
  'Unexpected Office kernel ABI version.',
);

const request = {
  protocol: 1,
  kind: 'layout',
  requestId: 21,
  revision: 34,
  page: {
    width: 794,
    height: 1123,
    marginTop: 80,
    marginRight: 80,
    marginBottom: 80,
    marginLeft: 80,
    headerHeight: 0,
    footerHeight: 0,
    pageGap: 24,
  },
  blocks: [
    { id: 'one', height: 500 },
    { id: 'two', height: 500 },
  ],
};
const input = new TextEncoder().encode(JSON.stringify(request));
const pointer = exports.office_kernel_alloc(input.byteLength);
try {
  new Uint8Array(exports.memory.buffer, pointer, input.byteLength).set(input);
  assert(
    exports.office_kernel_layout(pointer, input.byteLength) === 0,
    'Office kernel layout call failed.',
  );
} finally {
  exports.office_kernel_dealloc(pointer, input.byteLength);
}

const output = new Uint8Array(
  exports.memory.buffer,
  exports.office_kernel_result_pointer(),
  exports.office_kernel_result_length(),
).slice();
const result = JSON.parse(new TextDecoder().decode(output)) as {
  engine?: string;
  pages?: unknown[];
  requestId?: number;
  revision?: number;
};
assert(result.engine === 'wasm', 'Office kernel did not use the WASM engine.');
assert(
  result.pages?.length === 2,
  'Office kernel pagination smoke test failed.',
);
assert(result.requestId === 21, 'Office kernel request ID was not preserved.');
assert(result.revision === 34, 'Office kernel revision was not preserved.');

console.log('Office kernel WASM ABI smoke test passed.');

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}
