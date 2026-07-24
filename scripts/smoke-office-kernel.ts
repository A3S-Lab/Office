import { OFFICE_KERNEL_PROTOCOL_VERSION } from '../src/internal/kernel/office-kernel-protocol';

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
  office_kernel_spreadsheet_calculation: (
    pointer: number,
    length: number,
  ) => number;
  office_kernel_spreadsheet_session_calculation: (
    pointer: number,
    length: number,
  ) => number;
  office_kernel_text_layout: (pointer: number, length: number) => number;
  office_kernel_result_pointer: () => number;
  office_kernel_result_length: () => number;
}

interface SpreadsheetSmokeResult {
  engine?: string;
  kind?: string;
  cells?: Array<{ value?: { kind?: string; value?: number } }>;
  calculationOrder?: unknown[];
  issues?: Array<{ code?: string }>;
  stats?: {
    updateKind?: string;
    calculationScope?: string;
    formulaCellCount?: number;
    dirtyFormulaCellCount?: number;
    evaluatedFormulaCellCount?: number;
    dependencyEdgeCount?: number;
  };
}

const wasm = await WebAssembly.instantiate(
  await Bun.file(
    new URL('../generated/office-kernel.wasm', import.meta.url),
  ).arrayBuffer(),
  {},
);
const exports = wasm.instance.exports as unknown as OfficeKernelWasmExports;

assert(
  exports.office_kernel_abi_version() === OFFICE_KERNEL_PROTOCOL_VERSION,
  'Unexpected Office kernel ABI version.',
);

const request = {
  protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
  kind: 'layout',
  requestId: 21,
  revision: 34,
  documentRevision: 55,
  startPageIndex: 0,
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
  documentRevision?: number;
  startPageIndex?: number;
};
assert(result.engine === 'wasm', 'Office kernel did not use the WASM engine.');
assert(
  result.pages?.length === 2,
  'Office kernel pagination smoke test failed.',
);
assert(result.requestId === 21, 'Office kernel request ID was not preserved.');
assert(result.revision === 34, 'Office kernel revision was not preserved.');
assert(
  result.documentRevision === 55,
  'Office kernel document revision was not preserved.',
);
assert(
  result.startPageIndex === 0,
  'Office kernel page offset was not preserved.',
);

const geometryRequest = {
  protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
  kind: 'presentationGeometry',
  requestId: 22,
  revision: 35,
  documentRevision: 56,
  operation: {
    type: 'alignToSlide',
    alignment: 'right',
  },
  elements: [
    {
      id: 'shape',
      x: 12,
      y: 18,
      width: 25,
      height: 20,
    },
  ],
};
const geometryInput = new TextEncoder().encode(JSON.stringify(geometryRequest));
const geometryPointer = exports.office_kernel_alloc(geometryInput.byteLength);
try {
  new Uint8Array(
    exports.memory.buffer,
    geometryPointer,
    geometryInput.byteLength,
  ).set(geometryInput);
  assert(
    exports.office_kernel_presentation_geometry(
      geometryPointer,
      geometryInput.byteLength,
    ) === 0,
    'Office kernel presentation geometry call failed.',
  );
} finally {
  exports.office_kernel_dealloc(geometryPointer, geometryInput.byteLength);
}
const geometryOutput = new Uint8Array(
  exports.memory.buffer,
  exports.office_kernel_result_pointer(),
  exports.office_kernel_result_length(),
).slice();
const geometryResult = JSON.parse(new TextDecoder().decode(geometryOutput)) as {
  engine?: string;
  kind?: string;
  elements?: Array<{ x?: number }>;
};
assert(
  geometryResult.kind === 'presentationGeometryResult',
  'Office kernel presentation response kind was not preserved.',
);
assert(
  geometryResult.engine === 'wasm',
  'Office presentation geometry did not use the WASM engine.',
);
assert(
  geometryResult.elements?.[0]?.x === 75,
  'Office presentation alignment result was incorrect.',
);

await registerFont(
  'noto-sans-regular',
  new URL(
    '../node_modules/@embedpdf/fonts-latin/fonts/NotoSans-Regular.ttf',
    import.meta.url,
  ),
);
await registerFont(
  'noto-sans-hans-regular',
  new URL(
    '../node_modules/@embedpdf/fonts-sc/fonts/NotoSansHans-Regular.otf',
    import.meta.url,
  ),
);
await registerFont(
  'noto-naskh-arabic-regular',
  new URL(
    '../node_modules/@embedpdf/fonts-arabic/fonts/NotoNaskhArabic-Regular.ttf',
    import.meta.url,
  ),
);
await registerFont(
  'noto-sans-hebrew-regular',
  new URL(
    '../node_modules/@embedpdf/fonts-hebrew/fonts/NotoSansHebrew-Regular.ttf',
    import.meta.url,
  ),
);

const bidiText = 'A3S مرحبًا 2026 שלום';
const rtlTabText = 'A3S\tשלום';
const textLayoutRequest = {
  protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
  kind: 'textLayout',
  requestId: 23,
  revision: 36,
  documentRevision: 57,
  paragraphs: [
    {
      id: 'paragraph',
      text: 'A3S 文档确定性排版',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 4,
          fontId: 'noto-sans-hans-regular',
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: false,
          kerning: true,
        },
        {
          startUtf16: 4,
          endUtf16: 11,
          fontId: 'noto-sans-hans-regular',
          fontSize: 22,
          lineHeight: 30,
          letterSpacing: 0,
          ligatures: false,
          kerning: true,
        },
      ],
      maxWidth: 70,
      firstLineMaxWidth: 56,
      direction: 'auto',
      whiteSpace: 'breakSpaces',
    },
    {
      id: 'tabbed-paragraph',
      text: 'A3S Office\tQ',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 12,
          fontId: 'noto-sans-hans-regular',
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: false,
          kerning: true,
        },
      ],
      maxWidth: 240,
      firstLineMaxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'breakSpaces',
      tabLayout: {
        origin: 0,
        firstLineIndent: 0,
        defaultInterval: 48,
        stops: [{ position: 96, alignment: 'right' }],
      },
    },
    {
      id: 'fallback-paragraph',
      text: 'A3S 文档',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 6,
          fontId: 'noto-sans-regular',
          fallbackFontIds: ['noto-sans-hans-regular'],
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: true,
          kerning: true,
        },
      ],
      maxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'normal',
    },
    {
      id: 'missing-fallback-paragraph',
      text: 'A3S 文档',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 6,
          fontId: 'noto-sans-regular',
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: true,
          kerning: true,
        },
      ],
      maxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'normal',
    },
    {
      id: 'cjk-metrics-paragraph',
      text: '文档',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 2,
          fontId: 'noto-sans-hans-regular',
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: true,
          kerning: true,
        },
      ],
      maxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'normal',
    },
    {
      id: 'fallback-cjk-metrics-paragraph',
      text: '文档',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 2,
          fontId: 'noto-sans-regular',
          fallbackFontIds: ['noto-sans-hans-regular'],
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: true,
          kerning: true,
        },
      ],
      maxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'normal',
    },
    {
      id: 'bidi-paragraph',
      text: bidiText,
      runs: [
        {
          startUtf16: 0,
          endUtf16: bidiText.length,
          fontId: 'noto-sans-regular',
          fallbackFontIds: [
            'noto-sans-hans-regular',
            'noto-naskh-arabic-regular',
            'noto-sans-hebrew-regular',
          ],
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: true,
          kerning: true,
        },
      ],
      maxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'normal',
    },
    {
      id: 'rtl-tab-paragraph',
      text: rtlTabText,
      runs: [
        {
          startUtf16: 0,
          endUtf16: rtlTabText.length,
          fontId: 'noto-sans-regular',
          fallbackFontIds: ['noto-sans-hebrew-regular'],
          fontSize: 14,
          lineHeight: 21,
          letterSpacing: 0,
          ligatures: true,
          kerning: true,
        },
      ],
      maxWidth: 240,
      direction: 'ltr',
      whiteSpace: 'breakSpaces',
      tabLayout: {
        origin: 0,
        firstLineIndent: 0,
        defaultInterval: 48,
        stops: [],
      },
    },
  ],
};
const textInput = new TextEncoder().encode(JSON.stringify(textLayoutRequest));
const textPointer = exports.office_kernel_alloc(textInput.byteLength);
try {
  new Uint8Array(exports.memory.buffer, textPointer, textInput.byteLength).set(
    textInput,
  );
  assert(
    exports.office_kernel_text_layout(textPointer, textInput.byteLength) === 0,
    'Office kernel text layout call failed.',
  );
} finally {
  exports.office_kernel_dealloc(textPointer, textInput.byteLength);
}
const textOutput = new Uint8Array(
  exports.memory.buffer,
  exports.office_kernel_result_pointer(),
  exports.office_kernel_result_length(),
).slice();
const textResult = JSON.parse(new TextDecoder().decode(textOutput)) as {
  engine?: string;
  kind?: string;
  layouts?: Array<{
    id?: string;
    glyphCount?: number;
    fallbackGlyphCount?: number;
    missingGlyphCount?: number;
    lines?: Array<{
      startUtf16?: number;
      endUtf16?: number;
      ascent?: number;
      descent?: number;
      height?: number;
      width?: number;
    }>;
  }>;
  unsupportedParagraphIds?: string[];
};
assert(
  textResult.kind === 'textLayoutResult',
  'Office text layout response kind was not preserved.',
);
assert(
  textResult.engine === 'wasm',
  'Office text layout did not use the WASM engine.',
);
assert(
  (textResult.layouts?.[0]?.lines?.length ?? 0) > 1,
  'Office text layout did not wrap the paragraph.',
);
assert(
  textResult.layouts?.[0]?.missingGlyphCount === 0,
  'Office text layout reported missing glyphs.',
);
assert(
  (textResult.layouts?.[0]?.glyphCount ?? 0) > 0,
  'Office text layout did not shape glyphs.',
);
const tabbedLayout = textResult.layouts?.find(
  (layout) => layout.id === 'tabbed-paragraph',
);
assert(
  tabbedLayout?.lines?.length === 1,
  'Office text layout did not keep the tabbed paragraph on one line.',
);
assert(
  Math.abs((tabbedLayout?.lines?.[0]?.width ?? 0) - 96) < 0.01,
  `Office text layout did not apply the right-aligned structured tab stop; received ${tabbedLayout?.lines?.[0]?.width ?? 'no width'}.`,
);
assert(
  textResult.layouts?.[0]?.lines?.some((line) => line.height === 30),
  'Office text layout did not preserve mixed-run line metrics.',
);
const fallbackLayout = textResult.layouts?.find(
  (layout) => layout.id === 'fallback-paragraph',
);
const missingFallbackLayout = textResult.layouts?.find(
  (layout) => layout.id === 'missing-fallback-paragraph',
);
const cjkMetricsLayout = textResult.layouts?.find(
  (layout) => layout.id === 'cjk-metrics-paragraph',
);
const fallbackCjkMetricsLayout = textResult.layouts?.find(
  (layout) => layout.id === 'fallback-cjk-metrics-paragraph',
);
const bidiLayout = textResult.layouts?.find(
  (layout) => layout.id === 'bidi-paragraph',
);
assert(
  fallbackLayout?.missingGlyphCount === 0,
  'Office text layout did not resolve CJK glyphs through the registered fallback face.',
);
assert(
  (fallbackLayout?.fallbackGlyphCount ?? 0) > 0,
  'Office text layout did not report glyphs shaped by the fallback face.',
);
assert(
  (missingFallbackLayout?.missingGlyphCount ?? 0) > 0,
  'Office text layout did not prove that the Latin primary face lacks the CJK glyphs.',
);
const fallbackLine = fallbackLayout?.lines?.[0];
const latinLine = missingFallbackLayout?.lines?.[0];
const cjkLine = cjkMetricsLayout?.lines?.[0];
assert(
  latinLine?.ascent !== cjkLine?.ascent ||
    latinLine?.descent !== cjkLine?.descent,
  'Office text layout fixtures do not distinguish Latin and CJK font metrics.',
);
assert(
  fallbackLine?.ascent ===
    Math.max(latinLine?.ascent ?? 0, cjkLine?.ascent ?? 0) &&
    fallbackLine?.descent ===
      Math.max(latinLine?.descent ?? 0, cjkLine?.descent ?? 0),
  'Office text layout did not include the resolved fallback face in line metrics.',
);
const fallbackCjkLine = fallbackCjkMetricsLayout?.lines?.[0];
assert(
  fallbackCjkMetricsLayout?.missingGlyphCount === 0 &&
    (fallbackCjkMetricsLayout?.fallbackGlyphCount ?? 0) > 0 &&
    fallbackCjkLine?.ascent ===
      Math.max(latinLine?.ascent ?? 0, cjkLine?.ascent ?? 0) &&
    fallbackCjkLine?.descent ===
      Math.max(latinLine?.descent ?? 0, cjkLine?.descent ?? 0),
  'Office text layout did not preserve the primary strut and fallback metrics on a fallback-only line.',
);
assert(
  bidiLayout?.missingGlyphCount === 0 &&
    (bidiLayout?.fallbackGlyphCount ?? 0) > 0 &&
    (bidiLayout?.glyphCount ?? 0) > 0,
  'Office text layout did not shape mixed Latin, Arabic, numeric, and Hebrew directional runs with registered fallback faces.',
);
assert(
  textResult.unsupportedParagraphIds?.includes('rtl-tab-paragraph'),
  'Office text layout did not explicitly fall back for a tab paragraph containing resolved RTL text.',
);

const spreadsheetRequest = {
  protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
  kind: 'spreadsheetCalculation',
  requestId: 24,
  revision: 37,
  documentRevision: 58,
  sheets: [
    {
      id: 'sheet-1',
      name: 'Sheet 1',
      cells: [
        { row: 0, column: 0, value: { kind: 'number', value: 2 } },
        { row: 1, column: 0, value: { kind: 'number', value: 3 } },
        {
          row: 0,
          column: 1,
          formula: '=SUM(A1:A2)',
          value: { kind: 'blank' },
        },
        {
          row: 1,
          column: 1,
          formula: '=B1*2',
          value: { kind: 'blank' },
        },
      ],
    },
  ],
};
const spreadsheetResult = calculateSpreadsheetWithWasm(spreadsheetRequest);
assert(
  spreadsheetResult.kind === 'spreadsheetCalculationResult',
  'Office Spreadsheet response kind was not preserved.',
);
assert(
  spreadsheetResult.engine === 'wasm',
  'Office Spreadsheet calculation did not use the WASM engine.',
);
assert(
  spreadsheetResult.cells?.[0]?.value?.value === 5 &&
    spreadsheetResult.cells?.[1]?.value?.value === 10,
  'Office Spreadsheet dependency calculation returned incorrect values.',
);
assert(
  spreadsheetResult.calculationOrder?.length === 2 &&
    spreadsheetResult.issues?.length === 0,
  'Office Spreadsheet calculation did not preserve deterministic diagnostics.',
);

const spreadsheetSessionResult = calculateSpreadsheetSessionWithWasm({
  protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
  kind: 'spreadsheetSessionCalculation',
  requestId: 26,
  revision: 1,
  documentRevision: 1,
  update: {
    kind: 'replace',
    sheets: spreadsheetRequest.sheets,
  },
  calculation: { kind: 'workbook' },
});
assert(
  spreadsheetSessionResult.kind === 'spreadsheetSessionCalculationResult' &&
    spreadsheetSessionResult.stats?.updateKind === 'replace' &&
    spreadsheetSessionResult.stats?.evaluatedFormulaCellCount === 2 &&
    spreadsheetSessionResult.stats?.dependencyEdgeCount === 3,
  'Office Spreadsheet session did not initialize its dependency graph.',
);
const spreadsheetPatchResult = calculateSpreadsheetSessionWithWasm({
  protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
  kind: 'spreadsheetSessionCalculation',
  requestId: 27,
  revision: 2,
  documentRevision: 2,
  update: {
    kind: 'patch',
    baseDocumentRevision: 1,
    changes: [
      {
        kind: 'upsert',
        sheetId: 'sheet-1',
        row: 0,
        column: 0,
        value: { kind: 'number', value: 4 },
      },
    ],
  },
  calculation: { kind: 'dirty' },
});
assert(
  spreadsheetPatchResult.cells?.[0]?.value?.value === 7 &&
    spreadsheetPatchResult.cells?.[1]?.value?.value === 14 &&
    spreadsheetPatchResult.stats?.updateKind === 'patch' &&
    spreadsheetPatchResult.stats?.dirtyFormulaCellCount === 2 &&
    spreadsheetPatchResult.stats?.evaluatedFormulaCellCount === 2,
  'Office Spreadsheet session did not limit recalculation to the dirty dependency subgraph.',
);

const deepSpreadsheetResult = calculateSpreadsheetWithWasm({
  ...spreadsheetRequest,
  requestId: 25,
  sheets: [
    {
      id: 'sheet-1',
      name: 'Sheet 1',
      cells: Array.from({ length: 80 }, (_, row) => ({
        row,
        column: 0,
        formula: row === 0 ? '=1' : `=A${row}+1`,
        value: { kind: 'number', value: row + 1 },
      })),
    },
  ],
  targets: [{ sheetId: 'sheet-1', row: 79, column: 0 }],
});
assert(
  deepSpreadsheetResult.issues?.some(
    (issue) =>
      issue.code === 'office.kernel.spreadsheet.dependency_depth_exceeded',
  ),
  'Office Spreadsheet dependency depth was not bounded inside WASM.',
);

console.log('Office kernel WASM ABI smoke test passed.');

function calculateSpreadsheetWithWasm(
  request: unknown,
): SpreadsheetSmokeResult {
  const input = new TextEncoder().encode(JSON.stringify(request));
  const pointer = exports.office_kernel_alloc(input.byteLength);
  try {
    new Uint8Array(exports.memory.buffer, pointer, input.byteLength).set(input);
    assert(
      exports.office_kernel_spreadsheet_calculation(
        pointer,
        input.byteLength,
      ) === 0,
      'Office kernel Spreadsheet calculation call failed.',
    );
  } finally {
    exports.office_kernel_dealloc(pointer, input.byteLength);
  }
  const output = new Uint8Array(
    exports.memory.buffer,
    exports.office_kernel_result_pointer(),
    exports.office_kernel_result_length(),
  ).slice();
  return JSON.parse(new TextDecoder().decode(output)) as SpreadsheetSmokeResult;
}

function calculateSpreadsheetSessionWithWasm(
  request: unknown,
): SpreadsheetSmokeResult {
  const input = new TextEncoder().encode(JSON.stringify(request));
  const pointer = exports.office_kernel_alloc(input.byteLength);
  let status = 1;
  try {
    new Uint8Array(exports.memory.buffer, pointer, input.byteLength).set(input);
    status = exports.office_kernel_spreadsheet_session_calculation(
      pointer,
      input.byteLength,
    );
  } finally {
    exports.office_kernel_dealloc(pointer, input.byteLength);
  }
  const output = new Uint8Array(
    exports.memory.buffer,
    exports.office_kernel_result_pointer(),
    exports.office_kernel_result_length(),
  ).slice();
  const result = JSON.parse(
    new TextDecoder().decode(output),
  ) as SpreadsheetSmokeResult;
  assert(
    status === 0,
    `Office kernel Spreadsheet session calculation call failed: ${JSON.stringify(result)}`,
  );
  return result;
}

async function registerFont(id: string, url: URL): Promise<void> {
  const encodedId = new TextEncoder().encode(id);
  const data = new Uint8Array(await Bun.file(url).arrayBuffer());
  const idPointer = exports.office_kernel_alloc(encodedId.byteLength);
  const dataPointer = exports.office_kernel_alloc(data.byteLength);
  try {
    new Uint8Array(exports.memory.buffer, idPointer, encodedId.byteLength).set(
      encodedId,
    );
    new Uint8Array(exports.memory.buffer, dataPointer, data.byteLength).set(
      data,
    );
    assert(
      exports.office_kernel_register_font(
        idPointer,
        encodedId.byteLength,
        dataPointer,
        data.byteLength,
      ) === 0,
      `Office kernel rejected font '${id}'.`,
    );
  } finally {
    exports.office_kernel_dealloc(idPointer, encodedId.byteLength);
    exports.office_kernel_dealloc(dataPointer, data.byteLength);
  }
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}
