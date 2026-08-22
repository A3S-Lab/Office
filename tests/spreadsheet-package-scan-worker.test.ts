import { describe, expect, test } from '@rstest/core';
import { scanSpreadsheetPackageInWorker } from '../src/internal/features/work/work-spreadsheet-package-scan-worker-client';
import {
  createPlainXlsxWorkbookPlan,
  streamPlainXlsxWorksheet,
  type PlainXlsxCellChunk,
} from '../src/internal/features/work/work-xlsx-plain-fast-path';
import { scanXlsxWorksheetXml } from '../src/internal/features/work/work-xlsx-worksheet-scan';

describe('spreadsheet package scan worker', () => {
  test('classifies all worksheet gates in one namespace-safe scan', () => {
    expect(
      scanXlsxWorksheetXml(
        '<worksheet><sheetData><row r="1"><c><v>cols f pane customHeight</v></c></row></sheetData></worksheet>',
      ),
    ).toEqual({
      hasDirectCellStyles: false,
      hasDiagnosticFeatures: false,
      hasFormulaFeatures: false,
      hasImportedFeatures: false,
      hasRichTextCells: false,
      requiresSheetJsCellStyles: false,
    });

    expect(
      scanXlsxWorksheetXml(
        [
          '<x:worksheet xmlns:x="urn:test">',
          '<x:cols><x:col min="1" max="1"/></x:cols>',
          '<x:sheetData><x:row r="1" customHeight="1" ht="30">',
          '<x:c r="A1" s="5"><x:f>SUM(A2:A3)</x:f></x:c>',
          '</x:row></x:sheetData>',
          '<x:pane/><x:dataValidation/><x:drawing/>',
          '</x:worksheet>',
        ].join(''),
      ),
    ).toEqual({
      hasDirectCellStyles: true,
      hasDiagnosticFeatures: true,
      hasFormulaFeatures: true,
      hasImportedFeatures: true,
      hasRichTextCells: false,
      requiresSheetJsCellStyles: true,
    });

    expect(
      scanXlsxWorksheetXml(
        '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>Rich</t></r></is></c></row></sheetData></worksheet>',
      ),
    ).toMatchObject({ hasRichTextCells: true });
  });

  test('returns typed scans and terminates after transferring the source', async () => {
    const worker = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() =>
          instance.emitMessage({
            fastPath: false,
            kind: 'success',
            worksheets: {
              'xl/worksheets/sheet1.xml': {
                hasDirectCellStyles: false,
                hasDiagnosticFeatures: false,
                hasFormulaFeatures: false,
                hasImportedFeatures: false,
                hasRichTextCells: false,
                requiresSheetJsCellStyles: false,
              },
            },
          }),
        );
      },
      async () => {
        await expect(
          scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
        ).resolves.toEqual({
          plainWorksheets: null,
          workbook: null,
          worksheets: {
            'xl/worksheets/sheet1.xml': {
              hasDirectCellStyles: false,
              hasDiagnosticFeatures: false,
              hasFormulaFeatures: false,
              hasImportedFeatures: false,
              hasRichTextCells: false,
              requiresSheetJsCellStyles: false,
            },
          },
        });
        return RecordingPackageScanWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
    expect(worker?.transferCount).toBe(1);
  });

  test('reconstructs provisional columnar cells only after authentication', async () => {
    let candidateCount = 0;
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            columnCount: 1,
            kind: 'plain-worksheet-start',
            name: 'Plain',
            rowCount: 1,
          });
          instance.emitMessage({
            chunk: {
              coordinates: new Uint32Array([0]),
              kinds: new Uint8Array([3]),
              numericValues: new Float64Array(),
              startRow: 0,
              textValues: ['A1'],
            },
            kind: 'plain-cells',
            name: 'Plain',
          });
          instance.emitMessage({
            columnCount: 1,
            dense: true,
            kind: 'worksheet',
            name: 'Plain',
            populatedCellCount: 1,
            properties: { '!ref': 'A1:A1' },
            rowCount: 1,
          });
          instance.emitMessage({
            fastPath: true,
            kind: 'success',
            worksheets: {
              'xl/worksheets/sheet1.xml': {
                hasDirectCellStyles: false,
                hasDiagnosticFeatures: false,
                hasFormulaFeatures: false,
                hasImportedFeatures: false,
                hasRichTextCells: false,
                requiresSheetJsCellStyles: false,
              },
            },
          });
        });
      },
      () =>
        scanSpreadsheetPackageInWorker(new ArrayBuffer(16), undefined, () => {
          candidateCount += 1;
        }),
    );

    expect(candidateCount).toBe(1);
    expect(result?.workbook).toEqual({
      SheetNames: ['Plain'],
      Sheets: {
        Plain: Object.assign([], { '!ref': 'A1:A1' }),
      },
    });
    expect(result?.plainWorksheets).toEqual({
      Plain: {
        columnCount: 1,
        data: [[{ v: 'A1' }]],
        populatedCellCount: 1,
        rowCount: 1,
      },
    });
    expect(Object.isFrozen(result?.plainWorksheets?.Plain?.data[0]?.[0])).toBe(
      true,
    );
  });

  test('authenticates logical geometry before cell chunks arrive', async () => {
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            columnCount: 10,
            kind: 'plain-worksheet-start',
            name: 'Plain',
            rowCount: 100_000,
          });
          instance.emitMessage({
            chunk: {
              coordinates: new Uint32Array([0]),
              kinds: new Uint8Array([0]),
              numericValues: new Float64Array([1]),
              startRow: 0,
              textValues: [],
            },
            kind: 'plain-cells',
            name: 'Plain',
          });
          instance.emitMessage({
            columnCount: 10,
            dense: true,
            kind: 'worksheet',
            name: 'Plain',
            populatedCellCount: 1,
            properties: { '!ref': 'A1:J100000' },
            rowCount: 100_000,
          });
          instance.emitMessage({
            fastPath: true,
            kind: 'success',
            worksheets: {},
          });
        });
      },
      () => scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
    );

    expect(result?.plainWorksheets?.Plain).toMatchObject({
      columnCount: 10,
      populatedCellCount: 1,
      rowCount: 100_000,
    });
    expect(result?.plainWorksheets?.Plain?.data).toEqual([[{ v: 1 }]]);
  });

  test('rejects duplicate coordinates using the monotonic stream invariant', async () => {
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            columnCount: 1,
            kind: 'plain-worksheet-start',
            name: 'Plain',
            rowCount: 1,
          });
          instance.emitMessage({
            chunk: {
              coordinates: new Uint32Array([0, 0]),
              kinds: new Uint8Array([0, 0]),
              numericValues: new Float64Array([1, 2]),
              startRow: 0,
              textValues: [],
            },
            kind: 'plain-cells',
            name: 'Plain',
          });
          instance.emitMessage({
            fastPath: true,
            kind: 'success',
            worksheets: {},
          });
        });
      },
      () => scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
    );

    expect(result).toEqual({
      plainWorksheets: null,
      workbook: null,
      worksheets: {},
    });
  });

  test('rejects malformed columnar payloads without exposing provisional cells', async () => {
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            columnCount: 1,
            kind: 'plain-worksheet-start',
            name: 'Plain',
            rowCount: 1,
          });
          instance.emitMessage({
            chunk: {
              coordinates: new Uint32Array([0]),
              kinds: new Uint8Array([3]),
              numericValues: new Float64Array([1]),
              startRow: 0,
              textValues: ['ambiguous'],
            },
            kind: 'plain-cells',
            name: 'Plain',
          });
          instance.emitMessage({
            fastPath: true,
            kind: 'success',
            worksheets: {},
          });
        });
      },
      () => scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
    );

    expect(result).toEqual({
      plainWorksheets: null,
      workbook: null,
      worksheets: {},
    });
  });

  test('rejects duplicate workbook declarations before authentication', async () => {
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          emitAuthenticatedPlainWorksheet(instance);
        });
      },
      () => scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
    );

    expect(result?.workbook).toBeNull();
    expect(result?.plainWorksheets).toBeNull();
  });

  test('rejects worksheet extents that disagree with the authenticated reference', async () => {
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            columnCount: 2,
            kind: 'plain-worksheet-start',
            name: 'Plain',
            rowCount: 1,
          });
          instance.emitMessage({
            chunk: {
              coordinates: new Uint32Array([0]),
              kinds: new Uint8Array([0]),
              numericValues: new Float64Array([1]),
              startRow: 0,
              textValues: [],
            },
            kind: 'plain-cells',
            name: 'Plain',
          });
          instance.emitMessage({
            columnCount: 2,
            dense: true,
            kind: 'worksheet',
            name: 'Plain',
            populatedCellCount: 1,
            properties: { '!ref': 'A1:A1' },
            rowCount: 1,
          });
          instance.emitMessage({
            fastPath: true,
            kind: 'success',
            worksheets: {},
          });
        });
      },
      () => scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
    );

    expect(result?.workbook).toBeNull();
    expect(result?.plainWorksheets).toBeNull();
  });

  test('rejects cell chunks after a worksheet terminator', async () => {
    const result = await withPackageScanWorker(
      (instance) => {
        queueMicrotask(() => {
          instance.emitMessage({
            kind: 'workbook',
            workbook: { SheetNames: ['Plain'] },
          });
          instance.emitMessage({
            columnCount: 1,
            kind: 'plain-worksheet-start',
            name: 'Plain',
            rowCount: 1,
          });
          instance.emitMessage({
            columnCount: 1,
            dense: true,
            kind: 'worksheet',
            name: 'Plain',
            populatedCellCount: 1,
            properties: { '!ref': 'A1:A1' },
            rowCount: 1,
          });
          instance.emitMessage({
            chunk: {
              coordinates: new Uint32Array([0]),
              kinds: new Uint8Array([0]),
              numericValues: new Float64Array([1]),
              startRow: 0,
              textValues: [],
            },
            kind: 'plain-cells',
            name: 'Plain',
          });
          instance.emitMessage({
            fastPath: true,
            kind: 'success',
            worksheets: {},
          });
        });
      },
      () => scanSpreadsheetPackageInWorker(new ArrayBuffer(16)),
    );

    expect(result?.workbook).toBeNull();
    expect(result?.plainWorksheets).toBeNull();
  });

  test('cancels only its active worker with the caller signal', async () => {
    const controller = new AbortController();
    const worker = await withPackageScanWorker(
      () => undefined,
      async () => {
        const pending = scanSpreadsheetPackageInWorker(
          new ArrayBuffer(16),
          controller.signal,
        );
        controller.abort();
        await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
        return RecordingPackageScanWorker.instance;
      },
    );

    expect(worker?.terminated).toBe(true);
  });

  test('builds a minimal workbook plan only for a closed plain-XLSX package', () => {
    const plan = createPlainXlsxWorkbookPlan(plainXlsxPackageParts());

    expect(plan?.sheets).toEqual([
      {
        name: 'Plain & safe',
        partPath: 'xl/worksheets/sheet1.xml',
      },
    ]);
    expect(plan?.workbook).toMatchObject({
      SheetNames: ['Plain & safe'],
      Workbook: {
        Names: [],
        Sheets: [{ Hidden: 0, id: 'rId1', name: 'Plain & safe' }],
      },
    });

    expect(
      createPlainXlsxWorkbookPlan({
        ...plainXlsxPackageParts(),
        packagePaths: [
          ...plainXlsxPackageParts().packagePaths,
          'xl/styles.xml',
        ],
      }),
    ).toBeNull();
    expect(
      createPlainXlsxWorkbookPlan({
        ...plainXlsxPackageParts(),
        workbook: plainXlsxPackageParts().workbook.replace(
          '</workbook>',
          '<definedNames/></workbook>',
        ),
      }),
    ).toBeNull();
  });

  test('streams bounded dense rows for primitive inline worksheet cells', () => {
    const chunks: PlainXlsxCellChunk[] = [];
    const result = streamPlainXlsxWorksheet(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:D3"/>',
        '<sheetData>',
        '<row r="1">',
        '<c r="A1" t="inlineStr"><is><t>A &amp; B</t></is></c>',
        '<c r="B1"><v>-1.25e2</v></c>',
        '<c r="C1" t="b"><v>1</v></c>',
        '<c r="D1" t="e"><v>#DIV/0!</v></c>',
        '</row>',
        '<row r="3"><c r="A3" t="inlineStr"><is><t>Tail</t></is></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
      (chunk) => chunks.push(chunk),
    );

    expect(result).toEqual({
      columnCount: 4,
      populatedCellCount: 5,
      properties: { '!ref': 'A1:D3' },
      rowCount: 3,
    });
    expect(chunks).toEqual([
      {
        coordinates: new Uint32Array([0, 1, 2, 3, 32_768]),
        kinds: new Uint8Array([3, 0, 1, 2, 3]),
        numericValues: new Float64Array([-125, 1, 7]),
        startRow: 0,
        textValues: ['A & B', 'Tail'],
      },
    ]);
  });

  test('bounds columnar chunks by both row span and populated cells', () => {
    const rowChunks: PlainXlsxCellChunk[] = [];
    const rowXml = Array.from(
      { length: 257 },
      (_, index) =>
        `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`,
    ).join('');
    expect(
      streamPlainXlsxWorksheet(plainWorksheet('A1:A257', rowXml), (chunk) =>
        rowChunks.push(chunk),
      ),
    ).toMatchObject({ populatedCellCount: 257 });
    expect(
      rowChunks.map((chunk) => [chunk.startRow, chunk.coordinates.length]),
    ).toEqual([
      [0, 256],
      [256, 1],
    ]);

    const cellChunks: PlainXlsxCellChunk[] = [];
    const cells = Array.from({ length: 4_097 }, (_, column) => {
      const address = spreadsheetCellAddress(0, column);
      return `<c r="${address}"><v>${column}</v></c>`;
    }).join('');
    expect(
      streamPlainXlsxWorksheet(
        plainWorksheet('A1:FAO1', `<row r="1">${cells}</row>`),
        (chunk) => cellChunks.push(chunk),
      ),
    ).toMatchObject({ populatedCellCount: 4_097 });
    expect(cellChunks.map((chunk) => chunk.coordinates.length)).toEqual([
      4_096, 1,
    ]);
  });

  test('rejects unsupported or ambiguous worksheet XML instead of guessing', () => {
    const worksheet = (cell: string) =>
      [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A1"/><sheetData><row r="1">',
        cell,
        '</row></sheetData></worksheet>',
      ].join('');

    expect(
      streamPlainXlsxWorksheet(
        worksheet('<c r="A1"><f>1+1</f><v>2</v></c>'),
        () => undefined,
      ),
    ).toBeNull();
    expect(
      streamPlainXlsxWorksheet(
        worksheet('<c r="A1" s="1"><v>2</v></c>'),
        () => undefined,
      ),
    ).toBeNull();
    expect(
      streamPlainXlsxWorksheet(
        worksheet('<c r="A1" t="s"><v>0</v></c>'),
        () => undefined,
      ),
    ).toBeNull();
    expect(
      streamPlainXlsxWorksheet(
        worksheet('<c r="A1" t="inlineStr"><is><t>broken &copy;</t></is></c>'),
        () => undefined,
      ),
    ).toBeNull();
    expect(
      streamPlainXlsxWorksheet(
        worksheet(
          '<c r="A1" t="inlineStr"><is><t>broken &#1junk;</t></is></c>',
        ),
        () => undefined,
      ),
    ).toBeNull();
    expect(
      streamPlainXlsxWorksheet(
        worksheet('<c r="A1" t="inlineStr"><is><t><r>nested</r></t></is></c>'),
        () => undefined,
      ),
    ).toBeNull();
    for (const number of [
      '',
      '+',
      '.',
      '1e',
      '1e+',
      'Infinity',
      '0x10',
      ' 1',
      '1 ',
      '1_0',
    ]) {
      expect(
        streamPlainXlsxWorksheet(
          worksheet(`<c r="A1"><v>${number}</v></c>`),
          () => undefined,
        ),
      ).toBeNull();
    }
  });

  test('authenticates cell coordinates without accepting ambiguous address syntax', () => {
    const worksheet = (address: string) =>
      plainWorksheet(
        'A1:A1',
        `<row r="1"><c r="${address}"><v>1</v></c></row>`,
      );

    for (const address of [
      'a1',
      'A0',
      'A01',
      'A2',
      'A1x',
      'AAAA1',
      'XFE1',
      'A1048577',
    ]) {
      expect(
        streamPlainXlsxWorksheet(worksheet(address), () => undefined),
      ).toBeNull();
    }

    expect(
      streamPlainXlsxWorksheet(
        plainWorksheet(
          'XFD1048576:XFD1048576',
          '<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>',
        ),
        () => undefined,
      ),
    ).toMatchObject({
      columnCount: 16_384,
      populatedCellCount: 1,
      rowCount: 1_048_576,
    });
  });

  test('rejects ambiguous worksheet row numbers', () => {
    for (const row of ['0', '01', '1048577']) {
      expect(
        streamPlainXlsxWorksheet(
          plainWorksheet(
            'A1:A1',
            `<row r="${row}"><c r="A1"><v>1</v></c></row>`,
          ),
          () => undefined,
        ),
      ).toBeNull();
    }
  });
});

function plainXlsxPackageParts() {
  return {
    contentTypes: [
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
      '</Types>',
    ].join(''),
    packagePaths: [
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml',
    ],
    rootRelationships: [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
      '</Relationships>',
    ].join(''),
    workbook: [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<sheets><sheet name="Plain &amp; safe" sheetId="1" r:id="rId1"/></sheets>',
      '</workbook>',
    ].join(''),
    workbookRelationships: [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
      '</Relationships>',
    ].join(''),
  };
}

function plainWorksheet(reference: string, rows: string): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="${reference}"/><sheetData>`,
    rows,
    '</sheetData></worksheet>',
  ].join('');
}

function spreadsheetCellAddress(row: number, column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return `${label}${row + 1}`;
}

function emitAuthenticatedPlainWorksheet(
  instance: RecordingPackageScanWorker,
): void {
  instance.emitMessage({
    columnCount: 1,
    kind: 'plain-worksheet-start',
    name: 'Plain',
    rowCount: 1,
  });
  instance.emitMessage({
    chunk: {
      coordinates: new Uint32Array([0]),
      kinds: new Uint8Array([0]),
      numericValues: new Float64Array([1]),
      startRow: 0,
      textValues: [],
    },
    kind: 'plain-cells',
    name: 'Plain',
  });
  instance.emitMessage({
    columnCount: 1,
    dense: true,
    kind: 'worksheet',
    name: 'Plain',
    populatedCellCount: 1,
    properties: { '!ref': 'A1:A1' },
    rowCount: 1,
  });
  instance.emitMessage({
    fastPath: true,
    kind: 'success',
    worksheets: {},
  });
}

type WorkerPost = (worker: RecordingPackageScanWorker) => void;

class RecordingPackageScanWorker {
  static instance: RecordingPackageScanWorker | null = null;
  static post: WorkerPost = () => undefined;

  readonly listeners = new Map<string, Set<EventListener>>();
  terminated = false;
  transferCount = 0;

  constructor() {
    RecordingPackageScanWorker.instance = this;
  }

  addEventListener(type: string, listener: EventListener): void {
    const listeners = this.listeners.get(type) ?? new Set<EventListener>();
    listeners.add(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type: string, listener: EventListener): void {
    this.listeners.get(type)?.delete(listener);
  }

  postMessage(_message: unknown, options?: StructuredSerializeOptions): void {
    this.transferCount = options?.transfer?.length ?? 0;
    RecordingPackageScanWorker.post(this);
  }

  terminate(): void {
    this.terminated = true;
  }

  emitMessage(data: unknown): void {
    const event = new MessageEvent('message', { data });
    for (const listener of this.listeners.get('message') ?? []) listener(event);
  }
}

async function withPackageScanWorker<T>(
  post: WorkerPost,
  run: () => Promise<T>,
): Promise<T> {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  RecordingPackageScanWorker.instance = null;
  RecordingPackageScanWorker.post = post;
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: RecordingPackageScanWorker as unknown as typeof Worker,
  });
  try {
    return await run();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'Worker', descriptor);
    else Reflect.deleteProperty(globalThis, 'Worker');
  }
}
