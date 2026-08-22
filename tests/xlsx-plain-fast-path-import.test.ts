import { expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';

test('adopts authenticated columnar cells as the controlled matrix', async () => {
  const bytes = await primitiveWorksheetPackage();
  const workers = await withPlainFastPathWorkers(async () => {
    const artifact = await importOfficeFile(
      new File([bytes], 'primitive.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      {
        artifactId: 'artifact-reserved-spreadsheet',
        spreadsheetSheetIds: ['sheet-reserved-spreadsheet'],
      },
    );
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }
    const data = artifact.content.sheets[0]?.data;
    expect(artifact.id).toBe('artifact-reserved-spreadsheet');
    expect(artifact.content.sheets[0]?.id).toBe('sheet-reserved-spreadsheet');
    expect(data?.[0]?.slice(0, 4)).toEqual([
      { v: 'Alpha' },
      { v: 42 },
      { v: true },
      { ct: { t: 'e' }, m: '#DIV/0!', v: 7 },
    ]);
    expect(Object.isFrozen(data?.[0]?.[0])).toBe(true);
    expect(Object.isFrozen(data?.[0])).toBe(true);
    expect(Object.isFrozen(data)).toBe(true);
    return PlainFastPathWorker.instances;
  });

  const importWorkers = workers.filter(
    (worker) => worker.name === 'a3s-office-spreadsheet-import',
  );
  const packageWorkers = workers.filter(
    (worker) => worker.name === 'a3s-office-spreadsheet-package-scan',
  );
  expect(importWorkers).toHaveLength(1);
  expect(importWorkers[0]?.terminated).toBe(true);
  expect(packageWorkers).toHaveLength(1);
  expect(packageWorkers[0]?.terminated).toBe(true);
});

async function primitiveWorksheetPackage(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    [
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
      '</Types>',
    ].join(''),
  );
  zip.file(
    '_rels/.rels',
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/workbook.xml',
    [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<sheets><sheet name="Plain" sheetId="1" r:id="rId1"/></sheets>',
      '</workbook>',
    ].join(''),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    [
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A1:D1"/><sheetData><row r="1">',
      '<c r="A1" t="inlineStr"><is><t>Alpha</t></is></c>',
      '<c r="B1"><v>42</v></c>',
      '<c r="C1" t="b"><v>1</v></c>',
      '<c r="D1" t="e"><v>#DIV/0!</v></c>',
      '</row></sheetData></worksheet>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

class PlainFastPathWorker {
  static instances: PlainFastPathWorker[] = [];

  readonly listeners = new Map<string, Set<EventListener>>();
  readonly name: string;
  terminated = false;

  constructor(_url: URL, options?: WorkerOptions) {
    this.name = options?.name ?? '';
    PlainFastPathWorker.instances.push(this);
  }

  addEventListener(type: string, listener: EventListener): void {
    const listeners = this.listeners.get(type) ?? new Set<EventListener>();
    listeners.add(listener);
    this.listeners.set(type, listeners);
  }

  removeEventListener(type: string, listener: EventListener): void {
    this.listeners.get(type)?.delete(listener);
  }

  postMessage(): void {
    if (this.name !== 'a3s-office-spreadsheet-package-scan') return;
    queueMicrotask(() => {
      this.emitMessage({
        kind: 'workbook',
        workbook: {
          SheetNames: ['Plain'],
          Workbook: {
            Names: [],
            Sheets: [
              {
                Hidden: 0,
                id: 'rId1',
                name: 'Plain',
                sheetId: '1',
                sheetid: '1',
              },
            ],
          },
        },
      });
      this.emitMessage({
        columnCount: 4,
        kind: 'plain-worksheet-start',
        name: 'Plain',
        rowCount: 1,
      });
      this.emitMessage({
        chunk: {
          coordinates: new Uint32Array([0, 1, 2, 3]),
          kinds: new Uint8Array([3, 0, 1, 2]),
          numericValues: new Float64Array([42, 1, 7]),
          startRow: 0,
          textValues: ['Alpha'],
        },
        kind: 'plain-cells',
        name: 'Plain',
      });
      this.emitMessage({
        columnCount: 4,
        dense: true,
        kind: 'worksheet',
        name: 'Plain',
        populatedCellCount: 4,
        properties: { '!ref': 'A1:D1' },
        rowCount: 1,
      });
      this.emitMessage({
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
  }

  terminate(): void {
    this.terminated = true;
  }

  private emitMessage(data: unknown): void {
    const event = new MessageEvent('message', { data });
    for (const listener of this.listeners.get('message') ?? []) listener(event);
  }
}

async function withPlainFastPathWorkers<T>(run: () => Promise<T>): Promise<T> {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  PlainFastPathWorker.instances = [];
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: PlainFastPathWorker as unknown as typeof Worker,
  });
  try {
    return await run();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'Worker', descriptor);
    else Reflect.deleteProperty(globalThis, 'Worker');
  }
}
