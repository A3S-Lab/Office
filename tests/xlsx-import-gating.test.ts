import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import type { WorkBook, WorkSheet } from 'xlsx';
import { importOfficeFile } from '../src/core';
import { analyzeSpreadsheetCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';
import { readXlsxFormulaFeaturesFromPackage } from '../src/internal/features/work/work-xlsx-formulas';
import { readXlsxSheetFeaturesFromPackage } from '../src/internal/features/work/work-xlsx-interop';
import { xlsxWorksheetRequiresSheetJsCellStyles } from '../src/internal/features/work/work-xlsx-style-gate';
import { scanXlsxWorksheetXml } from '../src/internal/features/work/work-xlsx-worksheet-scan';

describe('XLSX import feature gating', () => {
  test('detects only worksheet layout that requires SheetJS styles', () => {
    expect(
      xlsxWorksheetRequiresSheetJsCellStyles(
        '<worksheet><sheetData><row r="1"><c><v>cols</v></c></row></sheetData></worksheet>',
      ),
    ).toBe(false);
    expect(
      xlsxWorksheetRequiresSheetJsCellStyles(
        '<worksheet><cols><col min="1" max="1" width="20"/></cols></worksheet>',
      ),
    ).toBe(true);
    expect(
      xlsxWorksheetRequiresSheetJsCellStyles(
        '<worksheet><sheetData><row r="7" ht="30" customHeight="1"/></sheetData></worksheet>',
      ),
    ).toBe(true);
  });

  test('does not build a worksheet DOM when cell data has no XML-only features', async () => {
    const bytes = await plainWorksheetPackage();
    const archive = await OoxmlPackage.load(bytes);
    const worksheet = {
      '!ref': 'A1:A2',
      A1: { t: 's', v: 'Alpha' },
      A2: { t: 's', v: 'Beta' },
    } as WorkSheet;
    const workbook: WorkBook = {
      SheetNames: ['Plain'],
      Sheets: { Plain: worksheet },
    };
    const worksheetScans = {
      'xl/worksheets/sheet1.xml': scanXlsxWorksheetXml(plainWorksheetXml()),
    };
    const readArchiveText = archive.text.bind(archive);
    let worksheetTextReads = 0;
    archive.text = async (partPath) => {
      if (partPath.startsWith('xl/worksheets/')) worksheetTextReads += 1;
      return readArchiveText(partPath);
    };
    const original = globalThis.DOMParser;
    let worksheetParseCount = 0;
    class RecordingDomParser extends original {
      override parseFromString(source: string, type: DOMParserSupportedType) {
        if (source.includes('<sheetData')) worksheetParseCount += 1;
        return super.parseFromString(source, type);
      }
    }
    Object.defineProperty(globalThis, 'DOMParser', {
      configurable: true,
      value: RecordingDomParser,
    });
    try {
      const sheetFeatures = await readXlsxSheetFeaturesFromPackage(
        archive,
        worksheetScans,
      );
      const formulaFeatures = await readXlsxFormulaFeaturesFromPackage(
        archive,
        worksheetScans,
      );
      await analyzeSpreadsheetCompatibility(
        new File([bytes], 'plain.xlsx'),
        'xlsx',
        workbook,
        archive,
        undefined,
        { formulaFeatures, worksheetScans },
      );

      expect(sheetFeatures.get('Plain')).toMatchObject({
        charts: [],
        conditionalFormats: [],
        images: [],
        validations: [],
      });
      expect(formulaFeatures.sheets.get('Plain')).toMatchObject({
        formulas: [],
        ranges: [],
      });
      expect(worksheetTextReads).toBe(0);
      expect(worksheetParseCount).toBe(0);
    } finally {
      Object.defineProperty(globalThis, 'DOMParser', {
        configurable: true,
        value: original,
      });
    }
  });

  test('preserves custom row and column layout without a styles part', async () => {
    const bytes = await layoutWorksheetPackage();
    const artifact = await importOfficeFile(
      new File([bytes], 'layout.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
    );
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }

    expect(artifact.content.sheets[0]?.config).toMatchObject({
      columnlen: { 0: 120 },
      rowlen: { 0: 30 },
    });
  });

  test('restarts SheetJS after a package candidate rejects formula cells', async () => {
    const bytes = await formulaWorksheetPackage();
    const importWorkers = await withCandidateRejectingWorkers(async () => {
      const artifact = await importOfficeFile(
        new File([bytes], 'formula-fallback.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
      );
      if (artifact.content.type !== 'spreadsheet') {
        throw new Error('Expected a spreadsheet artifact.');
      }
      expect(artifact.content.sheets[0]?.data?.[0]?.[0]).toMatchObject({
        f: '=1+1',
        v: 2,
      });
      return CandidateRejectingWorker.instances.filter(
        (worker) => worker.name === 'a3s-office-spreadsheet-import',
      );
    });

    expect(importWorkers).toHaveLength(2);
    expect(importWorkers[0]?.terminated).toBe(true);
    expect(importWorkers[1]?.terminated).toBe(true);
  });
});

async function layoutWorksheetPackage(): Promise<ArrayBuffer> {
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
      '<sheets><sheet name="Layout" sheetId="1" r:id="rId1"/></sheets>',
      '</workbook>',
    ].join(''),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1"',
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"',
      ' Target="worksheets/sheet1.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    [
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A1"/>',
      '<cols><col min="1" max="1" width="20" customWidth="1"/></cols>',
      '<sheetData><row r="1" ht="30" customHeight="1">',
      '<c r="A1" t="inlineStr"><is><t>Layout</t></is></c>',
      '</row></sheetData>',
      '</worksheet>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function plainWorksheetPackage(): Promise<ArrayBuffer> {
  const zip = new JSZip();
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
      '<Relationship Id="rId1"',
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"',
      ' Target="worksheets/sheet1.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file('xl/worksheets/sheet1.xml', plainWorksheetXml());
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function formulaWorksheetPackage(): Promise<ArrayBuffer> {
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
      '<sheets><sheet name="Formula" sheetId="1" r:id="rId1"/></sheets>',
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
      '<dimension ref="A1"/><sheetData><row r="1">',
      '<c r="A1"><f>1+1</f><v>2</v></c>',
      '</row></sheetData></worksheet>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

function plainWorksheetXml(): string {
  return [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:A2"/>',
    '<sheetData>',
    '<row r="1"><c r="A1" t="inlineStr"><is><t>Alpha</t></is></c></row>',
    '<row r="2"><c r="A2" t="inlineStr"><is><t>Beta</t></is></c></row>',
    '</sheetData>',
    '</worksheet>',
  ].join('');
}

class CandidateRejectingWorker {
  static instances: CandidateRejectingWorker[] = [];

  readonly listeners = new Map<string, Set<EventListener>>();
  readonly name: string;
  terminated = false;

  constructor(_url: URL, options?: WorkerOptions) {
    this.name = options?.name ?? '';
    CandidateRejectingWorker.instances.push(this);
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
    if (this.name === 'a3s-office-spreadsheet-package-scan') {
      queueMicrotask(() => {
        this.emitMessage({
          kind: 'workbook',
          workbook: { SheetNames: ['Formula'] },
        });
        this.emitMessage({ kind: 'fast-path-rejected' });
        this.emitMessage({
          fastPath: false,
          kind: 'success',
          worksheets: {
            'xl/worksheets/sheet1.xml': {
              hasDirectCellStyles: false,
              hasDiagnosticFeatures: false,
              hasFormulaFeatures: true,
              hasImportedFeatures: false,
              requiresSheetJsCellStyles: false,
            },
          },
        });
      });
      return;
    }
    const importWorkerCount = CandidateRejectingWorker.instances.filter(
      (worker) => worker.name === 'a3s-office-spreadsheet-import',
    ).length;
    if (importWorkerCount > 1) {
      queueMicrotask(() => this.emitMessage({ kind: 'failure' }));
    }
  }

  terminate(): void {
    this.terminated = true;
  }

  private emitMessage(data: unknown): void {
    const event = new MessageEvent('message', { data });
    for (const listener of this.listeners.get('message') ?? []) listener(event);
  }
}

async function withCandidateRejectingWorkers<T>(
  run: () => Promise<T>,
): Promise<T> {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  CandidateRejectingWorker.instances = [];
  Object.defineProperty(globalThis, 'Worker', {
    configurable: true,
    value: CandidateRejectingWorker as unknown as typeof Worker,
  });
  try {
    return await run();
  } finally {
    if (descriptor) Object.defineProperty(globalThis, 'Worker', descriptor);
    else Reflect.deleteProperty(globalThis, 'Worker');
  }
}
