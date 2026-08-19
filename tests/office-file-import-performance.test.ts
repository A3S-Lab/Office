import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { checkCellIsLocked, type Context } from '@fortune-sheet/core';
import * as XLSX from 'xlsx';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
  type OfficeFileImportProgress,
} from '../src/core';
import {
  sheetProtectionAuthority,
  withEditableRange,
  withSheetProtection,
} from '../src/internal/features/work/work-spreadsheet-protection';
import { WorkFileImportController } from '../src/internal/features/work/work-file-import';
import { documentModelUsesWindowing } from '../src/internal/features/work/work-document-windowing';

describe('office file import performance', () => {
  test('yields through the browser scheduler when it is available', async () => {
    const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'scheduler');
    let yieldCount = 0;
    Object.defineProperty(globalThis, 'scheduler', {
      configurable: true,
      value: {
        yield: async () => {
          yieldCount += 1;
        },
      },
    });
    try {
      await new WorkFileImportController({}, 0).yieldToMainThread();
      expect(yieldCount).toBe(1);
    } finally {
      if (descriptor) {
        Object.defineProperty(globalThis, 'scheduler', descriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'scheduler');
      }
    }
  });

  test('formats only cells that need non-General display text', async () => {
    const worksheet = XLSX.utils.aoa_to_sheet([
      [0.25, new Date('2026-08-18T00:00:00.000Z')],
      ['Plain text', 42],
    ]);
    if (!worksheet.A1 || !worksheet.B1) {
      throw new Error('Expected formatted worksheet cells.');
    }
    worksheet.A1.z = '0.0%';
    worksheet.B1.z = 'yyyy-mm-dd';
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Formatted');
    const bytes = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    }) as ArrayBuffer;

    const artifact = await importOfficeFile(
      new File([bytes], 'formatted.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
    );
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }

    expect(artifact.content.sheets[0]?.data?.[0]?.[0]).toMatchObject({
      m: '25.0%',
      v: 0.25,
    });
    expect(artifact.content.sheets[0]?.data?.[0]?.[1]).toMatchObject({
      m: '2026-08-18',
    });
    expect(artifact.content.sheets[0]?.data?.[1]?.[0]).toEqual({
      v: 'Plain text',
    });
    expect(artifact.content.sheets[0]?.data?.[1]?.[1]).toEqual({ v: 42 });
  });

  test('reads an XLSX source once and keeps a maximum worksheet range sparse', async () => {
    const buffer = await sparseWorkbookBuffer();
    const source = new File([buffer], 'maximum-range.xlsx', {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    let sourceReads = 0;
    source.arrayBuffer = async () => {
      sourceReads += 1;
      return buffer.slice(0);
    };
    const progress: OfficeFileImportProgress[] = [];

    const artifact = await importOfficeFile(source, {
      onProgress: (event) => progress.push(event),
    });

    expect(sourceReads).toBe(1);
    expect(artifact.content.type).toBe('spreadsheet');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }
    const sheet = artifact.content.sheets[0];
    expect(sheet).toMatchObject({
      name: 'Sparse',
      row: 1_048_576,
      column: 16_384,
    });
    expect(sheet?.data).toHaveLength(1_048_576);
    expect(Object.keys(sheet?.data ?? [])).toEqual(['0']);
    expect(Object.keys(sheet?.data?.[0] ?? [])).toEqual(['0']);
    expect(sheet?.data?.[0]?.[0]).toMatchObject({
      v: 'Anchor',
    });

    expect(progress[0]).toMatchObject({
      stage: 'reading',
      progress: 0,
      stageProgress: 0,
    });
    expect(progress.some(({ stage }) => stage === 'parsing')).toBe(true);
    expect(progress.some(({ stage }) => stage === 'analyzing')).toBe(true);
    expect(progress.at(-1)).toMatchObject({
      stage: 'finalizing',
      progress: 1,
      stageProgress: 1,
    });
    expect(
      progress.every(
        (event, index) =>
          index === 0 || event.progress >= (progress[index - 1]?.progress ?? 0),
      ),
    ).toBe(true);
  });

  test('stops an import when its AbortSignal is cancelled between stages', async () => {
    const controller = new AbortController();
    const pending = importOfficeFile(
      new File(['# Large report'], 'report.md', { type: 'text/markdown' }),
      {
        signal: controller.signal,
        onProgress: ({ stage }) => {
          if (stage === 'parsing') controller.abort();
        },
      },
    );

    await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
  });

  test('bypasses full HTML conversion for a large plain DOCX', async () => {
    const paragraphCount = 20_000;
    const buffer = await largePlainDocumentBuffer(paragraphCount);
    performance.clearMeasures();

    const artifact = await importOfficeFile(
      new File([buffer], 'large-plain.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }

    expect(artifact.content.model).toBeDefined();
    expect(
      documentModelUsesWindowing(
        artifact.content.model?.root ?? { type: 'doc' },
      ),
    ).toBe(true);
    expect(artifact.content.html.match(/<p>/g)).toHaveLength(paragraphCount);
    expect(artifact.compatibility?.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.large-document-windowing' }),
    );
    expect(
      performance.getEntriesByName('a3s-office.document.import-total').at(-1),
    ).toMatchObject({ detail: { path: 'large-simple' } });
  });

  test('exports and reimports a maximum sparse worksheet', async () => {
    const artifact = createArtifact('blank-spreadsheet');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }
    const data = artifact.content.sheets[0]?.data ?? [];
    data.length = 1_048_576;
    data[0] = [];
    data[0].length = 16_384;
    data[0][0] = { v: 'Anchor', m: 'Anchor' };
    data[1_048_575] = [];
    data[1_048_575].length = 16_384;
    data[1_048_575][16_383] = { v: 'Tail', m: 'Tail' };
    artifact.content.sheets[0] = {
      ...artifact.content.sheets[0],
      name: 'Sparse',
      row: 1_048_576,
      column: 16_384,
      data,
    };

    const exported = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const worksheetXml = await archive
      .file('xl/worksheets/sheet1.xml')
      ?.async('text');
    expect(worksheetXml).toContain('ref="A1:XFD1048576"');
    expect(worksheetXml).toContain('r="1"');
    expect(worksheetXml).toContain('r="1048576"');
    expect(worksheetXml).toContain('r="XFD1048576"');

    const imported = await importOfficeFile(
      new File([exported], 'sparse-roundtrip.xlsx', { type: exported.type }),
    );
    expect(imported.content.type).toBe('spreadsheet');
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }
    const sheet = imported.content.sheets[0];
    expect(sheet).toMatchObject({ row: 1_048_576, column: 16_384 });
    expect(Object.keys(sheet?.data ?? [])).toEqual(['0', '1048575']);
    expect(Object.keys(sheet?.data?.[0] ?? [])).toEqual(['0']);
    expect(Object.keys(sheet?.data?.[1_048_575] ?? [])).toEqual(['16383']);
    expect(sheet?.data?.[1_048_575]?.[16_383]).toMatchObject({
      v: 'Tail',
    });
  });

  test('round-trips maximum compact validation and protection ranges', async () => {
    const artifact = createArtifact('blank-spreadsheet');
    if (artifact.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }
    const validationItem = {
      type: 'dropdown',
      type2: '',
      rangeTxt: 'A1:XFD1048576',
      value1: 'Ready,Blocked',
      value2: '',
      validity: '',
      remote: false,
      prohibitInput: true,
      hintShow: true,
      hintValue: 'Choose a workflow state.',
    };
    const data = artifact.content.sheets[0]?.data ?? [];
    data.length = 1_048_576;
    data[0] = [];
    data[0].length = 16_384;
    data[0][0] = { v: 'Ready', m: 'Ready' };
    let sheet = {
      ...artifact.content.sheets[0],
      id: 'sheet-1',
      name: 'Sparse rules',
      row: 1_048_576,
      column: 16_384,
      data,
      dataValidationRanges: [
        {
          ranges: [{ row: [0, 1_048_575], column: [0, 16_383] }],
          item: validationItem,
        },
      ],
    };
    sheet = withSheetProtection(
      withEditableRange(sheet, null, {
        name: 'Inputs',
        sqref: 'B2:B1048576',
      }),
      true,
    );
    artifact.content.sheets = [sheet];

    const exported = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const worksheetXml =
      (await archive.file('xl/worksheets/sheet1.xml')?.async('text')) ?? '';
    expect(worksheetXml).toContain('sqref="A1:XFD1048576"');
    expect(worksheetXml).toContain('sqref="B2:B1048576"');
    expect(worksheetXml).toContain('<sheetProtection');

    const imported = await importOfficeFile(
      new File([exported], 'compact-rules.xlsx', { type: exported.type }),
    );
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected a spreadsheet artifact.');
    }
    const importedSheet = imported.content.sheets[0];
    expect(importedSheet?.dataValidationRanges).toEqual([
      {
        ranges: [{ row: [0, 1_048_575], column: [0, 16_383] }],
        item: expect.objectContaining({
          type: 'dropdown',
          value1: 'Ready,Blocked',
          prohibitInput: true,
          hintShow: true,
        }),
      },
    ]);
    expect(Object.keys(importedSheet?.dataVerification ?? {})).toEqual([]);
    const authority = sheetProtectionAuthority(importedSheet);
    expect(authority.sheet).toBe(1);
    expect(authority.cellProtectionRanges).toContainEqual({
      range: { row: [1, 1_048_575], column: [1, 1] },
      locked: false,
      hidden: false,
    });
    const context = {
      currentSheetId: importedSheet?.id,
      luckysheetfile: [importedSheet],
    } as unknown as Context;
    expect(checkCellIsLocked(context, 800_000, 1, importedSheet?.id)).toBe(
      false,
    );
    expect(checkCellIsLocked(context, 800_000, 2, importedSheet?.id)).toBe(
      true,
    );
  });
});

async function sparseWorkbookBuffer(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
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
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/workbook.xml',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<sheets><sheet name="Sparse" sheetId="1" r:id="rId1"/></sheets>',
      '</workbook>',
    ].join(''),
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A1:XFD1048576"/>',
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Anchor</t></is></c></row></sheetData>',
      '</worksheet>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

async function largePlainDocumentBuffer(
  paragraphCount: number,
): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
      '</Types>',
    ].join(''),
  );
  zip.file(
    '_rels/.rels',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'word/document.xml',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
      Array.from(
        { length: paragraphCount },
        (_, index) => `<w:p><w:r><w:t>Line ${index + 1}</w:t></w:r></w:p>`,
      ).join(''),
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
      '</w:body></w:document>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
