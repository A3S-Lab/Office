import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';

describe('Spreadsheet XLSX date-system retention', () => {
  test('retains a 1904 workbook through dynamic filtering, export, and reopen', async () => {
    const source = await workbook1904WithJanuaryFilter();
    const imported = await importWorkFile(
      new File([source], 'dates-1904.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
    );
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected a Spreadsheet artifact.');
    }

    expect(imported.content.dateSystem).toBe('1904');
    expect(imported.content.sheets[0]?.data?.[1]?.[0]).toMatchObject({
      ct: { t: 'd' },
      v: 0,
    });
    expect(imported.content.sheets[0]?.data?.[3]?.[0]).toMatchObject({
      ct: { t: 'd' },
      f: '=A2',
      v: 0,
    });
    expect(imported.content.sheets[0]?.filter?.['0']).toMatchObject({
      caljs: {
        criteria: { type: 'dynamic', kind: 'month-1' },
      },
      rowhidden: { '2': 0 },
    });
    expect(imported.content.sheets[0]?.config?.rowhidden).toEqual({ '2': 0 });

    const exported = await createWorkArtifactBlob(imported);
    const exportedBytes = await exported.arrayBuffer();
    const archive = await OoxmlPackage.load(exportedBytes);
    const workbookXml = await archive.text('xl/workbook.xml');
    expect(workbookXml).toMatch(/<workbookPr[^>]*date1904="(?:1|true)"/);

    const raw = XLSX.read(exportedBytes, {
      type: 'array',
      cellDates: false,
      cellFormula: true,
      cellNF: true,
    });
    expect(raw.Workbook?.WBProps?.date1904).toBe(true);
    expect(raw.Sheets.Dates?.A2?.v).toBe(0);
    expect(raw.Sheets.Dates?.A3?.v).toBe(31);
    expect(raw.Sheets.Dates?.A4).toMatchObject({ f: 'A2', v: 0 });

    const reopened = await importWorkFile(
      new File([exported], 'dates-1904-reopened.xlsx', {
        type: exported.type,
      }),
    );
    if (reopened.content.type !== 'spreadsheet') {
      throw new Error('Expected a reopened Spreadsheet artifact.');
    }
    expect(reopened.content.dateSystem).toBe('1904');
    expect(reopened.content.sheets[0]?.filter?.['0']).toMatchObject({
      caljs: {
        criteria: { type: 'dynamic', kind: 'month-1' },
      },
      rowhidden: { '2': 0 },
    });
    expect(reopened.content.sheets[0]?.config?.rowhidden).toEqual({ '2': 0 });
  });
});

async function workbook1904WithJanuaryFilter(): Promise<ArrayBuffer> {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([['Date'], [0], [31]]);
  worksheet.A2.z = 'yyyy-mm-dd';
  worksheet.A3.z = 'yyyy-mm-dd';
  worksheet.A4 = { f: 'A2', t: 'n', v: 0, z: 'yyyy-mm-dd' };
  worksheet['!ref'] = 'A1:A4';
  worksheet['!autofilter'] = { ref: 'A1:A3' };
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Dates');
  workbook.Workbook = { WBProps: { date1904: true } };
  const bytes = XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
    compression: true,
  }) as ArrayBuffer;
  const zip = await JSZip.loadAsync(bytes);
  const worksheetEntry = zip.file('xl/worksheets/sheet1.xml');
  if (!worksheetEntry) throw new Error('Expected a worksheet package part.');
  const worksheetXml = await worksheetEntry.async('text');
  const withFilter = worksheetXml.replace(
    /<autoFilter ref="A1:A3"\s*\/>/,
    '<autoFilter ref="A1:A3"><filterColumn colId="0"><dynamicFilter type="M1"/></filterColumn></autoFilter>',
  );
  if (withFilter === worksheetXml) {
    throw new Error('Expected a self-closing worksheet AutoFilter.');
  }
  zip.file('xl/worksheets/sheet1.xml', withFilter);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
