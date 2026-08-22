import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { importOfficeFile } from '../src/core';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import {
  descendants,
  OoxmlPackage,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  createXlsxRichTextReadContext,
  MAX_XLSX_RICH_TEXT_CELL_CHARACTERS,
  MAX_XLSX_RICH_TEXT_CELLS,
  MAX_XLSX_RICH_TEXT_RUNS,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  readXlsxRichTextCells,
} from '../src/internal/features/work/work-xlsx-rich-text';

describe('XLSX rich-text cells', () => {
  test('reads shared and inline rich strings with native run formatting', () => {
    const sharedStrings = parseXml(
      [
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<si><r><rPr><rFont val="Aptos Display"/><b/><color theme="4"/>',
        '<sz val="12"/><u val="double"/></rPr><t>Alpha</t></r>',
        '<r><rPr><i/></rPr><t xml:space="preserve"> shared </t></r></si>',
        '<si><t>Plain shared string</t></si>',
        '</sst>',
      ].join(''),
      'shared strings',
    );
    const worksheet = parseXml(
      [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c>',
        '<c r="B1" t="s"><v>1</v></c>',
        '<c r="C1" t="inlineStr"><is>',
        '<r><rPr><i/><strike/><color rgb="FF00AA66"/></rPr>',
        '<t xml:space="preserve"> inline </t></r>',
        '<r><rPr><u val="singleAccounting"/></rPr><t>text</t></r>',
        '</is></c></row></sheetData></worksheet>',
      ].join(''),
      'worksheet',
    );
    const context = createXlsxRichTextReadContext({
      sharedStrings,
      styles: null,
      theme: null,
    });

    expect(readXlsxRichTextCells(worksheet, context)).toEqual([
      {
        column: 0,
        row: 0,
        runs: [
          {
            a3sXlsxColorOrigin: {
              baseColor: '#4472c4',
              index: 4,
              kind: 'theme',
              renderedColor: '#4472c4',
            },
            bl: 1,
            fc: '#4472c4',
            ff: 'Aptos Display',
            fs: 12,
            un: 2,
            v: 'Alpha',
          },
          { it: 1, v: ' shared ' },
        ],
        text: 'Alpha shared ',
      },
      {
        column: 2,
        row: 0,
        runs: [
          { cl: 1, fc: '#00aa66', it: 1, v: ' inline ' },
          { un: 3, v: 'text' },
        ],
        text: ' inline text',
      },
    ]);
  });

  test('fails closed for invalid coordinates, references, and bounded payloads', () => {
    const oversizedText = 'x'.repeat(MAX_XLSX_RICH_TEXT_CELL_CHARACTERS + 1);
    const tooManyRuns = Array.from(
      { length: MAX_XLSX_RICH_TEXT_RUNS_PER_CELL + 1 },
      () => '<r><t>x</t></r>',
    ).join('');
    const worksheet = parseXml(
      [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1">',
        '<c r="A0" t="inlineStr"><is><r><t>invalid</t></r></is></c>',
        '<c r="A1" t="s"><v>999</v></c>',
        `<c r="B1" t="inlineStr"><is><r><t>${oversizedText}</t></r></is></c>`,
        `<c r="C1" t="inlineStr"><is>${tooManyRuns}</is></c>`,
        '<c r="D1" t="inlineStr"><is><r><rPr><sz val="999"/></rPr><t>safe</t></r></is></c>',
        '</row></sheetData></worksheet>',
      ].join(''),
      'bounded worksheet',
    );
    const context = createXlsxRichTextReadContext({
      sharedStrings: null,
      styles: null,
      theme: null,
    });

    expect(readXlsxRichTextCells(worksheet, context)).toEqual([
      {
        column: 3,
        row: 0,
        runs: [{ v: 'safe' }],
        text: 'safe',
      },
    ]);
  });

  test('shares bounded cell and run materialization across worksheet reads', () => {
    const worksheet = parseXml(
      [
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetData><row r="1">',
        '<c r="A1" t="inlineStr"><is><r><t>first</t></r><r><t> cell</t></r></is></c>',
        '<c r="B1" t="inlineStr"><is><r><t>second</t></r></is></c>',
        '</row></sheetData></worksheet>',
      ].join(''),
      'budgeted worksheet',
    );
    const context = createXlsxRichTextReadContext({
      sharedStrings: null,
      styles: null,
      theme: null,
    });
    expect(context.remainingCells).toBe(MAX_XLSX_RICH_TEXT_CELLS);
    expect(context.remainingRuns).toBe(MAX_XLSX_RICH_TEXT_RUNS);
    context.remainingCells = 1;
    context.remainingRuns = 2;

    expect(readXlsxRichTextCells(worksheet, context)).toEqual([
      {
        column: 0,
        row: 0,
        runs: [{ v: 'first' }, { v: ' cell' }],
        text: 'first cell',
      },
    ]);
    expect(context.remainingCells).toBe(0);
    expect(context.remainingRuns).toBe(0);
    expect(readXlsxRichTextCells(worksheet, context)).toEqual([]);
  });

  test('imports shared and inline runs, exports native inline strings, and reopens them', async () => {
    const source = await nativeRichTextWorkbook();
    const imported = await importOfficeFile(
      new File([source], 'native-rich-text.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
    );
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected an imported spreadsheet.');
    }

    expect(imported.content.sheets[0]?.data?.[0]?.[0]).toMatchObject({
      ct: {
        t: 'inlineStr',
        s: [
          {
            bl: 1,
            fc: '#4f81bd',
            ff: 'Aptos Display',
            fs: 12,
            un: 2,
            v: 'Alpha',
          },
          { it: 1, v: ' shared ' },
        ],
      },
      v: 'Alpha shared ',
    });
    expect(imported.content.sheets[0]?.data?.[0]?.[1]).toMatchObject({
      ct: {
        t: 'inlineStr',
        s: [
          { cl: 1, fc: '#00aa66', it: 1, v: ' inline ' },
          { un: 3, v: 'text' },
        ],
      },
      v: ' inline text',
    });

    const exported = await createWorkArtifactBlob(imported);
    const archive = await OoxmlPackage.load(await exported.arrayBuffer());
    const worksheet = await archive.xml('xl/worksheets/sheet1.xml');
    const cells = new Map(
      descendants(worksheet, 'c').map((cell) => [cell.getAttribute('r'), cell]),
    );
    const first = cells.get('A1');
    const second = cells.get('B1');
    if (!first || !second)
      throw new Error('Expected exported rich-text cells.');
    const firstXml = new XMLSerializer().serializeToString(first);
    const secondXml = new XMLSerializer().serializeToString(second);
    expect(firstXml).toContain('t="inlineStr"');
    expect(firstXml).toContain('<rFont val="Aptos Display"');
    expect(firstXml).toContain('<b');
    expect(firstXml).toContain('<color theme="4"');
    expect(firstXml).toContain('<u val="double"');
    expect(firstXml).toContain('xml:space="preserve"> shared </t>');
    expect(secondXml).toContain('<strike');
    expect(secondXml).toContain('rgb="FF00AA66"');
    expect(secondXml).toContain('val="singleAccounting"');

    const reopened = await importWorkFile(
      new File([exported], 'native-rich-text-reopened.xlsx', {
        type: exported.type,
      }),
    );
    if (reopened.content.type !== 'spreadsheet') {
      throw new Error('Expected a reopened spreadsheet.');
    }
    expect(reopened.content.sheets[0]?.data?.[0]?.slice(0, 2)).toMatchObject([
      {
        ct: {
          t: 'inlineStr',
          s: [
            { bl: 1, v: 'Alpha' },
            { it: 1, v: ' shared ' },
          ],
        },
        v: 'Alpha shared ',
      },
      {
        ct: {
          t: 'inlineStr',
          s: [
            { cl: 1, fc: '#00aa66', it: 1, v: ' inline ' },
            { un: 3, v: 'text' },
          ],
        },
        v: ' inline text',
      },
    ]);
  });
});

async function nativeRichTextWorkbook(): Promise<ArrayBuffer> {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([['shared placeholder', 'inline placeholder']]),
    'Rich text',
  );
  const seed = XLSX.write(workbook, {
    bookType: 'xlsx',
    compression: true,
    type: 'array',
  }) as ArrayBuffer;
  const zip = await JSZip.loadAsync(seed);
  const sheet = zip.file('xl/worksheets/sheet1.xml');
  if (!sheet) throw new Error('Expected a worksheet part.');
  const worksheet = (await sheet.async('text'))
    .replace(/<c r="A1"[\s\S]*?<\/c>/, '<c r="A1" t="s"><v>0</v></c>')
    .replace(
      /<c r="B1"[\s\S]*?<\/c>/,
      [
        '<c r="B1" t="inlineStr"><is>',
        '<r><rPr><i/><strike/><color rgb="FF00AA66"/></rPr>',
        '<t xml:space="preserve"> inline </t></r>',
        '<r><rPr><u val="singleAccounting"/></rPr><t>text</t></r>',
        '</is></c>',
      ].join(''),
    );
  zip.file('xl/worksheets/sheet1.xml', worksheet);
  zip.file(
    'xl/sharedStrings.xml',
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">',
      '<si><r><rPr><rFont val="Aptos Display"/><b/><color theme="4"/>',
      '<sz val="12"/><u val="double"/></rPr><t>Alpha</t></r>',
      '<r><rPr><i/></rPr><t xml:space="preserve"> shared </t></r></si>',
      '</sst>',
    ].join(''),
  );

  const relationships = zip.file('xl/_rels/workbook.xml.rels');
  const contentTypes = zip.file('[Content_Types].xml');
  if (!relationships || !contentTypes) {
    throw new Error('Expected workbook package metadata.');
  }
  zip.file(
    'xl/_rels/workbook.xml.rels',
    (await relationships.async('text')).replace(
      '</Relationships>',
      '<Relationship Id="rIdRichText" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
    ),
  );
  zip.file(
    '[Content_Types].xml',
    (await contentTypes.async('text')).replace(
      '</Types>',
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
