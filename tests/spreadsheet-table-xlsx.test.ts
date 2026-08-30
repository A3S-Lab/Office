import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { OoxmlPackage } from '../src/internal/features/work/work-ooxml-package';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import {
  patchXlsxSpreadsheetTables,
  readXlsxWorksheetTables,
} from '../src/internal/features/work/work-xlsx-tables';

describe('spreadsheet table XLSX interop', () => {
  test('imports a typed ListObject with filters and display options', async () => {
    const buffer = await tableWorkbook();
    const archive = await OoxmlPackage.load(buffer);
    const tables = await readXlsxWorksheetTables(
      archive,
      'xl/worksheets/sheet1.xml',
    );

    expect(tables).toEqual([
      expect.objectContaining({
        name: 'Sales',
        displayName: 'SalesView',
        ooxmlId: 7,
        range: { row: [0, 3], column: [0, 2] },
        columns: [{ name: 'Region' }, { name: 'Qty' }, { name: 'State' }],
        filters: [
          {
            column: 2,
            criteria: {
              type: 'values',
              values: ['Ready', 'Blocked'],
              includeBlanks: true,
            },
          },
        ],
        headerRow: true,
        totalsRow: false,
        style: { family: 'medium', number: 4 },
        showFirstColumn: true,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      }),
    ]);
  });

  test('writes table parts, worksheet relationships, tableParts, and content types', async () => {
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sales',
          tables: [
            {
              id: 'table-local',
              ooxmlId: 9,
              name: 'Inventory',
              range: { row: [1, 4], column: [1, 3] },
              columns: [{ name: 'Item' }, { name: 'Units' }, { name: 'Cost' }],
              filters: [
                {
                  column: 1,
                  criteria: { type: 'greater-than', value: '5' },
                },
              ],
              headerRow: true,
              totalsRow: true,
              style: { family: 'dark', number: 2 },
              showFirstColumn: false,
              showLastColumn: true,
              showRowStripes: false,
              showColumnStripes: true,
            },
          ],
        },
      ],
    };
    const patched = await patchXlsxSpreadsheetTables(
      await blankWorkbook(),
      content,
    );
    const zip = await JSZip.loadAsync(patched);
    const worksheet =
      (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? '';
    const relationships =
      (await zip.file('xl/worksheets/_rels/sheet1.xml.rels')?.async('text')) ??
      '';
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    const contentTypes =
      (await zip.file('[Content_Types].xml')?.async('text')) ?? '';

    expect(worksheet).toContain('<tableParts count="1">');
    expect(worksheet).toContain('r:id="rId1"');
    expect(relationships).toContain('/relationships/table');
    expect(relationships).toContain('Target="../tables/table1.xml"');
    expect(table).toContain('id="9"');
    expect(table).toContain('name="Inventory"');
    expect(table).toContain('ref="B2:D5"');
    expect(table).toContain('totalsRowCount="1"');
    expect(table).toContain('<customFilter operator="greaterThan" val="5"/>');
    expect(table).toContain('name="TableStyleDark2"');
    expect(table).toContain('showLastColumn="1"');
    expect(contentTypes).toContain('/xl/tables/table1.xml');
    expect(contentTypes).toContain(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    );

    const roundTrip = await readXlsxWorksheetTables(
      await OoxmlPackage.load(patched),
      'xl/worksheets/sheet1.xml',
    );
    expect(roundTrip[0]).toMatchObject({
      name: 'Inventory',
      range: { row: [1, 4], column: [1, 3] },
      totalsRow: true,
      style: { family: 'dark', number: 2 },
    });
  });

  test('round-trips a bounded calculated-column formula without its leading equals sign', async () => {
    const buffer = await tableWorkbookWithCalculatedColumn();
    const imported = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );

    expect(imported[0]?.columns).toEqual([
      { name: 'Region' },
      { name: 'Qty' },
      { name: 'State', calculatedFormula: '=[@Qty]*2' },
    ]);

    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sales',
          tables: imported,
        },
      ],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).toContain(
      '<calculatedColumnFormula>[@Qty]*2</calculatedColumnFormula>',
    );
    expect(table).not.toContain(
      '<calculatedColumnFormula>=[@Qty]*2</calculatedColumnFormula>',
    );

    const roundTrip = await readXlsxWorksheetTables(
      await OoxmlPackage.load(patched),
      'xl/worksheets/sheet1.xml',
    );
    expect(roundTrip[0]?.columns[2]).toEqual({
      name: 'State',
      calculatedFormula: '=[@Qty]*2',
    });
  });

  test('drops unsafe calculated-column metadata at the XLSX boundary', async () => {
    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sales',
          tables: [
            {
              id: 'table-unsafe',
              name: 'Sales',
              range: { row: [0, 2], column: [0, 2] },
              columns: [
                { name: 'Region' },
                { name: 'Qty' },
                {
                  name: 'State',
                  calculatedFormula: '=HYPERLINK("https://example.test")',
                },
              ],
              filters: [],
              headerRow: true,
              totalsRow: false,
              style: { family: 'medium', number: 2 },
              showFirstColumn: false,
              showLastColumn: false,
              showRowStripes: true,
              showColumnStripes: false,
            },
          ],
        },
      ],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).not.toContain('calculatedColumnFormula');
  });

  test('round-trips native and custom totals-row metadata', async () => {
    const buffer = await tableWorkbookWithTotalsRow();
    const imported = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );
    expect(imported[0]?.columns).toEqual([
      { name: 'Region', totalsLabel: 'Total' },
      { name: 'Qty', totalsFunction: 'sum' },
      {
        name: 'State',
        totalsFunction: 'custom',
        totalsFormula: '=SUM(Sales[Qty])',
      },
    ]);

    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sales', tables: imported }],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).toContain('totalsRowLabel="Total"');
    expect(table).toContain('totalsRowFunction="sum"');
    expect(table).toContain(
      '<totalsRowFormula>SUM(Sales[Qty])</totalsRowFormula>',
    );
    expect(table).not.toContain(
      '<totalsRowFormula>=SUM(Sales[Qty])</totalsRowFormula>',
    );
  });

  test('ignores malformed, duplicate, and prototype-like filter columns', async () => {
    const buffer = await tableWorkbook(
      [
        '<filterColumn><dynamicFilter type="today"/></filterColumn>',
        '<filterColumn colId="0"><dynamicFilter type="constructor"/></filterColumn>',
        '<filterColumn colId="1"><top10 percent="1" val="101"/></filterColumn>',
        '<filterColumn colId="2"><filters><filter val="Ready"/></filters></filterColumn>',
        '<filterColumn colId="2"><filters><filter val="Blocked"/></filters></filterColumn>',
      ].join(''),
    );
    const tables = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );

    expect(tables[0]?.filters).toEqual([
      {
        column: 2,
        criteria: {
          type: 'values',
          values: ['Ready'],
          includeBlanks: false,
        },
      },
    ]);
  });

  test('round-trips Top/Bottom item and percentage filters', async () => {
    const buffer = await tableWorkbook(
      [
        '<filterColumn colId="0"><top10 top="1" percent="0" val="2"/></filterColumn>',
        '<filterColumn colId="1"><top10 top="0" percent="1" val="50"/></filterColumn>',
        '<filterColumn colId="2"><top10 top="1" percent="1" val="25"/></filterColumn>',
      ].join(''),
    );
    const tables = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );

    expect(tables[0]?.filters).toEqual([
      { column: 0, criteria: { type: 'top', count: 2 } },
      { column: 1, criteria: { type: 'bottom-percent', percent: 50 } },
      { column: 2, criteria: { type: 'top-percent', percent: 25 } },
    ]);

    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sales', tables }],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).toContain('<top10 percent="0" top="1" val="2"/>');
    expect(table).toContain('<top10 percent="1" top="0" val="50"/>');
    expect(table).toContain('<top10 percent="1" top="1" val="25"/>');

    const roundTrip = await readXlsxWorksheetTables(
      await OoxmlPackage.load(patched),
      'xl/worksheets/sheet1.xml',
    );
    expect(roundTrip[0]?.filters).toEqual(tables[0]?.filters);
  });

  test('round-trips arbitrary WPS wildcard expressions without flattening their meaning', async () => {
    const buffer = await tableWorkbook(
      [
        '<filterColumn colId="0"><customFilters>',
        '<customFilter operator="equal" val="K?ng*"/>',
        '</customFilters></filterColumn>',
        '<filterColumn colId="1"><customFilters and="1">',
        '<customFilter operator="equal" val="*risk??"/>',
        '<customFilter operator="notEqual" val="*do?e"/>',
        '</customFilters></filterColumn>',
        '<filterColumn colId="2"><customFilters>',
        '<customFilter operator="equal" val="King~*"/>',
        '</customFilters></filterColumn>',
      ].join(''),
    );
    const tables = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );

    expect(tables[0]?.filters).toEqual([
      {
        column: 0,
        criteria: { type: 'matches-wildcard', value: 'K?ng*' },
      },
      {
        column: 1,
        criteria: {
          type: 'compound',
          conjunction: 'and',
          conditions: [
            { type: 'matches-wildcard', value: '*risk??' },
            { type: 'does-not-match-wildcard', value: '*do?e' },
          ],
        },
      },
      {
        column: 2,
        criteria: { type: 'equals', value: 'King*' },
      },
    ]);

    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sales', tables }],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).toContain('<customFilter operator="equal" val="K?ng*"/>');
    expect(table).toContain(
      '<customFilters and="1"><customFilter operator="equal" val="*risk??"/><customFilter operator="notEqual" val="*do?e"/></customFilters>',
    );
    expect(table).toContain('<customFilter operator="equal" val="King~*"/>');

    const roundTrip = await readXlsxWorksheetTables(
      await OoxmlPackage.load(patched),
      'xl/worksheets/sheet1.xml',
    );
    expect(roundTrip[0]?.filters).toEqual(tables[0]?.filters);
  });

  test('round-trips literal wildcard characters without changing filter meaning', async () => {
    const buffer = await tableWorkbook(
      [
        '<filterColumn colId="0"><customFilters>',
        '<customFilter operator="equal" val="literal~*~?~~value"/>',
        '</customFilters></filterColumn>',
        '<filterColumn colId="1"><customFilters>',
        '<customFilter operator="notEqual" val="other~*~?~~value"/>',
        '</customFilters></filterColumn>',
        '<filterColumn colId="2"><customFilters>',
        '<customFilter operator="notEqual" val="unsupported*"/>',
        '</customFilters></filterColumn>',
      ].join(''),
    );
    const tables = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );

    expect(tables[0]?.filters).toEqual([
      {
        column: 0,
        criteria: { type: 'equals', value: 'literal*?~value' },
      },
      {
        column: 1,
        criteria: { type: 'not-equals', value: 'other*?~value' },
      },
      {
        column: 2,
        criteria: { type: 'does-not-begin-with', value: 'unsupported' },
      },
    ]);

    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sales', tables }],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).toContain(
      '<customFilter operator="equal" val="literal~*~?~~value"/>',
    );
    expect(table).toContain(
      '<customFilter operator="notEqual" val="other~*~?~~value"/>',
    );
    expect(table).toContain(
      '<customFilter operator="notEqual" val="unsupported*"/>',
    );

    const roundTrip = await readXlsxWorksheetTables(
      await OoxmlPackage.load(patched),
      'xl/worksheets/sheet1.xml',
    );
    expect(roundTrip[0]?.filters).toEqual(tables[0]?.filters);
  });

  test('round-trips compound and negative prefix or suffix custom filters', async () => {
    const buffer = await tableWorkbook(
      [
        '<filterColumn colId="0"><customFilters and="1">',
        '<customFilter operator="greaterThan" val="100"/>',
        '<customFilter operator="lessThan" val="200"/>',
        '</customFilters></filterColumn>',
        '<filterColumn colId="1"><customFilters and="0">',
        '<customFilter operator="equal" val="*risk*"/>',
        '<customFilter operator="notEqual" val="blocked*"/>',
        '</customFilters></filterColumn>',
        '<filterColumn colId="2"><customFilters>',
        '<customFilter operator="notEqual" val="*archived"/>',
        '</customFilters></filterColumn>',
      ].join(''),
    );
    const tables = await readXlsxWorksheetTables(
      await OoxmlPackage.load(buffer),
      'xl/worksheets/sheet1.xml',
    );

    expect(tables[0]?.filters).toEqual([
      {
        column: 0,
        criteria: {
          type: 'compound',
          conjunction: 'and',
          conditions: [
            { type: 'greater-than', value: '100' },
            { type: 'less-than', value: '200' },
          ],
        },
      },
      {
        column: 1,
        criteria: {
          type: 'compound',
          conjunction: 'or',
          conditions: [
            { type: 'contains', value: 'risk' },
            { type: 'does-not-begin-with', value: 'blocked' },
          ],
        },
      },
      {
        column: 2,
        criteria: { type: 'does-not-end-with', value: 'archived' },
      },
    ]);

    const patched = await patchXlsxSpreadsheetTables(await blankWorkbook(), {
      type: 'spreadsheet',
      sheets: [{ id: 'sheet-1', name: 'Sales', tables }],
    });
    const zip = await JSZip.loadAsync(patched);
    const table = (await zip.file('xl/tables/table1.xml')?.async('text')) ?? '';
    expect(table).toContain(
      '<customFilters and="1"><customFilter operator="greaterThan" val="100"/><customFilter operator="lessThan" val="200"/></customFilters>',
    );
    expect(table).toContain(
      '<customFilters and="0"><customFilter operator="equal" val="*risk*"/><customFilter operator="notEqual" val="blocked*"/></customFilters>',
    );
    expect(table).toContain(
      '<customFilter operator="notEqual" val="*archived"/>',
    );

    const roundTrip = await readXlsxWorksheetTables(
      await OoxmlPackage.load(patched),
      'xl/worksheets/sheet1.xml',
    );
    expect(roundTrip[0]?.filters).toEqual(tables[0]?.filters);
  });
});

async function tableWorkbook(
  filterColumns = [
    '<filterColumn colId="2">',
    '<filters blank="1"><filter val="Ready"/><filter val="Blocked"/></filters>',
    '</filterColumn>',
  ].join(''),
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await blankWorkbook());
  zip.file(
    'xl/worksheets/sheet1.xml',
    [
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<dimension ref="A1:C4"/><sheetData/>',
      '<tableParts count="1"><tablePart r:id="rId5"/></tableParts>',
      '</worksheet>',
    ].join(''),
  );
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    [
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      '<Relationship Id="rId5"',
      ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"',
      ' Target="../tables/table7.xml"/>',
      '</Relationships>',
    ].join(''),
  );
  zip.file(
    'xl/tables/table7.xml',
    [
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ' id="7" name="Sales" displayName="SalesView" ref="A1:C4"',
      ' headerRowCount="1" totalsRowCount="0" totalsRowShown="0">',
      `<autoFilter ref="A1:C4">${filterColumns}</autoFilter>`,
      '<tableColumns count="3"><tableColumn id="1" name="Region"/>',
      '<tableColumn id="2" name="Qty"/><tableColumn id="3" name="State"/>',
      '</tableColumns>',
      '<tableStyleInfo name="TableStyleMedium4" showFirstColumn="1"',
      ' showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
      '</table>',
    ].join(''),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function tableWorkbookWithCalculatedColumn(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await tableWorkbook());
  const table = (await zip.file('xl/tables/table7.xml')?.async('text')) ?? '';
  zip.file(
    'xl/tables/table7.xml',
    table.replace(
      '<tableColumn id="3" name="State"/>',
      '<tableColumn id="3" name="State"><calculatedColumnFormula>[@Qty]*2</calculatedColumnFormula></tableColumn>',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function tableWorkbookWithTotalsRow(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await tableWorkbook());
  const table = (await zip.file('xl/tables/table7.xml')?.async('text')) ?? '';
  zip.file(
    'xl/tables/table7.xml',
    table
      .replace(
        '<tableColumn id="1" name="Region"/>',
        '<tableColumn id="1" name="Region" totalsRowLabel="Total"/>',
      )
      .replace(
        '<tableColumn id="2" name="Qty"/>',
        '<tableColumn id="2" name="Qty" totalsRowFunction="sum"/>',
      )
      .replace(
        '<tableColumn id="3" name="State"/>',
        '<tableColumn id="3" name="State" totalsRowFunction="custom"><totalsRowFormula>SUM(Sales[Qty])</totalsRowFormula></tableColumn>',
      ),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function blankWorkbook(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    [
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml"',
      ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/worksheets/sheet1.xml"',
      ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
      '</Types>',
    ].join(''),
  );
  zip.file(
    'xl/workbook.xml',
    [
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
      '<sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets>',
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
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D8"/><sheetData/></worksheet>',
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
