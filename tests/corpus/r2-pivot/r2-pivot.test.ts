import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { OoxmlPackage } from '../../../src/internal/features/work/work-ooxml-package';
import { WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS } from '../../../src/internal/features/work/work-spreadsheet-pivot-styles';
import { inspectXlsxPivotTables } from '../../../src/internal/features/work/work-xlsx-pivots';

const SPREADSHEET_NS =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELATIONSHIP_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIP_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NS =
  'http://schemas.openxmlformats.org/package/2006/content-types';

describe('R2 pivot depth diagnostics and styles', () => {
  test('exposes a bounded Traditional Office pivot-style catalog', () => {
    expect(WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS.length).toBeGreaterThanOrEqual(
      20,
    );
    expect(
      WORK_SPREADSHEET_PIVOT_STYLE_OPTIONS.map((option) => option.value),
    ).toEqual(expect.arrayContaining(['PivotStyleLight16', 'PivotStyleDark28']));
  });

  test('fail-closes calculated pivot-cache fields with a dedicated code', async () => {
    const buffer = await buildPivotWorkbook({
      cacheFields: `<cacheFields count="2">
  <cacheField name="Category" numFmtId="0"><sharedItems count="1"><s v="A"/></sharedItems></cacheField>
  <cacheField name="Margin" numFmtId="0" formula="Amount*0.1"><sharedItems/></cacheField>
</cacheFields>`,
    });
    const result = await inspectXlsxPivotTables(
      await OoxmlPackage.load(buffer),
    );
    expect(result.tables).toEqual([]);
    expect(result.unsupported).toEqual([
      expect.objectContaining({
        code: 'xlsx.pivots.calculated-fields',
      }),
    ]);
  });

  test('fail-closes grouped pivot-cache fields separately from calculated fields', async () => {
    const buffer = await buildPivotWorkbook({
      cacheFields: `<cacheFields count="1">
  <cacheField name="Category" numFmtId="0">
    <sharedItems count="1"><s v="A"/></sharedItems>
    <fieldGroup><rangePr groupBy="days"/></fieldGroup>
  </cacheField>
</cacheFields>`,
    });
    const result = await inspectXlsxPivotTables(
      await OoxmlPackage.load(buffer),
    );
    expect(result.tables).toEqual([]);
    expect(result.unsupported).toEqual([
      expect.objectContaining({
        code: 'xlsx.pivots.grouping',
      }),
    ]);
  });

  test('fail-closes pivot slicers and timelines without editable import', async () => {
    const buffer = await buildPivotWorkbook({
      cacheFields: `<cacheFields count="1">
  <cacheField name="Category" numFmtId="0"><sharedItems count="1"><s v="A"/></sharedItems></cacheField>
</cacheFields>`,
      withSlicerParts: true,
    });
    const result = await inspectXlsxPivotTables(
      await OoxmlPackage.load(buffer),
    );
    expect(result.unsupported).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'xlsx.pivots.slicers',
        }),
      ]),
    );
  });
});

async function buildPivotWorkbook(options: {
  cacheFields: string;
  withSlicerParts?: boolean;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELATIONSHIP_NS}">
  <Relationship Id="rId1" Type="${RELATIONSHIP_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${RELATIONSHIP_NS}">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="1" r:id="rId2"/>
  </pivotCaches>
</workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELATIONSHIP_NS}">
  <Relationship Id="rId1" Type="${RELATIONSHIP_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIP_NS}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${RELATIONSHIP_NS}">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Category</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>10</v></c></row>
  </sheetData>
</worksheet>`,
  );
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELATIONSHIP_NS}">
  <Relationship Id="rId1" Type="${RELATIONSHIP_NS}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/pivotTables/pivotTable1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="${SPREADSHEET_NS}" name="Pivot1" cacheId="1" dataCaption="Values">
  <location ref="D1:E3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="-2"/></colFields>
  <dataFields count="1"><dataField name="Sum of Amount" fld="1" baseField="0" baseItem="0"/></dataFields>
</pivotTableDefinition>`,
  );
  zip.file(
    'xl/pivotTables/_rels/pivotTable1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELATIONSHIP_NS}">
  <Relationship Id="rId1" Type="${RELATIONSHIP_NS}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/pivotCache/pivotCacheDefinition1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="${SPREADSHEET_NS}" xmlns:r="${RELATIONSHIP_NS}" r:id="rId1" refreshOnLoad="0" recordCount="1">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B2" sheet="Sheet1"/>
  </cacheSource>
  ${options.cacheFields}
</pivotCacheDefinition>`,
  );
  zip.file(
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_RELATIONSHIP_NS}">
  <Relationship Id="rId1" Type="${RELATIONSHIP_NS}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
  );
  zip.file(
    'xl/pivotCache/pivotCacheRecords1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="${SPREADSHEET_NS}" count="1">
  <r><s v="A"/><n v="10"/></r>
</pivotCacheRecords>`,
  );
  if (options.withSlicerParts) {
    zip.file(
      'xl/slicers/slicer1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="${SPREADSHEET_NS}"><slicer name="Category"/></slicers>`,
    );
    zip.file(
      'xl/slicerCaches/slicerCache1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="${SPREADSHEET_NS}" name="Slicer_Category" sourceName="Category"/>`,
    );
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}
