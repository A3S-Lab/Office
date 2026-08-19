import JSZip from 'jszip';

const ROW_COUNT = 100_000;
const TABLE_COLUMN_COUNT = 3;
const SPREADSHEET_COLUMN_COUNT = 10;
const fixtureDirectory = new URL('./fixtures/', import.meta.url);
const documentsOnly = process.argv.includes('--documents-only');
const spreadsheetsOnly = process.argv.includes('--spreadsheets-only');

const manifest: Record<string, unknown> = {
  generatedAt: new Date().toISOString(),
  rowCount: ROW_COUNT,
  fixtures: [],
};

if (!spreadsheetsOnly) {
  await generateFixture('document-text-100k.docx', () => textDocumentXml());
  await generateFixture('document-table-100k.docx', () => tableDocumentXml());
}
if (!documentsOnly) {
  await generateSpreadsheetFixture('spreadsheet-table-100k-x10.xlsx');
}

await Bun.write(
  new URL('manifest.json', fixtureDirectory),
  `${JSON.stringify(manifest, null, 2)}\n`,
);

console.log(JSON.stringify(manifest, null, 2));

async function generateFixture(
  name: string,
  createDocumentXml: () => string,
): Promise<void> {
  const startedAt = performance.now();
  const archive = new JSZip();
  archive.file('[Content_Types].xml', docxContentTypes());
  archive.file('_rels/.rels', docxRootRelationships());
  const documentXml = createDocumentXml();
  archive.file('word/document.xml', documentXml);
  const bytes = await archive.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await Bun.write(new URL(name, fixtureDirectory), bytes);
  (manifest.fixtures as unknown[]).push({
    name,
    compressedBytes: bytes.byteLength,
    documentXmlBytes: Buffer.byteLength(documentXml),
    generationMs: round(performance.now() - startedAt),
    rows: ROW_COUNT,
    columns: name.includes('table') ? TABLE_COLUMN_COUNT : 1,
  });
}

async function generateSpreadsheetFixture(name: string): Promise<void> {
  const startedAt = performance.now();
  const archive = new JSZip();
  archive.file('[Content_Types].xml', xlsxContentTypes());
  archive.file('_rels/.rels', xlsxRootRelationships());
  archive.file('xl/workbook.xml', workbookXml());
  archive.file('xl/_rels/workbook.xml.rels', workbookRelationships());
  const worksheetXml = spreadsheetXml();
  archive.file('xl/worksheets/sheet1.xml', worksheetXml);
  const bytes = await archive.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await Bun.write(new URL(name, fixtureDirectory), bytes);
  (manifest.fixtures as unknown[]).push({
    name,
    compressedBytes: bytes.byteLength,
    worksheetXmlBytes: Buffer.byteLength(worksheetXml),
    generationMs: round(performance.now() - startedAt),
    rows: ROW_COUNT,
    columns: SPREADSHEET_COLUMN_COUNT,
    populatedCells: ROW_COUNT * SPREADSHEET_COLUMN_COUNT,
  });
}

function textDocumentXml(): string {
  const paragraphs = new Array<string>(ROW_COUNT);
  for (let index = 0; index < ROW_COUNT; index += 1) {
    const row = String(index + 1).padStart(6, '0');
    paragraphs[index] =
      `<w:p><w:r><w:t>Line ${row} - A3S Office 100k text benchmark.</w:t></w:r></w:p>`;
  }
  return wordDocument(paragraphs.join(''));
}

function tableDocumentXml(): string {
  const rows = new Array<string>(ROW_COUNT);
  for (let index = 0; index < ROW_COUNT; index += 1) {
    const row = index + 1;
    rows[index] = [
      '<w:tr>',
      tableCell(String(row)),
      tableCell(`Record ${String(row).padStart(6, '0')}`),
      tableCell(String((row * 17) % 10_000)),
      '</w:tr>',
    ].join('');
  }
  return wordDocument(
    [
      '<w:tbl>',
      '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>',
      '<w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="5200"/><w:gridCol w:w="2200"/></w:tblGrid>',
      rows.join(''),
      '</w:tbl>',
    ].join(''),
  );
}

function tableCell(value: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${value}</w:t></w:r></w:p></w:tc>`;
}

function spreadsheetXml(): string {
  const rows = new Array<string>(ROW_COUNT);
  const numericColumns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
  for (let index = 0; index < ROW_COUNT; index += 1) {
    const row = index + 1;
    const cells = [
      `<c r="A${row}" t="inlineStr"><is><t>Record ${String(row).padStart(6, '0')}</t></is></c>`,
    ];
    for (let column = 0; column < numericColumns.length; column += 1) {
      cells.push(
        `<c r="${numericColumns[column]}${row}"><v>${row * (column + 1)}</v></c>`,
      );
    }
    rows[index] = `<row r="${row}">${cells.join('')}</row>`;
  }
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<dimension ref="A1:J${ROW_COUNT}"/>`,
    '<sheetData>',
    rows.join(''),
    '</sheetData>',
    '</worksheet>',
  ].join('');
}

function wordDocument(body: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:body>',
    body,
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
    '</w:body>',
    '</w:document>',
  ].join('');
}

function docxContentTypes(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '</Types>',
  ].join('');
}

function docxRootRelationships(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
    '</Relationships>',
  ].join('');
}

function xlsxContentTypes(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '</Types>',
  ].join('');
}

function xlsxRootRelationships(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '</Relationships>',
  ].join('');
}

function workbookXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<sheets><sheet name="100k rows" sheetId="1" r:id="rId1"/></sheets>',
    '</workbook>',
  ].join('');
}

function workbookRelationships(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
    '</Relationships>',
  ].join('');
}

function round(value: number): number {
  return Math.round(value * 10) / 10;
}
