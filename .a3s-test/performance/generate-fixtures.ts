import { jsPDF } from 'jspdf';
import JSZip from 'jszip';

const ROW_COUNT = 100_000;
const TABLE_COLUMN_COUNT = 3;
const SPREADSHEET_COLUMN_COUNT = 10;
const PRESENTATION_SLIDE_COUNT = 1_000;
const PRESENTATION_ELEMENTS_PER_SLIDE = 9;
const PDF_PAGE_COUNT = 1_000;
const FIXTURE_ARCHIVE_DATE = new Date('2026-08-19T00:00:00.000Z');
const fixtureDirectory = new URL('./fixtures/', import.meta.url);
const fixtureSelection = Object.freeze({
  documents: process.argv.includes('--documents-only'),
  pdf: process.argv.includes('--pdf-only'),
  presentations: process.argv.includes('--presentations-only'),
  spreadsheets: process.argv.includes('--spreadsheets-only'),
});
const selectedFixtureFamilies = Object.entries(fixtureSelection)
  .filter(([, selected]) => selected)
  .map(([family]) => family);

if (selectedFixtureFamilies.length > 1) {
  throw new Error('Select at most one performance fixture family.');
}

const manifest: Record<string, unknown> = {
  generatedAt: new Date().toISOString(),
  pdfPageCount: PDF_PAGE_COUNT,
  presentationElementsPerSlide: PRESENTATION_ELEMENTS_PER_SLIDE,
  presentationSlideCount: PRESENTATION_SLIDE_COUNT,
  rowCount: ROW_COUNT,
  fixtures: [],
};

if (shouldGenerate('documents')) {
  await generateFixture('document-text-100k.docx', () => textDocumentXml());
  await generateFixture('document-table-100k.docx', () => tableDocumentXml());
}
if (shouldGenerate('spreadsheets')) {
  await generateSpreadsheetFixture('spreadsheet-table-100k-x10.xlsx');
}
if (shouldGenerate('presentations')) {
  await generatePresentationFixture('presentation-slides-1000.pptx');
}
if (shouldGenerate('pdf')) {
  await generatePdfFixture('pdf-pages-1000.pdf');
}

await Bun.write(
  new URL('manifest.json', fixtureDirectory),
  `${JSON.stringify(manifest, null, 2)}\n`,
);

console.log(JSON.stringify(manifest, null, 2));

function shouldGenerate(family: keyof typeof fixtureSelection): boolean {
  return (
    selectedFixtureFamilies.length === 0 || fixtureSelection[family] === true
  );
}

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

async function generatePdfFixture(name: string): Promise<void> {
  const startedAt = performance.now();
  const pdf = new jsPDF({
    compress: true,
    format: 'a4',
    orientation: 'portrait',
    unit: 'pt',
  });
  pdf.setCreationDate(new Date('2026-08-19T00:00:00.000Z'));
  pdf.setFileId('00000000000000000000000000001000');
  pdf.setProperties({
    author: 'A3S Lab',
    creator: 'A3S Office performance fixtures',
    subject: 'Deterministic 1,000-page PDF windowing benchmark',
    title: 'A3S Office 1,000-page PDF benchmark',
  });
  for (let page = 1; page <= PDF_PAGE_COUNT; page += 1) {
    if (page > 1) pdf.addPage('a4', 'portrait');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(24);
    pdf.text(`A3S Office PDF page ${page}`, 72, 96);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(12);
    pdf.text(
      `Deterministic page marker ${String(page).padStart(4, '0')}.`,
      72,
      124,
    );
  }
  const bytes = pdf.output('arraybuffer');
  await Bun.write(new URL(name, fixtureDirectory), bytes);
  (manifest.fixtures as unknown[]).push({
    name,
    compressedBytes: bytes.byteLength,
    generationMs: round(performance.now() - startedAt),
    pages: PDF_PAGE_COUNT,
  });
}

async function generatePresentationFixture(name: string): Promise<void> {
  const startedAt = performance.now();
  const archive = new JSZip();
  addFixedArchiveFile(archive, '[Content_Types].xml', pptxContentTypes());
  addFixedArchiveFile(archive, '_rels/.rels', pptxRootRelationships());
  addFixedArchiveFile(
    archive,
    'ppt/presentation.xml',
    presentationDocumentXml(),
  );
  addFixedArchiveFile(
    archive,
    'ppt/_rels/presentation.xml.rels',
    presentationRelationshipsXml(),
  );
  let slideXmlBytes = 0;
  for (let slide = 1; slide <= PRESENTATION_SLIDE_COUNT; slide += 1) {
    const xml = presentationSlideXml(slide);
    slideXmlBytes += Buffer.byteLength(xml);
    addFixedArchiveFile(archive, `ppt/slides/slide${slide}.xml`, xml);
  }
  const bytes = await archive.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await Bun.write(new URL(name, fixtureDirectory), bytes);
  (manifest.fixtures as unknown[]).push({
    name,
    compressedBytes: bytes.byteLength,
    elements: PRESENTATION_SLIDE_COUNT * PRESENTATION_ELEMENTS_PER_SLIDE,
    elementsPerSlide: PRESENTATION_ELEMENTS_PER_SLIDE,
    generationMs: round(performance.now() - startedAt),
    slideXmlBytes,
    slides: PRESENTATION_SLIDE_COUNT,
  });
}

function addFixedArchiveFile(
  archive: JSZip,
  path: string,
  content: string,
): void {
  archive.file(path, content, {
    createFolders: false,
    date: FIXTURE_ARCHIVE_DATE,
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

function pptxContentTypes(): string {
  const slides = Array.from(
    { length: PRESENTATION_SLIDE_COUNT },
    (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    slides,
    '</Types>',
  ].join('');
}

function pptxRootRelationships(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>',
    '</Relationships>',
  ].join('');
}

function presentationDocumentXml(): string {
  const slideIds = Array.from(
    { length: PRESENTATION_SLIDE_COUNT },
    (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`,
  ).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    `<p:sldIdLst>${slideIds}</p:sldIdLst>`,
    '<p:sldSz cx="12192000" cy="6858000"/>',
    '<p:notesSz cx="6858000" cy="9144000"/>',
    '</p:presentation>',
  ].join('');
}

function presentationRelationshipsXml(): string {
  const relationships = Array.from(
    { length: PRESENTATION_SLIDE_COUNT },
    (_, index) =>
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`,
  ).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    relationships,
    '</Relationships>',
  ].join('');
}

function presentationSlideXml(slide: number): string {
  const colors = ['E8F0FE', 'E8F5E9', 'FFF3E0', 'F3E5F5', 'E0F7FA', 'FCE4EC'];
  const cards = Array.from({ length: 6 }, (_, index) => {
    const column = index % 3;
    const row = Math.floor(index / 3);
    return presentationShapeXml({
      fill: colors[index] ?? 'F2F4F7',
      fontSize: 16,
      height: 1_250_000,
      id: index + 4,
      name: `Metric ${index + 1}`,
      text: `Metric ${index + 1}: ${(slide * (index + 3)) % 10_000}`,
      width: 3_250_000,
      x: 650_000 + column * 3_750_000,
      y: 2_100_000 + row * 1_650_000,
    });
  }).join('');
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    `<p:cSld name="Benchmark slide ${slide}"><p:spTree>`,
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>',
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/><a:chOff x="0" y="0"/><a:chExt cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>',
    presentationShapeXml({
      fill: 'FFFFFF',
      fontSize: 30,
      height: 700_000,
      id: 2,
      name: 'Title',
      text: `A3S Office presentation slide ${slide}`,
      width: 10_800_000,
      x: 650_000,
      y: 450_000,
    }),
    presentationShapeXml({
      fill: 'FFFFFF',
      fontSize: 15,
      height: 450_000,
      id: 3,
      name: 'Subtitle',
      text: `Deterministic scene marker ${String(slide).padStart(4, '0')}`,
      width: 10_800_000,
      x: 650_000,
      y: 1_250_000,
    }),
    cards,
    presentationShapeXml({
      fill: '172033',
      fontSize: 12,
      height: 360_000,
      id: 10,
      name: 'Footer',
      text: `Slide ${slide} of ${PRESENTATION_SLIDE_COUNT}`,
      textColor: 'FFFFFF',
      width: 10_800_000,
      x: 650_000,
      y: 6_050_000,
    }),
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
    '</p:sld>',
  ].join('');
}

function presentationShapeXml({
  fill,
  fontSize,
  height,
  id,
  name,
  text,
  textColor = '172033',
  width,
  x,
  y,
}: {
  fill: string;
  fontSize: number;
  height: number;
  id: number;
  name: string;
  text: string;
  textColor?: string;
  width: number;
  x: number;
  y: number;
}): string {
  return [
    '<p:sp>',
    `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`,
    '<p:spPr>',
    `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>`,
    '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>',
    `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`,
    '<a:ln><a:noFill/></a:ln>',
    '</p:spPr>',
    '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p>',
    `<a:r><a:rPr lang="en-US" sz="${fontSize * 100}"/><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill><a:t>${escapeXml(text)}</a:t></a:r>`,
    '<a:endParaRPr lang="en-US"/>',
    '</a:p></p:txBody>',
    '</p:sp>',
  ].join('');
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function round(value: number): number {
  return Math.round(value * 10) / 10;
}
