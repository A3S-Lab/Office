import { Buffer } from 'node:buffer';
import { mkdir } from 'node:fs/promises';
import path from 'node:path';
import {
  AlignmentType,
  BorderStyle,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  Document,
  DocumentGridType,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlignTable,
  WidthType,
} from 'docx';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';

const fixtureDirectory = path.resolve(
  import.meta.dirname,
  '../.a3s-test/fixtures',
);
const pdfPath = path.join(fixtureDirectory, 'pdf-thumbnail-keyboard.pdf');
const picturePath = path.join(fixtureDirectory, 'word-picture.png');
const documentComparisonPath = path.join(
  fixtureDirectory,
  'word-document-comparison-revised.html',
);
const longDocumentPath = path.join(
  fixtureDirectory,
  'word-page-navigation-120.docx',
);
const longRevisionDocumentPath = path.join(
  fixtureDirectory,
  'word-revisions-120.docx',
);
const longCommentDocumentPath = path.join(
  fixtureDirectory,
  'word-comments-120.docx',
);
const multiPageTableDocumentPath = path.join(
  fixtureDirectory,
  'word-multi-page-table.docx',
);
const themeTableDocumentPath = path.join(
  fixtureDirectory,
  'word-theme-table.docx',
);
const styledTableDocumentPath = path.join(
  fixtureDirectory,
  'word-styled-table.docx',
);
const wpsLayoutDocumentPath = path.join(
  fixtureDirectory,
  'word-wps-layout.docx',
);
const wpsFontMatrixDocumentPath = path.join(
  fixtureDirectory,
  'word-wps-font-matrix.docx',
);
const wpsCjkFontMatrixDocumentPath = path.join(
  fixtureDirectory,
  'word-wps-cjk-font-matrix.docx',
);
const wpsGridMatrixDocumentPath = path.join(
  fixtureDirectory,
  'word-wps-grid-matrix.docx',
);
const wpsScriptMatrixDocumentPath = path.join(
  fixtureDirectory,
  'word-wps-script-matrix.docx',
);

await mkdir(fixtureDirectory, { recursive: true });
await Bun.write(
  documentComparisonPath,
  [
    '<h1>发布审阅基线</h1>',
    '<p><strong>基线版本：</strong>1.1</p>',
    '<h2>比较步骤</h2>',
    '<p>打开“审阅”，选择“比较文档”，再导入 DOCX、HTML 或 TXT 修订版本并生成修订。</p>',
    '<h2>审阅范围</h2>',
    '<p style="text-align: center">当前基线包含架构说明、协作协议与发布说明。</p>',
    '<p>新增的发布摘要必须在交付前确认。</p>',
    '<h2>决策要求</h2>',
    '<p>发布前必须逐项接受或拒绝所有确定性修订。</p>',
  ].join(''),
);
await Bun.write(pdfPath, createPdfThumbnailKeyboardFixture());
await Bun.write(longDocumentPath, await createLongWordNavigationFixture());
await Bun.write(
  longRevisionDocumentPath,
  await createLongWordRevisionFixture(),
);
await Bun.write(longCommentDocumentPath, await createLongWordCommentFixture());
await Bun.write(
  multiPageTableDocumentPath,
  await createMultiPageWordTableFixture(),
);
await Bun.write(themeTableDocumentPath, await createThemeTableWordFixture());
await Bun.write(styledTableDocumentPath, await createStyledTableWordFixture());
await Bun.write(wpsLayoutDocumentPath, await createWpsLayoutWordFixture());
await Bun.write(
  wpsFontMatrixDocumentPath,
  await createWpsFontMatrixWordFixture(),
);
await Bun.write(
  wpsCjkFontMatrixDocumentPath,
  await createWpsCjkFontMatrixWordFixture(),
);
await Bun.write(
  wpsGridMatrixDocumentPath,
  await createWpsGridMatrixWordFixture(),
);
await Bun.write(
  wpsScriptMatrixDocumentPath,
  await createWpsScriptMatrixWordFixture(),
);
await Bun.write(
  picturePath,
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAYAAADMtn8nAAAB+klEQVR4nO3bwQnCUBRFwZRjT25TidUKukoDsQARMfj5OTCL2V94nOVbLut9B5qW2QOA4wQMYQKGsLeAn9sOnJSAIUzAECZgCBMwhAkYwgQMYQKGMAFDmIAhTMAQJmAIEzCECRjChgX86xvUbLMPAUcIWMCECVjAhAlYwIQJWMCECVjAhAlYwIQJWMCECVjAhAlYwIQJ+E8BX28P+EjAAiZMwAImTMACJkzAAiZMwAImTMACJkzAAiZMwAImTMACJkzAAiZMwAImTMACJkzAAiZMwCcPGGYQsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACRsWMDCegCFMwBAmYAgTMIQJGMIEDGEChjABQ5iAIUzAECZgCBMwhAkYwr4GDHQIGMIEDGEChrAXam5Zu0ZEGKIAAAAASUVORK5CYII=',
    'base64',
  ),
);

console.log(`Created ${pdfPath}`);
console.log(`Created ${documentComparisonPath}`);
console.log(`Created ${longDocumentPath}`);
console.log(`Created ${longRevisionDocumentPath}`);
console.log(`Created ${longCommentDocumentPath}`);
console.log(`Created ${multiPageTableDocumentPath}`);
console.log(`Created ${themeTableDocumentPath}`);
console.log(`Created ${styledTableDocumentPath}`);
console.log(`Created ${wpsLayoutDocumentPath}`);
console.log(`Created ${wpsFontMatrixDocumentPath}`);
console.log(`Created ${wpsCjkFontMatrixDocumentPath}`);
console.log(`Created ${wpsGridMatrixDocumentPath}`);
console.log(`Created ${wpsScriptMatrixDocumentPath}`);
console.log(`Created ${picturePath}`);

async function createLongWordNavigationFixture(): Promise<Buffer> {
  const pageCount = 120;
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic long-document page navigation fixture',
    title: 'A3S Office 120-page navigation fixture',
    sections: [
      {
        children: Array.from({ length: pageCount }, (_, index) => {
          const page = index + 1;
          const paddedPage = String(page).padStart(3, '0');
          return [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              pageBreakBefore: index > 0,
              children: [
                new TextRun({
                  text: `A3S Office long document - Page ${paddedPage}`,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: `Deterministic navigation marker ${paddedPage}.`,
                }),
              ],
            }),
          ];
        }).flat(),
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function createLongWordRevisionFixture(): Promise<Buffer> {
  const revisionCount = 120;
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic long Word revision-review fixture',
    title: 'A3S Office 120-revision review fixture',
    sections: [
      {
        children: Array.from({ length: revisionCount }, (_, index) => {
          const revision = String(index + 1).padStart(3, '0');
          return new Paragraph({
            children: [
              new TextRun({ text: `Revision context ${revision}: ` }),
              new TextRun({
                text: `Deterministic revision marker ${revision}.`,
              }),
            ],
          });
        }),
      },
    ],
  });
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document));
  const documentXmlFile = archive.file('word/document.xml');
  const settingsXmlFile = archive.file('word/settings.xml');
  if (!documentXmlFile || !settingsXmlFile) {
    throw new Error('Generated Word revision fixture is missing package XML.');
  }

  let documentXml = await documentXmlFile.async('string');
  for (let index = 0; index < revisionCount; index += 1) {
    const revision = String(index + 1).padStart(3, '0');
    const text = `Deterministic revision marker ${revision}.`;
    const run = `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    if (!documentXml.includes(run)) {
      throw new Error(`Unable to mark Word fixture revision ${revision}.`);
    }
    documentXml = documentXml.replace(
      run,
      `<w:ins w:id="${index + 1}" w:author="A3S Test" w:date="2026-08-02T00:00:00.000Z">${run}</w:ins>`,
    );
  }
  const fixtureDate = new Date('2026-08-02T00:00:00.000Z');
  archive.file('word/document.xml', documentXml, { date: fixtureDate });

  const settingsXml = await settingsXmlFile.async('string');
  archive.file(
    'word/settings.xml',
    settingsXml.includes('<w:trackRevisions')
      ? settingsXml
      : settingsXml.replace(
          '</w:settings>',
          '<w:trackRevisions/></w:settings>',
        ),
    { date: fixtureDate },
  );
  return Buffer.from(
    await archive.generateAsync({
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      type: 'uint8array',
    }),
  );
}

async function createLongWordCommentFixture(): Promise<Buffer> {
  const commentCount = 120;
  const fixtureDate = new Date('2026-08-02T00:00:00.000Z');
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic long Word comment-review fixture',
    title: 'A3S Office 120-comment review fixture',
    comments: {
      children: Array.from({ length: commentCount }, (_, index) => {
        const comment = String(index + 1).padStart(3, '0');
        return {
          id: index,
          author: 'A3S Test',
          initials: 'AT',
          date: fixtureDate,
          children: [
            new Paragraph({
              children: [new TextRun({ text: `Review comment ${comment}.` })],
            }),
          ],
        };
      }),
    },
    sections: [
      {
        children: Array.from({ length: commentCount }, (_, index) => {
          const comment = String(index + 1).padStart(3, '0');
          return new Paragraph({
            pageBreakBefore: index > 0 && index % 8 === 0,
            children: [
              new TextRun({ text: `Comment context ${comment}: ` }),
              new CommentRangeStart(index),
              new TextRun({
                text: `Deterministic comment marker ${comment}.`,
              }),
              new CommentRangeEnd(index),
              new CommentReference(index),
            ],
          });
        }),
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function createMultiPageWordTableFixture(): Promise<Buffer> {
  const paragraphCount = 120;
  const bodyParagraphs = Array.from({ length: paragraphCount }, (_, index) => {
    const marker = String(index + 1).padStart(3, '0');
    return new Paragraph({
      children: [
        new TextRun({
          text: `Deterministic multi-page table row ${marker}.`,
        }),
      ],
    });
  });
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic multi-page Word table-row fixture',
    title: 'A3S Office multi-page table fixture',
    sections: [
      {
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Multi-page table layout' })],
          }),
          new Table({
            rows: [
              new TableRow({
                tableHeader: true,
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: 'Repeated heading' })],
                      }),
                    ],
                  }),
                ],
              }),
              new TableRow({
                cantSplit: false,
                children: [new TableCell({ children: bodyParagraphs })],
              }),
            ],
          }),
          new Paragraph({
            children: [new TextRun({ text: 'After multi-page table.' })],
          }),
        ],
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function createThemeTableWordFixture(): Promise<Buffer> {
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic Word theme-table fixture',
    title: 'A3S Office theme table fixture',
    sections: [
      {
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Theme table rendering' })],
          }),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: 'Theme border and fill' }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          new Paragraph({
            children: [new TextRun({ text: 'After theme table.' })],
          }),
        ],
      },
    ],
  });
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document));
  const documentXmlFile = archive.file('word/document.xml');
  const contentTypesFile = archive.file('[Content_Types].xml');
  const relationshipsFile = archive.file('word/_rels/document.xml.rels');
  if (!documentXmlFile || !contentTypesFile || !relationshipsFile) {
    throw new Error('Expected a complete DOCX package for the theme fixture.');
  }
  const [documentXml, contentTypes, relationships] = await Promise.all([
    documentXmlFile.async('string'),
    contentTypesFile.async('string'),
    relationshipsFile.async('string'),
  ]);
  const themedBorders = [
    '<w:tblBorders>',
    '<w:top w:val="single" w:sz="9" w:color="4472C4" w:themeColor="accent1" w:themeTint="80"/>',
    '<w:left w:val="dotted" w:sz="6" w:color="A5A5A5" w:themeColor="accent3"/>',
    '<w:bottom w:val="dashed" w:sz="6" w:color="ED7D31" w:themeColor="accent2"/>',
    '<w:right w:val="double" w:sz="12" w:color="4472C4" w:themeColor="accent1" w:themeShade="BF"/>',
    '<w:insideH w:val="nil"/>',
    '<w:insideV w:val="nil"/>',
    '</w:tblBorders>',
  ].join('');
  const withBorders = documentXml.replace(
    /<w:tblBorders>[\s\S]*?<\/w:tblBorders>/,
    themedBorders,
  );
  if (withBorders === documentXml) {
    throw new Error('Failed to apply deterministic theme table borders.');
  }
  const withCellShading = withBorders.replace(
    '<w:tc>',
    [
      '<w:tc>',
      '<w:tcPr>',
      '<w:shd w:val="clear" w:fill="ED7D31" w:themeFill="accent2" w:themeFillTint="99"/>',
      '</w:tcPr>',
    ].join(''),
  );
  if (withCellShading === withBorders) {
    throw new Error('Failed to apply deterministic theme table shading.');
  }
  archive.file('word/document.xml', withCellShading);
  archive.file('word/theme/theme1.xml', wordThemeFixtureXml());
  archive.file(
    '[Content_Types].xml',
    contentTypes.replace(
      '</Types>',
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    ),
  );
  archive.file(
    'word/_rels/document.xml.rels',
    relationships.replace(
      '</Relationships>',
      '<Relationship Id="rIdA3STheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
    ),
  );
  return archive.generateAsync({ type: 'nodebuffer' });
}

async function createStyledTableWordFixture(): Promise<Buffer> {
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic Word conditional table-style fixture',
    title: 'A3S Office styled table fixture',
    sections: [
      {
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Styled report table' })],
          }),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: 'Region' })],
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: 'Revenue' })],
                      }),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: 'North' })],
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: '$120,000' })],
                      }),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: 'Total' })],
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [new TextRun({ text: '$120,000' })],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          new Paragraph({
            children: [new TextRun({ text: 'After styled table.' })],
          }),
        ],
      },
    ],
  });
  const archive = await JSZip.loadAsync(await Packer.toBuffer(document));
  const documentXmlFile = archive.file('word/document.xml');
  const stylesXmlFile = archive.file('word/styles.xml');
  const contentTypesFile = archive.file('[Content_Types].xml');
  const relationshipsFile = archive.file('word/_rels/document.xml.rels');
  if (
    !documentXmlFile ||
    !stylesXmlFile ||
    !contentTypesFile ||
    !relationshipsFile
  ) {
    throw new Error('Expected a complete DOCX package for the styled table.');
  }
  const [documentXml, stylesXml, contentTypes, relationships] =
    await Promise.all([
      documentXmlFile.async('string'),
      stylesXmlFile.async('string'),
      contentTypesFile.async('string'),
      relationshipsFile.async('string'),
    ]);
  const tableProperties = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(documentXml)?.[0];
  if (!tableProperties) {
    throw new Error('Failed to locate the styled-table properties.');
  }
  const conditionalTableProperties = tableProperties
    .replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/, '')
    .replace(/<w:tblW\b[^>]*\/>/, '')
    .replace(/<w:tblLayout\b[^>]*\/>/, '')
    .replace(/<w:jc\b[^>]*\/>/, '')
    .replace(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/, '')
    .replace(
      '<w:tblPr>',
      [
        '<w:tblPr>',
        '<w:tblStyle w:val="A3SReportTable"/>',
        '<w:tblW w:type="pct" w:w="3125"/>',
        '<w:jc w:val="center"/>',
        '<w:tblLayout w:type="autofit"/>',
        '<w:tblCellMar>',
        '<w:top w:type="dxa" w:w="0"/>',
        '<w:left w:type="dxa" w:w="144"/>',
        '<w:bottom w:type="dxa" w:w="72"/>',
        '<w:right w:type="dxa" w:w="216"/>',
        '</w:tblCellMar>',
        '<w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
      ].join(''),
    );
  const cellMargins =
    '<w:tcMar><w:top w:type="dxa" w:w="120"/><w:right w:type="dxa" w:w="240"/></w:tcMar>';
  const tableGeometryXml = documentXml.replace(
    tableProperties,
    conditionalTableProperties,
  );
  const geometryDocumentXml = tableGeometryXml.includes('<w:tcPr>')
    ? tableGeometryXml.replace('<w:tcPr>', `<w:tcPr>${cellMargins}`)
    : tableGeometryXml.replace(
        '<w:tc>',
        `<w:tc><w:tcPr>${cellMargins}</w:tcPr>`,
      );
  if (!geometryDocumentXml.includes(cellMargins)) {
    throw new Error('Failed to add the styled-table cell margins.');
  }
  archive.file('word/document.xml', geometryDocumentXml);
  archive.file(
    'word/styles.xml',
    stylesXml.replace(
      '</w:styles>',
      `${wordTableStyleFixtureXml()}</w:styles>`,
    ),
  );
  archive.file('word/theme/theme1.xml', wordThemeFixtureXml());
  archive.file(
    '[Content_Types].xml',
    contentTypes.includes('/word/theme/theme1.xml')
      ? contentTypes
      : contentTypes.replace(
          '</Types>',
          '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
        ),
  );
  archive.file(
    'word/_rels/document.xml.rels',
    relationships.includes('/relationships/theme')
      ? relationships
      : relationships.replace(
          '</Relationships>',
          '<Relationship Id="rIdA3SStyledTableTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
        ),
  );
  return archive.generateAsync({ type: 'nodebuffer' });
}

async function createWpsLayoutWordFixture(): Promise<Buffer> {
  const bodyRun = (
    text: string,
    options: { bold?: boolean; color?: string } = {},
  ) =>
    new TextRun({
      text,
      font: 'Arial',
      size: 21,
      bold: options.bold ?? false,
      italics: false,
      color: options.color ?? '000000',
    });
  const tableParagraph = (
    text: string,
    options: {
      alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
      bold?: boolean;
      color?: string;
    } = {},
  ) =>
    new Paragraph({
      alignment: options.alignment ?? AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: 240 },
      children: [bodyRun(text, { bold: options.bold, color: options.color })],
    });
  const cell = (
    text: string,
    options: {
      alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
      bold?: boolean;
      color?: string;
      fill: string;
      width: number;
    },
  ) =>
    new TableCell({
      width: { size: options.width, type: WidthType.DXA },
      verticalAlign: VerticalAlignTable.CENTER,
      margins: { top: 120, right: 240, bottom: 120, left: 240 },
      shading: {
        type: ShadingType.CLEAR,
        color: 'auto',
        fill: options.fill,
      },
      children: [
        tableParagraph(text, {
          alignment: options.alignment,
          bold: options.bold,
          color: options.color,
        }),
      ],
    });
  const border = {
    style: BorderStyle.SINGLE,
    size: 6,
    color: '4472C4',
  } as const;
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic WPS and A3S Office layout-parity fixture',
    title: 'A3S Office WPS layout fixture',
    styles: {
      default: {
        document: {
          run: {
            font: 'Arial',
            size: 21,
            bold: false,
            italics: false,
            color: '000000',
          },
          paragraph: {
            spacing: { before: 0, after: 160, line: 252 },
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
              header: 720,
              footer: 720,
              gutter: 0,
            },
          },
        },
        children: [
          new Paragraph({
            spacing: { before: 0, after: 240, line: 384 },
            children: [
              new TextRun({
                text: 'Quarterly revenue report',
                font: 'Arial',
                size: 32,
                bold: false,
                italics: false,
                color: '2E74B5',
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 160, line: 252 },
            children: [
              bodyRun(
                'Direct formatting keeps this reference stable in WPS and A3S Office.',
              ),
            ],
          }),
          new Table({
            width: { size: 5000, type: WidthType.DXA },
            columnWidths: [3000, 2000],
            alignment: AlignmentType.CENTER,
            layout: TableLayoutType.FIXED,
            margins: { top: 120, right: 240, bottom: 120, left: 240 },
            borders: {
              top: border,
              right: border,
              bottom: border,
              left: border,
              insideHorizontal: border,
              insideVertical: border,
            },
            rows: [
              new TableRow({
                tableHeader: true,
                cantSplit: true,
                children: [
                  cell('Region', {
                    bold: true,
                    color: 'FFFFFF',
                    fill: '4472C4',
                    width: 3000,
                  }),
                  cell('Revenue', {
                    alignment: AlignmentType.RIGHT,
                    bold: true,
                    color: 'FFFFFF',
                    fill: '4472C4',
                    width: 2000,
                  }),
                ],
              }),
              new TableRow({
                cantSplit: true,
                children: [
                  cell('North', { fill: 'FCE4D6', width: 3000 }),
                  cell('$120,000', {
                    alignment: AlignmentType.RIGHT,
                    fill: 'FCE4D6',
                    width: 2000,
                  }),
                ],
              }),
              new TableRow({
                cantSplit: true,
                children: [
                  cell('Total', {
                    bold: true,
                    fill: 'E7E6E6',
                    width: 3000,
                  }),
                  cell('$120,000', {
                    alignment: AlignmentType.RIGHT,
                    bold: true,
                    fill: 'E7E6E6',
                    width: 2000,
                  }),
                ],
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 160, after: 0, line: 252 },
            children: [bodyRun('End of report.')],
          }),
        ],
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function createWpsFontMatrixWordFixture(): Promise<Buffer> {
  const samples = [
    ['Arial', 21],
    ['Arial', 32],
    ['Times New Roman', 21],
    ['Times New Roman', 32],
    ['Calibri', 21],
    ['Calibri', 32],
    ['Segoe UI', 21],
    ['Segoe UI', 32],
    ['Microsoft YaHei', 21],
    ['Microsoft YaHei', 32],
  ] as const;
  return createWpsTextMatrixWordFixture(
    'WPS Latin and Chinese system-font line metrics',
    samples.flatMap(([font, size]) =>
      [240, 240, 360].map((line, index) => ({
        font,
        label: `${font} ${size / 2}pt ${index + 1} line ${line}`,
        line,
        size,
      })),
    ),
  );
}

async function createWpsCjkFontMatrixWordFixture(): Promise<Buffer> {
  const fonts = [
    'Microsoft YaHei',
    'SimSun',
    'SimHei',
    'FangSong',
    'KaiTi',
    'DengXian',
  ] as const;
  const chineseSample = '\u4e2d\u6587\u6d4b\u8bd5';
  return createWpsTextMatrixWordFixture(
    'WPS CJK system-font line metrics',
    fonts.flatMap((font) =>
      ['Latin sample', chineseSample].flatMap((sample) =>
        [240, 240, 360].map((line, index) => ({
          font,
          label: `${font} ${sample} ${index + 1}`,
          line,
          size: 21,
        })),
      ),
    ),
  );
}

async function createWpsGridMatrixWordFixture(): Promise<Buffer> {
  const samples = [
    ['Arial', 21],
    ['Arial', 32],
    ['Segoe UI', 21],
    ['Segoe UI', 32],
    ['Microsoft YaHei', 21],
    ['Microsoft YaHei', 32],
  ] as const;
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic WPS document-grid layout fixture',
    title: 'A3S Office WPS document-grid matrix',
    sections: [
      {
        properties: {
          grid: { type: DocumentGridType.LINES, linePitch: 240 },
          page: wpsMatrixPageProperties(),
        },
        children: samples.flatMap(([font, size]) =>
          Array.from(
            { length: 3 },
            (_, index) =>
              new Paragraph({
                spacing: { before: 0, after: 0, line: 240 },
                children: [
                  new TextRun({
                    text: `${font} ${size / 2}pt grid row ${index + 1}`,
                    font,
                    size,
                    bold: false,
                    italics: false,
                    color: '000000',
                  }),
                ],
              }),
          ),
        ),
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function createWpsScriptMatrixWordFixture(): Promise<Buffer> {
  const font = {
    ascii: 'Segoe UI',
    hAnsi: 'Calibri',
    eastAsia: 'Microsoft YaHei',
    cs: 'Arial',
  } as const;
  const scriptSamples = [
    {
      bidirectional: false,
      label: 'Latin script sample',
      rightToLeft: false,
    },
    {
      bidirectional: false,
      label: '\u4e2d\u6587\u5b57\u4f53\u6d4b\u8bd5',
      rightToLeft: false,
    },
    {
      bidirectional: true,
      label:
        '\u0627\u062e\u062a\u0628\u0627\u0631 \u0627\u0644\u062e\u0637 \u0627\u0644\u0639\u0631\u0628\u064a',
      rightToLeft: true,
    },
    {
      bidirectional: true,
      label:
        '\u05d1\u05d3\u05d9\u05e7\u05ea \u05d2\u05d5\u05e4\u05df \u05e2\u05d1\u05e8\u05d9',
      rightToLeft: true,
    },
  ] as const;
  const scriptParagraphs = scriptSamples.flatMap(
    ({ bidirectional, label, rightToLeft }) =>
      [24, 32].flatMap((size) =>
        [240, 240, 360].map(
          (line) =>
            new Paragraph({
              bidirectional,
              spacing: { before: 0, after: 0, line },
              children: [
                new TextRun({
                  text: label,
                  font,
                  size,
                  sizeComplexScript: size,
                  rightToLeft,
                  bold: false,
                  boldComplexScript: false,
                  italics: false,
                  italicsComplexScript: false,
                  color: '000000',
                }),
              ],
            }),
        ),
      ),
  );
  const mixedParagraphs = Array.from(
    { length: 6 },
    (_, index) =>
      new Paragraph({
        spacing: { before: 0, after: 0, line: index === 5 ? 360 : 240 },
        children: [
          new TextRun({
            text: `Mixed ${index + 1} `,
            font,
            size: 21,
            bold: index % 2 === 0,
            italics: index % 2 !== 0,
            color: '000000',
          }),
          new TextRun({
            text: '\u0627\u0644\u0639\u0631\u0628\u064a\u0629 ',
            font,
            size: 21,
            sizeComplexScript: 21,
            rightToLeft: true,
            boldComplexScript: index % 2 === 0,
            italicsComplexScript: index % 2 !== 0,
            color: '000000',
          }),
          new TextRun({
            text: '\u05e2\u05d1\u05e8\u05d9\u05ea ',
            font,
            size: 21,
            sizeComplexScript: 21,
            rightToLeft: true,
            boldComplexScript: index % 2 === 0,
            italicsComplexScript: index % 2 !== 0,
            color: '000000',
          }),
          new TextRun({
            text: '\u4e2d\u6587',
            font,
            size: 21,
            bold: index % 2 === 0,
            italics: index % 2 !== 0,
            color: '000000',
          }),
        ],
      }),
  );
  const document = new Document({
    creator: 'A3S Lab',
    description:
      'Deterministic WPS Arabic, Hebrew, CJK, Latin, and mixed-run layout fixture',
    title: 'A3S Office WPS script and mixed-formatting matrix',
    sections: [
      {
        properties: { page: wpsMatrixPageProperties() },
        children: [...scriptParagraphs, ...mixedParagraphs],
      },
    ],
  });
  return Packer.toBuffer(document);
}

async function createWpsTextMatrixWordFixture(
  title: string,
  samples: readonly {
    font: string;
    label: string;
    line: number;
    size: number;
  }[],
): Promise<Buffer> {
  const document = new Document({
    creator: 'A3S Lab',
    description: 'Deterministic WPS system-font layout fixture',
    title,
    sections: [
      {
        properties: { page: wpsMatrixPageProperties() },
        children: samples.map(
          ({ font, label, line, size }) =>
            new Paragraph({
              spacing: { before: 0, after: 0, line },
              children: [
                new TextRun({
                  text: label,
                  font,
                  size,
                  bold: false,
                  italics: false,
                  color: '000000',
                }),
              ],
            }),
        ),
      },
    ],
  });
  return Packer.toBuffer(document);
}

function wpsMatrixPageProperties() {
  return {
    size: { width: 11_906, height: 16_838 },
    margin: {
      top: 1_440,
      right: 1_440,
      bottom: 1_440,
      left: 1_440,
      header: 720,
      footer: 720,
      gutter: 0,
    },
  };
}

function wordTableStyleFixtureXml(): string {
  return [
    '<w:style w:type="table" w:customStyle="1" w:styleId="A3SBaseTable">',
    '<w:name w:val="A3S Base Table"/>',
    '<w:rPr><w:color w:themeColor="dk1"/></w:rPr>',
    '<w:tblPr><w:tblBorders>',
    '<w:top w:val="single" w:sz="6" w:themeColor="accent1"/>',
    '<w:right w:val="single" w:sz="6" w:themeColor="accent1"/>',
    '<w:bottom w:val="single" w:sz="6" w:themeColor="accent1"/>',
    '<w:left w:val="single" w:sz="6" w:themeColor="accent1"/>',
    '<w:insideH w:val="dashed" w:sz="6" w:themeColor="accent3"/>',
    '<w:insideV w:val="nil"/>',
    '</w:tblBorders></w:tblPr>',
    '<w:tcPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:tcPr>',
    '<w:tblStylePr w:type="wholeTable"><w:pPr>',
    '<w:spacing w:after="80"/><w:keepLines/><w:widowControl/>',
    '</w:pPr></w:tblStylePr>',
    '<w:tblStylePr w:type="band1Horz"><w:tcPr>',
    '<w:shd w:val="clear" w:themeFill="accent2" w:themeFillTint="99"/>',
    '</w:tcPr></w:tblStylePr>',
    '<w:tblStylePr w:type="firstRow"><w:pPr>',
    '<w:jc w:val="center"/><w:spacing w:after="0"/><w:keepNext/>',
    '</w:pPr><w:rPr><w:b/><w:color w:themeColor="lt1"/></w:rPr><w:tcPr>',
    '<w:shd w:val="clear" w:themeFill="accent1"/>',
    '<w:tcBorders><w:bottom w:val="double" w:sz="12" w:themeColor="lt1"/></w:tcBorders>',
    '</w:tcPr></w:tblStylePr>',
    '</w:style>',
    '<w:style w:type="table" w:customStyle="1" w:styleId="A3SReportTable">',
    '<w:name w:val="A3S Report Table"/>',
    '<w:basedOn w:val="A3SBaseTable"/>',
    '<w:tblStylePr w:type="lastRow"><w:pPr>',
    '<w:jc w:val="right"/><w:ind w:right="180"/>',
    '</w:pPr><w:rPr><w:b/></w:rPr><w:tcPr>',
    '<w:shd w:val="clear" w:themeFill="accent3"/>',
    '</w:tcPr></w:tblStylePr>',
    '</w:style>',
  ].join('');
}

function wordThemeFixtureXml(): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="A3S Office Test Theme">',
    '<a:themeElements>',
    '<a:clrScheme name="A3S Office">',
    '<a:dk1><a:srgbClr val="000000"/></a:dk1>',
    '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>',
    '<a:dk2><a:srgbClr val="172033"/></a:dk2>',
    '<a:lt2><a:srgbClr val="F7F8FA"/></a:lt2>',
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>',
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>',
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>',
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>',
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>',
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>',
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>',
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>',
    '</a:clrScheme>',
    '<a:fontScheme name="A3S Office">',
    '<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>',
    '<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>',
    '</a:fontScheme>',
    '<a:fmtScheme name="A3S Office"/>',
    '</a:themeElements>',
    '</a:theme>',
  ].join('');
}

function createPdfThumbnailKeyboardFixture(): ArrayBuffer {
  const pdf = new jsPDF({
    compress: true,
    format: 'a4',
    orientation: 'portrait',
    unit: 'pt',
  });
  pdf.setCreationDate(new Date('2026-01-01T00:00:00.000Z'));
  pdf.setFileId('00000000000000000000000000000001');
  pdf.setProperties({
    author: 'A3S Lab',
    creator: 'A3S Office E2E tests',
    subject: 'PDF thumbnail keyboard navigation',
    title: 'A3S Office PDF keyboard fixture',
  });

  for (let page = 1; page <= 4; page += 1) {
    if (page > 1) pdf.addPage('a4', 'portrait');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(24);
    pdf.text(`A3S Office - Page ${page}`, 72, 96);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(12);
    pdf.text('Deterministic PDF thumbnail keyboard fixture.', 72, 124);
  }

  return pdf.output('arraybuffer');
}
