import { Buffer } from 'node:buffer';
import { mkdir } from 'node:fs/promises';
import path from 'node:path';
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from 'docx';
import JSZip from 'jszip';
import { jsPDF } from 'jspdf';

const fixtureDirectory = path.resolve(
  import.meta.dirname,
  '../.a3s-test/fixtures',
);
const pdfPath = path.join(fixtureDirectory, 'pdf-thumbnail-keyboard.pdf');
const picturePath = path.join(fixtureDirectory, 'word-picture.png');
const longDocumentPath = path.join(
  fixtureDirectory,
  'word-page-navigation-120.docx',
);
const longRevisionDocumentPath = path.join(
  fixtureDirectory,
  'word-revisions-120.docx',
);

await mkdir(fixtureDirectory, { recursive: true });
await Bun.write(pdfPath, createPdfThumbnailKeyboardFixture());
await Bun.write(longDocumentPath, await createLongWordNavigationFixture());
await Bun.write(
  longRevisionDocumentPath,
  await createLongWordRevisionFixture(),
);
await Bun.write(
  picturePath,
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAYAAADMtn8nAAAB+klEQVR4nO3bwQnCUBRFwZRjT25TidUKukoDsQARMfj5OTCL2V94nOVbLut9B5qW2QOA4wQMYQKGsLeAn9sOnJSAIUzAECZgCBMwhAkYwgQMYQKGMAFDmIAhTMAQJmAIEzCECRjChgX86xvUbLMPAUcIWMCECVjAhAlYwIQJWMCECVjAhAlYwIQJWMCECVjAhAlYwIQJ+E8BX28P+EjAAiZMwAImTMACJkzAAiZMwAImTMACJkzAAiZMwAImTMACJkzAAiZMwAImTMACJkzAAiZMwCcPGGYQsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACROwgAkTsIAJE7CACRsWMDCegCFMwBAmYAgTMIQJGMIEDGEChjABQ5iAIUzAECZgCBMwhAkYwr4GDHQIGMIEDGEChrAXam5Zu0ZEGKIAAAAASUVORK5CYII=',
    'base64',
  ),
);

console.log(`Created ${pdfPath}`);
console.log(`Created ${longDocumentPath}`);
console.log(`Created ${longRevisionDocumentPath}`);
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
