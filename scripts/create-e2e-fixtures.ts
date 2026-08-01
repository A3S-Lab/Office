import { mkdir } from 'node:fs/promises';
import path from 'node:path';
import { jsPDF } from 'jspdf';

const fixtureDirectory = path.resolve(
  import.meta.dirname,
  '../.a3s-test/fixtures',
);
const pdfPath = path.join(fixtureDirectory, 'pdf-thumbnail-keyboard.pdf');

await mkdir(fixtureDirectory, { recursive: true });
await Bun.write(pdfPath, createPdfThumbnailKeyboardFixture());

console.log(`Created ${pdfPath}`);

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
