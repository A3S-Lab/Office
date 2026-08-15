import { expect, type Page } from '@playwright/test';
import { jsPDF } from 'jspdf';

export async function openPdfFixture(
  page: Page,
  options: { pageCount?: number } = {},
): Promise<void> {
  await page
    .locator('input[aria-label="打开 Office 或 PDF 文件"]')
    .setInputFiles({
      name: 'visual-fixture.pdf',
      mimeType: 'application/pdf',
      buffer: createPdfFixture(options.pageCount ?? 1),
    });
}

export async function waitForPdfFixture(page: Page): Promise<void> {
  await expect(page.locator('.work-pdf-embed')).toHaveAttribute(
    'data-ready',
    'true',
    { timeout: 50_000 },
  );
  await expect
    .poll(
      () =>
        page
          .locator('.work-pdf-native-viewer img')
          .evaluateAll((images) =>
            images.some(
              (image) =>
                image instanceof HTMLImageElement &&
                image.complete &&
                image.naturalWidth > 0,
            ),
          ),
      { timeout: 50_000 },
    )
    .toBe(true);
}

function createPdfFixture(pageCount: number): Buffer {
  const pdf = new jsPDF({
    compress: true,
    format: 'a4',
    orientation: 'portrait',
    unit: 'pt',
  });
  pdf.setCreationDate(new Date('2026-01-01T00:00:00.000Z'));
  pdf.setFileId('A3S0FF1CE00000000000000000000001');
  pdf.setProperties({
    author: 'A3S Lab',
    creator: 'A3S Office visual tests',
    subject: 'Deterministic PDF editor fixture',
    title: 'A3S Office',
  });
  pdf.setFont('helvetica', 'bold');
  pdf.setFontSize(24);
  pdf.text('A3S Office', 72, 96);
  pdf.setFont('helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.text('PDF editor visual fixture', 72, 124);
  pdf.setDrawColor(40, 103, 216);
  pdf.setFillColor(238, 244, 255);
  pdf.roundedRect(72, 158, 451, 92, 8, 8, 'FD');
  pdf.setTextColor(34, 52, 82);
  pdf.text(
    'Typed toolbar, PDFium canvas, annotations, search, and save.',
    92,
    194,
  );
  pdf.text(
    'This page is generated in memory by the visual regression test.',
    92,
    218,
  );
  for (let pageNumber = 2; pageNumber <= pageCount; pageNumber += 1) {
    pdf.addPage('a4', 'portrait');
    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(24);
    pdf.text(`A3S Office - Page ${pageNumber}`, 72, 96);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(12);
    pdf.text('Deterministic PDF navigation fixture.', 72, 124);
  }
  return Buffer.from(pdf.output('arraybuffer'));
}
