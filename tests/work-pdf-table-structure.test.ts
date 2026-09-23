import { expect, test } from '@rstest/core';
import { jsPDF } from 'jspdf';
import { applyWorkPdfDocumentStructure } from '../src/internal/features/work/work-pdf-structure';
import {
  appendWorkPdfTableStructEntries,
  collectWorkPdfTableStructs,
} from '../src/internal/features/work/work-pdf-table-structure';

test('collects simple table row and cell structure', () => {
  const root = document.createElement('div');
  root.innerHTML =
    '<table><tr><th>Party</th><th>Obligation</th></tr><tr><td>Ada</td><td>Deliver</td></tr></table>';
  document.body.append(root);
  const table = root.querySelector('table')!;
  stubBoundingRect(table, { height: 80, left: 4, top: 8, width: 200 });

  const structs = collectWorkPdfTableStructs(root, {
    height: 200,
    left: 0,
    top: 0,
    width: 400,
  });
  root.remove();

  expect(structs).toEqual([
    {
      rows: [
        {
          cells: [
            { alt: 'Party', header: true },
            { alt: 'Obligation', header: true },
          ],
        },
        {
          cells: [
            { alt: 'Ada', header: false },
            { alt: 'Deliver', header: false },
          ],
        },
      ],
    },
  ]);
});

test('writes Table TR TH TD StructElems into the PDF structure tree', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfTableStructEntries(pdf, [
    {
      rows: [
        {
          cells: [
            { alt: 'Party', header: true },
            { alt: 'Obligation', header: true },
          ],
        },
        {
          cells: [
            { alt: 'Ada', header: false },
            { alt: 'Deliver', header: false },
          ],
        },
      ],
    },
  ]);
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en',
    title: 'Tabled doc',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('/S /Table');
  expect(ascii).toContain('/S /TR');
  expect(ascii).toContain('/S /TH');
  expect(ascii).toContain('/S /TD');
  expect(ascii).toContain('/Alt (Party)');
  expect(ascii).toContain('/Alt (Deliver)');
  const objects = [...ascii.matchAll(/(\d+) 0 obj\n([\s\S]*?)\nendobj/g)].map(
    (match) => ({ id: Number(match[1]), body: match[2] ?? '' }),
  );
  const table = objects.find((obj) => obj.body.includes('/S /Table'));
  const document = objects.find((obj) => obj.body.includes('/S /Document'));
  expect(table).toBeTruthy();
  expect(document).toBeTruthy();
  expect(document!.body).toContain(`${table!.id} 0 R`);
});

function stubBoundingRect(
  element: HTMLElement,
  box: { height: number; left: number; top: number; width: number },
): void {
  element.getBoundingClientRect = () =>
    ({
      bottom: box.top + box.height,
      height: box.height,
      left: box.left,
      right: box.left + box.width,
      top: box.top,
      width: box.width,
      x: box.left,
      y: box.top,
      toJSON() {
        return this;
      },
    }) as DOMRect;
}
