import { expect, test } from '@rstest/core';
import { jsPDF } from 'jspdf';
import {
  appendWorkPdfFigureStructEntries,
  collectWorkPdfFigureBoxes,
} from '../src/internal/features/work/work-pdf-figure-structure';
import { applyWorkPdfDocumentStructure } from '../src/internal/features/work/work-pdf-structure';

test('collects img boxes with alt or title and skips empty alt', () => {
  const root = document.createElement('div');
  root.innerHTML = [
    '<img alt="Architecture diagram" src="a.png" />',
    '<img title="Fallback title" src="b.png" />',
    '<img alt="" src="c.png" />',
    '<img src="d.png" />',
  ].join('');
  document.body.append(root);
  const images = Array.from(root.querySelectorAll('img'));
  stubBoundingRect(images[0]!, { height: 40, left: 8, top: 12, width: 60 });
  stubBoundingRect(images[1]!, { height: 30, left: 8, top: 60, width: 50 });
  stubBoundingRect(images[2]!, { height: 20, left: 8, top: 100, width: 40 });
  stubBoundingRect(images[3]!, { height: 20, left: 8, top: 130, width: 40 });

  const boxes = collectWorkPdfFigureBoxes(root, {
    height: 200,
    left: 0,
    top: 0,
    width: 400,
  });
  root.remove();

  expect(boxes).toEqual([
    {
      alt: 'Architecture diagram',
      height: 40,
      width: 60,
      x: 8,
      y: 12,
    },
    {
      alt: 'Fallback title',
      height: 30,
      width: 50,
      x: 8,
      y: 60,
    },
  ]);
});

test('writes Figure StructElems with MCID marked-content into the PDF', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  // Ensure a page content stream exists before emitting Figure BDC markers.
  pdf.text(' ', 1, 1);
  appendWorkPdfFigureStructEntries(pdf, [
    {
      alt: 'Architecture diagram',
      height: 40,
      width: 60,
      x: 8,
      y: 12,
    },
  ]);
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en',
    title: 'Figured doc',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('/S /Figure');
  expect(ascii).toContain('/Alt (Architecture diagram)');
  expect(ascii).toMatch(
    /\/Figure << \/Alt \(Architecture diagram\).*\/MCID 0 >> BDC/,
  );
  const objects = [...ascii.matchAll(/(\d+) 0 obj\n([\s\S]*?)\nendobj/g)].map(
    (match) => ({ id: Number(match[1]), body: match[2] ?? '' }),
  );
  const figure = objects.find((obj) => obj.body.includes('/S /Figure'));
  const document = objects.find((obj) => obj.body.includes('/S /Document'));
  expect(figure).toBeTruthy();
  expect(document).toBeTruthy();
  expect(document!.body).toContain(`${figure!.id} 0 R`);
  expect(figure!.body).toMatch(/\/K 0\b/);
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
