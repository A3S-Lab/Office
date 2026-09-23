import { expect, test } from '@rstest/core';
import { jsPDF } from 'jspdf';
import {
  appendWorkPdfExternalLinkAnnotations,
  collectWorkPdfExternalLinkBoxes,
} from '../src/internal/features/work/work-pdf-link-annotations';
import { applyWorkPdfDocumentStructure } from '../src/internal/features/work/work-pdf-structure';

test('collects http(s) link boxes and skips internal or unsafe hrefs', () => {
  const root = document.createElement('div');
  root.innerHTML = [
    '<a href="https://a3s.dev/office">Product</a>',
    '<a href="http://example.com/path">Example</a>',
    '<a href="#Architecture">Internal</a>',
    '<a href="javascript:alert(1)">Script</a>',
    '<a href="ftp://files.example">Ftp</a>',
  ].join('');
  document.body.append(root);
  const anchors = Array.from(root.querySelectorAll('a'));
  stubBoundingRect(anchors[0]!, { height: 16, left: 10, top: 20, width: 80 });
  stubBoundingRect(anchors[1]!, { height: 16, left: 10, top: 40, width: 90 });
  stubBoundingRect(anchors[2]!, { height: 16, left: 10, top: 60, width: 70 });
  stubBoundingRect(anchors[3]!, { height: 16, left: 10, top: 80, width: 70 });
  stubBoundingRect(anchors[4]!, { height: 16, left: 10, top: 100, width: 70 });

  const boxes = collectWorkPdfExternalLinkBoxes(root, {
    height: 200,
    left: 0,
    top: 0,
    width: 400,
  });
  root.remove();

  expect(boxes).toEqual([
    {
      alt: 'Product',
      height: 16,
      href: 'https://a3s.dev/office',
      width: 80,
      x: 10,
      y: 20,
    },
    {
      alt: 'Example',
      height: 16,
      href: 'http://example.com/path',
      width: 90,
      x: 10,
      y: 40,
    },
  ]);
});

test('writes URI link annotations, Link StructElems, and OBJR association', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfExternalLinkAnnotations(
    pdf,
    [
      {
        alt: 'Product site',
        height: 20,
        href: 'https://a3s.dev/office',
        width: 100,
        x: 10,
        y: 30,
      },
    ],
    { height: 280, left: 0, top: 0, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en',
    title: 'Linked doc',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('/Subtype /Link');
  expect(ascii).toContain('/URI (https://a3s.dev/office)');
  expect(ascii).toContain('/S /Link');
  expect(ascii).toContain('/Alt (Product site)');
  expect(ascii).toContain('/Type /OBJR');
  const objects = [...ascii.matchAll(/(\d+) 0 obj\n([\s\S]*?)\nendobj/g)].map(
    (match) => ({ id: Number(match[1]), body: match[2] ?? '' }),
  );
  const annot = objects.find(
    (obj) =>
      obj.body.includes('/Subtype /Link') &&
      obj.body.includes('/URI (https://a3s.dev/office)'),
  );
  const link = objects.find(
    (obj) =>
      obj.body.includes('/S /Link') && obj.body.includes('/Type /StructElem'),
  );
  const document = objects.find((obj) => obj.body.includes('/S /Document'));
  expect(annot).toBeTruthy();
  expect(link).toBeTruthy();
  expect(document).toBeTruthy();
  expect(document!.body).toContain(`${link!.id} 0 R`);
  expect(link!.body).toContain(`/Obj ${annot!.id} 0 R`);
  expect(ascii).toMatch(new RegExp(`/Annots \\[\\s*${annot!.id} 0 R`));
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
