import { expect, test } from '@rstest/core';
import { jsPDF } from 'jspdf';
import {
  appendWorkPdfInvisibleTextLayer,
  collectWorkPdfTextRuns,
  workPdfTextRunsFromClientRects,
} from '../src/internal/features/work/work-pdf-text-layer';

test('maps in-page client rects into searchable Latin text runs', () => {
  const runs = workPdfTextRunsFromClientRects(
    'Quarterly plan',
    [{ left: 120, top: 80, width: 140, height: 18 }],
    { left: 100, top: 50, width: 500, height: 700 },
    16,
  );
  expect(runs).toEqual([
    {
      fontSize: 16,
      height: 18,
      text: 'Quarterly plan',
      width: 140,
      x: 20,
      y: 30,
    },
  ]);
});

test('omits empty, out-of-page, and non-Latin runs instead of inventing geometry', () => {
  expect(
    workPdfTextRunsFromClientRects(
      '   ',
      [{ left: 0, top: 0, width: 10, height: 10 }],
      { left: 0, top: 0, width: 100, height: 100 },
      12,
    ),
  ).toEqual([]);
  expect(
    workPdfTextRunsFromClientRects(
      'Hello',
      [{ left: 400, top: 0, width: 20, height: 12 }],
      { left: 0, top: 0, width: 100, height: 100 },
      12,
    ),
  ).toEqual([]);
  expect(
    workPdfTextRunsFromClientRects(
      '中文标题',
      [{ left: 10, top: 10, width: 40, height: 16 }],
      { left: 0, top: 0, width: 200, height: 200 },
      14,
    ),
  ).toEqual([]);
});

test('collects text runs from a positioned page DOM', () => {
  document.body.innerHTML = `
    <div id="page" style="position: relative; width: 400px; height: 600px;">
      <p id="line" style="position: absolute; left: 24px; top: 40px; font-size: 16px; margin: 0;">Executive summary</p>
    </div>
  `;
  const page = document.getElementById('page');
  const line = document.getElementById('line');
  if (!(page instanceof HTMLElement) || !(line instanceof HTMLElement)) {
    throw new Error('Expected page fixtures.');
  }
  Object.defineProperty(window, 'getComputedStyle', {
    configurable: true,
    value: () => ({
      color: 'rgb(0, 0, 0)',
      fontSize: '16px',
      fontStyle: 'normal',
      fontWeight: '400',
    }),
  });
  const original = document.createRange.bind(document);
  document.createRange = () => {
    const range = original();
    range.getClientRects = () =>
      [
        {
          bottom: 56,
          height: 16,
          left: 24,
          right: 180,
          top: 40,
          width: 156,
          x: 24,
          y: 40,
          toJSON() {
            return this;
          },
        },
      ] as unknown as DOMRectList;
    return range;
  };
  try {
    const runs = collectWorkPdfTextRuns(page, {
      height: 600,
      left: 0,
      top: 0,
      width: 400,
    });
    expect(runs).toEqual([
      expect.objectContaining({
        fontSize: 16,
        text: 'Executive summary',
        x: 24,
        y: 40,
      }),
    ]);
  } finally {
    document.createRange = original;
  }
});

test('writes an invisible text layer that remains extractable from PDF bytes', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfInvisibleTextLayer(
    pdf,
    [
      {
        fontSize: 16,
        height: 18,
        text: 'Searchable milestone',
        width: 120,
        x: 24,
        y: 36,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = pdf.output('arraybuffer')
    ? Buffer.from(pdf.output('arraybuffer')).toString('latin1')
    : '';
  expect(ascii.includes('Searchable') || ascii.includes('(Searchable')).toBe(
    true,
  );
});
