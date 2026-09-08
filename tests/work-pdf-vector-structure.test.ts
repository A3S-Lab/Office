import { expect, test } from '@rstest/core';
import { jsPDF } from 'jspdf';
import {
  applyWorkPdfDocumentStructure,
  collectWorkPdfOutlineEntriesFromRoot,
  normalizePdfLanguage,
} from '../src/internal/features/work/work-pdf-structure';
import {
  appendWorkPdfVectorTextLayer,
  clearWorkPdfTextRunsOnCanvas,
  workPdfFontStyleFromCss,
  workPdfTextColorFromCss,
} from '../src/internal/features/work/work-pdf-vector-text';

test('parses CSS color and font style for vector text runs', () => {
  expect(workPdfTextColorFromCss('rgb(17, 34, 51)')).toBe('#112233');
  expect(workPdfTextColorFromCss('#abc')).toBe('#aabbcc');
  expect(
    workPdfFontStyleFromCss({ fontStyle: 'italic', fontWeight: '700' }),
  ).toBe('bolditalic');
  expect(
    workPdfFontStyleFromCss({ fontStyle: 'normal', fontWeight: '400' }),
  ).toBe('normal');
});

test('clears measured Latin run rectangles on the raster canvas', () => {
  const fillRectCalls: Array<[number, number, number, number]> = [];
  const canvas = {
    width: 200,
    height: 100,
    getContext(kind: string) {
      if (kind !== '2d') return null;
      return {
        fillStyle: '',
        restore() {},
        save() {},
        fillRect(x: number, y: number, w: number, h: number) {
          fillRectCalls.push([x, y, w, h]);
        },
      };
    },
  } as unknown as HTMLCanvasElement;
  clearWorkPdfTextRunsOnCanvas(
    canvas,
    [
      {
        fontSize: 12,
        height: 10,
        text: 'Hello',
        width: 40,
        x: 10,
        y: 20,
      },
    ],
    { height: 50, width: 100 },
    '#ffffff',
  );
  expect(fillRectCalls.length).toBe(1);
  const [x, y, w, h] = fillRectCalls[0];
  // scaleX=2, scaleY=2; padX=max(1,1)=1, padY=max(1,0.7)=1
  expect(x).toBe(19);
  expect(y).toBe(39);
  expect(w).toBe(82);
  expect(h).toBe(22);
});

test('writes a visible vector text layer extractable from PDF bytes', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorTextLayer(
    pdf,
    [
      {
        color: '#112233',
        fontSize: 16,
        fontStyle: 'bold',
        height: 18,
        text: 'Vector milestone',
        width: 120,
        x: 24,
        y: 36,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Vector milestone');
});

test('normalizes PDF language tags and rejects junk', () => {
  expect(normalizePdfLanguage('en-US')).toBe('en-US');
  expect(normalizePdfLanguage(' zh ')).toBe('zh');
  expect(normalizePdfLanguage('not a tag!!!')).toBeNull();
  expect(normalizePdfLanguage(undefined)).toBeNull();
});

test('collects heading outline entries within page bounds', () => {
  document.body.innerHTML = `
    <div id="root" class="work-pdf-export-page">
      <h1 id="on-page">Introduction</h1>
      <h2 id="off-page">Appendix</h2>
    </div>
  `;
  const root = document.getElementById('root');
  const onPage = document.getElementById('on-page');
  const offPage = document.getElementById('off-page');
  if (
    !(root instanceof HTMLElement) ||
    !(onPage instanceof HTMLElement) ||
    !(offPage instanceof HTMLElement)
  ) {
    throw new Error('Expected outline fixtures.');
  }
  onPage.getBoundingClientRect = () =>
    ({
      bottom: 40,
      height: 24,
      left: 10,
      right: 120,
      top: 16,
      width: 110,
      x: 10,
      y: 16,
      toJSON() {
        return this;
      },
    }) as DOMRect;
  offPage.getBoundingClientRect = () =>
    ({
      bottom: 900,
      height: 24,
      left: 10,
      right: 120,
      top: 876,
      width: 110,
      x: 10,
      y: 876,
      toJSON() {
        return this;
      },
    }) as DOMRect;
  expect(
    collectWorkPdfOutlineEntriesFromRoot(root, 2, {
      height: 200,
      left: 0,
      top: 0,
      width: 400,
    }),
  ).toEqual([{ pageNumber: 2, title: 'Introduction' }]);
});

test('applies title, language, and outline bookmarks to the PDF document', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: 'a4',
    compress: false,
  });
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en-US',
    outline: [{ pageNumber: 1, title: 'Overview' }],
    title: 'Quarterly plan',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Quarterly plan');
  expect(ascii).toContain('A3S Work');
  expect(ascii).toContain('Overview');
});
