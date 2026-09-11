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
  stubBoundingRect(onPage, { left: 10, top: 16, width: 110, height: 24 });
  stubBoundingRect(offPage, { left: 10, top: 876, width: 110, height: 24 });
  expect(
    collectWorkPdfOutlineEntriesFromRoot(root, 2, {
      height: 200,
      left: 0,
      top: 0,
      width: 400,
    }),
  ).toEqual([{ level: 1, pageNumber: 2, title: 'Introduction' }]);
});

test('admits outline-level paragraphs into PDF bookmarks', () => {
  document.body.innerHTML = `
    <div id="root" class="work-pdf-export-page">
      <p id="outline" data-office-outline-level="0">Writer outline title</p>
    </div>
  `;
  const root = document.getElementById('root');
  const outline = document.getElementById('outline');
  if (!(root instanceof HTMLElement) || !(outline instanceof HTMLElement)) {
    throw new Error('Expected outline-level fixture.');
  }
  stubBoundingRect(outline, { left: 8, top: 20, width: 160, height: 18 });
  const entries = collectWorkPdfOutlineEntriesFromRoot(root, 1, {
    height: 200,
    left: 0,
    top: 0,
    width: 400,
  });
  expect(entries).toEqual([
    { level: 1, pageNumber: 1, title: 'Writer outline title' },
  ]);

  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: 'a4',
    compress: false,
  });
  applyWorkPdfDocumentStructure(pdf, {
    outline: entries,
    title: 'Outline-only export',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Writer outline title');
});

test('collects mixed h1 and outline-level paragraphs', () => {
  document.body.innerHTML = `
    <div id="root" class="work-pdf-export-page">
      <h1 id="heading">Chapter</h1>
      <p id="outline" data-office-outline-level="1">Section via outline</p>
      <p data-office-outline-level="9">Body-level ignored</p>
      <p data-office-outline-level="abc">Invalid ignored</p>
    </div>
  `;
  const root = document.getElementById('root');
  const heading = document.getElementById('heading');
  const outline = document.getElementById('outline');
  if (
    !(root instanceof HTMLElement) ||
    !(heading instanceof HTMLElement) ||
    !(outline instanceof HTMLElement)
  ) {
    throw new Error('Expected mixed outline fixtures.');
  }
  stubBoundingRect(heading, { left: 8, top: 12, width: 140, height: 22 });
  stubBoundingRect(outline, { left: 8, top: 40, width: 180, height: 18 });
  const entries = collectWorkPdfOutlineEntriesFromRoot(root, 1, {
    height: 240,
    left: 0,
    top: 0,
    width: 400,
  });
  expect(entries).toEqual([
    { level: 1, pageNumber: 1, title: 'Chapter' },
    { level: 2, pageNumber: 1, title: 'Section via outline' },
  ]);

  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: 'a4',
    compress: false,
  });
  applyWorkPdfDocumentStructure(pdf, { outline: entries });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Chapter');
  expect(ascii).toContain('Section via outline');
});

test('drops off-page outline-level paragraphs', () => {
  document.body.innerHTML = `
    <div id="root" class="work-pdf-export-page">
      <p id="on-page" data-office-outline-level="2">On page</p>
      <p id="off-page" data-office-outline-level="2">Off page</p>
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
    throw new Error('Expected off-page outline fixtures.');
  }
  stubBoundingRect(onPage, { left: 10, top: 30, width: 100, height: 16 });
  stubBoundingRect(offPage, { left: 10, top: 900, width: 100, height: 16 });
  expect(
    collectWorkPdfOutlineEntriesFromRoot(root, 1, {
      height: 200,
      left: 0,
      top: 0,
      width: 400,
    }),
  ).toEqual([{ level: 3, pageNumber: 1, title: 'On page' }]);
});

test('skips empty outline titles and caps entry count', () => {
  const paragraphs = Array.from({ length: 520 }, (_, index) => {
    if (index === 0) {
      return '<p data-office-outline-level="0">   </p>';
    }
    if (index === 1) {
      return '<p data-office-outline-level="0"></p>';
    }
    return `<p id="entry-${index}" data-office-outline-level="0">Title ${index}</p>`;
  }).join('');
  document.body.innerHTML = `<div id="root" class="work-pdf-export-page">${paragraphs}</div>`;
  const root = document.getElementById('root');
  if (!(root instanceof HTMLElement)) {
    throw new Error('Expected capped outline fixture.');
  }
  for (const node of Array.from(
    root.querySelectorAll<HTMLElement>('p[data-office-outline-level]'),
  )) {
    stubBoundingRect(node, { left: 4, top: 8, width: 80, height: 12 });
  }
  const entries = collectWorkPdfOutlineEntriesFromRoot(root, 1, {
    height: 400,
    left: 0,
    top: 0,
    width: 400,
  });
  expect(entries).toHaveLength(512);
  expect(entries[0]?.title).toBe('Title 2');
  expect(entries.at(-1)?.title).toBe('Title 513');
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
