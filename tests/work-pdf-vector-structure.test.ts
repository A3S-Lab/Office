import { expect, test } from '@rstest/core';
import { jsPDF } from 'jspdf';
import {
  applyWorkPdfDocumentStructure,
  collectWorkPdfOutlineEntriesFromRoot,
  encodePdfActualTextOperand,
  normalizePdfLanguage,
  workPdfStructRoleFromOutlineLevel,
} from '../src/internal/features/work/work-pdf-structure';
import {
  appendWorkPdfVectorHighlightLayer,
  appendWorkPdfVectorParagraphBorderLayer,
  appendWorkPdfVectorUnderlineLayer,
  clearWorkPdfParagraphBorderStripsOnCanvas,
  clearWorkPdfUnderlineStripsOnCanvas,
  collectWorkPdfParagraphBorderBoxes,
  workPdfParagraphBordersFromElement,
  workPdfRunHighlightFromElement,
  workPdfRunUnderlineFromElement,
} from '../src/internal/features/work/work-pdf-vector-paint';
import { documentParagraphBordersDomAttributes } from '../src/internal/features/work/work-document-paragraph-borders';
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
  expect(ascii).toContain('/ActualText (Vector milestone)');
  expect(ascii).toContain('/MCID 0');
  expect(ascii).toContain('/Span << /ActualText');
  expect(ascii).toContain('BDC');
  expect(ascii).toContain('EMC');
  expect(ascii).toContain('/MarkInfo << /Marked true >>');
  expect(ascii).toContain('/StructParents 0');
  expect(ascii).toContain('/ParentTree');
  expect(ascii).toContain('/S /Span');
  expect(ascii).toMatch(/\/K\s+0\b/);
});

test('encodes ActualText operands for Latin and Unicode runs', () => {
  expect(encodePdfActualTextOperand('Hello (world)')).toBe(
    '(Hello \\(world\\))',
  );
  expect(encodePdfActualTextOperand('路径')).toBe('<FEFF8DEF5F84>');
});

test('resolves Writer underline marks into PDF stroke kinds', () => {
  document.body.innerHTML = `
    <p>
      <u id="single" data-office-underline-style="single" data-office-underline-color="#112233">One</u>
      <u id="double" data-office-underline-style="double">Two</u>
      <u id="thick" data-office-underline-style="thick">Three</u>
      <u id="none" data-office-underline-style="none">Off</u>
      <span id="plain">Plain</span>
    </p>
  `;
  const single = document.getElementById('single');
  const double = document.getElementById('double');
  const thick = document.getElementById('thick');
  const none = document.getElementById('none');
  const plain = document.getElementById('plain');
  if (
    !(single instanceof HTMLElement) ||
    !(double instanceof HTMLElement) ||
    !(thick instanceof HTMLElement) ||
    !(none instanceof HTMLElement) ||
    !(plain instanceof HTMLElement)
  ) {
    throw new Error('Expected underline fixtures.');
  }
  Object.defineProperty(window, 'getComputedStyle', {
    configurable: true,
    value: (element: Element) => ({
      color: 'rgb(0, 0, 0)',
      textDecorationColor: 'currentcolor',
      textDecorationLine:
        element === plain ? 'none' : element === none ? 'none' : 'underline',
      textDecorationStyle: 'solid',
      textDecorationThickness: 'auto',
    }),
  });
  expect(workPdfRunUnderlineFromElement(single)).toEqual({
    color: '#112233',
    kind: 'single',
  });
  expect(workPdfRunUnderlineFromElement(double)).toEqual({
    color: '#000000',
    kind: 'double',
  });
  expect(workPdfRunUnderlineFromElement(thick)).toEqual({
    color: '#000000',
    kind: 'thick',
  });
  expect(workPdfRunUnderlineFromElement(none)).toBeNull();
  expect(workPdfRunUnderlineFromElement(plain)).toBeNull();
});

test('resolves Writer highlight marks into PDF fill colors', () => {
  document.body.innerHTML = `
    <p>
      <span id="yellow" data-office-highlight="yellow">One</span>
      <span id="cyan" data-office-highlight="cyan">Two</span>
      <span id="none" data-office-highlight="none">Off</span>
      <span id="plain">Plain</span>
      <span id="css-yellow" style="background-color: #ffff00">Css</span>
    </p>
  `;
  const yellow = document.getElementById('yellow');
  const cyan = document.getElementById('cyan');
  const none = document.getElementById('none');
  const plain = document.getElementById('plain');
  const cssYellow = document.getElementById('css-yellow');
  if (
    !(yellow instanceof HTMLElement) ||
    !(cyan instanceof HTMLElement) ||
    !(none instanceof HTMLElement) ||
    !(plain instanceof HTMLElement) ||
    !(cssYellow instanceof HTMLElement)
  ) {
    throw new Error('Expected highlight fixtures.');
  }
  Object.defineProperty(window, 'getComputedStyle', {
    configurable: true,
    value: (element: Element) => ({
      backgroundColor:
        element === cssYellow
          ? 'rgb(255, 255, 0)'
          : element === yellow
            ? 'rgb(255, 255, 0)'
            : element === cyan
              ? 'rgb(0, 255, 255)'
              : 'rgba(0, 0, 0, 0)',
    }),
  });
  expect(workPdfRunHighlightFromElement(yellow)).toEqual({
    color: '#ffff00',
  });
  expect(workPdfRunHighlightFromElement(cyan)).toEqual({
    color: '#00ffff',
  });
  expect(workPdfRunHighlightFromElement(none)).toBeNull();
  expect(workPdfRunHighlightFromElement(plain)).toBeNull();
  expect(workPdfRunHighlightFromElement(cssYellow)).toEqual({
    color: '#ffff00',
  });
});

test('paints underlined vector runs as PDF path operators', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  const runs = [
    {
      color: '#112233',
      fontSize: 16,
      fontStyle: 'normal' as const,
      height: 18,
      text: 'Underlined',
      underline: { color: '#112233', kind: 'double' as const },
      width: 90,
      x: 20,
      y: 40,
    },
  ];
  appendWorkPdfVectorTextLayer(
    pdf,
    runs,
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  appendWorkPdfVectorUnderlineLayer(
    pdf,
    runs,
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Underlined');
  // jsPDF flips y from top origin: pageHeight - y
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/110\.\s+[\d.]+\s+l/);
  expect(ascii).toContain('0.07 0.13 0.2 RG');
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(2);
});

test('paints highlighted vector runs as PDF fill operators', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  const runs = [
    {
      color: '#112233',
      fontSize: 16,
      fontStyle: 'normal' as const,
      height: 18,
      highlight: { color: '#ffff00' },
      text: 'Highlighted',
      width: 90,
      x: 20,
      y: 40,
    },
  ];
  appendWorkPdfVectorHighlightLayer(
    pdf,
    runs,
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  appendWorkPdfVectorTextLayer(
    pdf,
    runs,
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Highlighted');
  // jsPDF normalizes fill RGB and flips y from top origin.
  expect(ascii).toContain('1. 1. 0. rg');
  expect(ascii).toMatch(/20\.\s+240\.\s+90\.\s+-18\.\s+re/);
  expect(ascii).toMatch(/\nf\n/);
});

test('resolves Writer paragraph borders into PDF stroke plans', () => {
  const attributes = documentParagraphBordersDomAttributes({
    top: { style: 'single', color: { value: '#112233' }, size: 12 },
    bottom: { style: 'double', color: { value: '#445566' }, size: 18 },
    left: { style: 'dashed', color: { value: '#778899' }, size: 8 },
    right: { style: 'thick', color: { value: '#aabbcc' }, size: 24 },
  });
  const betweenBarAttributes = documentParagraphBordersDomAttributes({
    between: { style: 'dotted', color: { value: '#334455' }, size: 12 },
    bar: { style: 'single', color: { value: '#556677' }, size: 18 },
  });
  const artAttributes = documentParagraphBordersDomAttributes({
    top: { style: 'apples', color: { value: '#112233' }, size: 12 },
  });
  const zigZagAttributes = documentParagraphBordersDomAttributes({
    top: { style: 'zigZag', color: { value: '#112233' }, size: 12 },
    bottom: { style: 'zigZagStitch', color: { value: '#445566' }, size: 12 },
  });
  const sawtoothAttributes = documentParagraphBordersDomAttributes({
    top: { style: 'sawtooth', color: { value: '#112233' }, size: 12 },
    bottom: { style: 'sharksTeeth', color: { value: '#445566' }, size: 12 },
  });
  const triangleAttributes = documentParagraphBordersDomAttributes({
    top: { style: 'triangles', color: { value: '#112233' }, size: 12 },
    bottom: { style: 'triangle1', color: { value: '#445566' }, size: 12 },
    left: { style: 'triangle2', color: { value: '#778899' }, size: 12 },
  });
  const waveBetweenAttributes = documentParagraphBordersDomAttributes({
    between: { style: 'wave', color: { value: '#112233' }, size: 12 },
    bar: { style: 'threeDEmboss', color: { value: '#445566' }, size: 14 },
  });
  document.body.innerHTML = `
    <div id="root">
      <p id="bordered" data-office-paragraph-borders='${attributes['data-office-paragraph-borders']}' style="${attributes.style}">Bordered</p>
      <p id="between-bar" data-office-paragraph-borders='${betweenBarAttributes['data-office-paragraph-borders']}' style="${betweenBarAttributes.style}">Between bar</p>
      <p id="art" data-office-paragraph-borders='${artAttributes['data-office-paragraph-borders']}' style="${artAttributes.style}">Art</p>
      <p id="zigzag" data-office-paragraph-borders='${zigZagAttributes['data-office-paragraph-borders']}' style="${zigZagAttributes.style}">Zigzag</p>
      <p id="sawtooth" data-office-paragraph-borders='${sawtoothAttributes['data-office-paragraph-borders']}' style="${sawtoothAttributes.style}">Sawtooth</p>
      <p id="triangles" data-office-paragraph-borders='${triangleAttributes['data-office-paragraph-borders']}' style="${triangleAttributes.style}">Triangles</p>
      <p id="wave-between" data-office-paragraph-borders='${waveBetweenAttributes['data-office-paragraph-borders']}' style="${waveBetweenAttributes.style}">Wave between</p>
      <p id="plain">Plain</p>
      <p id="nil" data-office-paragraph-borders='{"top":{"style":"nil"}}'>Nil</p>
    </div>
  `;
  const bordered = document.getElementById('bordered');
  const betweenBar = document.getElementById('between-bar');
  const art = document.getElementById('art');
  const zigzag = document.getElementById('zigzag');
  const sawtooth = document.getElementById('sawtooth');
  const triangles = document.getElementById('triangles');
  const waveBetween = document.getElementById('wave-between');
  const plain = document.getElementById('plain');
  const nil = document.getElementById('nil');
  if (
    !(bordered instanceof HTMLElement) ||
    !(betweenBar instanceof HTMLElement) ||
    !(art instanceof HTMLElement) ||
    !(zigzag instanceof HTMLElement) ||
    !(sawtooth instanceof HTMLElement) ||
    !(triangles instanceof HTMLElement) ||
    !(waveBetween instanceof HTMLElement) ||
    !(plain instanceof HTMLElement) ||
    !(nil instanceof HTMLElement)
  ) {
    throw new Error('Expected paragraph border fixtures.');
  }
  expect(workPdfParagraphBordersFromElement(bordered)).toEqual({
    top: { color: '#112233', kind: 'single', width: 2 },
    bottom: { color: '#445566', kind: 'double', width: 3 },
    left: { color: '#778899', kind: 'dashed', width: 8 / 6 },
    right: { color: '#aabbcc', kind: 'thick', width: 4 },
  });
  expect(workPdfParagraphBordersFromElement(betweenBar)).toEqual({
    between: { color: '#334455', kind: 'dotted', width: 2 },
    bar: { color: '#556677', kind: 'thick', width: 3 },
  });
  expect(workPdfParagraphBordersFromElement(art)).toBeNull();
  expect(workPdfParagraphBordersFromElement(zigzag)).toEqual({
    top: { color: '#112233', kind: 'zigZag', width: 16 },
    bottom: { color: '#445566', kind: 'zigZagStitch', width: 16 },
  });
  expect(workPdfParagraphBordersFromElement(sawtooth)).toEqual({
    top: { color: '#112233', kind: 'sawtooth', width: 16 },
    bottom: { color: '#445566', kind: 'sharksTeeth', width: 16 },
  });
  expect(workPdfParagraphBordersFromElement(triangles)).toEqual({
    top: { color: '#112233', kind: 'triangles', width: 16 },
    bottom: { color: '#445566', kind: 'triangle1', width: 16 },
    left: { color: '#778899', kind: 'triangle2', width: 16 },
  });
  expect(workPdfParagraphBordersFromElement(waveBetween)).toEqual({
    between: { color: '#112233', kind: 'wave', width: 2 },
    bar: { color: '#445566', kind: 'threeDEmboss', width: 14 / 6 },
  });
  expect(workPdfParagraphBordersFromElement(plain)).toBeNull();
  expect(workPdfParagraphBordersFromElement(nil)).toBeNull();
});

test('resolves threeD, inset, and outset paragraph borders into relief plans', () => {
  const attributes = documentParagraphBordersDomAttributes({
    top: { style: 'threeDEmboss', color: { value: '#112233' }, size: 12 },
    bottom: { style: 'threeDEngrave', color: { value: '#445566' }, size: 12 },
    left: { style: 'inset', color: { value: '#778899' }, size: 12 },
    right: { style: 'outset', color: { value: '#aabbcc' }, size: 12 },
  });
  document.body.innerHTML = `
    <div id="root">
      <p id="relief" data-office-paragraph-borders='${attributes['data-office-paragraph-borders']}' style="${attributes.style}">Relief</p>
    </div>
  `;
  const relief = document.getElementById('relief');
  if (!(relief instanceof HTMLElement)) {
    throw new Error('Expected relief border fixture.');
  }
  expect(workPdfParagraphBordersFromElement(relief)).toEqual({
    top: { color: '#112233', kind: 'threeDEmboss', width: 2 },
    bottom: { color: '#445566', kind: 'threeDEngrave', width: 2 },
    left: { color: '#778899', kind: 'inset', width: 2 },
    right: { color: '#aabbcc', kind: 'outset', width: 2 },
  });
});

test('collects measured paragraph border boxes within page bounds', () => {
  const attributes = documentParagraphBordersDomAttributes({
    top: { style: 'single', color: { value: '#112233' }, size: 12 },
    bottom: { style: 'single', color: { value: '#112233' }, size: 12 },
  });
  const betweenBarAttributes = documentParagraphBordersDomAttributes({
    between: { style: 'dashed', color: { value: '#334455' }, size: 12 },
    bar: { style: 'single', color: { value: '#556677' }, size: 12 },
  });
  document.body.innerHTML = `
    <div id="root" class="work-pdf-export-page">
      <p id="on-page" data-office-paragraph-borders='${attributes['data-office-paragraph-borders']}' style="${attributes.style}">On page</p>
      <p id="between-bar" data-office-paragraph-borders='${betweenBarAttributes['data-office-paragraph-borders']}' style="${betweenBarAttributes.style}">Between bar</p>
      <p id="off-page" data-office-paragraph-borders='${attributes['data-office-paragraph-borders']}' style="${attributes.style}">Off page</p>
    </div>
  `;
  const root = document.getElementById('root');
  const onPage = document.getElementById('on-page');
  const betweenBar = document.getElementById('between-bar');
  const offPage = document.getElementById('off-page');
  if (
    !(root instanceof HTMLElement) ||
    !(onPage instanceof HTMLElement) ||
    !(betweenBar instanceof HTMLElement) ||
    !(offPage instanceof HTMLElement)
  ) {
    throw new Error('Expected border box fixtures.');
  }
  stubBoundingRect(onPage, { left: 12, top: 24, width: 160, height: 28 });
  stubBoundingRect(betweenBar, { left: 12, top: 60, width: 160, height: 28 });
  stubBoundingRect(offPage, { left: 12, top: 900, width: 160, height: 28 });
  expect(
    collectWorkPdfParagraphBorderBoxes(root, {
      height: 200,
      left: 0,
      top: 0,
      width: 400,
    }),
  ).toEqual([
    {
      edges: {
        top: { color: '#112233', kind: 'single', width: 2 },
        bottom: { color: '#112233', kind: 'single', width: 2 },
      },
      height: 28,
      width: 160,
      x: 12,
      y: 24,
    },
    {
      edges: {
        between: { color: '#334455', kind: 'dashed', width: 2 },
        bar: { color: '#556677', kind: 'single', width: 2 },
      },
      height: 28,
      width: 160,
      x: 12,
      y: 60,
    },
  ]);
});

test('paints paragraph borders as PDF path operators', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          top: { color: '#112233', kind: 'single', width: 2 },
          bottom: { color: '#112233', kind: 'double', width: 2 },
          left: { color: '#112233', kind: 'dashed', width: 2 },
          right: { color: '#112233', kind: 'single', width: 2 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('0.07 0.13 0.2 RG');
  // top/left/right + double bottom (2) => at least 5 stroked segments
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(5);
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/120\.\s+[\d.]+\s+l/);
});

test('paints between and bar paragraph borders as PDF path operators', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          between: { color: '#334455', kind: 'single', width: 2 },
          bar: { color: '#556677', kind: 'thick', width: 3 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  // between → bottom horizontal; bar → left vertical
  expect(ascii).toContain('0.2 0.27 0.33 RG');
  expect(ascii).toContain('0.33 0.4 0.47 RG');
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(2);
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/120\.\s+[\d.]+\s+l/);
});

test('paints wave and doubleWave paragraph borders as sine polylines', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          top: { color: '#112233', kind: 'wave', width: 2 },
          bottom: { color: '#445566', kind: 'doubleWave', width: 2 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toMatch(/0\.07\s+0\.13\s+0\.2\s+RG/);
  expect(ascii).toMatch(/0\.27\s+0\.33\s+0\.4\s+RG/);
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  // wave: many segments; doubleWave: two parallel polylines
  expect(strokeCount).toBeGreaterThanOrEqual(24);
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/\s+l\n/);
});

test('paints threeD emboss and engrave borders as dual-tone offsets', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          top: { color: '#112233', kind: 'threeDEmboss', width: 2 },
          bottom: { color: '#445566', kind: 'threeDEngrave', width: 2 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  // emboss: light then dark; engrave: dark then light (shifted from base)
  expect(ascii).toMatch(/0\.5[89]\s+0\.6[01]\s+0\.6[34]\s+RG/);
  expect(ascii).toMatch(/0\.0[23]\s+0\.0[56]\s+0\.0[89]\s+RG/);
  expect(ascii).toMatch(/0\.1[12]\s+0\.1[45]\s+0\.1[78]\s+RG/);
  expect(ascii).toMatch(/0\.6[67]\s+0\.7\s+0\.7[23]\s+RG/);
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(4);
});

test('paints zigZag and zigZagStitch art borders as chevron polylines', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          top: { color: '#112233', kind: 'zigZag', width: 2 },
          bottom: { color: '#445566', kind: 'zigZagStitch', width: 2 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toMatch(/0\.07\s+0\.13\s+0\.2\s+RG/);
  expect(ascii).toMatch(/0\.27\s+0\.33\s+0\.4\s+RG/);
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(20);
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/\s+l\n/);
});

test('paints sawtooth and sharksTeeth art borders as triangular teeth', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          top: { color: '#112233', kind: 'sawtooth', width: 2 },
          bottom: { color: '#445566', kind: 'sharksTeeth', width: 2 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toMatch(/0\.07\s+0\.13\s+0\.2\s+RG/);
  expect(ascii).toMatch(/0\.27\s+0\.33\s+0\.4\s+RG/);
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(20);
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/\s+l\n/);
});

test('paints triangles, triangle1, and triangle2 art borders as closed triangles', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  appendWorkPdfVectorParagraphBorderLayer(
    pdf,
    [
      {
        edges: {
          top: { color: '#112233', kind: 'triangles', width: 2 },
          bottom: { color: '#445566', kind: 'triangle1', width: 2 },
          left: { color: '#778899', kind: 'triangle2', width: 2 },
        },
        height: 40,
        width: 100,
        x: 20,
        y: 30,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toMatch(/0\.07\s+0\.13\s+0\.2\s+RG/);
  expect(ascii).toMatch(/0\.27\s+0\.33\s+0\.4\s+RG/);
  expect(ascii).toMatch(/0\.47\s+0\.53\s+0\.6\s+RG/);
  const strokeCount = (ascii.match(/\nS\n/g) ?? []).length;
  expect(strokeCount).toBeGreaterThanOrEqual(30);
  expect(ascii).toMatch(/20\.\s+[\d.]+\s+m/);
  expect(ascii).toMatch(/\s+l\n/);
});

test('clears border strips on the raster canvas before vector paint', () => {
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
  clearWorkPdfParagraphBorderStripsOnCanvas(
    canvas,
    [
      {
        edges: {
          top: { color: '#000000', kind: 'single', width: 2 },
          bottom: { color: '#000000', kind: 'single', width: 2 },
          between: { color: '#000000', kind: 'dotted', width: 2 },
          bar: { color: '#000000', kind: 'single', width: 2 },
        },
        height: 20,
        width: 40,
        x: 10,
        y: 15,
      },
    ],
    { height: 50, width: 100 },
    '#ffffff',
  );
  // top + bottom + between(bottom) + bar(left)
  expect(fillRectCalls.length).toBe(4);
});

test('clears a strip under underlined runs before vector paint', () => {
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
  clearWorkPdfUnderlineStripsOnCanvas(
    canvas,
    [
      {
        color: '#000000',
        fontSize: 12,
        fontStyle: 'normal',
        height: 10,
        text: 'Hi',
        underline: { color: '#000000', kind: 'single' },
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
  // scaleX=2, scaleY=2; strip starts at y+height*0.75
  expect(x).toBe(19);
  expect(y).toBe(55);
  expect(w).toBe(82);
  expect(h).toBe(7);
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

test('maps outline levels to PDF structure roles', () => {
  expect(workPdfStructRoleFromOutlineLevel(1)).toBe('H1');
  expect(workPdfStructRoleFromOutlineLevel(6)).toBe('H6');
  expect(workPdfStructRoleFromOutlineLevel(7)).toBe('P');
  expect(workPdfStructRoleFromOutlineLevel(0)).toBe('P');
});

test('applies title, language, outline bookmarks, MarkInfo, and StructTreeRoot stubs', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: 'a4',
    compress: false,
  });
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en-US',
    outline: [
      { level: 1, pageNumber: 1, title: 'Overview' },
      { level: 2, pageNumber: 1, title: 'Details' },
      { level: 7, pageNumber: 1, title: 'Note' },
    ],
    title: 'Quarterly plan',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('Quarterly plan');
  expect(ascii).toContain('A3S Work');
  expect(ascii).toContain('Overview');
  expect(ascii).toContain('/MarkInfo << /Marked true >>');
  expect(ascii).toContain('/ViewerPreferences << /DisplayDocTitle true >>');
  expect(ascii).toContain('/Type /StructTreeRoot');
  expect(ascii).toContain('/S /Document');
  expect(ascii).toContain('/Lang (en-US)');
  expect(ascii).toContain('/S /H1');
  expect(ascii).toContain('/S /H2');
  expect(ascii).toContain('/S /P');
  expect(ascii).toContain('/Alt (Overview)');
  expect(ascii).toContain('/Alt (Details)');
  expect(ascii).toContain('/Alt (Note)');
  const catalog = ascii.match(/\/Type \/Catalog[\s\S]*?endobj/);
  expect(catalog?.[0]).toContain('/StructTreeRoot');
  expect(catalog?.[0]).toMatch(/\/StructTreeRoot \d+ 0 R/);
});

test('links vector-run MCIDs through ParentTree and page StructParents', () => {
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
        fontSize: 14,
        fontStyle: 'normal',
        height: 16,
        text: 'First run',
        width: 80,
        x: 20,
        y: 40,
      },
      {
        color: '#112233',
        fontSize: 14,
        fontStyle: 'normal',
        height: 16,
        text: 'Second run',
        width: 90,
        x: 20,
        y: 60,
      },
    ],
    { height: 280, width: 200 },
    { pageHeightPoints: 280, pageWidthPoints: 200 },
  );
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en',
    outline: [{ level: 1, pageNumber: 1, title: 'Heading' }],
    title: 'MCID parent tree',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('/MCID 0');
  expect(ascii).toContain('/MCID 1');
  expect(ascii).toContain('/StructParents 0');
  expect(ascii).toContain('/ParentTree');
  expect(ascii).toContain('/ParentTreeNextKey 1');
  expect(ascii).toContain('/S /H1');
  expect(ascii).toContain('/S /Span');
  expect(ascii).toContain('/Alt (First run)');
  expect(ascii).toContain('/Alt (Second run)');
  expect(ascii).toMatch(/\/Nums\s*\[\s*0\s*\[/);
  const pageDict = ascii.match(/\/Type \/Page[\s\S]*?endobj/);
  expect(pageDict?.[0]).toContain('/StructParents 0');
});

test('emits Document StructTreeRoot stub when outline is empty', () => {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'pt',
    format: [200, 280],
    compress: false,
  });
  applyWorkPdfDocumentStructure(pdf, {
    language: 'en',
    title: 'Empty outline',
  });
  const ascii = Buffer.from(pdf.output('arraybuffer')).toString('latin1');
  expect(ascii).toContain('/Type /StructTreeRoot');
  expect(ascii).toContain('/S /Document');
  expect(ascii).toContain('/Lang (en)');
  expect(ascii).toContain('/MarkInfo << /Marked true >>');
  expect(ascii).not.toContain('/S /H1');
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
