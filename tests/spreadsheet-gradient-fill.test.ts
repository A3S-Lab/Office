import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeSpreadsheetCollaborationBinding,
  initializeOfficeSpreadsheetCollaboration,
  type OfficeArtifact,
} from '../src/core';
import { applySpreadsheetCellFormat } from '../src/internal/features/work/editors/spreadsheet-cell-format';
import {
  captureSpreadsheetFormatPattern,
  spreadsheetFormatPainterBatches,
} from '../src/internal/features/work/editors/spreadsheet-format-painter';
import {
  pasteSpreadsheetSpecialCell,
  spreadsheetPasteCellInvalid,
} from '../src/internal/features/work/editors/spreadsheet-paste-special-cell';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
import {
  beginSpreadsheetNativeFillRender,
  finishSpreadsheetNativeFillRender,
} from '../src/internal/features/work/work-spreadsheet-native-fill-canvas';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  readXlsxDirectCellStyles,
  XlsxDirectCellStyleWriter,
} from '../src/internal/features/work/work-xlsx-cell-styles';
import { createXlsxColorResolver } from '../src/internal/features/work/work-xlsx-colors';
import {
  activeXlsxGradientFill,
  MAX_XLSX_GRADIENT_STOPS,
  readXlsxGradientFill,
  withXlsxGradientFill,
  XLSX_GRADIENT_FILL_CELL_KEY,
  xlsxGradientFill,
  type XlsxGradientFill,
} from '../src/internal/features/work/work-xlsx-gradient-fill';

describe('native XLSX gradient fills', () => {
  test('imports linear and path gradients with semantic stop colors', () => {
    const worksheet = parseXml(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="2"/></row></sheetData></worksheet>',
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><gradientFill degree="90.5"><stop position="0"><color theme="4" tint="0.25"/></stop><stop position="0.5"><color indexed="2"/></stop><stop position="1"><color auto="1"/></stop></gradientFill></fill><fill><gradientFill type="path" left="0.2" right="0.8" top="0.1" bottom="0.9"><stop position="0"><color rgb="FF112233"/></stop><stop position="1"><color theme="5"/></stop></gradientFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFill="1"/></cellXfs></styleSheet>',
      'xl/styles.xml',
    );
    const theme = parseXml(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="CC3300"/></a:accent2></a:clrScheme></a:themeElements></a:theme>',
      'xl/theme/theme1.xml',
    );

    expect(
      readXlsxDirectCellStyles(worksheet, styles, theme).map(
        ({ gradientFill, style }) => ({ gradientFill, style }),
      ),
    ).toEqual([
      {
        gradientFill: {
          degree: 90.5,
          stops: [
            {
              color: '#538cc6',
              colorOrigin: {
                baseColor: '#336699',
                index: 4,
                kind: 'theme',
                renderedColor: '#538cc6',
                tint: 0.25,
              },
              position: 0,
            },
            {
              color: '#ff0000',
              colorOrigin: {
                baseColor: '#ff0000',
                index: 2,
                kind: 'indexed',
                renderedColor: '#ff0000',
              },
              position: 0.5,
            },
            {
              color: '#000000',
              colorOrigin: {
                baseColor: '#000000',
                kind: 'automatic',
                renderedColor: '#000000',
              },
              position: 1,
            },
          ],
          type: 'linear',
        },
        style: { bg: '#538cc6' },
      },
      {
        gradientFill: {
          bottom: 0.9,
          left: 0.2,
          right: 0.8,
          stops: [
            { color: '#112233', position: 0 },
            {
              color: '#cc3300',
              colorOrigin: {
                baseColor: '#cc3300',
                index: 5,
                kind: 'theme',
                renderedColor: '#cc3300',
              },
              position: 1,
            },
          ],
          top: 0.1,
          type: 'path',
        },
        style: { bg: '#112233' },
      },
    ]);
  });

  test('fails closed for malformed, ambiguous, or over-budget gradients', () => {
    const resolver = createXlsxColorResolver(null, null);
    const gradient = (source: string) =>
      parseXml(
        `<gradientFill xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${source}</gradientFill>`,
        'xl/styles.xml',
      ).documentElement;
    const stop = (position: string, color = 'FF112233') =>
      `<stop position="${position}"><color rgb="${color}"/></stop>`;

    expect(
      readXlsxGradientFill(
        gradient(`degree="not-a-number">${stop('0')}${stop('1')}`),
        resolver,
      ),
    ).toBeUndefined();
    expect(
      readXlsxGradientFill(gradient(`>${stop('-0.1')}${stop('1')}`), resolver),
    ).toBeUndefined();
    expect(
      readXlsxGradientFill(gradient(`>${stop('0.8')}${stop('0.2')}`), resolver),
    ).toBeUndefined();
    expect(
      readXlsxGradientFill(
        gradient(`>${stop('0', 'not-a-color')}${stop('1')}`),
        resolver,
      ),
    ).toBeUndefined();
    expect(
      readXlsxGradientFill(gradient(`>${stop('0')}`), resolver),
    ).toBeUndefined();
    expect(
      readXlsxGradientFill(
        gradient(`type="path" left="0.8" right="0.2">${stop('0')}${stop('1')}`),
        resolver,
      ),
    ).toBeUndefined();
    expect(
      readXlsxGradientFill(
        gradient(
          `>${Array.from({ length: MAX_XLSX_GRADIENT_STOPS + 1 }, (_, index) => stop(String(index / MAX_XLSX_GRADIENT_STOPS))).join('')}`,
        ),
        resolver,
      ),
    ).toBeUndefined();
    expect(
      xlsxGradientFill({
        a3sXlsxGradientFill: {
          degree: 45,
          stops: [
            { color: '#112233', position: 0.8 },
            { color: '#445566', position: 0.2 },
          ],
          type: 'linear',
        },
      } as Cell),
    ).toBeUndefined();
  });

  test('renders linear and path gradients in the first visible-cell background paint', () => {
    const linearCalls: string[] = [];
    const linearContext = fakeGradientCanvasContext(linearCalls);
    const nativeFillRect = linearContext.fillRect;
    beginSpreadsheetNativeFillRender(
      { kind: 'gradient', value: gradientFill() },
      linearContext,
    );
    linearContext.fillRect(2, 3, 40, 20);
    finishSpreadsheetNativeFillRender(linearContext);
    linearContext.fillRect(50, 3, 10, 10);

    expect(linearCalls).toContain('linear:2:13:42:13');
    expect(linearCalls).toContain('stop:0:#2463eb');
    expect(linearCalls).toContain('stop:1:#ffffff');
    expect(linearCalls.at(-1)).toBe('fill:#2463eb:50:3:10:10');
    expect(linearContext.fillRect).toBe(nativeFillRect);

    const pathCalls: string[] = [];
    const pathContext = fakeGradientCanvasContext(pathCalls);
    beginSpreadsheetNativeFillRender(
      {
        kind: 'gradient',
        value: {
          bottom: 0.75,
          left: 0.25,
          right: 0.75,
          stops: [
            { color: '#000000', position: 0 },
            { color: '#ffffff', position: 1 },
          ],
          top: 0.25,
          type: 'path',
        },
      },
      pathContext,
    );
    pathContext.fillRect(0, 0, 40, 20);
    finishSpreadsheetNativeFillRender(pathContext);
    const pathPaints = pathCalls.filter((call) => call.startsWith('fill:'));
    expect(pathPaints[0]).toBe('fill:#ffffff:0:0:40:20');
    expect(pathPaints.at(-1)).toBe('fill:#000000:10:5:20:10');
    expect(pathPaints.length).toBeLessThanOrEqual(97);
  });

  test('keeps gradient metadata through unrelated edits and removes it for a new fill', () => {
    const source = contentWithGradientFill();
    const bold = applySpreadsheetCellFormat(source, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 0] },
      patch: { bold: true },
    });
    expect(activeXlsxGradientFill(bold?.sheets[0]?.data?.[0]?.[0])).toEqual(
      gradientFill(),
    );

    const solid = bold
      ? applySpreadsheetCellFormat(bold, {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [0, 0] },
          patch: { fill: { color: '#ff0000', kind: 'solid' } },
        })
      : null;
    const solidCell = solid?.sheets[0]?.data?.[0]?.[0];
    expect(solidCell?.bg).toBe('#ff0000');
    expect(xlsxGradientFill(solidCell)).toBeUndefined();
  });

  test('propagates gradient metadata with Format Painter and Paste Special formats', () => {
    const source = withXlsxGradientFill(
      { bg: '#2463eb', fc: '#112233', v: 'Gradient' },
      gradientFill(),
    );
    const pattern = captureSpreadsheetFormatPattern([[source]]);
    if (!pattern) throw new Error('Expected a format pattern.');
    const batches = spreadsheetFormatPainterBatches(pattern, {
      row: [3, 3],
      column: [4, 4],
    });
    expect(
      batches.find(
        ({ attribute }) => attribute === XLSX_GRADIENT_FILL_CELL_KEY,
      ),
    ).toMatchObject({ value: gradientFill() });

    const pasted = pasteSpreadsheetSpecialCell({
      source: { cell: source, borders: {} },
      destination: { bg: '#ffeecc', v: 'Keep' },
      content: 'formats',
      operation: 'none',
      rowOffset: 0,
      columnOffset: 0,
    });
    expect(pasted).not.toBeNull();
    expect(pasted).not.toBe(spreadsheetPasteCellInvalid);
    expect(activeXlsxGradientFill(pasted as Cell)).toEqual(gradientFill());
    expect((pasted as Cell).v).toBe('Keep');
  });

  test('synchronizes gradient metadata as a native collaboration cell field', () => {
    const session = createOfficeCollaborationSession({
      artifactId: 'gradient-fill',
      document: new Y.Doc(),
      kind: 'spreadsheet',
    });
    initializeOfficeSpreadsheetCollaboration(
      session,
      contentWithGradientFill(),
    );
    const binding = createOfficeSpreadsheetCollaborationBinding(session);
    const shared = binding.content();
    expect(activeXlsxGradientFill(shared.sheets[0]?.data?.[0]?.[0])).toEqual(
      gradientFill(),
    );

    const next = applySpreadsheetCellFormat(shared, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 0] },
      patch: { italic: true },
    });
    if (!next) throw new Error('Expected the collaboration edit to apply.');
    expect(binding.replace(shared, next)).toBe(true);
    expect(
      activeXlsxGradientFill(binding.content().sheets[0]?.data?.[0]?.[0]),
    ).toEqual(gradientFill());
    binding.destroy();
  });

  test('retains exact geometry, stops, and semantic colors after edit, export, and reopen', async () => {
    const imported = await importWorkFile(await gradientFillWorkbookFile());
    if (imported.content.type !== 'spreadsheet') {
      throw new Error('Expected a Spreadsheet artifact.');
    }
    const sheet = imported.content.sheets[0];
    const cell = sheet?.data?.[0]?.[0];
    if (!sheet || !cell) throw new Error('Expected the gradient-filled cell.');
    const metadata = xlsxGradientFill(cell);
    const storedMetadata = (cell as unknown as Record<string, unknown>)[
      XLSX_GRADIENT_FILL_CELL_KEY
    ] as XlsxGradientFill;
    expect(Object.isFrozen(cell)).toBe(true);
    expect(metadata).toBeDefined();
    expect(Object.isFrozen(storedMetadata)).toBe(true);
    expect(Object.isFrozen(storedMetadata.stops)).toBe(true);
    expect(Object.isFrozen(storedMetadata.stops[0]?.colorOrigin)).toBe(true);

    const edited = applySpreadsheetCellFormat(imported.content, {
      sheetId: sheet.id,
      range: { row: [0, 0], column: [0, 0] },
      patch: { bold: true },
    });
    if (!edited) throw new Error('Expected the unrelated edit to apply.');
    imported.content = edited;

    const blob = await createWorkArtifactBlob(imported);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const stylesSource = await archive.file('xl/styles.xml')?.async('text');
    if (!stylesSource) throw new Error('Expected native XLSX styles.');
    const styles = parseXml(stylesSource, 'xl/styles.xml');
    const retained = descendants(styles, 'gradientFill').find(
      (gradient) => attribute(gradient, 'degree') === '90.5',
    );
    if (!retained) throw new Error('Expected the retained gradient.');
    const stops = descendants(retained, 'stop');
    expect(stops.map((stop) => attribute(stop, 'position'))).toEqual([
      '0',
      '0.5',
      '1',
    ]);
    const colors = stops.map((stop) => directChild(stop, 'color'));
    expect(attribute(colors[0] ?? retained, 'theme')).toBe('4');
    expect(attribute(colors[0] ?? retained, 'tint')).toBe('0.25');
    expect(attribute(colors[1] ?? retained, 'indexed')).toBe('2');
    expect(attribute(colors[2] ?? retained, 'auto')).toBe('1');

    const reopened = await importWorkFile(
      new File([blob], 'gradient-fill-reopened.xlsx', { type: blob.type }),
    );
    if (reopened.content.type !== 'spreadsheet') {
      throw new Error('Expected a reopened Spreadsheet artifact.');
    }
    const reopenedCell = reopened.content.sheets[0]?.data?.[0]?.[0];
    expect(reopenedCell?.bl).toBe(1);
    expect(activeXlsxGradientFill(reopenedCell)).toMatchObject({
      degree: 90.5,
      stops: [
        {
          color: '#538cc6',
          colorOrigin: { kind: 'theme', index: 4, tint: 0.25 },
          position: 0,
        },
        {
          color: '#ff0000',
          colorOrigin: { kind: 'indexed', index: 2 },
          position: 0.5,
        },
        {
          color: '#000000',
          colorOrigin: { kind: 'automatic' },
          position: 1,
        },
      ],
      type: 'linear',
    });
  });

  test('writes exact path geometry and ordered stops', () => {
    const styles = emptyStylesDocument();
    const writer = new XlsxDirectCellStyleWriter(styles);
    const cell = withXlsxGradientFill(
      { bg: '#112233' },
      {
        bottom: 0.9,
        left: 0.2,
        right: 0.8,
        stops: [
          { color: '#112233', position: 0 },
          { color: '#445566', position: 0.4 },
          { color: '#ffffff', position: 1 },
        ],
        top: 0.1,
        type: 'path',
      },
    );

    expect(writer.styleId(0, cell)).toBe(1);
    const gradient = descendants(styles, 'gradientFill')[0];
    if (!gradient) throw new Error('Expected the generated path gradient.');
    expect({
      bottom: attribute(gradient, 'bottom'),
      left: attribute(gradient, 'left'),
      right: attribute(gradient, 'right'),
      top: attribute(gradient, 'top'),
      type: attribute(gradient, 'type'),
    }).toEqual({
      bottom: '0.9',
      left: '0.2',
      right: '0.8',
      top: '0.1',
      type: 'path',
    });
    expect(
      descendants(gradient, 'stop').map((stop) => ({
        color: attribute(directChild(stop, 'color') ?? stop, 'rgb'),
        position: attribute(stop, 'position'),
      })),
    ).toEqual([
      { color: 'FF112233', position: '0' },
      { color: 'FF445566', position: '0.4' },
      { color: 'FFFFFFFF', position: '1' },
    ]);
  });

  test('round-trips authored linear and path gradients from typed cell-format patches', async () => {
    const linear = {
      degree: 315.25,
      stops: [
        { color: '#1d4ed8', position: 0 },
        { color: '#67e8f9', position: 0.35 },
        { color: '#ffffff', position: 1 },
      ],
      type: 'linear' as const,
    };
    const path = {
      bottom: 0.9,
      left: 0.15,
      right: 0.85,
      stops: [
        { color: '#b42318', position: 0 },
        { color: '#fff2cc', position: 0.55 },
        { color: '#ffffff', position: 1 },
      ],
      top: 0.1,
      type: 'path' as const,
    };
    const content = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Authored gradients',
          data: [[{ v: 'Linear' }, { v: 'Path' }]],
        },
      ],
    } satisfies WorkSpreadsheetContent;
    const withLinear = applySpreadsheetCellFormat(content, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 0] },
      patch: { fill: { kind: 'gradient', value: linear } },
    });
    const authored = withLinear
      ? applySpreadsheetCellFormat(withLinear, {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [1, 1] },
          patch: { fill: { kind: 'gradient', value: path } },
        })
      : null;
    if (!authored) throw new Error('Expected authored gradients to apply.');
    const now = Date.now();
    const artifact: OfficeArtifact = {
      id: 'authored-gradients',
      kind: 'spreadsheet',
      title: 'Authored gradients',
      favorite: false,
      createdAt: now,
      updatedAt: now,
      lastOpenedAt: now,
      revision: 1,
      content: authored,
    };

    const blob = await createWorkArtifactBlob(artifact);
    const reopened = await importWorkFile(
      new File([blob], 'authored-gradients.xlsx', { type: blob.type }),
    );
    if (reopened.content.type !== 'spreadsheet') {
      throw new Error('Expected a reopened Spreadsheet artifact.');
    }
    expect(
      activeXlsxGradientFill(reopened.content.sheets[0]?.data?.[0]?.[0]),
    ).toEqual(linear);
    expect(
      activeXlsxGradientFill(reopened.content.sheets[0]?.data?.[0]?.[1]),
    ).toEqual(path);
  });

  test('falls back to literal stop colors when semantic palette identities conflict', () => {
    const styles = emptyStylesDocument();
    const writer = new XlsxDirectCellStyleWriter(styles, null, {
      indexed: new Map(),
      theme: new Map(),
    });
    const cell = withXlsxGradientFill(
      { bg: '#2463eb' },
      {
        ...gradientFill(),
        stops: [
          {
            color: '#2463eb',
            colorOrigin: {
              baseColor: '#2463eb',
              index: 4,
              kind: 'theme',
              renderedColor: '#2463eb',
            },
            position: 0,
          },
          { color: '#ffffff', position: 1 },
        ],
      },
    );

    expect(writer.styleId(0, cell)).toBe(1);
    const gradient = descendants(styles, 'gradientFill')[0];
    const color = gradient
      ? directChild(descendants(gradient, 'stop')[0] ?? gradient, 'color')
      : undefined;
    if (!color) throw new Error('Expected the generated gradient stop.');
    expect(attribute(color, 'theme')).toBeNull();
    expect(attribute(color, 'rgb')).toBe('FF2463EB');
  });
});

function emptyStylesDocument(): Document {
  return parseXml(
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>',
    'xl/styles.xml',
  );
}

function gradientFill(): XlsxGradientFill {
  return {
    degree: 0,
    stops: [
      { color: '#2463eb', position: 0 },
      { color: '#ffffff', position: 1 },
    ],
    type: 'linear',
  };
}

function contentWithGradientFill(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Gradient fill',
        row: 1,
        column: 1,
        data: [
          [
            withXlsxGradientFill(
              { bg: '#2463eb', v: 'Gradient' },
              gradientFill(),
            ),
          ],
        ],
      },
    ],
  };
}

function fakeGradientCanvasContext(calls: string[]): CanvasRenderingContext2D {
  const state = { fillStyle: '#ffffff' as string | CanvasGradient };
  const stack: Array<typeof state> = [];
  const context = {
    createLinearGradient: (
      startX: number,
      startY: number,
      endX: number,
      endY: number,
    ) => {
      calls.push(`linear:${startX}:${startY}:${endX}:${endY}`);
      const gradient = {
        addColorStop: (position: number, color: string) =>
          calls.push(`stop:${position}:${color}`),
      } as unknown as CanvasGradient;
      return gradient;
    },
    fillRect(x: number, y: number, width: number, height: number) {
      calls.push(
        `fill:${typeof state.fillStyle === 'string' ? state.fillStyle : 'gradient'}:${x}:${y}:${width}:${height}`,
      );
    },
    restore: () => {
      const previous = stack.pop();
      if (previous) Object.assign(state, previous);
    },
    save: () => stack.push({ ...state }),
    get fillStyle() {
      return state.fillStyle;
    },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) {
      state.fillStyle = value as string | CanvasGradient;
    },
  };
  return context as unknown as CanvasRenderingContext2D;
}

async function gradientFillWorkbookFile(): Promise<File> {
  const XLSX = await import('xlsx');
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([['Gradient']]),
    'Gradient',
  );
  const archive = await JSZip.loadAsync(
    XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer,
  );
  const worksheetSource = await archive
    .file('xl/worksheets/sheet1.xml')
    ?.async('text');
  const themeSource = await archive.file('xl/theme/theme1.xml')?.async('text');
  if (!worksheetSource || !themeSource) {
    throw new Error('Expected generated worksheet and theme parts.');
  }
  const worksheet = parseXml(worksheetSource, 'xl/worksheets/sheet1.xml');
  descendants(worksheet, 'c')[0]?.setAttribute('s', '1');
  archive.file(
    'xl/worksheets/sheet1.xml',
    new XMLSerializer().serializeToString(worksheet),
  );
  archive.file(
    'xl/styles.xml',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><gradientFill degree="90.5"><stop position="0"><color theme="4" tint="0.25"/></stop><stop position="0.5"><color indexed="2"/></stop><stop position="1"><color auto="1"/></stop></gradientFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs></styleSheet>',
  );
  const theme = parseXml(themeSource, 'xl/theme/theme1.xml');
  const scheme = descendants(theme, 'clrScheme')[0];
  if (!scheme) throw new Error('Expected a theme color scheme.');
  directChild(scheme, 'accent1')?.firstElementChild?.setAttribute(
    'val',
    '336699',
  );
  archive.file(
    'xl/theme/theme1.xml',
    new XMLSerializer().serializeToString(theme),
  );
  return new File(
    [await archive.generateAsync({ type: 'arraybuffer' })],
    'gradient-fill.xlsx',
    {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  );
}
