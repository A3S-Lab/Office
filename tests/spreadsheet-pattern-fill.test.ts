import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeSpreadsheetCollaborationBinding,
  initializeOfficeSpreadsheetCollaboration,
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
  beginSpreadsheetNativeFillRender,
  finishSpreadsheetNativeFillRender,
} from '../src/internal/features/work/work-spreadsheet-native-fill-canvas';
import {
  createWorkArtifactBlob,
  importWorkFile,
} from '../src/internal/features/work/work-file-io';
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
import {
  activeXlsxPatternFill,
  withXlsxPatternFill,
  XLSX_PATTERN_FILL_CELL_KEY,
  xlsxPatternFill,
  xlsxPatternFillTypes,
  type XlsxPatternFill,
} from '../src/internal/features/work/work-xlsx-pattern-fill';

describe('native XLSX pattern fills', () => {
  test('imports non-solid patterns with semantic colors', () => {
    const worksheet = parseXml(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="2"/><c r="C1" s="3"/></row></sheetData></worksheet>',
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGrid"><fgColor theme="4" tint="0.25"/><bgColor indexed="2"/></patternFill></fill><fill><patternFill patternType="gray0625"><fgColor auto="1"/><bgColor rgb="FFF4F6F8"/></patternFill></fill><fill><patternFill patternType="lightTrellis"><fgColor rgb="FF112233"/><bgColor theme="5"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" applyFill="1"/></cellXfs></styleSheet>',
      'xl/styles.xml',
    );
    const theme = parseXml(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="CC3300"/></a:accent2></a:clrScheme></a:themeElements></a:theme>',
      'xl/theme/theme1.xml',
    );

    const fills = readXlsxDirectCellStyles(worksheet, styles, theme);
    expect(
      fills.map(({ patternFill, style }) => ({ patternFill, style })),
    ).toEqual([
      {
        patternFill: {
          backgroundColor: '#ff0000',
          backgroundColorOrigin: {
            baseColor: '#ff0000',
            index: 2,
            kind: 'indexed',
            renderedColor: '#ff0000',
          },
          foregroundColor: '#538cc6',
          foregroundColorOrigin: {
            baseColor: '#336699',
            index: 4,
            kind: 'theme',
            renderedColor: '#538cc6',
            tint: 0.25,
          },
          patternType: 'darkGrid',
        },
        style: { bg: '#ff0000' },
      },
      {
        patternFill: {
          backgroundColor: '#f4f6f8',
          foregroundColor: '#000000',
          foregroundColorOrigin: {
            baseColor: '#000000',
            kind: 'automatic',
            renderedColor: '#000000',
          },
          patternType: 'gray0625',
        },
        style: { bg: '#f4f6f8' },
      },
      {
        patternFill: {
          backgroundColor: '#cc3300',
          backgroundColorOrigin: {
            baseColor: '#cc3300',
            index: 5,
            kind: 'theme',
            renderedColor: '#cc3300',
          },
          foregroundColor: '#112233',
          patternType: 'lightTrellis',
        },
        style: { bg: '#cc3300' },
      },
    ]);
  });

  test('recognizes the complete OOXML non-solid pattern vocabulary', () => {
    const cells = xlsxPatternFillTypes
      .map(
        (_patternType, index) =>
          `<c r="${String.fromCharCode(65 + index)}1" s="${index + 1}"/>`,
      )
      .join('');
    const fills = xlsxPatternFillTypes
      .map(
        (patternType) =>
          `<fill><patternFill patternType="${patternType}"><fgColor rgb="FF123456"/><bgColor rgb="FFF4F6F8"/></patternFill></fill>`,
      )
      .join('');
    const xfs = xlsxPatternFillTypes
      .map(
        (_patternType, index) =>
          `<xf numFmtId="0" fontId="0" fillId="${index + 2}" borderId="0" applyFill="1"/>`,
      )
      .join('');
    const worksheet = parseXml(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cells}</row></sheetData></worksheet>`,
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="${xlsxPatternFillTypes.length + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fills}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="${xlsxPatternFillTypes.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>${xfs}</cellXfs></styleSheet>`,
      'xl/styles.xml',
    );

    expect(
      readXlsxDirectCellStyles(worksheet, styles).map(
        ({ patternFill }) => patternFill?.patternType,
      ),
    ).toEqual(xlsxPatternFillTypes);
  });

  test('fails closed for unsupported or malformed pattern metadata', () => {
    const worksheet = parseXml(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="2"/><c r="C1" s="3"/></row></sheetData></worksheet>',
      'xl/worksheets/sheet1.xml',
    );
    const styles = parseXml(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="futurePattern"><fgColor rgb="FF112233"/></patternFill></fill><fill><gradientFill degree="45"><stop position="0"><color rgb="FF112233"/></stop></gradientFill></fill><fill><patternFill patternType="darkGrid"><fgColor rgb="not-a-color"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="1" borderId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" applyFill="1"/></cellXfs></styleSheet>',
      'xl/styles.xml',
    );

    expect(readXlsxDirectCellStyles(worksheet, styles)).toEqual([]);
    expect(
      xlsxPatternFill({
        a3sXlsxPatternFill: {
          patternType: 'futurePattern',
          foregroundColor: '#112233',
          backgroundColor: '#ffffff',
        },
      } as Cell),
    ).toBeUndefined();
  });

  test('renders the pattern after the native background and restores Canvas state', () => {
    const calls: string[] = [];
    const context = fakeCanvasContext(calls);
    const fill = patternFill();
    const nativeFillRect = context.fillRect;

    beginSpreadsheetNativeFillRender({ kind: 'pattern', value: fill }, context);
    context.fillStyle = fill.backgroundColor;
    context.fillRect(2, 3, 40, 20);
    finishSpreadsheetNativeFillRender(context);
    context.fillRect(50, 3, 10, 10);

    expect(calls[0]).toBe('fill:#ffffff:2:3:40:20');
    expect(calls.some((call) => call.startsWith('stroke:#2463eb:'))).toBe(true);
    expect(calls.at(-1)).toBe('fill:#ffffff:50:3:10:10');
    expect(context.fillRect).toBe(nativeFillRect);
  });

  test('keeps pattern metadata through unrelated edits and removes it for a new fill', () => {
    const source = contentWithPatternFill();
    const bold = applySpreadsheetCellFormat(source, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 0] },
      patch: { bold: true },
    });
    const boldCell = bold?.sheets[0]?.data?.[0]?.[0];
    expect(activeXlsxPatternFill(boldCell)).toEqual(patternFill());

    const solid = bold
      ? applySpreadsheetCellFormat(bold, {
          sheetId: 'sheet-1',
          range: { row: [0, 0], column: [0, 0] },
          patch: { fillColor: '#ff0000' },
        })
      : null;
    const solidCell = solid?.sheets[0]?.data?.[0]?.[0];
    expect(solidCell?.bg).toBe('#ff0000');
    expect(xlsxPatternFill(solidCell)).toBeUndefined();
  });

  test('propagates pattern metadata with Format Painter and Paste Special formats', () => {
    const source = withXlsxPatternFill(
      { bg: '#ffffff', fc: '#112233', v: 'Pattern' },
      patternFill(),
    );
    const pattern = captureSpreadsheetFormatPattern([[source]]);
    if (!pattern) throw new Error('Expected a format pattern.');
    const batches = spreadsheetFormatPainterBatches(pattern, {
      row: [3, 3],
      column: [4, 4],
    });
    expect(
      batches.find(({ attribute }) => attribute === XLSX_PATTERN_FILL_CELL_KEY),
    ).toMatchObject({ value: patternFill() });

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
    expect(activeXlsxPatternFill(pasted as Cell)).toEqual(patternFill());
    expect((pasted as Cell).v).toBe('Keep');
  });

  test('synchronizes pattern metadata as a native collaboration cell field', () => {
    const session = createOfficeCollaborationSession({
      artifactId: 'pattern-fill',
      document: new Y.Doc(),
      kind: 'spreadsheet',
    });
    const initial = contentWithPatternFill();
    initializeOfficeSpreadsheetCollaboration(session, initial);
    const binding = createOfficeSpreadsheetCollaborationBinding(session);
    const shared = binding.content();
    expect(activeXlsxPatternFill(shared.sheets[0]?.data?.[0]?.[0])).toEqual(
      patternFill(),
    );

    const next = applySpreadsheetCellFormat(shared, {
      sheetId: 'sheet-1',
      range: { row: [0, 0], column: [0, 0] },
      patch: { italic: true },
    });
    if (!next) throw new Error('Expected the collaboration edit to apply.');
    expect(binding.replace(shared, next)).toBe(true);
    expect(
      activeXlsxPatternFill(binding.content().sheets[0]?.data?.[0]?.[0]),
    ).toEqual(patternFill());
    binding.destroy();
  });

  test('retains the exact pattern and semantic colors after edit, export, and reopen', async () => {
    const imported = await importWorkFile(await patternFillWorkbookFile());
    if (imported.content.type !== 'spreadsheet')
      throw new Error('Expected a Spreadsheet artifact.');
    const sheet = imported.content.sheets[0];
    const cell = sheet?.data?.[0]?.[0];
    if (!sheet || !cell) throw new Error('Expected the pattern-filled cell.');
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
    const retained = descendants(styles, 'patternFill').find(
      (pattern) => attribute(pattern, 'patternType') === 'darkGrid',
    );
    const foreground = retained ? directChild(retained, 'fgColor') : undefined;
    const background = retained ? directChild(retained, 'bgColor') : undefined;
    if (!foreground || !background)
      throw new Error('Expected the retained pattern colors.');
    expect(attribute(foreground, 'theme')).toBe('4');
    expect(attribute(foreground, 'tint')).toBe('0.25');
    expect(attribute(background, 'indexed')).toBe('2');

    const reopened = await importWorkFile(
      new File([blob], 'pattern-fill-reopened.xlsx', { type: blob.type }),
    );
    if (reopened.content.type !== 'spreadsheet')
      throw new Error('Expected a reopened Spreadsheet artifact.');
    const reopenedCell = reopened.content.sheets[0]?.data?.[0]?.[0];
    expect(reopenedCell?.bl).toBe(1);
    expect(activeXlsxPatternFill(reopenedCell)).toMatchObject({
      backgroundColor: '#ff0000',
      foregroundColor: '#538cc6',
      patternType: 'darkGrid',
      foregroundColorOrigin: { kind: 'theme', index: 4, tint: 0.25 },
      backgroundColorOrigin: { kind: 'indexed', index: 2 },
    });
  });

  test('falls back to literal colors when semantic palette identities conflict', () => {
    const styles = parseXml(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>',
      'xl/styles.xml',
    );
    const writer = new XlsxDirectCellStyleWriter(styles, null, {
      indexed: new Map(),
      theme: new Map(),
    });
    const cell = withXlsxPatternFill(
      { bg: '#ffffff' },
      {
        ...patternFill(),
        foregroundColorOrigin: {
          kind: 'theme',
          index: 4,
          baseColor: '#2463eb',
          renderedColor: '#2463eb',
        },
      },
    );

    expect(writer.styleId(0, cell)).toBe(1);
    const pattern = descendants(styles, 'patternFill').find(
      (candidate) => attribute(candidate, 'patternType') === 'darkGrid',
    );
    const foreground = pattern ? directChild(pattern, 'fgColor') : undefined;
    if (!foreground) throw new Error('Expected the generated foreground.');
    expect(attribute(foreground, 'theme')).toBeNull();
    expect(attribute(foreground, 'rgb')).toBe('FF2463EB');
  });
});

function patternFill(): XlsxPatternFill {
  return {
    backgroundColor: '#ffffff',
    foregroundColor: '#2463eb',
    patternType: 'darkGrid',
  };
}

function contentWithPatternFill(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Pattern fill',
        row: 1,
        column: 1,
        data: [
          [withXlsxPatternFill({ bg: '#ffffff', v: 'Pattern' }, patternFill())],
        ],
      },
    ],
  };
}

function fakeCanvasContext(calls: string[]): CanvasRenderingContext2D {
  const state = {
    fillStyle: '#000000',
    lineWidth: 1,
    strokeStyle: '#000000',
  };
  const stack: Array<typeof state> = [];
  const context = {
    beginPath: () => calls.push('begin'),
    clip: () => calls.push('clip'),
    closePath: () => calls.push('close'),
    fillRect(x: number, y: number, width: number, height: number) {
      calls.push(
        `fill:${String(state.fillStyle)}:${x}:${y}:${width}:${height}`,
      );
    },
    lineTo: (x: number, y: number) => calls.push(`line:${x}:${y}`),
    moveTo: (x: number, y: number) => calls.push(`move:${x}:${y}`),
    rect: (x: number, y: number, width: number, height: number) =>
      calls.push(`rect:${x}:${y}:${width}:${height}`),
    restore: () => {
      calls.push('restore');
      const previous = stack.pop();
      if (previous) Object.assign(state, previous);
    },
    save: () => {
      calls.push('save');
      stack.push({ ...state });
    },
    setLineDash: (dash: number[]) => calls.push(`dash:${dash.join(',')}`),
    stroke: () =>
      calls.push(
        `stroke:${String(state.strokeStyle)}:${String(state.lineWidth)}`,
      ),
    get fillStyle() {
      return state.fillStyle;
    },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) {
      state.fillStyle = value;
    },
    get lineWidth() {
      return state.lineWidth;
    },
    set lineWidth(value: number) {
      state.lineWidth = value;
    },
    get strokeStyle() {
      return state.strokeStyle;
    },
    set strokeStyle(value: string | CanvasGradient | CanvasPattern) {
      state.strokeStyle = value;
    },
  };
  return context as unknown as CanvasRenderingContext2D;
}

async function patternFillWorkbookFile(): Promise<File> {
  const XLSX = await import('xlsx');
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([['Pattern']]),
    'Pattern',
  );
  const archive = await JSZip.loadAsync(
    XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer,
  );
  const worksheetSource = await archive
    .file('xl/worksheets/sheet1.xml')
    ?.async('text');
  const themeSource = await archive.file('xl/theme/theme1.xml')?.async('text');
  if (!worksheetSource || !themeSource)
    throw new Error('Expected generated worksheet and theme parts.');
  const worksheet = parseXml(worksheetSource, 'xl/worksheets/sheet1.xml');
  descendants(worksheet, 'c')[0]?.setAttribute('s', '1');
  archive.file(
    'xl/worksheets/sheet1.xml',
    new XMLSerializer().serializeToString(worksheet),
  );
  archive.file(
    'xl/styles.xml',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGrid"><fgColor theme="4" tint="0.25"/><bgColor indexed="2"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs></styleSheet>',
  );
  const theme = parseXml(themeSource, 'xl/theme/theme1.xml');
  const scheme = descendants(theme, 'clrScheme')[0];
  if (!scheme) throw new Error('Expected a theme color scheme.');
  const accent1 = directChild(scheme, 'accent1');
  accent1?.firstElementChild?.setAttribute('val', '336699');
  archive.file(
    'xl/theme/theme1.xml',
    new XMLSerializer().serializeToString(theme),
  );
  return new File(
    [await archive.generateAsync({ type: 'arraybuffer' })],
    'pattern-fill.xlsx',
    {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
  );
}
