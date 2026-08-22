import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  applySpreadsheetRichTextSelectionFormat,
  canApplySpreadsheetRichTextSelectionFormat,
  formatSpreadsheetRichTextSelection,
  spreadsheetRichTextSelectionToggleValue,
} from '../src/internal/features/work/editors/spreadsheet-rich-text-selection-format';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet rich-text selection format', () => {
  test('splits crossed runs immutably and invalidates only changed color origins', () => {
    const origin = {
      baseColor: '#4472c4',
      index: 4,
      kind: 'theme',
      renderedColor: '#4472c4',
    } as const;
    const source = {
      v: 'A3S Office',
      ct: {
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: origin,
            bl: 1,
            fc: '#4472c4',
            v: 'A3S',
          },
          { it: 1, v: ' Office' },
        ],
      },
    } satisfies Cell;
    const sourceRuns = source.ct.s;

    const formatted = formatSpreadsheetRichTextSelection(
      source,
      { start: 2, end: 7 },
      'fc',
      '#abc',
    );

    expect(formatted).toEqual({
      v: 'A3S Office',
      ct: {
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: origin,
            bl: 1,
            fc: '#4472c4',
            v: 'A3',
          },
          { bl: 1, fc: '#aabbcc', v: 'S' },
          { fc: '#aabbcc', it: 1, v: ' Off' },
          { it: 1, v: 'ice' },
        ],
      },
    });
    expect(formatted?.ct?.s).not.toBe(sourceRuns);
    expect(source.ct.s).toBe(sourceRuns);
    expect(source.ct.s[0]).toMatchObject({ v: 'A3S', fc: '#4472c4' });
  });

  test('converts a plain text cell while retaining its effective base font', () => {
    const source = {
      v: 'Alpha',
      m: 'Alpha',
      ff: 'Georgia',
      fs: 12,
      bl: 1,
      fc: '#172033',
      bg: '#fff2cc',
    } satisfies Cell;

    expect(
      formatSpreadsheetRichTextSelection(source, { start: 1, end: 4 }, 'it', 1),
    ).toEqual({
      v: 'Alpha',
      ff: 'Georgia',
      fs: 12,
      bl: 1,
      fc: '#172033',
      bg: '#fff2cc',
      ct: {
        t: 'inlineStr',
        s: [
          {
            bl: 1,
            fc: '#172033',
            ff: 'Georgia',
            fs: 12,
            v: 'A',
          },
          {
            bl: 1,
            fc: '#172033',
            ff: 'Georgia',
            fs: 12,
            it: 1,
            v: 'lph',
          },
          {
            bl: 1,
            fc: '#172033',
            ff: 'Georgia',
            fs: 12,
            v: 'a',
          },
        ],
      },
    });
    expect(source).toHaveProperty('m', 'Alpha');
    expect(source).not.toHaveProperty('ct');
  });

  test('coalesces equivalent adjacent runs after formatting', () => {
    const source = {
      v: 'ABC',
      ct: {
        t: 'inlineStr',
        s: [{ bl: 1, v: 'A' }, { bl: 1, v: 'B' }, { v: 'C' }],
      },
    } satisfies Cell;

    expect(
      formatSpreadsheetRichTextSelection(source, { start: 1, end: 3 }, 'bl', 1)
        ?.ct?.s,
    ).toEqual([{ bl: 1, v: 'ABC' }]);
  });

  test('toggles a selected mixed run on and a uniformly styled run off', () => {
    const source = {
      v: 'ABC',
      ct: {
        t: 'inlineStr',
        s: [{ bl: 1, un: 2, v: 'A' }, { v: 'B' }, { bl: 1, v: 'C' }],
      },
    } satisfies Cell;

    expect(
      spreadsheetRichTextSelectionToggleValue(
        source,
        { start: 0, end: 3 },
        'bl',
      ),
    ).toBe(1);
    expect(
      spreadsheetRichTextSelectionToggleValue(
        source,
        { start: 0, end: 1 },
        'bl',
      ),
    ).toBe(0);
    expect(
      spreadsheetRichTextSelectionToggleValue(
        source,
        { start: 0, end: 1 },
        'un',
      ),
    ).toBe(0);
  });

  test('rejects formulas, non-text values, unsafe values, and split surrogate pairs', () => {
    const text = { v: 'A😀B' } satisfies Cell;
    expect(
      formatSpreadsheetRichTextSelection(text, { start: 2, end: 3 }, 'bl', 1),
    ).toBeNull();
    expect(
      formatSpreadsheetRichTextSelection(text, { start: 1, end: 3 }, 'bl', 1)
        ?.ct?.s,
    ).toEqual([{ v: 'A' }, { bl: 1, v: '😀' }, { v: 'B' }]);
    expect(
      formatSpreadsheetRichTextSelection(
        { v: 'Formula', f: '="Formula"' },
        { start: 0, end: 1 },
        'bl',
        1,
      ),
    ).toBeNull();
    expect(
      formatSpreadsheetRichTextSelection(
        { v: 42 },
        { start: 0, end: 1 },
        'bl',
        1,
      ),
    ).toBeNull();
    expect(
      formatSpreadsheetRichTextSelection(
        { v: 'Alpha' },
        { start: 0, end: 5 },
        'fs',
        410,
      ),
    ).toBeNull();
    expect(
      formatSpreadsheetRichTextSelection(
        { v: 'Alpha' },
        { start: 0, end: 5 },
        'fc',
        'red',
      ),
    ).toBeNull();
  });

  test('keeps the 512-run cell boundary after a selection split', () => {
    const runs = Array.from({ length: 512 }, (_, index) => ({
      bl: (index % 2) as 0 | 1,
      v: index === 256 ? 'XYZ' : 'X',
    }));
    const source = {
      v: runs.map((run) => run.v).join(''),
      ct: { t: 'inlineStr', s: runs },
    } satisfies Cell;
    const start = runs
      .slice(0, 256)
      .reduce((total, run) => total + run.v.length, 0);

    expect(
      formatSpreadsheetRichTextSelection(
        source,
        { start: start + 1, end: start + 2 },
        'it',
        1,
      ),
    ).toBeNull();
  });

  test('updates matrix and sparse sheets through one bounded content request', () => {
    const matrixCell = { v: 'Matrix' } satisfies Cell;
    const sparseCell = { v: 'Sparse' } satisfies Cell;
    const content = {
      type: 'spreadsheet',
      sheets: [
        { id: 'matrix', name: 'Matrix', data: [[matrixCell]] },
        {
          id: 'sparse',
          name: 'Sparse',
          celldata: [{ r: 3, c: 2, v: sparseCell }],
        },
      ],
    } satisfies WorkSpreadsheetContent;

    const matrixRequest = {
      sheetId: 'matrix',
      row: 0,
      column: 0,
      selection: { start: 0, end: 6 },
      attribute: 'un',
      value: 4,
    } as const;
    expect(
      canApplySpreadsheetRichTextSelectionFormat(content, matrixRequest),
    ).toBe(true);
    const withMatrix = applySpreadsheetRichTextSelectionFormat(
      content,
      matrixRequest,
    );
    expect(withMatrix?.sheets[0]?.data?.[0]?.[0]?.ct?.s).toEqual([
      { un: 4, v: 'Matrix' },
    ]);

    const withSparse = applySpreadsheetRichTextSelectionFormat(content, {
      sheetId: 'sparse',
      row: 3,
      column: 2,
      selection: { start: 1, end: 5 },
      attribute: 'cl',
      value: 1,
    });
    expect(withSparse?.sheets[1]?.celldata?.[0]?.v.ct?.s).toEqual([
      { v: 'S' },
      { cl: 1, v: 'pars' },
      { v: 'e' },
    ]);
    expect(content.sheets[0]?.data?.[0]?.[0]).toBe(matrixCell);
    expect(content.sheets[1]?.celldata?.[0]?.v).toBe(sparseCell);
  });
});
