import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  reconcileSpreadsheetRichTextCellEdit,
  restoreSpreadsheetRichTextCellRuns,
  sameSpreadsheetRichTextCellText,
} from '../src/internal/features/work/work-xlsx-rich-text-edit';

const themeOrigin = {
  baseColor: '#4472c4',
  index: 4,
  kind: 'theme',
  renderedColor: '#4472c4',
} as const;

const indexedOrigin = {
  baseColor: '#159469',
  index: 10,
  kind: 'indexed',
  renderedColor: '#159469',
} as const;

describe('Spreadsheet rich-text edit reconciliation', () => {
  test('reconstructs a formula-bar insertion with inherited native run styles', () => {
    const previous = richTextCell();
    const changed = {
      bg: '#f4f7ff',
      ct: { fa: 'General', t: 'g' },
      v: 'NatXive rich text',
    } satisfies Cell;

    expect(reconcileSpreadsheetRichTextCellEdit(previous, changed)).toEqual({
      bg: '#f4f7ff',
      ct: {
        fa: 'General',
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: themeOrigin,
            bl: 1,
            fc: '#4472c4',
            ff: 'Aptos Display',
            fs: 14,
            v: 'NatXive ',
          },
          {
            a3sXlsxColorOrigin: indexedOrigin,
            fc: '#159469',
            ff: 'Georgia',
            fs: 12,
            it: 1,
            v: 'rich ',
          },
          { cl: 1, fc: '#a04896', un: 2, v: 'text' },
        ],
      },
      v: 'NatXive rich text',
    });
  });

  test('preserves unaffected styles across deletion and uses the replaced run for inserted text', () => {
    const previous = {
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, v: 'Alpha ' },
          { it: 1, v: 'Beta ' },
          { un: 2, v: 'Gamma' },
        ],
      },
      v: 'Alpha Beta Gamma',
    } satisfies Cell;

    expect(
      reconcileSpreadsheetRichTextCellEdit(previous, {
        v: 'Alpha Gamma',
      }),
    ).toMatchObject({
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, v: 'Alpha ' },
          { un: 2, v: 'Gamma' },
        ],
      },
      v: 'Alpha Gamma',
    });
    expect(
      reconcileSpreadsheetRichTextCellEdit(
        {
          ct: {
            t: 'inlineStr',
            s: [
              { bl: 1, v: 'A' },
              { it: 1, v: 'BC' },
              { un: 2, v: 'D' },
            ],
          },
          v: 'ABCD',
        },
        { v: 'AXD' },
      ),
    ).toMatchObject({
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, v: 'A' },
          { it: 1, v: 'X' },
          { un: 2, v: 'D' },
        ],
      },
      v: 'AXD',
    });
  });

  test('reconstructs source styles when changed live runs lose native color metadata', () => {
    const previous = richTextCell();
    const changed = {
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, v: 'Native ' },
          { it: 1, v: 'rich ' },
          { fc: '#172033', v: 'text!' },
        ],
      },
    } satisfies Cell;

    expect(reconcileSpreadsheetRichTextCellEdit(previous, changed)).toEqual({
      ct: {
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: themeOrigin,
            bl: 1,
            fc: '#4472c4',
            ff: 'Aptos Display',
            fs: 14,
            v: 'Native ',
          },
          {
            a3sXlsxColorOrigin: indexedOrigin,
            fc: '#159469',
            ff: 'Georgia',
            fs: 12,
            it: 1,
            v: 'rich ',
          },
          { cl: 1, fc: '#a04896', un: 2, v: 'text!' },
        ],
      },
      v: 'Native rich text!',
    });
  });

  test('applies only authenticated formatted paste runs inside the replaced range', () => {
    const previous = richTextCell();
    const current = {
      bg: '#f4f7ff',
      ct: { fa: 'General', t: 'g' },
      v: 'Native styled text',
    } satisfies Cell;

    expect(
      reconcileSpreadsheetRichTextCellEdit(previous, current, {
        end: 11,
        runs: [
          {
            bl: 1,
            fc: '#c00000',
            ff: 'Aptos Display',
            fs: 16,
            v: 'sty',
          },
          { it: 1, un: 2, v: 'led' },
        ],
        start: 7,
        text: 'styled',
      }),
    ).toEqual({
      bg: '#f4f7ff',
      ct: {
        fa: 'General',
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: themeOrigin,
            bl: 1,
            fc: '#4472c4',
            ff: 'Aptos Display',
            fs: 14,
            v: 'Native ',
          },
          {
            bl: 1,
            fc: '#c00000',
            ff: 'Aptos Display',
            fs: 16,
            v: 'sty',
          },
          { it: 1, un: 2, v: 'led' },
          {
            a3sXlsxColorOrigin: indexedOrigin,
            fc: '#159469',
            ff: 'Georgia',
            fs: 12,
            it: 1,
            v: ' ',
          },
          { cl: 1, fc: '#a04896', un: 2, v: 'text' },
        ],
      },
      v: 'Native styled text',
    });
  });

  test('creates native runs when formatted text is pasted into a plain or empty cell', () => {
    expect(
      reconcileSpreadsheetRichTextCellEdit(
        { bl: 1, fc: '#4472c4', v: 'Plain text' },
        { bl: 1, fc: '#4472c4', v: 'Plain red' },
        {
          end: 10,
          runs: [{ fc: '#c00000', it: 1, v: 'red' }],
          start: 6,
          text: 'red',
        },
      ),
    ).toMatchObject({
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, fc: '#4472c4', v: 'Plain ' },
          { fc: '#c00000', it: 1, v: 'red' },
        ],
      },
      v: 'Plain red',
    });

    expect(
      reconcileSpreadsheetRichTextCellEdit(
        null,
        { v: 'Fresh' },
        {
          end: 0,
          runs: [{ bl: 1, ff: 'Georgia', v: 'Fresh' }],
          start: 0,
          text: 'Fresh',
        },
      ),
    ).toEqual({
      ct: {
        s: [{ bl: 1, ff: 'Georgia', v: 'Fresh' }],
        t: 'inlineStr',
      },
      v: 'Fresh',
    });
  });

  test('ignores a formatted-paste intent that does not prove the emitted text', () => {
    const previous = richTextCell();
    const current = { v: 'Native plain text' } satisfies Cell;

    expect(
      reconcileSpreadsheetRichTextCellEdit(previous, current, {
        end: 11,
        runs: [{ bl: 1, fc: '#c00000', v: 'styled' }],
        start: 7,
        text: 'styled',
      }).ct?.s,
    ).toEqual([
      {
        a3sXlsxColorOrigin: themeOrigin,
        bl: 1,
        fc: '#4472c4',
        ff: 'Aptos Display',
        fs: 14,
        v: 'Native ',
      },
      {
        a3sXlsxColorOrigin: indexedOrigin,
        fc: '#159469',
        ff: 'Georgia',
        fs: 12,
        it: 1,
        v: 'plain ',
      },
      { cl: 1, fc: '#a04896', un: 2, v: 'text' },
    ]);
  });

  test('restores matching semantic colors on live rich runs and splits conflicting origins safely', () => {
    const previous = {
      ct: {
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: themeOrigin,
            fc: '#4472c4',
            v: 'Blue',
          },
          {
            a3sXlsxColorOrigin: {
              ...themeOrigin,
              index: 5,
            },
            fc: '#4472c4',
            v: ' color',
          },
        ],
      },
      v: 'Blue color',
    } satisfies Cell;
    const current = {
      ct: {
        t: 'inlineStr',
        s: [{ fc: '#4472c4', v: 'BluXe color' }],
      },
    } satisfies Cell;

    expect(reconcileSpreadsheetRichTextCellEdit(previous, current)).toEqual({
      ct: {
        t: 'inlineStr',
        s: [
          {
            a3sXlsxColorOrigin: themeOrigin,
            fc: '#4472c4',
            v: 'BluXe',
          },
          {
            a3sXlsxColorOrigin: { ...themeOrigin, index: 5 },
            fc: '#4472c4',
            v: ' color',
          },
        ],
      },
      v: 'BluXe color',
    });

    const recolored = reconcileSpreadsheetRichTextCellEdit(previous, {
      ct: {
        t: 'inlineStr',
        s: [{ fc: '#ff0000', v: 'Blue' }, { v: ' color' }],
      },
    });
    expect(recolored.ct?.s).toEqual([
      { fc: '#ff0000', v: 'Blue' },
      { v: ' color' },
    ]);
  });

  test('keeps surrogate-pair boundaries intact when inheriting an insertion style', () => {
    const previous = {
      ct: {
        t: 'inlineStr',
        s: [
          { bl: 1, v: 'A😀' },
          { it: 1, v: 'B' },
        ],
      },
      v: 'A😀B',
    } satisfies Cell;

    expect(
      reconcileSpreadsheetRichTextCellEdit(previous, { v: 'A😀!B' }).ct?.s,
    ).toEqual([
      { bl: 1, v: 'A😀!' },
      { it: 1, v: 'B' },
    ]);
  });

  test('fails closed for formulas, non-text replacements, and oversized text', () => {
    const previous = richTextCell();
    const formula = { f: '=1+1', v: 2 } satisfies Cell;
    const numeric = { v: 42 } satisfies Cell;
    const oversized = { v: 'x'.repeat(32_768) } satisfies Cell;

    expect(reconcileSpreadsheetRichTextCellEdit(previous, formula)).toBe(
      formula,
    );
    expect(reconcileSpreadsheetRichTextCellEdit(previous, numeric)).toBe(
      numeric,
    );
    expect(reconcileSpreadsheetRichTextCellEdit(previous, oversized)).toBe(
      oversized,
    );
  });

  test('proves only exact text-stable native rich-text callbacks', () => {
    const previous = richTextCell();

    expect(
      sameSpreadsheetRichTextCellText(previous, {
        ct: { fa: 'General', t: 'g' },
        v: 'Native rich text',
      }),
    ).toBe(true);
    expect(
      sameSpreadsheetRichTextCellText(previous, {
        ct: { t: 'inlineStr', s: [{ v: 'Native rich text' }] },
      }),
    ).toBe(true);
    expect(
      sameSpreadsheetRichTextCellText(previous, { v: 'Changed rich text' }),
    ).toBe(false);
    expect(
      sameSpreadsheetRichTextCellText(previous, {
        f: '=A1',
        v: 'Native rich text',
      }),
    ).toBe(false);
  });

  test('restores source runs for an unauthenticated text-stable live value', () => {
    const previous = richTextCell();

    expect(
      restoreSpreadsheetRichTextCellRuns(previous, {
        ct: {
          t: 'inlineStr',
          s: [
            { bl: 1, v: 'Native ' },
            { it: 1, v: 'rich ' },
            { fc: '#172033', v: 'text' },
          ],
        },
      }),
    ).toEqual(previous);
  });
});

function richTextCell(): Cell {
  return {
    v: 'Native rich text',
    ct: {
      t: 'inlineStr',
      s: [
        {
          a3sXlsxColorOrigin: themeOrigin,
          bl: 1,
          fc: '#4472c4',
          ff: 'Aptos Display',
          fs: 14,
          v: 'Native ',
        },
        {
          a3sXlsxColorOrigin: indexedOrigin,
          fc: '#159469',
          ff: 'Georgia',
          fs: 12,
          it: 1,
          v: 'rich ',
        },
        { cl: 1, fc: '#a04896', un: 2, v: 'text' },
      ],
    },
  };
}
