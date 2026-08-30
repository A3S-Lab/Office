import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  createSpreadsheetSortAppearanceRows,
  createSpreadsheetSortDirectAppearanceRows,
  spreadsheetSortAppearanceFields,
} from '../src/internal/features/work/editors/spreadsheet-sort-appearance';
import { withXlsxGradientFill } from '../src/internal/features/work/work-xlsx-gradient-fill';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

describe('spreadsheet sort appearances', () => {
  test('snapshots effective direct and conditional colors plus icon levels', () => {
    const sheet = appearanceSheet();
    const rows = sheet.data?.map((row) => [...row]) as (Cell | null)[][];
    const appearances = createSpreadsheetSortAppearanceRows(
      sheet,
      { row: [0, 3], column: [0, 1] },
      rows,
    );

    expect(appearances).toEqual([
      [
        { cellColor: '#112233', fontColor: '#ffffff', icon: null },
        { cellColor: null, fontColor: null, icon: null },
      ],
      [
        { cellColor: '#aabbcc', fontColor: '#0000ff', icon: null },
        {
          cellColor: null,
          fontColor: null,
          icon: { iconSet: '3TrafficLights1', index: 0 },
        },
      ],
      [
        { cellColor: '#fff2cc', fontColor: '#d84b4f', icon: null },
        {
          cellColor: null,
          fontColor: null,
          icon: { iconSet: '3TrafficLights1', index: 1 },
        },
      ],
      [
        { cellColor: null, fontColor: null, icon: null },
        {
          cellColor: null,
          fontColor: null,
          icon: { iconSet: '3TrafficLights1', index: 2 },
        },
      ],
    ]);
  });

  test('discovers stable per-column targets while excluding a retained header', () => {
    const sheet = appearanceSheet();
    const rows = sheet.data?.map((row) => [...row]) as (Cell | null)[][];
    const appearances = createSpreadsheetSortAppearanceRows(
      sheet,
      { row: [0, 3], column: [0, 1] },
      rows,
    );

    expect(
      spreadsheetSortAppearanceFields(
        appearances,
        { row: [0, 3], column: [0, 1] },
        'top-to-bottom',
        true,
      ),
    ).toEqual([
      {
        index: 0,
        cellColors: ['#aabbcc', '#fff2cc', null],
        fontColors: ['#0000ff', '#d84b4f', null],
        icons: [],
      },
      {
        index: 1,
        cellColors: [null],
        fontColors: [null],
        icons: [
          { iconSet: '3TrafficLights1', index: 0 },
          { iconSet: '3TrafficLights1', index: 1 },
          { iconSet: '3TrafficLights1', index: 2 },
        ],
      },
    ]);
  });

  test('discovers stable per-row targets for left-to-right sorting', () => {
    const appearances = createSpreadsheetSortDirectAppearanceRows([
      [
        { v: 'Plain' },
        { v: 'Red', bg: '#fce8e6' },
        { v: 'Red again', bg: '#fce8e6' },
      ],
      [
        { v: 'Automatic' },
        { v: 'Blue', fc: '#4472c4' },
        { v: 'Blue again', fc: '#4472c4' },
      ],
    ]);

    expect(
      spreadsheetSortAppearanceFields(
        appearances,
        { row: [4, 5], column: [2, 4] },
        'left-to-right',
        false,
      ),
    ).toEqual([
      {
        index: 4,
        cellColors: [null, '#fce8e6'],
        fontColors: [null],
        icons: [],
      },
      {
        index: 5,
        cellColors: [null],
        fontColors: [null, '#4472c4'],
        icons: [],
      },
    ]);
  });

  test('does not misrepresent a native gradient as one sortable solid color', () => {
    const rows = [
      [
        withXlsxGradientFill(
          { v: 'Gradient', bg: '#112233' },
          {
            type: 'linear',
            degree: 45,
            stops: [
              { position: 0, color: '#112233' },
              { position: 1, color: '#ddeeff' },
            ],
          },
        ),
      ],
      [{ v: 'No fill' }],
    ];
    const appearances = createSpreadsheetSortDirectAppearanceRows(rows);

    expect(appearances[0]?.[0]?.cellColor).toBeUndefined();
    expect(appearances[1]?.[0]?.cellColor).toBeNull();
    expect(
      spreadsheetSortAppearanceFields(
        appearances,
        { row: [0, 1], column: [0, 0] },
        'top-to-bottom',
        false,
      )[0]?.cellColors,
    ).toEqual([null]);
  });
});

function appearanceSheet(): WorkSpreadsheetSheet {
  return {
    id: 'sheet-1',
    name: 'Visuals',
    data: [
      [{ v: 'Task', bg: '#112233', fc: '#ffffff' }, { v: 'Score' }],
      [{ v: 'Alpha', bg: '#abc', fc: '#00f' }, { v: 10 }],
      [{ v: 'Beta' }, { v: 20 }],
      [{ v: 'Gamma' }, { v: 30 }],
    ],
    luckysheet_conditionformat_save: [
      {
        type: 'default',
        cellrange: [{ row: [2, 2], column: [0, 0] }],
        conditionName: 'textContains',
        conditionValue: ['Beta'],
        format: { cellColor: 'rgb(255, 242, 204)', textColor: '#d84b4f' },
      },
      {
        type: 'icons',
        cellrange: [{ row: [1, 3], column: [1, 1] }],
        format: {
          iconSet: '3TrafficLights1',
          showValue: true,
          reverse: false,
          percent: false,
          thresholds: [
            { type: 'min', gte: true },
            { type: 'num', value: 15, gte: true },
            { type: 'num', value: 25, gte: true },
          ],
        },
      },
    ],
  };
}
