import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import {
  sameSpreadsheetHistoryContent,
  spreadsheetContentWithSelection,
} from '../src/internal/features/work/editors/spreadsheet-editor-support';
import { useOfficeHistory } from '../src/internal/features/work/editors/use-office-history';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('treats calculated formula caches as part of their source edit', () => {
  const initial = workbook();
  const calculated = structuredClone(initial);
  const formula = calculated.sheets[0]?.data?.[0]?.[1];
  if (!formula) throw new Error('Formula fixture is missing.');
  formula.v = 6;
  formula.m = '6';
  formula.ct = { fa: 'General', t: 'n' };

  expect(sameSpreadsheetHistoryContent(initial, calculated)).toBe(true);

  const formatted = structuredClone(calculated);
  const formattedFormula = formatted.sheets[0]?.data?.[0]?.[1];
  if (!formattedFormula) throw new Error('Formula fixture is missing.');
  formattedFormula.ct = { fa: '0%', t: 'n' };
  expect(sameSpreadsheetHistoryContent(calculated, formatted)).toBe(false);

  const changedInput = structuredClone(calculated);
  const input = changedInput.sheets[0]?.data?.[0]?.[0];
  if (!input) throw new Error('Input fixture is missing.');
  input.v = 4;
  input.m = '4';
  expect(sameSpreadsheetHistoryContent(calculated, changedInput)).toBe(false);
});

test('undo skips the intermediate formula-cache state', () => {
  const initial = workbook();
  const edited = structuredClone(initial);
  const input = edited.sheets[0]?.data?.[0]?.[0];
  if (!input) throw new Error('Input fixture is missing.');
  input.v = 3;
  input.m = '3';
  const changes: WorkSpreadsheetContent[] = [];
  const { result, rerender } = renderHook(
    ({ value }) =>
      useOfficeHistory({
        content: value,
        onChange: (next) => changes.push(next),
        sameValue: sameSpreadsheetHistoryContent,
      }),
    { initialProps: { value: initial } },
  );

  rerender({ value: edited });
  const calculated = structuredClone(edited);
  const formula = calculated.sheets[0]?.data?.[0]?.[1];
  if (!formula) throw new Error('Formula fixture is missing.');
  formula.v = 6;
  formula.m = '6';
  formula.ct = { fa: 'General', t: 'n' };
  rerender({ value: calculated });

  expect(result.current.canUndo).toBe(true);
  act(() => expect(result.current.undo()).toBe(true));
  expect(changes).toEqual([initial]);
});

test('undo skips derived Fortune sheet payloads after a format change', () => {
  const initial = workbook();
  const formatted = structuredClone(initial);
  const formattedCell = formatted.sheets[0]?.data?.[0]?.[0];
  if (!formattedCell) throw new Error('Spreadsheet fixture is missing.');
  formattedCell.bl = 1;
  formatted.sheets[0].celldata = [
    { r: 0, c: 0, v: structuredClone(formattedCell) },
  ];
  const normalized = structuredClone(formatted);
  delete normalized.sheets[0].celldata;
  normalized.sheets[0].luckysheet_select_save = [
    { row: [0, 0], column: [0, 0] },
  ];
  const changes: WorkSpreadsheetContent[] = [];
  const { result, rerender } = renderHook(
    ({ value }) =>
      useOfficeHistory({
        content: value,
        onChange: (next) => changes.push(next),
        sameValue: sameSpreadsheetHistoryContent,
      }),
    { initialProps: { value: initial } },
  );

  rerender({ value: formatted });
  rerender({ value: normalized });

  expect(result.current.canUndo).toBe(true);
  act(() => expect(result.current.undo()).toBe(true));
  expect(changes).toEqual([initial]);
});

test('preserves the live cell selection when applying spreadsheet history', () => {
  const initial = workbook();
  const selected = spreadsheetContentWithSelection(initial, 'sheet-1', {
    row: [10, 10],
    column: [5, 5],
    row_focus: 10,
    column_focus: 5,
  });

  expect(initial.sheets[0].luckysheet_select_save).toBeUndefined();
  expect(selected.sheets[0].luckysheet_select_save).toEqual([
    {
      row: [10, 10],
      column: [5, 5],
      row_focus: 10,
      column_focus: 5,
    },
  ]);
});

function workbook(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [
          [
            { v: 2, m: '2' },
            { f: '=A1*2', v: 4, m: '4' },
          ],
        ],
      },
    ],
  };
}
