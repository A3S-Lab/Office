import { describe, expect, test } from '@rstest/core';
import {
  activateSpreadsheetSheet,
  addSpreadsheetSheet,
  adjacentSpreadsheetSheetId,
  deleteSpreadsheetSheet,
  duplicateSpreadsheetSheet,
  hideSpreadsheetSheet,
  moveSpreadsheetSheet,
  renameSpreadsheetSheet,
  setSpreadsheetSheetColor,
  spreadsheetSheetNameValidationMessage,
} from '../src/internal/features/work/editors/spreadsheet-sheet-model';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet sheet model', () => {
  test('creates a localized worksheet and activates it immediately', () => {
    const content = workbook();
    const next = addSpreadsheetSheet(content);

    expect(next).not.toBe(content);
    expect(content.sheets).toHaveLength(2);
    expect(next.sheets).toHaveLength(3);
    expect(next.sheets.map(({ name, status }) => ({ name, status }))).toEqual([
      { name: '执行看板', status: 0 },
      { name: '风险台账', status: 0 },
      { name: '工作表 3', status: 1 },
    ]);
    expect(next.sheets[2]).toMatchObject({
      id: 'sheet-3',
      order: 2,
      row: 60,
      column: 26,
    });
    expect(next.sheets[2]?.data).toHaveLength(60);
    expect(next.sheets[2]?.data?.every((row) => row.length === 26)).toBe(true);
    expect(next.sheets[2]?.data?.flat().every((cell) => cell === null)).toBe(
      true,
    );
  });

  test('activates adjacent visible worksheets with wraparound navigation', () => {
    const content = workbook();
    content.sheets.push({
      id: 'sheet-hidden',
      name: '隐藏页',
      hide: 1,
      order: 2,
      status: 0,
    });

    expect(adjacentSpreadsheetSheetId(content, 'sheet-1', -1)).toBe('sheet-2');
    expect(adjacentSpreadsheetSheetId(content, 'sheet-1', 1)).toBe('sheet-2');
    expect(adjacentSpreadsheetSheetId(content, 'sheet-2', 1)).toBe('sheet-1');

    const next = activateSpreadsheetSheet(content, 'sheet-2');
    expect(next?.sheets.map(({ id, status }) => ({ id, status }))).toEqual([
      { id: 'sheet-1', status: 0 },
      { id: 'sheet-2', status: 1 },
      { id: 'sheet-hidden', status: 0 },
    ]);
  });

  test('moves across hidden worksheets using the visible tab order', () => {
    const content = workbook();
    content.sheets[1].order = 2;
    content.sheets.splice(1, 0, {
      id: 'sheet-hidden',
      name: '隐藏页',
      hide: 1,
      order: 1,
      status: 0,
    });

    const moved = moveSpreadsheetSheet(content, 'sheet-1', 1);

    expect(moved?.sheets.map(({ id, order }) => ({ id, order }))).toEqual([
      { id: 'sheet-2', order: 0 },
      { id: 'sheet-hidden', order: 1 },
      { id: 'sheet-1', order: 2 },
    ]);
  });

  test('keeps worksheet lifecycle operations immutable and safe', () => {
    const content = workbook();
    const duplicated = duplicateSpreadsheetSheet(content, 'sheet-1');
    expect(duplicated?.sheets.at(-1)).toMatchObject({
      id: 'sheet-3',
      name: '执行看板 副本',
      status: 1,
    });

    const renamed = renameSpreadsheetSheet(content, 'sheet-2', '  风险清单  ');
    expect(renamed?.sheets[1]?.name).toBe('风险清单');
    expect(renameSpreadsheetSheet(content, 'sheet-2', '执行看板')).toBeNull();

    const colored = setSpreadsheetSheetColor(content, 'sheet-2', '#e06c53');
    expect(colored?.sheets[1]?.color).toBe('#e06c53');

    const moved = moveSpreadsheetSheet(content, 'sheet-2', -1);
    expect(moved?.sheets.map(({ id, order }) => ({ id, order }))).toEqual([
      { id: 'sheet-2', order: 0 },
      { id: 'sheet-1', order: 1 },
    ]);

    const hidden = hideSpreadsheetSheet(content, 'sheet-1', true);
    if (!hidden) throw new Error('Expected the active sheet to be hidden.');
    expect(hidden?.sheets).toEqual([
      expect.objectContaining({ id: 'sheet-1', hide: 1, status: 0 }),
      expect.objectContaining({ id: 'sheet-2', hide: 0, status: 1 }),
    ]);
    expect(hideSpreadsheetSheet(hidden, 'sheet-2', true)).toBeNull();

    const deleted = deleteSpreadsheetSheet(content, 'sheet-1');
    if (!deleted) throw new Error('Expected the active sheet to be deleted.');
    expect(deleted?.sheets).toEqual([
      expect.objectContaining({ id: 'sheet-2', order: 0, status: 1 }),
    ]);
    expect(deleteSpreadsheetSheet(deleted, 'sheet-2')).toBeNull();
  });

  test('describes worksheet rename validation without duplicating command rules', () => {
    const content = workbook();

    expect(
      spreadsheetSheetNameValidationMessage(
        content.sheets,
        'sheet-2',
        '执行看板',
      ),
    ).toBe('名称已存在');
    expect(
      spreadsheetSheetNameValidationMessage(content.sheets, 'sheet-2', '  '),
    ).toBe('请输入名称');
    expect(
      spreadsheetSheetNameValidationMessage(
        content.sheets,
        'sheet-2',
        '风险/清单',
      ),
    ).toBe('名称不能包含 \\ / ? * [ ] :');
    expect(
      spreadsheetSheetNameValidationMessage(
        content.sheets,
        'sheet-2',
        '风险清单',
      ),
    ).toBeNull();
  });
});

function workbook(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: '执行看板',
        order: 0,
        status: 1,
        data: [[{ v: 'A3S', m: 'A3S' }]],
      },
      {
        id: 'sheet-2',
        name: '风险台账',
        order: 1,
        status: 0,
        data: [],
      },
    ],
  };
}
