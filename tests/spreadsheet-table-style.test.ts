import { describe, expect, test } from '@rstest/core';
import {
  createSpreadsheetTableRenderResolver,
  spreadsheetTableStyleChoices,
  spreadsheetTableStylePalette,
} from '../src/internal/features/work/editors/spreadsheet-table-style';
import {
  beginSpreadsheetTableCellRender,
  finishSpreadsheetTableCellRender,
} from '../src/internal/features/work/editors/spreadsheet-table-render';
import type { WorkSpreadsheetTable } from '../src/internal/features/work/work-types';

describe('spreadsheet table styles', () => {
  test('exposes every OOXML built-in table style identity', () => {
    const choices = spreadsheetTableStyleChoices();
    expect(choices).toHaveLength(60);
    expect(
      choices.filter((choice) => choice.style.family === 'light'),
    ).toHaveLength(21);
    expect(
      choices.filter((choice) => choice.style.family === 'medium'),
    ).toHaveLength(28);
    expect(
      choices.filter((choice) => choice.style.family === 'dark'),
    ).toHaveLength(11);
    expect(new Set(choices.map((choice) => choice.ooxmlName)).size).toBe(60);
    expect(spreadsheetTableStylePalette({ family: 'none' })).toBeNull();
  });

  test('resolves header, row and column stripes lazily without materializing cells', () => {
    const table: WorkSpreadsheetTable = {
      id: 'table-1',
      name: 'Table1',
      range: { row: [10, 1_000_000], column: [2, 5] },
      columns: [{ name: 'A' }, { name: 'B' }, { name: 'C' }, { name: 'D' }],
      filters: [],
      headerRow: true,
      totalsRow: false,
      style: { family: 'medium', number: 2 },
      showFirstColumn: true,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: true,
    };
    const resolve = createSpreadsheetTableRenderResolver([table]);

    expect(resolve(9, 2)).toBeNull();
    expect(resolve(10, 2)).toMatchObject({
      bold: true,
      role: 'header',
      tableId: 'table-1',
    });
    expect(resolve(11, 2)).toMatchObject({ bold: true, role: 'body' });
    expect(resolve(11, 3)).toMatchObject({ role: 'body' });
    expect(resolve(12, 3)?.background).not.toBe(resolve(11, 3)?.background);
    expect(resolve(1_000_000, 5)).toMatchObject({ role: 'body' });
    expect(resolve(1_000_001, 5)).toBeNull();
  });

  test('keeps styling render-only while conditional formatting wins', () => {
    const resolve = createSpreadsheetTableRenderResolver([
      {
        id: 'table-1',
        name: 'Table1',
        range: { row: [0, 2], column: [0, 1] },
        columns: [{ name: 'A' }, { name: 'B' }],
        filters: [],
        headerRow: true,
        totalsRow: false,
        style: { family: 'dark', number: 3 },
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    ]);
    const cell = Object.freeze({ v: 'Header', fc: '#111111', bl: 0 });
    const calls: string[] = [];
    const context = {
      fillStyle: '#ffffff',
      font: '12px sans-serif',
      strokeStyle: '#000000',
      lineWidth: 1,
      save: () => calls.push('save'),
      beginPath: () => calls.push('begin'),
      moveTo: () => undefined,
      lineTo: () => undefined,
      closePath: () => undefined,
      stroke: () => calls.push('stroke'),
      restore: () => calls.push('restore'),
      fillText: () =>
        calls.push(`text:${String(context.fillStyle)}:${context.font}`),
    } as unknown as CanvasRenderingContext2D;
    const style = resolve(0, 0);

    beginSpreadsheetTableCellRender(
      cell,
      style,
      { cellColor: '#fff2cc', textColor: '#9c0006' },
      context,
    );
    expect(context.fillStyle).toBe('#fff2cc');
    context.fillText('Header', 0, 0);
    expect(cell).toEqual({ v: 'Header', fc: '#111111', bl: 0 });

    finishSpreadsheetTableCellRender(
      cell,
      { row: 0, column: 0, startX: 0, startY: 0, endX: 96, endY: 24 },
      style,
      context,
    );
    expect(cell).toEqual({ v: 'Header', fc: '#111111', bl: 0 });
    expect(calls).toEqual([
      'text:#9c0006:bold 12px sans-serif',
      'save',
      'begin',
      'stroke',
      'restore',
    ]);

    const sparseCell = { v: 'Body' };
    const bodyStyle = resolve(1, 0);
    beginSpreadsheetTableCellRender(sparseCell, bodyStyle, undefined, context);
    finishSpreadsheetTableCellRender(
      sparseCell,
      { row: 1, column: 0, startX: 0, startY: 24, endX: 96, endY: 48 },
      bodyStyle,
      context,
    );
    expect(sparseCell).toEqual({ v: 'Body' });
  });
});
