import { describe, expect, test } from '@rstest/core';
import {
  drawSpreadsheetDiagonalDownBorder,
  drawSpreadsheetDiagonalUpBorder,
} from '../src/internal/features/work/work-spreadsheet-diagonal-border-canvas';

describe('spreadsheet diagonal-up Canvas rendering', () => {
  test('draws diagonal-down from the upper-left to the lower-right', () => {
    const calls: unknown[][] = [];
    const context = canvasContext(calls);

    drawSpreadsheetDiagonalDownBorder(
      context,
      { startX: 10, startY: 20, endX: 110, endY: 44 },
      { color: '#d84b4f', style: '13' },
    );

    expect(context.strokeStyle).toBe('#d84b4f');
    expect(context.lineWidth).toBe(3);
    expect(calls).toContainEqual(['moveTo', 10.5, 20.5]);
    expect(calls).toContainEqual(['lineTo', 108.5, 42.5]);
  });

  test('draws from the lower-left to the upper-right with the native dash style', () => {
    const calls: unknown[][] = [];
    const context = canvasContext(calls);

    drawSpreadsheetDiagonalUpBorder(
      context,
      { startX: 10, startY: 20, endX: 110, endY: 44 },
      { color: '#2463eb', style: '10' },
    );

    expect(context.strokeStyle).toBe('#2463eb');
    expect(context.lineWidth).toBe(2);
    expect(calls).toContainEqual(['setLineDash', 6, 3, 1, 3]);
    expect(calls).toContainEqual(['moveTo', 10.5, 42.5]);
    expect(calls).toContainEqual(['lineTo', 108.5, 20.5]);
    expect(calls.filter(([name]) => name === 'stroke')).toHaveLength(1);
  });

  test('renders imported double diagonals as two parallel strokes', () => {
    const calls: unknown[][] = [];
    const context = canvasContext(calls);

    drawSpreadsheetDiagonalUpBorder(
      context,
      { startX: 0, startY: 0, endX: 20, endY: 20 },
      { color: '#172033', style: '7' },
    );

    expect(calls.filter(([name]) => name === 'stroke')).toHaveLength(2);
    expect(calls).toContainEqual(['moveTo', 0.5, 17.5]);
    expect(calls).toContainEqual(['moveTo', 0.5, 19.5]);
  });
});

function canvasContext(calls: unknown[][]): CanvasRenderingContext2D {
  return {
    beginPath: () => calls.push(['beginPath']),
    lineCap: 'butt',
    lineTo: (x, y) => calls.push(['lineTo', x, y]),
    lineWidth: 1,
    moveTo: (x, y) => calls.push(['moveTo', x, y]),
    restore: () => calls.push(['restore']),
    save: () => calls.push(['save']),
    setLineDash: (segments) => calls.push(['setLineDash', ...segments]),
    stroke: () => calls.push(['stroke']),
    strokeStyle: '#000000',
  } as CanvasRenderingContext2D;
}
