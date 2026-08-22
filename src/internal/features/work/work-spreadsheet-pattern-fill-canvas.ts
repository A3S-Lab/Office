import type { XlsxPatternFill } from './work-xlsx-pattern-fill';

export interface SpreadsheetPatternFillBounds {
  endX: number;
  endY: number;
  startX: number;
  startY: number;
}

export function drawSpreadsheetPatternFillOverlay(
  context: CanvasRenderingContext2D,
  bounds: SpreadsheetPatternFillBounds,
  fill: XlsxPatternFill,
): void {
  const left = Math.min(bounds.startX, bounds.endX) + 1;
  const top = Math.min(bounds.startY, bounds.endY) + 1;
  const right = Math.max(bounds.startX, bounds.endX) - 1;
  const bottom = Math.max(bounds.startY, bounds.endY) - 1;
  if (right <= left || bottom <= top) return;

  context.save();
  context.beginPath();
  context.rect(left, top, right - left, bottom - top);
  context.clip();
  context.strokeStyle = fill.foregroundColor;
  context.fillStyle = fill.foregroundColor;
  context.setLineDash([]);

  switch (fill.patternType) {
    case 'darkHorizontal':
      drawHorizontalLines(context, left, top, right, bottom, 4, 2);
      break;
    case 'lightHorizontal':
      drawHorizontalLines(context, left, top, right, bottom, 7, 1);
      break;
    case 'darkVertical':
      drawVerticalLines(context, left, top, right, bottom, 4, 2);
      break;
    case 'lightVertical':
      drawVerticalLines(context, left, top, right, bottom, 7, 1);
      break;
    case 'darkDown':
      drawDiagonalLines(context, left, top, right, bottom, 5, 2, 'down');
      break;
    case 'lightDown':
      drawDiagonalLines(context, left, top, right, bottom, 8, 1, 'down');
      break;
    case 'darkUp':
      drawDiagonalLines(context, left, top, right, bottom, 5, 2, 'up');
      break;
    case 'lightUp':
      drawDiagonalLines(context, left, top, right, bottom, 8, 1, 'up');
      break;
    case 'darkGrid':
      drawHorizontalLines(context, left, top, right, bottom, 5, 1.5);
      drawVerticalLines(context, left, top, right, bottom, 5, 1.5);
      break;
    case 'lightGrid':
      drawHorizontalLines(context, left, top, right, bottom, 8, 1);
      drawVerticalLines(context, left, top, right, bottom, 8, 1);
      break;
    case 'darkTrellis':
      drawDiagonalLines(context, left, top, right, bottom, 6, 1.5, 'down');
      drawDiagonalLines(context, left, top, right, bottom, 6, 1.5, 'up');
      break;
    case 'lightTrellis':
      drawDiagonalLines(context, left, top, right, bottom, 10, 1, 'down');
      drawDiagonalLines(context, left, top, right, bottom, 10, 1, 'up');
      break;
    case 'gray0625':
      drawPatternDots(context, left, top, right, bottom, 8, 1);
      break;
    case 'gray125':
      drawPatternDots(context, left, top, right, bottom, 6, 1);
      break;
    case 'lightGray':
      drawPatternDots(context, left, top, right, bottom, 4, 1.5);
      break;
    case 'mediumGray':
      drawDiagonalLines(context, left, top, right, bottom, 4, 1, 'down');
      drawDiagonalLines(context, left, top, right, bottom, 4, 1, 'up');
      break;
    case 'darkGray':
      context.fillRect(left, top, right - left, bottom - top);
      context.fillStyle = fill.backgroundColor;
      drawPatternDots(context, left, top, right, bottom, 4, 1.5);
      break;
  }
  context.restore();
}

function drawHorizontalLines(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  right: number,
  bottom: number,
  requestedSpacing: number,
  lineWidth: number,
): void {
  const spacing = boundedLineSpacing(bottom - top, requestedSpacing);
  context.lineWidth = lineWidth;
  context.beginPath();
  for (let y = alignedPatternStart(top, spacing); y <= bottom; y += spacing) {
    context.moveTo(left, crispPatternCoordinate(y, lineWidth));
    context.lineTo(right, crispPatternCoordinate(y, lineWidth));
  }
  context.stroke();
}

function drawVerticalLines(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  right: number,
  bottom: number,
  requestedSpacing: number,
  lineWidth: number,
): void {
  const spacing = boundedLineSpacing(right - left, requestedSpacing);
  context.lineWidth = lineWidth;
  context.beginPath();
  for (let x = alignedPatternStart(left, spacing); x <= right; x += spacing) {
    context.moveTo(crispPatternCoordinate(x, lineWidth), top);
    context.lineTo(crispPatternCoordinate(x, lineWidth), bottom);
  }
  context.stroke();
}

function drawDiagonalLines(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  right: number,
  bottom: number,
  requestedSpacing: number,
  lineWidth: number,
  direction: 'down' | 'up',
): void {
  const height = bottom - top;
  const spacing = boundedLineSpacing(right - left + height, requestedSpacing);
  context.lineWidth = lineWidth;
  context.beginPath();
  const first = alignedPatternStart(left - height, spacing);
  for (let x = first; x <= right; x += spacing) {
    if (direction === 'down') {
      context.moveTo(x, top);
      context.lineTo(x + height, bottom);
    } else {
      context.moveTo(x, bottom);
      context.lineTo(x + height, top);
    }
  }
  context.stroke();
}

function drawPatternDots(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  right: number,
  bottom: number,
  requestedSpacing: number,
  size: number,
): void {
  const area = Math.max(1, (right - left) * (bottom - top));
  const spacing = Math.max(requestedSpacing, Math.sqrt(area / 2_048));
  const startX = alignedPatternStart(left, spacing);
  const startY = alignedPatternStart(top, spacing);
  for (let y = startY; y <= bottom; y += spacing) {
    const offset = Math.round(y / spacing) % 2 ? spacing / 2 : 0;
    for (let x = startX + offset; x <= right; x += spacing) {
      context.fillRect(x, y, size, size);
    }
  }
}

function alignedPatternStart(value: number, spacing: number): number {
  return Math.floor(value / spacing) * spacing;
}

function boundedLineSpacing(span: number, requested: number): number {
  return Math.max(requested, span / 512);
}

function crispPatternCoordinate(value: number, lineWidth: number): number {
  return Math.round(value) + (Math.round(lineWidth) % 2 ? 0.5 : 0);
}
