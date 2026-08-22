export interface SpreadsheetDiagonalBorderBounds {
  endX: number;
  endY: number;
  startX: number;
  startY: number;
}

export interface SpreadsheetDiagonalCanvasLine {
  color: string;
  style: string;
}

export function drawSpreadsheetDiagonalDownBorder(
  context: CanvasRenderingContext2D,
  bounds: SpreadsheetDiagonalBorderBounds,
  line: SpreadsheetDiagonalCanvasLine,
): void {
  drawSpreadsheetDiagonalBorder(context, bounds, line, 'down');
}

export function drawSpreadsheetDiagonalUpBorder(
  context: CanvasRenderingContext2D,
  bounds: SpreadsheetDiagonalBorderBounds,
  line: SpreadsheetDiagonalCanvasLine,
): void {
  drawSpreadsheetDiagonalBorder(context, bounds, line, 'up');
}

function drawSpreadsheetDiagonalBorder(
  context: CanvasRenderingContext2D,
  bounds: SpreadsheetDiagonalBorderBounds,
  line: SpreadsheetDiagonalCanvasLine,
  direction: 'down' | 'up',
): void {
  const left = Math.round(bounds.startX) + 0.5;
  const top = Math.round(bounds.startY) + 0.5;
  const right = Math.round(bounds.endX) - 1.5;
  const bottom = Math.round(bounds.endY) - 1.5;
  if (right <= left || bottom <= top) return;

  context.save();
  context.strokeStyle = line.color;
  context.lineCap = 'butt';
  context.lineWidth = spreadsheetDiagonalLineWidth(line.style);
  context.setLineDash(spreadsheetDiagonalLineDash(line.style));
  const startY = direction === 'down' ? top : bottom;
  const endY = direction === 'down' ? bottom : top;
  if (line.style === '7') {
    drawLine(context, left, startY - 1, right, endY - 1);
    drawLine(context, left, startY + 1, right, endY + 1);
  } else {
    drawLine(context, left, startY, right, endY);
  }
  context.restore();
}

function drawLine(
  context: CanvasRenderingContext2D,
  startX: number,
  startY: number,
  endX: number,
  endY: number,
): void {
  context.beginPath();
  context.moveTo(startX, startY);
  context.lineTo(endX, endY);
  context.stroke();
}

function spreadsheetDiagonalLineWidth(style: string): number {
  if (style === '2') return 0.5;
  if (['8', '9', '10', '11', '12'].includes(style)) return 2;
  if (style === '13') return 3;
  return 1;
}

function spreadsheetDiagonalLineDash(style: string): number[] {
  if (style === '3') return [1, 2];
  if (style === '4' || style === '9') return [5, 3];
  if (style === '5' || style === '10' || style === '12') return [6, 3, 1, 3];
  if (style === '6' || style === '11') return [6, 3, 1, 3, 1, 3];
  return [];
}
