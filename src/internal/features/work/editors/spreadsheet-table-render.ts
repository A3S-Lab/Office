import type { Cell } from '@fortune-sheet/core';
import type { SpreadsheetTableCellRenderStyle } from './spreadsheet-table-style';

export interface SpreadsheetTableRenderCellInfo {
  column: number;
  endX: number;
  endY: number;
  row: number;
  startX: number;
  startY: number;
}

export interface SpreadsheetTableConditionalStyle {
  cellColor?: string;
  textColor?: string;
}

interface SpreadsheetTableContextRestore {
  fillText: CanvasRenderingContext2D['fillText'];
  fillTextOwned: boolean;
}

const pendingContextRestores = new WeakMap<
  CanvasRenderingContext2D,
  SpreadsheetTableContextRestore
>();

export function beginSpreadsheetTableCellRender(
  _cell: Cell | null,
  tableStyle: SpreadsheetTableCellRenderStyle | null,
  conditionalStyle: SpreadsheetTableConditionalStyle | undefined,
  context: CanvasRenderingContext2D,
): void {
  if (tableStyle) context.fillStyle = tableStyle.background;
  if (conditionalStyle?.cellColor) {
    context.fillStyle = conditionalStyle.cellColor;
  }
  if (
    !tableStyle ||
    pendingContextRestores.has(context) ||
    typeof context.fillText !== 'function'
  ) {
    return;
  }
  const fillText = context.fillText;
  pendingContextRestores.set(context, {
    fillText,
    fillTextOwned: Object.hasOwn(context, 'fillText'),
  });
  const textColor = conditionalStyle?.textColor ?? tableStyle.textColor;
  context.fillText = (text, x, y, maxWidth) => {
    const previousFillStyle = context.fillStyle;
    const previousFont = context.font;
    context.fillStyle = textColor;
    if (tableStyle.bold)
      context.font = boldSpreadsheetTableCanvasFont(context.font);
    try {
      if (maxWidth === undefined) fillText.call(context, text, x, y);
      else fillText.call(context, text, x, y, maxWidth);
    } finally {
      context.fillStyle = previousFillStyle;
      context.font = previousFont;
    }
  };
}

export function finishSpreadsheetTableCellRender(
  _cell: Cell | null,
  cellInfo: SpreadsheetTableRenderCellInfo,
  tableStyle: SpreadsheetTableCellRenderStyle | null,
  context: CanvasRenderingContext2D,
): void {
  if (tableStyle) drawSpreadsheetTableCellBorder(cellInfo, tableStyle, context);
  const restore = pendingContextRestores.get(context);
  if (!restore) return;
  if (restore.fillTextOwned) context.fillText = restore.fillText;
  else delete (context as Partial<CanvasRenderingContext2D>).fillText;
  pendingContextRestores.delete(context);
}

function boldSpreadsheetTableCanvasFont(font: string): string {
  return /(?:^|\s)(?:bold|[6-9]00)(?:\s|$)/i.test(font) ? font : `bold ${font}`;
}

function drawSpreadsheetTableCellBorder(
  cellInfo: SpreadsheetTableRenderCellInfo,
  tableStyle: SpreadsheetTableCellRenderStyle,
  context: CanvasRenderingContext2D,
): void {
  const left = Math.round(cellInfo.startX) + 0.5;
  const top = Math.round(cellInfo.startY) + 0.5;
  const right = Math.round(cellInfo.endX) - 0.5;
  const bottom = Math.round(cellInfo.endY) - 0.5;
  context.save();
  context.beginPath();
  context.strokeStyle = tableStyle.borderColor;
  context.lineWidth = tableStyle.role === 'header' ? 1 : 0.75;
  context.moveTo(left, top);
  context.lineTo(right, top);
  context.lineTo(right, bottom);
  context.lineTo(left, bottom);
  context.closePath();
  context.stroke();
  context.restore();
}
