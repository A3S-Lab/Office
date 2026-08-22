import { drawSpreadsheetGradientFill } from './work-spreadsheet-gradient-fill-canvas';
import { drawSpreadsheetPatternFillOverlay } from './work-spreadsheet-pattern-fill-canvas';
import type { XlsxNativeFill } from './work-xlsx-native-fill';

interface SpreadsheetNativeFillContextRestore {
  fillRect: CanvasRenderingContext2D['fillRect'];
  fillRectOwned: boolean;
}

const pendingContextRestores = new WeakMap<
  CanvasRenderingContext2D,
  SpreadsheetNativeFillContextRestore
>();

/**
 * Fortune Sheet paints a cell background immediately after beforeRenderCell.
 * Intercepting only that first viewport paint keeps native fills behind text
 * and avoids any fill work for cells outside the visible Canvas.
 */
export function beginSpreadsheetNativeFillRender(
  fill: XlsxNativeFill | undefined,
  context: CanvasRenderingContext2D,
): void {
  finishSpreadsheetNativeFillRender(context);
  if (!fill || typeof context.fillRect !== 'function') return;
  context.fillStyle =
    fill.kind === 'gradient'
      ? (fill.value.stops[0]?.color ?? '#ffffff')
      : fill.value.backgroundColor;
  const fillRect = context.fillRect;
  pendingContextRestores.set(context, {
    fillRect,
    fillRectOwned: Object.hasOwn(context, 'fillRect'),
  });
  context.fillRect = (x, y, width, height) => {
    restoreSpreadsheetNativeFillContext(context);
    const bounds = {
      endX: x + width,
      endY: y + height,
      startX: x,
      startY: y,
    };
    if (fill.kind === 'gradient') {
      drawSpreadsheetGradientFill(context, bounds, fill.value, fillRect);
      return;
    }
    context.fillStyle = fill.value.backgroundColor;
    fillRect.call(context, x, y, width, height);
    drawSpreadsheetPatternFillOverlay(context, bounds, fill.value);
  };
}

export function finishSpreadsheetNativeFillRender(
  context: CanvasRenderingContext2D,
): void {
  restoreSpreadsheetNativeFillContext(context);
}

function restoreSpreadsheetNativeFillContext(
  context: CanvasRenderingContext2D,
): void {
  const restore = pendingContextRestores.get(context);
  if (!restore) return;
  if (restore.fillRectOwned) context.fillRect = restore.fillRect;
  else delete (context as Partial<CanvasRenderingContext2D>).fillRect;
  pendingContextRestores.delete(context);
}
