import type {
  XlsxGradientFill,
  XlsxGradientStop,
} from './work-xlsx-gradient-fill';

export interface SpreadsheetGradientFillBounds {
  endX: number;
  endY: number;
  startX: number;
  startY: number;
}

type SpreadsheetGradientFillRect = CanvasRenderingContext2D['fillRect'];

const MAX_PATH_GRADIENT_CONTOURS = 96;

export function drawSpreadsheetGradientFill(
  context: CanvasRenderingContext2D,
  bounds: SpreadsheetGradientFillBounds,
  fill: XlsxGradientFill,
  fillRect: SpreadsheetGradientFillRect = context.fillRect,
): void {
  const left = Math.min(bounds.startX, bounds.endX);
  const top = Math.min(bounds.startY, bounds.endY);
  const right = Math.max(bounds.startX, bounds.endX);
  const bottom = Math.max(bounds.startY, bounds.endY);
  const width = right - left;
  const height = bottom - top;
  if (width <= 0 || height <= 0) return;

  context.save();
  if (fill.type === 'linear') {
    drawLinearGradient(context, left, top, width, height, fill, fillRect);
  } else {
    drawPathGradient(context, left, top, width, height, fill, fillRect);
  }
  context.restore();
}

function drawLinearGradient(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  width: number,
  height: number,
  fill: Extract<XlsxGradientFill, { type: 'linear' }>,
  fillRect: SpreadsheetGradientFillRect,
): void {
  if (typeof context.createLinearGradient !== 'function') {
    context.fillStyle = fill.stops[0]?.color ?? '#ffffff';
    fillRect.call(context, left, top, width, height);
    return;
  }
  const radians = (fill.degree * Math.PI) / 180;
  const horizontal = Math.cos(radians);
  const vertical = Math.sin(radians);
  const centerX = left + width / 2;
  const centerY = top + height / 2;
  const span = Math.abs(width * horizontal) + Math.abs(height * vertical);
  const gradient = context.createLinearGradient(
    centerX - (horizontal * span) / 2,
    centerY - (vertical * span) / 2,
    centerX + (horizontal * span) / 2,
    centerY + (vertical * span) / 2,
  );
  for (const stop of fill.stops) {
    gradient.addColorStop(stop.position, stop.color);
  }
  context.fillStyle = gradient;
  fillRect.call(context, left, top, width, height);
}

function drawPathGradient(
  context: CanvasRenderingContext2D,
  left: number,
  top: number,
  width: number,
  height: number,
  fill: Extract<XlsxGradientFill, { type: 'path' }>,
  fillRect: SpreadsheetGradientFillRect,
): void {
  const innerLeft = left + width * fill.left;
  const innerRight = left + width * fill.right;
  const innerTop = top + height * fill.top;
  const innerBottom = top + height * fill.bottom;
  const contours = Math.min(
    MAX_PATH_GRADIENT_CONTOURS,
    Math.max(2, Math.ceil(Math.max(width, height))),
  );

  for (let index = contours; index >= 0; index -= 1) {
    const progress = index / contours;
    const contourLeft = interpolate(innerLeft, left, progress);
    const contourRight = interpolate(innerRight, left + width, progress);
    const contourTop = interpolate(innerTop, top, progress);
    const contourBottom = interpolate(innerBottom, top + height, progress);
    context.fillStyle = gradientColorAt(fill.stops, progress);
    fillRect.call(
      context,
      contourLeft,
      contourTop,
      contourRight - contourLeft,
      contourBottom - contourTop,
    );
  }
}

function gradientColorAt(
  stops: readonly XlsxGradientStop[],
  position: number,
): string {
  const first = stops[0];
  const last = stops.at(-1);
  if (!first || !last || position <= first.position) {
    return first?.color ?? '#ffffff';
  }
  if (position >= last.position) return last.color;
  for (let index = 1; index < stops.length; index += 1) {
    const right = stops[index];
    const left = stops[index - 1];
    if (!left || !right || position > right.position) continue;
    const span = right.position - left.position;
    const progress = span > 0 ? (position - left.position) / span : 1;
    return interpolateRgb(left.color, right.color, progress);
  }
  return last.color;
}

function interpolateRgb(left: string, right: string, progress: number): string {
  const leftRgb = rgbChannels(left);
  const rightRgb = rgbChannels(right);
  if (!leftRgb || !rightRgb) return progress < 0.5 ? left : right;
  return `#${leftRgb
    .map((channel, index) =>
      Math.round(interpolate(channel, rightRgb[index] ?? channel, progress))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

function rgbChannels(color: string): [number, number, number] | null {
  if (!/^#[0-9a-f]{6}$/i.test(color)) return null;
  return [
    Number.parseInt(color.slice(1, 3), 16),
    Number.parseInt(color.slice(3, 5), 16),
    Number.parseInt(color.slice(5, 7), 16),
  ];
}

function interpolate(left: number, right: number, progress: number): number {
  return left + (right - left) * progress;
}
