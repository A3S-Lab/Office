import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';
import { useEffect, useRef } from 'react';
import { drawSpreadsheetGradientFill } from '../work-spreadsheet-gradient-fill-canvas';
import { drawSpreadsheetPatternFillOverlay } from '../work-spreadsheet-pattern-fill-canvas';
import type { XlsxPatternFillType } from '../work-xlsx-pattern-fill';
import type { SpreadsheetCellFillFormat } from './spreadsheet-cell-fill-format';

export function SpreadsheetFormatCellsFillPreview({
  fill,
  patternLabels,
}: {
  fill: SpreadsheetCellFillFormat;
  patternLabels: Readonly<Record<XlsxPatternFillType, string>>;
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const render = () => renderSpreadsheetFillPreview(canvas, fill);
    render();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(render);
    observer.observe(canvas);
    return () => observer.disconnect();
  }, [fill]);

  return (
    <figure
      className={`work-spreadsheet-format-cells-fill-preview ${fill.kind}`}
    >
      <canvas
        ref={canvasRef}
        role="img"
        aria-label={spreadsheetFillPreviewLabel(fill, patternLabels)}
      />
      <figcaption>
        {spreadsheetFillPreviewCaption(fill, patternLabels)}
      </figcaption>
    </figure>
  );
}

function renderSpreadsheetFillPreview(
  canvas: HTMLCanvasElement,
  fill: SpreadsheetCellFillFormat,
): void {
  let context: CanvasRenderingContext2D | null = null;
  try {
    context = canvas.getContext('2d');
  } catch {
    return;
  }
  if (!context) return;
  const bounds = canvas.getBoundingClientRect();
  const width = Math.max(1, Math.round(bounds.width || 220));
  const height = Math.max(1, Math.round(bounds.height || 132));
  const scale = Math.min(2, Math.max(1, window.devicePixelRatio || 1));
  const pixelWidth = Math.round(width * scale);
  const pixelHeight = Math.round(height * scale);
  if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
  if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
  context.setTransform(scale, 0, 0, scale, 0, 0);
  context.clearRect(0, 0, width, height);
  if (fill.kind === 'none') return;
  if (fill.kind === 'solid') {
    context.fillStyle = fill.color;
    context.fillRect(0, 0, width, height);
    return;
  }
  if (fill.kind === 'pattern') {
    context.fillStyle = fill.value.backgroundColor;
    context.fillRect(0, 0, width, height);
    drawSpreadsheetPatternFillOverlay(
      context,
      { endX: width, endY: height, startX: 0, startY: 0 },
      fill.value,
    );
    return;
  }
  drawSpreadsheetGradientFill(
    context,
    { endX: width, endY: height, startX: 0, startY: 0 },
    fill.value,
  );
}

function spreadsheetFillPreviewLabel(
  fill: SpreadsheetCellFillFormat,
  patternLabels: Readonly<Record<XlsxPatternFillType, string>>,
): string {
  const messages = resolveOfficeMessages();
  if (fill.kind === 'none') {
    return officeMessage(messages, 'spreadsheet.formatCells.fill.preview.none');
  }
  if (fill.kind === 'solid') {
    return officeMessage(messages, 'spreadsheet.formatCells.fill.preview.solid', {
      color: fill.color,
    });
  }
  if (fill.kind === 'pattern') {
    return officeMessage(messages, 'spreadsheet.formatCells.fill.preview.pattern', {
      pattern: patternLabels[fill.value.patternType],
    });
  }
  const type = officeMessage(
    messages,
    fill.value.type === 'linear'
      ? 'spreadsheet.formatCells.fill.gradient.linear'
      : 'spreadsheet.formatCells.fill.gradient.path',
  );
  return officeMessage(messages, 'spreadsheet.formatCells.fill.preview.gradient', {
    type,
    n: String(fill.value.stops.length),
  });
}

function spreadsheetFillPreviewCaption(
  fill: SpreadsheetCellFillFormat,
  patternLabels: Readonly<Record<XlsxPatternFillType, string>>,
): string {
  const messages = resolveOfficeMessages();
  if (fill.kind === 'none') {
    return officeMessage(messages, 'spreadsheet.formatCells.fill.summary.none');
  }
  if (fill.kind === 'solid') return fill.color.toUpperCase();
  if (fill.kind === 'pattern') return patternLabels[fill.value.patternType];
  return fill.value.type === 'linear'
    ? officeMessage(messages, 'spreadsheet.formatCells.fill.summary.linear', {
        degree: formatNumber(fill.value.degree),
        n: String(fill.value.stops.length),
      })
    : officeMessage(messages, 'spreadsheet.formatCells.fill.summary.path', {
        n: String(fill.value.stops.length),
      });
}

function formatNumber(value: number): string {
  return Number.isInteger(value)
    ? String(value)
    : String(Number(value.toFixed(4)));
}
