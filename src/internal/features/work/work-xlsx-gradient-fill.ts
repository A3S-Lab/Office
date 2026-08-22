import type { Cell } from '@fortune-sheet/core';
import { attribute, directChildren } from './work-ooxml-package';
import {
  normalizeXlsxSemanticColorOrigin,
  readXlsxSemanticColorOrigin,
  type XlsxSemanticColorOrigin,
  xlsxSemanticColorOriginKey,
} from './work-xlsx-cell-style-origin';
import { xlsxRgbColor } from './work-xlsx-cell-style-xml';
import { resolveXlsxColor, type XlsxColorResolver } from './work-xlsx-colors';

export const XLSX_GRADIENT_FILL_CELL_KEY = 'a3sXlsxGradientFill' as const;
export const MAX_XLSX_GRADIENT_STOPS = 256;

export interface XlsxGradientStop {
  color: string;
  colorOrigin?: XlsxSemanticColorOrigin;
  position: number;
}

export type XlsxGradientFill =
  | {
      degree: number;
      stops: XlsxGradientStop[];
      type: 'linear';
    }
  | {
      bottom: number;
      left: number;
      right: number;
      stops: XlsxGradientStop[];
      top: number;
      type: 'path';
    };

type CellWithXlsxGradientFill = Cell & {
  [XLSX_GRADIENT_FILL_CELL_KEY]?: unknown;
};

export function readXlsxGradientFill(
  element: Element | undefined,
  colors: XlsxColorResolver,
): XlsxGradientFill | undefined {
  if (!element) return undefined;
  const type = attribute(element, 'type') ?? 'linear';
  if (type !== 'linear' && type !== 'path') return undefined;
  const children = directChildren(element);
  if (
    children.length < 2 ||
    children.length > MAX_XLSX_GRADIENT_STOPS ||
    children.some((child) => child.localName !== 'stop')
  ) {
    return undefined;
  }

  const stops: XlsxGradientStop[] = [];
  for (const stop of children) {
    const position = unitInterval(attribute(stop, 'position'));
    const stopChildren = directChildren(stop);
    const colorElement = stopChildren[0];
    if (
      position === null ||
      stopChildren.length !== 1 ||
      colorElement?.localName !== 'color'
    ) {
      return undefined;
    }
    const renderedColor = resolveXlsxColor(colorElement, colors);
    const color = normalizedColor(renderedColor);
    if (!color || (stops.at(-1)?.position ?? 0) > position) return undefined;
    const colorOrigin = readXlsxSemanticColorOrigin(colorElement, colors);
    stops.push({
      color,
      ...(colorOrigin ? { colorOrigin } : {}),
      position,
    });
  }

  if (type === 'linear') {
    const degree = optionalFiniteNumber(attribute(element, 'degree'), 0);
    return degree === null ? undefined : { degree, stops, type };
  }
  const left = optionalUnitInterval(attribute(element, 'left'), 0);
  const right = optionalUnitInterval(attribute(element, 'right'), 0);
  const top = optionalUnitInterval(attribute(element, 'top'), 0);
  const bottom = optionalUnitInterval(attribute(element, 'bottom'), 0);
  if (
    left === null ||
    right === null ||
    top === null ||
    bottom === null ||
    left > right ||
    top > bottom
  ) {
    return undefined;
  }
  return { bottom, left, right, stops, top, type };
}

export function withXlsxGradientFill(
  cell: Cell,
  fill: XlsxGradientFill | undefined,
): Cell {
  const normalized = normalizeXlsxGradientFill(fill);
  return normalized
    ? ({
        ...cell,
        [XLSX_GRADIENT_FILL_CELL_KEY]: normalized,
      } as CellWithXlsxGradientFill)
    : cell;
}

export function deleteXlsxGradientFill(cell: Cell): void {
  delete (cell as CellWithXlsxGradientFill)[XLSX_GRADIENT_FILL_CELL_KEY];
}

export function xlsxGradientFill(
  cell: Cell | null | undefined,
): XlsxGradientFill | undefined {
  return normalizeXlsxGradientFill(
    (cell as CellWithXlsxGradientFill | null | undefined)?.[
      XLSX_GRADIENT_FILL_CELL_KEY
    ],
  );
}

export function activeXlsxGradientFill(
  cell: Cell | null | undefined,
): XlsxGradientFill | undefined {
  const fill = xlsxGradientFill(cell);
  return fill &&
    normalizedColor(cell?.bg) === xlsxGradientFillFallbackColor(fill)
    ? fill
    : undefined;
}

export function normalizeXlsxGradientFill(
  value: unknown,
): XlsxGradientFill | undefined {
  if (!isRecord(value) || !Array.isArray(value.stops)) return undefined;
  if (value.stops.length < 2 || value.stops.length > MAX_XLSX_GRADIENT_STOPS) {
    return undefined;
  }
  const stops: XlsxGradientStop[] = [];
  for (const valueStop of value.stops) {
    if (!isRecord(valueStop)) return undefined;
    const position = normalizedUnitInterval(valueStop.position);
    const color = normalizedColor(valueStop.color);
    if (
      position === null ||
      !color ||
      (stops.at(-1)?.position ?? 0) > position
    ) {
      return undefined;
    }
    const colorOrigin = normalizeXlsxSemanticColorOrigin(valueStop.colorOrigin);
    stops.push({
      color,
      ...(colorOrigin ? { colorOrigin } : {}),
      position,
    });
  }
  if (value.type === 'linear') {
    const degree = finiteNumber(value.degree);
    return degree === null ? undefined : { degree, stops, type: 'linear' };
  }
  if (value.type !== 'path') return undefined;
  const left = normalizedUnitInterval(value.left);
  const right = normalizedUnitInterval(value.right);
  const top = normalizedUnitInterval(value.top);
  const bottom = normalizedUnitInterval(value.bottom);
  if (
    left === null ||
    right === null ||
    top === null ||
    bottom === null ||
    left > right ||
    top > bottom
  ) {
    return undefined;
  }
  return { bottom, left, right, stops, top, type: 'path' };
}

export function xlsxGradientFillFallbackColor(fill: XlsxGradientFill): string {
  return fill.stops[0]?.color ?? '#ffffff';
}

export function xlsxGradientFillSemanticColors(
  fill: XlsxGradientFill | undefined,
): XlsxSemanticColorOrigin[] {
  return (fill?.stops ?? []).flatMap((stop) =>
    stop.colorOrigin ? [stop.colorOrigin] : [],
  );
}

export function xlsxGradientFillKey(
  fill: XlsxGradientFill | undefined,
): string {
  if (!fill) return '';
  const geometry =
    fill.type === 'linear'
      ? `linear:${fill.degree}`
      : `path:${fill.left}:${fill.right}:${fill.top}:${fill.bottom}`;
  return `${geometry}:${fill.stops
    .map(
      (stop) =>
        `${stop.position}:${stop.color}:${xlsxSemanticColorOriginKey(stop.colorOrigin)}`,
    )
    .join('|')}`;
}

function unitInterval(value: string | null): number | null {
  if (value === null || !value.trim()) return null;
  return normalizedUnitInterval(Number(value));
}

function optionalUnitInterval(
  value: string | null,
  fallback: number,
): number | null {
  return value === null ? fallback : unitInterval(value);
}

function normalizedUnitInterval(value: unknown): number | null {
  return typeof value === 'number' &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 1
    ? value
    : null;
}

function optionalFiniteNumber(
  value: string | null,
  fallback: number,
): number | null {
  if (value === null) return fallback;
  if (!value.trim()) return null;
  return finiteNumber(Number(value));
}

function finiteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function normalizedColor(value: unknown): string | null {
  const rgb = xlsxRgbColor(value);
  return rgb ? `#${rgb.slice(-6).toLowerCase()}` : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
