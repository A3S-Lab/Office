import type { Image } from '@fortune-sheet/core';
import type { WorkSpreadsheetChartLayout } from './work-spreadsheet-chart-layout';

export interface WorkSpreadsheetImage extends Image {
  name?: string;
  altText?: string;
  contentType?: string;
}

export type WorkSpreadsheetChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'radar'
  | 'scatter'
  | 'bubble'
  | 'combination';
export type WorkSpreadsheetRadarStyle = 'standard' | 'marker' | 'filled';
export type WorkSpreadsheetScatterStyle =
  | 'marker'
  | 'line'
  | 'lineMarker'
  | 'smooth'
  | 'smoothMarker';
export type WorkSpreadsheetBubbleSizeRepresents = 'area' | 'width';
export type WorkSpreadsheetCombinationSeriesType = 'column' | 'line' | 'area';
export type WorkSpreadsheetChartAxisGroup = 'primary' | 'secondary';
export type WorkSpreadsheetChartAxisPosition =
  | 'bottom'
  | 'left'
  | 'top'
  | 'right';
export type WorkSpreadsheetChartLineDash = 'solid' | 'dash' | 'dot' | 'dashDot';
export type WorkSpreadsheetChartMarkerSymbol =
  | 'none'
  | 'circle'
  | 'square'
  | 'diamond'
  | 'triangle'
  | 'plus'
  | 'x'
  | 'star';

export interface WorkSpreadsheetChartMarkerStyle {
  symbol?: WorkSpreadsheetChartMarkerSymbol;
  size?: number;
  fillColor?: string;
  lineColor?: string;
}

export interface WorkSpreadsheetChartSeriesStyle {
  fillColor?: string;
  fillTransparency?: number;
  lineColor?: string;
  lineWidth?: number;
  lineDash?: WorkSpreadsheetChartLineDash;
  marker?: WorkSpreadsheetChartMarkerStyle;
}

export interface WorkSpreadsheetChartAxis {
  title?: string;
  titleReference?: string;
  reverseOrder?: boolean;
  labelPosition?: 'nextTo' | 'high' | 'low' | 'none';
  majorTickMark?: 'none' | 'inside' | 'outside' | 'cross';
  labelInterval?: number;
  minimum?: number;
  maximum?: number;
  majorUnit?: number;
  showMajorGridlines?: boolean;
  numberFormat?: string;
  numberFormatSourceLinked?: boolean;
}

export interface WorkSpreadsheetChartAxes {
  bottom?: WorkSpreadsheetChartAxis;
  left?: WorkSpreadsheetChartAxis;
  top?: WorkSpreadsheetChartAxis;
  right?: WorkSpreadsheetChartAxis;
}

export type WorkSpreadsheetDataLabelPosition =
  | 'bestFit'
  | 'center'
  | 'insideBase'
  | 'insideEnd'
  | 'outsideEnd'
  | 'left'
  | 'right'
  | 'above'
  | 'below';

export interface WorkSpreadsheetDataLabels {
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercentage?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  position?: WorkSpreadsheetDataLabelPosition;
}

export type WorkSpreadsheetErrorBarDirection = 'x' | 'y';
export type WorkSpreadsheetErrorBarType = 'both' | 'plus' | 'minus';
export type WorkSpreadsheetErrorBarValueType =
  | 'fixedValue'
  | 'percentage'
  | 'standardDeviation'
  | 'standardError'
  | 'custom';

export interface WorkSpreadsheetErrorBars {
  direction: WorkSpreadsheetErrorBarDirection;
  barType: WorkSpreadsheetErrorBarType;
  valueType: WorkSpreadsheetErrorBarValueType;
  value?: number;
  showEndCaps?: boolean;
  plusValues?: number[];
  plusReference?: string;
  minusValues?: number[];
  minusReference?: string;
}

export type WorkSpreadsheetTrendlineType =
  | 'linear'
  | 'exponential'
  | 'logarithmic'
  | 'polynomial'
  | 'power'
  | 'movingAverage';

export interface WorkSpreadsheetTrendline {
  type: WorkSpreadsheetTrendlineType;
  name?: string;
  order?: number;
  period?: number;
  forward?: number;
  backward?: number;
  intercept?: number;
  displayEquation?: boolean;
  displayRSquared?: boolean;
}

export function normalizeWorkSpreadsheetDoughnutHoleSize(
  value: unknown,
): number {
  const size = Number(value);
  if (!Number.isFinite(size)) return 50;
  return Math.min(90, Math.max(10, Math.round(size)));
}

export function normalizeWorkSpreadsheetRadarStyle(
  value: unknown,
): WorkSpreadsheetRadarStyle {
  return value === 'standard' || value === 'filled' || value === 'marker'
    ? value
    : 'marker';
}

export function normalizeWorkSpreadsheetScatterStyle(
  value: unknown,
): WorkSpreadsheetScatterStyle {
  return value === 'line' ||
    value === 'lineMarker' ||
    value === 'smooth' ||
    value === 'smoothMarker' ||
    value === 'marker'
    ? value
    : 'marker';
}

export function normalizeWorkSpreadsheetBubbleScale(value: unknown): number {
  const scale = Number(value);
  if (!Number.isFinite(scale)) return 100;
  return Math.min(300, Math.max(0, Math.round(scale)));
}

export function normalizeWorkSpreadsheetBubbleSizeRepresents(
  value: unknown,
): WorkSpreadsheetBubbleSizeRepresents {
  return value === 'width' || value === 'w' ? 'width' : 'area';
}

export function workSpreadsheetChartUsesNumericXAxis(
  type: WorkSpreadsheetChartType,
): boolean {
  return type === 'scatter' || type === 'bubble';
}

export function workSpreadsheetChartSupportsAxes(
  type: WorkSpreadsheetChartType,
): boolean {
  return type !== 'pie' && type !== 'doughnut';
}

export function normalizeWorkSpreadsheetCombinationSeriesType(
  value: unknown,
): WorkSpreadsheetCombinationSeriesType {
  return value === 'line' || value === 'area' || value === 'column'
    ? value
    : 'column';
}

export function normalizeWorkSpreadsheetChartAxisGroup(
  value: unknown,
): WorkSpreadsheetChartAxisGroup {
  return value === 'secondary' ? 'secondary' : 'primary';
}

export function workSpreadsheetCombinationSeriesTypeLabel(
  type: WorkSpreadsheetCombinationSeriesType,
): string {
  if (type === 'line') return '折线图';
  if (type === 'area') return '面积图';
  return '柱形图';
}

export function workSpreadsheetChartSupportsTrendlines(
  type: WorkSpreadsheetChartType,
): boolean {
  return type !== 'pie' && type !== 'doughnut' && type !== 'radar';
}

export function workSpreadsheetChartSupportsErrorBars(
  type: WorkSpreadsheetChartType,
): boolean {
  return type !== 'pie' && type !== 'doughnut' && type !== 'radar';
}

export function normalizeWorkSpreadsheetErrorBars(
  source: WorkSpreadsheetErrorBars,
  chartType: WorkSpreadsheetChartType,
): WorkSpreadsheetErrorBars {
  const direction =
    workSpreadsheetChartUsesNumericXAxis(chartType) && source.direction === 'x'
      ? 'x'
      : 'y';
  const barType =
    source.barType === 'plus' || source.barType === 'minus'
      ? source.barType
      : 'both';
  const valueType =
    source.valueType === 'percentage' ||
    source.valueType === 'standardDeviation' ||
    source.valueType === 'standardError' ||
    source.valueType === 'custom' ||
    source.valueType === 'fixedValue'
      ? source.valueType
      : 'fixedValue';
  const fallbackValue = valueType === 'percentage' ? 5 : 1;
  const numericValue = Number(source.value);
  const value = Number.isFinite(numericValue)
    ? Math.max(0, numericValue)
    : fallbackValue;
  const plusValues = source.plusValues?.map(normalizedErrorBarAmount);
  const minusValues = source.minusValues?.map(normalizedErrorBarAmount);
  return {
    direction,
    barType,
    valueType,
    ...(valueType === 'fixedValue' ||
    valueType === 'percentage' ||
    valueType === 'standardDeviation'
      ? { value }
      : {}),
    ...(source.showEndCaps === false ? { showEndCaps: false } : {}),
    ...(valueType === 'custom' && barType !== 'minus' && plusValues?.length
      ? { plusValues }
      : {}),
    ...(valueType === 'custom' &&
    barType !== 'minus' &&
    source.plusReference?.trim()
      ? { plusReference: source.plusReference.trim().replace(/^=/, '') }
      : {}),
    ...(valueType === 'custom' && barType !== 'plus' && minusValues?.length
      ? { minusValues }
      : {}),
    ...(valueType === 'custom' &&
    barType !== 'plus' &&
    source.minusReference?.trim()
      ? { minusReference: source.minusReference.trim().replace(/^=/, '') }
      : {}),
  };
}

export function workSpreadsheetErrorBarValueTypeLabel(
  type: WorkSpreadsheetErrorBarValueType,
): string {
  if (type === 'percentage') return '百分比';
  if (type === 'standardDeviation') return '标准差';
  if (type === 'standardError') return '标准误差';
  if (type === 'custom') return '自定义';
  return '固定值';
}

export function workSpreadsheetErrorBarTypeLabel(
  type: WorkSpreadsheetErrorBarType,
): string {
  if (type === 'plus') return '正向';
  if (type === 'minus') return '负向';
  return '双向';
}

function normalizedErrorBarAmount(value: number): number {
  return Number.isFinite(value) ? Math.max(0, value) : 0;
}

export function normalizeWorkSpreadsheetDataLabelPosition(
  value: unknown,
): WorkSpreadsheetDataLabelPosition {
  return value === 'center' ||
    value === 'insideBase' ||
    value === 'insideEnd' ||
    value === 'outsideEnd' ||
    value === 'left' ||
    value === 'right' ||
    value === 'above' ||
    value === 'below' ||
    value === 'bestFit'
    ? value
    : 'bestFit';
}

export function normalizeWorkSpreadsheetDataLabels(
  source: WorkSpreadsheetDataLabels,
  chartType: WorkSpreadsheetChartType,
): WorkSpreadsheetDataLabels {
  const position = source.position
    ? normalizeWorkSpreadsheetDataLabelPosition(source.position)
    : undefined;
  const separator =
    typeof source.separator === 'string'
      ? source.separator.slice(0, 64)
      : undefined;
  return {
    ...(source.showValue === true ? { showValue: true } : {}),
    ...(source.showCategoryName === true ? { showCategoryName: true } : {}),
    ...(source.showSeriesName === true ? { showSeriesName: true } : {}),
    ...(source.showPercentage === true &&
    (chartType === 'pie' || chartType === 'doughnut')
      ? { showPercentage: true }
      : {}),
    ...(source.showBubbleSize === true && chartType === 'bubble'
      ? { showBubbleSize: true }
      : {}),
    ...(separator !== undefined ? { separator } : {}),
    ...(position ? { position } : {}),
  };
}

export function workSpreadsheetDataLabelPositionLabel(
  position: WorkSpreadsheetDataLabelPosition,
): string {
  if (position === 'center') return '居中';
  if (position === 'insideBase') return '内侧基部';
  if (position === 'insideEnd') return '内侧末端';
  if (position === 'outsideEnd') return '外侧末端';
  if (position === 'left') return '左侧';
  if (position === 'right') return '右侧';
  if (position === 'above') return '上方';
  if (position === 'below') return '下方';
  return '最佳匹配';
}

export function normalizeWorkSpreadsheetTrendlineType(
  value: unknown,
): WorkSpreadsheetTrendlineType {
  return value === 'exponential' ||
    value === 'logarithmic' ||
    value === 'polynomial' ||
    value === 'power' ||
    value === 'movingAverage' ||
    value === 'linear'
    ? value
    : 'linear';
}

export function normalizeWorkSpreadsheetTrendline(
  trendline: WorkSpreadsheetTrendline,
): WorkSpreadsheetTrendline {
  const type = normalizeWorkSpreadsheetTrendlineType(trendline.type);
  const name = trendline.name?.trim().slice(0, 255);
  const order = normalizedInteger(trendline.order, 2, 6, 2);
  const period = normalizedInteger(trendline.period, 2, 255, 2);
  const forward = normalizedNonNegativeNumber(trendline.forward);
  const backward = normalizedNonNegativeNumber(trendline.backward);
  const intercept = Number.isFinite(trendline.intercept)
    ? Number(trendline.intercept)
    : undefined;
  return {
    type,
    ...(name ? { name } : {}),
    ...(type === 'polynomial' ? { order } : {}),
    ...(type === 'movingAverage' ? { period } : {}),
    ...(forward > 0 ? { forward } : {}),
    ...(backward > 0 ? { backward } : {}),
    ...(intercept !== undefined ? { intercept } : {}),
    ...(trendline.displayEquation === true ? { displayEquation: true } : {}),
    ...(trendline.displayRSquared === true ? { displayRSquared: true } : {}),
  };
}

export function workSpreadsheetTrendlineTypeLabel(
  type: WorkSpreadsheetTrendlineType,
): string {
  if (type === 'exponential') return '指数';
  if (type === 'logarithmic') return '对数';
  if (type === 'polynomial') return '多项式';
  if (type === 'power') return '幂';
  if (type === 'movingAverage') return '移动平均';
  return '线性';
}

function normalizedInteger(
  value: unknown,
  minimum: number,
  maximum: number,
  fallback: number,
): number {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(maximum, Math.max(minimum, Math.round(number)));
}

function normalizedNonNegativeNumber(value: unknown): number {
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, number) : 0;
}

export function workSpreadsheetChartTypeLabel(
  type: WorkSpreadsheetChartType,
): string {
  if (type === 'bar') return '条形图';
  if (type === 'line') return '折线图';
  if (type === 'pie') return '饼图';
  if (type === 'doughnut') return '圆环图';
  if (type === 'area') return '面积图';
  if (type === 'radar') return '雷达图';
  if (type === 'scatter') return '散点图';
  if (type === 'bubble') return '气泡图';
  if (type === 'combination') return '组合图';
  return '柱形图';
}

export interface WorkSpreadsheetChartSeries {
  name: string;
  values: number[];
  nameReference?: string;
  valuesReference?: string;
  xValues?: number[];
  xValuesReference?: string;
  bubbleSizes?: number[];
  bubbleSizesReference?: string;
  chartType?: WorkSpreadsheetCombinationSeriesType;
  axisGroup?: WorkSpreadsheetChartAxisGroup;
  dataLabels?: WorkSpreadsheetDataLabels;
  errorBars?: WorkSpreadsheetErrorBars[];
  trendlines?: WorkSpreadsheetTrendline[];
  style?: WorkSpreadsheetChartSeriesStyle;
}

export interface WorkSpreadsheetChart extends WorkSpreadsheetChartLayout {
  id: string;
  name: string;
  altText?: string;
  type: WorkSpreadsheetChartType;
  title?: string;
  titleReference?: string;
  axes?: WorkSpreadsheetChartAxes;
  categories: string[];
  categoryReference?: string;
  series: WorkSpreadsheetChartSeries[];
  showLegend: boolean;
  doughnutHoleSize?: number;
  radarStyle?: WorkSpreadsheetRadarStyle;
  scatterStyle?: WorkSpreadsheetScatterStyle;
  bubbleScale?: number;
  showNegativeBubbles?: boolean;
  bubbleSizeRepresents?: WorkSpreadsheetBubbleSizeRepresents;
  left: number;
  top: number;
  width: number;
  height: number;
}
