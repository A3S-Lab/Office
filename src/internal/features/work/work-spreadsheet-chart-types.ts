import type { Image } from '@fortune-sheet/core';
import type { OfficeMessageCatalog } from '../../i18n/office-messages';
import { officeMessage, resolveOfficeMessages } from '../../i18n/office-locale';
import type { WorkSpreadsheetChartLayout } from './work-spreadsheet-chart-layout';
import type { WorkSpreadsheetImageTransform } from './work-xlsx-image-transform';

export type { WorkSpreadsheetImageTransform } from './work-xlsx-image-transform';

export interface WorkSpreadsheetImage extends Image {
  name?: string;
  altText?: string;
  contentType?: string;
  /** Bounded 90-degree / flip transform preserved on native XLSX round trips. */
  transform?: WorkSpreadsheetImageTransform;
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
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (type === 'line') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.line');
  }
  if (type === 'area') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.area');
  }
  return officeMessage(messages, 'spreadsheet.chart.typeShort.column');
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
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (type === 'percentage') {
    return officeMessage(
      messages,
      'spreadsheet.chart.errorBar.valueType.percentage',
    );
  }
  if (type === 'standardDeviation') {
    return officeMessage(
      messages,
      'spreadsheet.chart.errorBar.valueType.standardDeviation',
    );
  }
  if (type === 'standardError') {
    return officeMessage(
      messages,
      'spreadsheet.chart.errorBar.valueType.standardError',
    );
  }
  if (type === 'custom') {
    return officeMessage(
      messages,
      'spreadsheet.chart.errorBar.valueType.custom',
    );
  }
  return officeMessage(
    messages,
    'spreadsheet.chart.errorBar.valueType.fixedValue',
  );
}

export function workSpreadsheetErrorBarTypeLabel(
  type: WorkSpreadsheetErrorBarType,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (type === 'plus') {
    return officeMessage(messages, 'spreadsheet.chart.errorBar.barType.plus');
  }
  if (type === 'minus') {
    return officeMessage(messages, 'spreadsheet.chart.errorBar.barType.minus');
  }
  return officeMessage(messages, 'spreadsheet.chart.errorBar.barType.both');
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
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (position === 'center') {
    return officeMessage(
      messages,
      'spreadsheet.chart.dataLabel.position.center',
    );
  }
  if (position === 'insideBase') {
    return officeMessage(
      messages,
      'spreadsheet.chart.dataLabel.position.insideBase',
    );
  }
  if (position === 'insideEnd') {
    return officeMessage(
      messages,
      'spreadsheet.chart.dataLabel.position.insideEnd',
    );
  }
  if (position === 'outsideEnd') {
    return officeMessage(
      messages,
      'spreadsheet.chart.dataLabel.position.outsideEnd',
    );
  }
  if (position === 'left') {
    return officeMessage(messages, 'spreadsheet.chart.dataLabel.position.left');
  }
  if (position === 'right') {
    return officeMessage(messages, 'spreadsheet.chart.dataLabel.position.right');
  }
  if (position === 'above') {
    return officeMessage(messages, 'spreadsheet.chart.dataLabel.position.above');
  }
  if (position === 'below') {
    return officeMessage(messages, 'spreadsheet.chart.dataLabel.position.below');
  }
  return officeMessage(
    messages,
    'spreadsheet.chart.dataLabel.position.bestFit',
  );
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
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (type === 'exponential') {
    return officeMessage(messages, 'spreadsheet.chart.trendline.type.exponential');
  }
  if (type === 'logarithmic') {
    return officeMessage(messages, 'spreadsheet.chart.trendline.type.logarithmic');
  }
  if (type === 'polynomial') {
    return officeMessage(messages, 'spreadsheet.chart.trendline.type.polynomial');
  }
  if (type === 'power') {
    return officeMessage(messages, 'spreadsheet.chart.trendline.type.power');
  }
  if (type === 'movingAverage') {
    return officeMessage(
      messages,
      'spreadsheet.chart.trendline.type.movingAverage',
    );
  }
  return officeMessage(messages, 'spreadsheet.chart.trendline.type.linear');
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
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (type === 'bar') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.bar');
  }
  if (type === 'line') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.line');
  }
  if (type === 'pie') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.pie');
  }
  if (type === 'doughnut') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.doughnut');
  }
  if (type === 'area') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.area');
  }
  if (type === 'radar') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.radar');
  }
  if (type === 'scatter') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.scatter');
  }
  if (type === 'bubble') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.bubble');
  }
  if (type === 'combination') {
    return officeMessage(messages, 'spreadsheet.chart.typeShort.combination');
  }
  return officeMessage(messages, 'spreadsheet.chart.typeShort.column');
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
