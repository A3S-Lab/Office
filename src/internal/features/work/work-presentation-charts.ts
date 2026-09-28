import {
  officeMessage,
  resolveOfficeMessages,
  type OfficeMessageCatalog,
} from '../../i18n/office-locale';
import { presentationChartAxesForType } from './work-presentation-chart-axes';
import {
  normalizeWorkSpreadsheetChartLayout,
  type WorkSpreadsheetChartLayout,
  workSpreadsheetChartSupportsSeriesAnalysis,
} from './work-spreadsheet-chart-layout';
import { normalizeWorkSpreadsheetChartSeriesStyle } from './work-spreadsheet-chart-series-style';
import { createWorkId } from './work-templates';
import type {
  WorkSlideBubbleSizeRepresents,
  WorkSlideChart,
  WorkSlideChartDataLabelPosition,
  WorkSlideChartDataLabels,
  WorkSlideChartLegendPosition,
  WorkSlideChartSeries,
  WorkSlideChartSeriesStyle,
  WorkSlideChartType,
  WorkSlideElement,
  WorkSlideRadarStyle,
  WorkSlideScatterStyle,
} from './work-types';
import {
  normalizeWorkSpreadsheetErrorBars,
  normalizeWorkSpreadsheetTrendline,
  workSpreadsheetChartSupportsErrorBars,
  workSpreadsheetChartSupportsTrendlines,
} from './work-types';

const MAX_CHART_ITEMS = 256;
const chartNumberFormatter = new Intl.NumberFormat('zh-CN', {
  maximumFractionDigits: 6,
});
const chartPercentageFormatter = new Intl.NumberFormat('zh-CN', {
  maximumFractionDigits: 1,
});

export function createPresentationChartElement(
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): WorkSlideElement {
  return {
    id: createWorkId('element'),
    type: 'chart',
    x: 18,
    y: 20,
    width: 64,
    height: 52,
    text: '',
    fontSize: 14,
    color: '#172033',
    fill: '#ffffff',
    bold: false,
    align: 'center',
    borderColor: '#d9dee8',
    borderWidth: 1,
    altText: officeMessage(messages, 'presentation.chart.sample.altText'),
    chart: {
      type: 'column',
      title: officeMessage(messages, 'presentation.chart.sample.title'),
      categories: [
        officeMessage(messages, 'presentation.chart.sample.q1'),
        officeMessage(messages, 'presentation.chart.sample.q2'),
        officeMessage(messages, 'presentation.chart.sample.q3'),
      ],
      series: [
        {
          name: officeMessage(messages, 'presentation.chart.sample.series1'),
          values: [32, 48, 61],
        },
      ],
      showLegend: true,
      legendPosition: 'right',
    },
  };
}

export function presentationChartTypeLabel(
  type: WorkSlideChartType,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (type === 'bar') {
    return officeMessage(messages, 'presentation.chart.typeShort.bar');
  }
  if (type === 'line') {
    return officeMessage(messages, 'presentation.chart.typeShort.line');
  }
  if (type === 'pie') {
    return officeMessage(messages, 'presentation.chart.typeShort.pie');
  }
  if (type === 'doughnut') {
    return officeMessage(messages, 'presentation.chart.typeShort.doughnut');
  }
  if (type === 'area') {
    return officeMessage(messages, 'presentation.chart.typeShort.area');
  }
  if (type === 'radar') {
    return officeMessage(messages, 'presentation.chart.typeShort.radar');
  }
  if (type === 'scatter') {
    return officeMessage(messages, 'presentation.chart.typeShort.scatter');
  }
  if (type === 'bubble') {
    return officeMessage(messages, 'presentation.chart.typeShort.bubble');
  }
  return officeMessage(messages, 'presentation.chart.typeShort.column');
}

export function withPresentationChartType(
  chart: WorkSlideChart,
  type: WorkSlideChartType,
): WorkSlideChart {
  const {
    axes,
    bubbleScale,
    bubbleSizeRepresents,
    categoryAxisTitle,
    dataLabels,
    doughnutHoleSize,
    radarStyle,
    scatterStyle,
    series,
    showNegativeBubbles,
    valueAxisTitle,
    ...base
  } = chart;
  const nextAxes = presentationChartAxesForType(
    { ...chart, axes, categoryAxisTitle, valueAxisTitle },
    type,
  );
  const nextDataLabels = dataLabels
    ? normalizePresentationChartDataLabels(dataLabels, type)
    : undefined;
  if (nextDataLabels && !presentationChartDataLabelsHaveContent(nextDataLabels))
    nextDataLabels.showValue = true;
  const nextCategories =
    presentationChartUsesNumericXAxis(type) &&
    !presentationChartUsesNumericXAxis(chart.type)
      ? chart.categories.map((_, index) => String(index + 1))
      : chart.categories;
  const nextSeries = series.map(
    ({ bubbleSizes, errorBars, trendlines, ...current }) => ({
      ...current,
      ...(type === 'bubble'
        ? {
            bubbleSizes:
              bubbleSizes?.map(finitePresentationChartNumber) ??
              current.values.map(() => 1),
          }
        : {}),
      ...(workSpreadsheetChartSupportsErrorBars(type) && errorBars?.length
        ? { errorBars: normalizePresentationChartErrorBars(errorBars, type) }
        : {}),
      ...(workSpreadsheetChartSupportsTrendlines(type) && trendlines?.length
        ? { trendlines: trendlines.map(normalizeWorkSpreadsheetTrendline) }
        : {}),
    }),
  );
  return withPresentationChartLayout({
    ...base,
    categories: nextCategories,
    series: nextSeries,
    type,
    ...(nextAxes ? { axes: nextAxes } : {}),
    ...(nextDataLabels ? { dataLabels: nextDataLabels } : {}),
    ...(type === 'doughnut'
      ? { doughnutHoleSize: normalizeDoughnutHoleSize(doughnutHoleSize) }
      : {}),
    ...(type === 'radar'
      ? { radarStyle: normalizeRadarStyle(radarStyle) }
      : {}),
    ...(type === 'scatter'
      ? { scatterStyle: normalizePresentationScatterStyle(scatterStyle) }
      : {}),
    ...(type === 'bubble'
      ? {
          bubbleScale: normalizePresentationBubbleScale(bubbleScale),
          showNegativeBubbles: showNegativeBubbles === true,
          bubbleSizeRepresents:
            normalizePresentationBubbleSizeRepresents(bubbleSizeRepresents),
        }
      : {}),
  });
}

export function withPresentationChartLayout(
  chart: WorkSlideChart,
  change: Partial<
    WorkSpreadsheetChartLayout & Pick<WorkSlideChart, 'showLegend'>
  > = {},
): WorkSlideChart {
  const source = { ...chart, ...change };
  const normalized = normalizeWorkSpreadsheetChartLayout(source);
  const supportsAnalysis = workSpreadsheetChartSupportsSeriesAnalysis({
    ...source,
    ...normalized,
  });
  const {
    gapWidth: _gapWidth,
    grouping: _grouping,
    legendOverlay: _legendOverlay,
    legendPosition: _legendPosition,
    overlap: _overlap,
    series: _series,
    smoothLines: _smoothLines,
    ...base
  } = source;
  const series = source.series.map((item) => {
    const { errorBars, style, trendlines, ...seriesBase } = item;
    const normalizedStyle = normalizePresentationChartSeriesStyle(
      style,
      source.type,
    );
    return {
      ...seriesBase,
      ...(normalizedStyle ? { style: normalizedStyle } : {}),
      ...(supportsAnalysis &&
      workSpreadsheetChartSupportsErrorBars(source.type) &&
      errorBars?.length
        ? {
            errorBars: normalizePresentationChartErrorBars(
              errorBars,
              source.type,
            ),
          }
        : {}),
      ...(supportsAnalysis &&
      workSpreadsheetChartSupportsTrendlines(source.type) &&
      trendlines?.length
        ? { trendlines: trendlines.map(normalizeWorkSpreadsheetTrendline) }
        : {}),
    };
  });
  return {
    ...base,
    series,
    ...(source.legendPosition !== undefined
      ? { legendPosition: normalized.legendPosition }
      : {}),
    ...(source.legendOverlay !== undefined
      ? { legendOverlay: normalized.legendOverlay }
      : {}),
    ...(normalized.grouping !== undefined && source.grouping !== undefined
      ? { grouping: normalized.grouping }
      : {}),
    ...(normalized.gapWidth !== undefined && source.gapWidth !== undefined
      ? { gapWidth: normalized.gapWidth }
      : {}),
    ...(normalized.overlap !== undefined && source.overlap !== undefined
      ? { overlap: normalized.overlap }
      : {}),
    ...(normalized.smoothLines !== undefined && source.smoothLines !== undefined
      ? { smoothLines: normalized.smoothLines }
      : {}),
  };
}

export function normalizePresentationChartSeriesStyle(
  style: WorkSlideChartSeriesStyle | undefined,
  type: WorkSlideChartType,
): WorkSlideChartSeriesStyle | undefined {
  const normalized = normalizeWorkSpreadsheetChartSeriesStyle(style);
  if (!normalized) return undefined;
  if (presentationChartSupportsSeriesMarkers(type)) return normalized;
  const { marker: _marker, ...withoutMarker } = normalized;
  return Object.keys(withoutMarker).length ? withoutMarker : undefined;
}

export function withPresentationChartSeriesStyle(
  chart: WorkSlideChart,
  seriesIndex: number,
  style: WorkSlideChartSeriesStyle | undefined,
): WorkSlideChart {
  return withPresentationChartLayout({
    ...chart,
    series: chart.series.map((series, index) =>
      index === seriesIndex
        ? {
            ...series,
            style: normalizePresentationChartSeriesStyle(style, chart.type),
          }
        : series,
    ),
  });
}

export function presentationChartSupportsSeriesMarkers(
  type: WorkSlideChartType,
): boolean {
  return type === 'line' || type === 'radar' || type === 'scatter';
}

export function presentationChartHasCustomSeriesStyles(
  chart: WorkSlideChart,
): boolean {
  return chart.series.some((series) =>
    Boolean(normalizePresentationChartSeriesStyle(series.style, chart.type)),
  );
}

export function parsePresentationChartCategories(value: string): string[] {
  return value
    .split(/\r?\n/)
    .map((item) => item.trim().slice(0, 255))
    .filter(Boolean)
    .slice(0, MAX_CHART_ITEMS);
}

export function parsePresentationChartValues(value: string): number[] {
  return value
    .split(/[\s,;，；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .slice(0, MAX_CHART_ITEMS)
    .map((item) => {
      const number = Number(item);
      return Number.isFinite(number) ? number : 0;
    });
}

export function parsePresentationChartValueDraft(
  value: string,
): number[] | null {
  const items = value
    .split(/[\s,;，；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .slice(0, MAX_CHART_ITEMS);
  const values = items.map(Number);
  return values.every(Number.isFinite) ? values : null;
}

export function parsePresentationChartXValues(value: string): string[] {
  return parsePresentationChartValues(value).map((item) => String(item));
}

export function createPresentationChartSeries(
  chart: WorkSlideChart,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): WorkSlideChartSeries {
  const values = Array.from(
    { length: Math.max(1, chart.categories.length) },
    () => 0,
  );
  return {
    name: officeMessage(messages, 'presentation.chart.seriesFallback', {
      index: String(chart.series.length + 1),
    }),
    values,
    ...(chart.type === 'bubble' ? { bubbleSizes: values.map(() => 1) } : {}),
  };
}

export function withPresentationChartSeriesAnalysis(
  chart: WorkSlideChart,
  seriesIndex: number,
  analysis: Pick<WorkSlideChartSeries, 'errorBars' | 'trendlines'>,
): WorkSlideChart {
  return withPresentationChartLayout({
    ...chart,
    series: chart.series.map((series, index) => {
      if (index !== seriesIndex) return series;
      const {
        errorBars: _errorBars,
        trendlines: _trendlines,
        ...base
      } = series;
      const errorBars = workSpreadsheetChartSupportsErrorBars(chart.type)
        ? normalizePresentationChartErrorBars(
            analysis.errorBars ?? [],
            chart.type,
          )
        : undefined;
      const trendlines = workSpreadsheetChartSupportsTrendlines(chart.type)
        ? analysis.trendlines?.map(normalizeWorkSpreadsheetTrendline)
        : undefined;
      return {
        ...base,
        ...(errorBars?.length ? { errorBars } : {}),
        ...(trendlines?.length ? { trendlines } : {}),
      };
    }),
  });
}

export function presentationChartTrendlineCount(chart: WorkSlideChart): number {
  if (
    !workSpreadsheetChartSupportsSeriesAnalysis(chart) ||
    !workSpreadsheetChartSupportsTrendlines(chart.type)
  ) {
    return 0;
  }
  return chart.series.reduce(
    (count, series) => count + (series.trendlines?.length ?? 0),
    0,
  );
}

export function presentationChartErrorBarCount(chart: WorkSlideChart): number {
  if (
    !workSpreadsheetChartSupportsSeriesAnalysis(chart) ||
    !workSpreadsheetChartSupportsErrorBars(chart.type)
  ) {
    return 0;
  }
  return chart.series.reduce(
    (count, series) => count + (series.errorBars?.length ?? 0),
    0,
  );
}

function normalizePresentationChartErrorBars(
  sources: NonNullable<WorkSlideChartSeries['errorBars']>,
  type: WorkSlideChartType,
): NonNullable<WorkSlideChartSeries['errorBars']> {
  const directions = new Set<string>();
  return sources.flatMap((source) => {
    const errorBars = normalizeWorkSpreadsheetErrorBars(source, type);
    if (directions.has(errorBars.direction)) return [];
    directions.add(errorBars.direction);
    return [errorBars];
  });
}

export function normalizeDoughnutHoleSize(value: number | undefined): number {
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(90, Math.max(10, Math.round(number)))
    : 50;
}

export function normalizeRadarStyle(
  value: WorkSlideRadarStyle | undefined,
): WorkSlideRadarStyle {
  return value === 'marker' || value === 'filled' ? value : 'standard';
}

export function normalizePresentationScatterStyle(
  value: unknown,
): WorkSlideScatterStyle {
  return value === 'marker' ||
    value === 'line' ||
    value === 'smooth' ||
    value === 'smoothMarker' ||
    value === 'lineMarker'
    ? value
    : 'lineMarker';
}

export function normalizePresentationBubbleScale(value: unknown): number {
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(300, Math.max(5, Math.round(number)))
    : 100;
}

export function normalizePresentationBubbleSizeRepresents(
  value: unknown,
): WorkSlideBubbleSizeRepresents {
  return value === 'width' ? 'width' : 'area';
}

export function presentationChartUsesNumericXAxis(
  type: WorkSlideChartType,
): boolean {
  return type === 'scatter' || type === 'bubble';
}

export function presentationChartXValues(chart: WorkSlideChart): number[] {
  const length = Math.max(
    chart.categories.length,
    ...chart.series.map((series) => series.values.length),
  );
  return Array.from({ length }, (_, index) => {
    const value = Number(chart.categories[index]);
    return Number.isFinite(value) ? value : index + 1;
  });
}

export function presentationChartBubbleSizes(
  series: WorkSlideChartSeries,
): number[] {
  return series.values.map((_, index) =>
    finitePresentationChartNumber(series.bubbleSizes?.[index] ?? 1),
  );
}

export function normalizePresentationChartLegendPosition(
  value: unknown,
): WorkSlideChartLegendPosition {
  return value === 'left' ||
    value === 'top' ||
    value === 'bottom' ||
    value === 'topRight' ||
    value === 'right'
    ? value
    : 'right';
}

export function presentationChartLegendPositionLabel(
  position: WorkSlideChartLegendPosition,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (position === 'left') {
    return officeMessage(messages, 'presentation.chart.legend.left');
  }
  if (position === 'top') {
    return officeMessage(messages, 'presentation.chart.legend.top');
  }
  if (position === 'bottom') {
    return officeMessage(messages, 'presentation.chart.legend.bottom');
  }
  if (position === 'topRight') {
    return officeMessage(messages, 'presentation.chart.legend.topRight');
  }
  return officeMessage(messages, 'presentation.chart.legend.right');
}

export function presentationChartShowsLegend(chart: WorkSlideChart): boolean {
  return chart.showLegend ?? chart.series.length > 1;
}

export function presentationChartSupportsAxisTitles(
  chart: Pick<WorkSlideChart, 'type'>,
): boolean {
  return chart.type !== 'pie' && chart.type !== 'doughnut';
}

export function presentationChartDataLabelPositions(
  type: WorkSlideChartType,
): readonly WorkSlideChartDataLabelPosition[] {
  if (type === 'pie' || type === 'doughnut')
    return ['bestFit', 'center', 'insideEnd', 'outsideEnd'];
  if (
    type === 'line' ||
    type === 'radar' ||
    type === 'scatter' ||
    type === 'bubble'
  ) {
    return ['above', 'below', 'left', 'right', 'center'];
  }
  return ['outsideEnd', 'insideEnd', 'insideBase', 'center'];
}

export function normalizePresentationChartDataLabelPosition(
  value: unknown,
  type: WorkSlideChartType,
): WorkSlideChartDataLabelPosition {
  const positions = presentationChartDataLabelPositions(type);
  if (
    typeof value === 'string' &&
    positions.includes(value as WorkSlideChartDataLabelPosition)
  ) {
    return value as WorkSlideChartDataLabelPosition;
  }
  if (type === 'pie' || type === 'doughnut') return 'bestFit';
  if (
    type === 'line' ||
    type === 'radar' ||
    type === 'scatter' ||
    type === 'bubble'
  )
    return 'above';
  return 'outsideEnd';
}

export function normalizePresentationChartDataLabels(
  source: WorkSlideChartDataLabels,
  type: WorkSlideChartType,
): WorkSlideChartDataLabels {
  const separator =
    typeof source.separator === 'string'
      ? source.separator.slice(0, 64)
      : undefined;
  return {
    ...(source.showValue === true ? { showValue: true } : {}),
    ...(source.showCategoryName === true ? { showCategoryName: true } : {}),
    ...(source.showSeriesName === true ? { showSeriesName: true } : {}),
    ...(source.showPercentage === true &&
    (type === 'pie' || type === 'doughnut')
      ? { showPercentage: true }
      : {}),
    ...(source.showBubbleSize === true && type === 'bubble'
      ? { showBubbleSize: true }
      : {}),
    ...(separator !== undefined ? { separator } : {}),
    position: normalizePresentationChartDataLabelPosition(
      source.position,
      type,
    ),
  };
}

export function withPresentationChartDataLabels(
  chart: WorkSlideChart,
  dataLabels: WorkSlideChartDataLabels | undefined,
): WorkSlideChart {
  const { dataLabels: _current, ...base } = chart;
  return dataLabels
    ? {
        ...base,
        dataLabels: normalizePresentationChartDataLabels(
          dataLabels,
          chart.type,
        ),
      }
    : base;
}

export function presentationChartHasDataLabels(chart: WorkSlideChart): boolean {
  return chart.dataLabels !== undefined;
}

export function presentationChartDataLabelPositionLabel(
  position: WorkSlideChartDataLabelPosition,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  if (position === 'center') {
    return officeMessage(messages, 'presentation.chart.dataLabel.position.center');
  }
  if (position === 'insideBase') {
    return officeMessage(
      messages,
      'presentation.chart.dataLabel.position.insideBase',
    );
  }
  if (position === 'insideEnd') {
    return officeMessage(
      messages,
      'presentation.chart.dataLabel.position.insideEnd',
    );
  }
  if (position === 'outsideEnd') {
    return officeMessage(
      messages,
      'presentation.chart.dataLabel.position.outsideEnd',
    );
  }
  if (position === 'left') {
    return officeMessage(messages, 'presentation.chart.dataLabel.position.left');
  }
  if (position === 'right') {
    return officeMessage(messages, 'presentation.chart.dataLabel.position.right');
  }
  if (position === 'above') {
    return officeMessage(messages, 'presentation.chart.dataLabel.position.above');
  }
  if (position === 'below') {
    return officeMessage(messages, 'presentation.chart.dataLabel.position.below');
  }
  return officeMessage(
    messages,
    'presentation.chart.dataLabel.position.bestFit',
  );
}

export function presentationChartDataLabelText(
  chart: WorkSlideChart,
  seriesIndex: number,
  pointIndex: number,
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  const series = chart.series[seriesIndex];
  if (!series || !chart.dataLabels) return '';
  const labels = normalizePresentationChartDataLabels(
    chart.dataLabels,
    chart.type,
  );
  const parts: string[] = [];
  if (labels.showSeriesName)
    parts.push(
      series.name.trim() ||
        officeMessage(messages, 'presentation.chart.seriesFallback', {
          index: String(seriesIndex + 1),
        }),
    );
  if (labels.showCategoryName)
    parts.push(
      chart.categories[pointIndex]?.trim() ||
        officeMessage(messages, 'presentation.chart.categoryFallback', {
          index: String(pointIndex + 1),
        }),
    );
  if (labels.showValue)
    parts.push(formatPresentationChartNumber(series.values[pointIndex]));
  if (labels.showPercentage) {
    const values = series.values.map((value) =>
      Math.max(0, finitePresentationChartNumber(value)),
    );
    const total = values.reduce((sum, value) => sum + value, 0);
    if (total > 0)
      parts.push(
        `${chartPercentageFormatter.format((values[pointIndex] / total) * 100)}%`,
      );
  }
  if (labels.showBubbleSize)
    parts.push(formatPresentationChartNumber(series.bubbleSizes?.[pointIndex]));
  return parts.join(labels.separator ?? ', ');
}

function presentationChartDataLabelsHaveContent(
  labels: WorkSlideChartDataLabels,
): boolean {
  return Boolean(
    labels.showValue ||
      labels.showCategoryName ||
      labels.showSeriesName ||
      labels.showPercentage ||
      labels.showBubbleSize,
  );
}

function formatPresentationChartNumber(value: number | undefined): string {
  return chartNumberFormatter.format(finitePresentationChartNumber(value));
}

function finitePresentationChartNumber(value: number | undefined): number {
  return Number.isFinite(value) ? Number(value) : 0;
}
