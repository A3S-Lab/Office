import { normalizeWorkSpreadsheetChartAxes } from '../work-spreadsheet-chart-axis';
import { normalizeWorkSpreadsheetChartLayout } from '../work-spreadsheet-chart-layout';
import { parseSpreadsheetChartReference } from '../work-spreadsheet-charts';
import {
  normalizeWorkSpreadsheetBubbleScale,
  normalizeWorkSpreadsheetBubbleSizeRepresents,
  normalizeWorkSpreadsheetChartAxisGroup,
  normalizeWorkSpreadsheetCombinationSeriesType,
  normalizeWorkSpreadsheetDoughnutHoleSize,
  normalizeWorkSpreadsheetErrorBars,
  normalizeWorkSpreadsheetRadarStyle,
  normalizeWorkSpreadsheetScatterStyle,
  type WorkSpreadsheetChart,
  type WorkSpreadsheetChartSeries,
  type WorkSpreadsheetChartType,
  type WorkSpreadsheetCombinationSeriesType,
  type WorkSpreadsheetContent,
  workSpreadsheetChartSupportsErrorBars,
  workSpreadsheetChartUsesNumericXAxis,
} from '../work-types';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';

export interface ChartListItem {
  sheetId: string;
  sheetName: string;
  chart: WorkSpreadsheetChart;
}

export interface ChartDraft extends Omit<WorkSpreadsheetChart, 'series'> {
  sheetId: string;
  series: WorkSpreadsheetChartSeries[];
}

export function chartDraft(item: ChartListItem): ChartDraft {
  return {
    ...item.chart,
    sheetId: item.sheetId,
    axes: item.chart.axes
      ? {
          bottom: item.chart.axes.bottom
            ? { ...item.chart.axes.bottom }
            : undefined,
          left: item.chart.axes.left ? { ...item.chart.axes.left } : undefined,
          top: item.chart.axes.top ? { ...item.chart.axes.top } : undefined,
          right: item.chart.axes.right
            ? { ...item.chart.axes.right }
            : undefined,
        }
      : undefined,
    series: item.chart.series.map((series) => ({
      ...series,
      values: [...series.values],
      xValues: series.xValues ? [...series.xValues] : undefined,
      bubbleSizes: series.bubbleSizes ? [...series.bubbleSizes] : undefined,
      dataLabels: series.dataLabels ? { ...series.dataLabels } : undefined,
      errorBars: series.errorBars?.map((errorBars) => ({
        ...errorBars,
        plusValues: errorBars.plusValues
          ? [...errorBars.plusValues]
          : undefined,
        minusValues: errorBars.minusValues
          ? [...errorBars.minusValues]
          : undefined,
      })),
      trendlines: series.trendlines?.map((trendline) => ({ ...trendline })),
      style: series.style
        ? {
            ...series.style,
            marker: series.style.marker
              ? { ...series.style.marker }
              : undefined,
          }
        : undefined,
    })),
  };
}

export function chartDraftWithType(
  draft: ChartDraft,
  type: WorkSpreadsheetChartType,
): ChartDraft {
  const numericXAxis = workSpreadsheetChartUsesNumericXAxis(type);
  const numericCategories = draft.categories.map(strictNumericCategory);
  const categoriesAreNumeric =
    numericCategories.length > 0 &&
    numericCategories.every((value) => value !== null);
  const typedSeries = numericXAxis
    ? draft.series.map((item) => {
        if (item.xValuesReference?.trim() || item.xValues?.length) return item;
        return {
          ...item,
          xValues: item.values.map(
            (_, index) => numericCategories[index] ?? index + 1,
          ),
          xValuesReference: categoriesAreNumeric
            ? draft.categoryReference
            : undefined,
        };
      })
    : type === 'combination'
      ? draft.series.map((item, index) => ({
          ...item,
          chartType: item.chartType
            ? normalizeWorkSpreadsheetCombinationSeriesType(item.chartType)
            : combinationType(index),
          axisGroup: item.axisGroup
            ? normalizeWorkSpreadsheetChartAxisGroup(item.axisGroup)
            : index === 0
              ? 'primary'
              : 'secondary',
        }))
      : draft.series;
  const supportsErrorBars = workSpreadsheetChartSupportsErrorBars(type);
  const series = typedSeries.map((item) => {
    if (!supportsErrorBars) return { ...item, errorBars: undefined };
    const errorBars = item.errorBars
      ?.filter((source) => numericXAxis || source.direction !== 'x')
      .map((source) => normalizeWorkSpreadsheetErrorBars(source, type));
    return { ...item, errorBars: errorBars?.length ? errorBars : undefined };
  });
  const hasSecondaryAxes =
    type === 'combination' &&
    series.some(
      (item) =>
        normalizeWorkSpreadsheetChartAxisGroup(item.axisGroup) === 'secondary',
    );
  return {
    ...draft,
    type,
    series,
    axes: normalizeWorkSpreadsheetChartAxes(draft.axes, type, hasSecondaryAxes),
    ...normalizeWorkSpreadsheetChartLayout({ ...draft, type }),
    doughnutHoleSize:
      type === 'doughnut'
        ? normalizeWorkSpreadsheetDoughnutHoleSize(draft.doughnutHoleSize)
        : draft.doughnutHoleSize,
    radarStyle:
      type === 'radar'
        ? normalizeWorkSpreadsheetRadarStyle(draft.radarStyle)
        : draft.radarStyle,
    scatterStyle:
      type === 'scatter'
        ? normalizeWorkSpreadsheetScatterStyle(draft.scatterStyle)
        : draft.scatterStyle,
    bubbleScale:
      type === 'bubble'
        ? normalizeWorkSpreadsheetBubbleScale(draft.bubbleScale)
        : draft.bubbleScale,
    showNegativeBubbles:
      type === 'bubble'
        ? draft.showNegativeBubbles === true
        : draft.showNegativeBubbles,
    bubbleSizeRepresents:
      type === 'bubble'
        ? normalizeWorkSpreadsheetBubbleSizeRepresents(
            draft.bubbleSizeRepresents,
          )
        : draft.bubbleSizeRepresents,
  };
}

function strictNumericCategory(value: string): number | null {
  const text = value.trim();
  if (!text) return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

export function newChartSeries(
  type: WorkSpreadsheetChartType,
  index: number,
): WorkSpreadsheetChartSeries {
  const catalog = resolveOfficeMessages();
  return {
    name: officeMessage(catalog, 'spreadsheet.chartDraft.seriesName', {
      index: String(index + 1),
    }),
    values: [],
    valuesReference: '',
    ...(workSpreadsheetChartUsesNumericXAxis(type)
      ? { xValues: [], xValuesReference: '' }
      : {}),
    ...(type === 'bubble' ? { bubbleSizes: [], bubbleSizesReference: '' } : {}),
    ...(type === 'combination'
      ? {
          chartType: combinationType(index),
          axisGroup: index === 0 ? 'primary' : 'secondary',
        }
      : {}),
  };
}

function combinationType(index: number): WorkSpreadsheetCombinationSeriesType {
  return index === 0 ? 'column' : 'line';
}

export function chartKey(item: ChartListItem | undefined): string | null {
  return item ? `${item.sheetId}:${item.chart.id}` : null;
}

export function replaceSeries(
  series: WorkSpreadsheetChartSeries[],
  index: number,
  change: Partial<WorkSpreadsheetChartSeries>,
): WorkSpreadsheetChartSeries[] {
  return series.map((item, candidate) =>
    candidate === index ? { ...item, ...change } : item,
  );
}

export function validateSeriesErrorBars(
  content: WorkSpreadsheetContent,
  ownerSheet: WorkSpreadsheetContent['sheets'][number],
  series: WorkSpreadsheetChartSeries,
  seriesIndex: number,
  chartType: WorkSpreadsheetChartType,
): string | null {
  const directions = new Set<string>();
  for (const [errorBarIndex, errorBars] of (series.errorBars ?? []).entries()) {
    const catalog = resolveOfficeMessages();
    const prefix = officeMessage(catalog, 'spreadsheet.chartDraft.errorBar.prefix', {
      series: String(seriesIndex + 1),
      errorBar: String(errorBarIndex + 1),
    });
    if (
      !workSpreadsheetChartUsesNumericXAxis(chartType) &&
      errorBars.direction === 'x'
    ) {
      return officeMessage(catalog, 'spreadsheet.chartDraft.errorBar.xNotAllowed', {
        prefix,
      });
    }
    if (directions.has(errorBars.direction)) {
      return officeMessage(
        catalog,
        'spreadsheet.chartDraft.errorBar.duplicateDirection',
        { series: String(seriesIndex + 1) },
      );
    }
    directions.add(errorBars.direction);

    if (
      (errorBars.valueType === 'fixedValue' ||
        errorBars.valueType === 'percentage' ||
        errorBars.valueType === 'standardDeviation') &&
      (typeof errorBars.value !== 'number' ||
        !Number.isFinite(errorBars.value) ||
        errorBars.value < 0)
    ) {
      return officeMessage(catalog, 'spreadsheet.chartDraft.errorBar.nonNegative', {
        prefix,
      });
    }
    if (errorBars.valueType !== 'custom') continue;

    if (errorBars.barType !== 'minus') {
      const issue = validateCustomErrorBarSource(
        content,
        ownerSheet,
        errorBars.plusReference,
        errorBars.plusValues,
        officeMessage(catalog, 'spreadsheet.chartDraft.errorBar.plus', { prefix }),
      );
      if (issue) return issue;
    }
    if (errorBars.barType !== 'plus') {
      const issue = validateCustomErrorBarSource(
        content,
        ownerSheet,
        errorBars.minusReference,
        errorBars.minusValues,
        officeMessage(catalog, 'spreadsheet.chartDraft.errorBar.minus', { prefix }),
      );
      if (issue) return issue;
    }
  }
  return null;
}

export function validateChartAxes(
  content: WorkSpreadsheetContent,
  ownerSheet: WorkSpreadsheetContent['sheets'][number],
  axes: WorkSpreadsheetChart['axes'],
  hasSecondaryAxes: boolean,
): string | null {
  const catalog = resolveOfficeMessages();
  const entries = [
    [officeMessage(catalog, 'spreadsheet.chartDraft.axis.bottom'), axes?.bottom],
    [officeMessage(catalog, 'spreadsheet.chartDraft.axis.left'), axes?.left],
    ...(hasSecondaryAxes
      ? ([
          [officeMessage(catalog, 'spreadsheet.chartDraft.axis.top'), axes?.top],
          [officeMessage(catalog, 'spreadsheet.chartDraft.axis.right'), axes?.right],
        ] as const)
      : []),
  ] as const;
  for (const [label, axis] of entries) {
    if ((axis?.title?.length ?? 0) > 255)
      return officeMessage(catalog, 'spreadsheet.chartDraft.axis.titleTooLong', {
        label,
      });
    if (
      axis?.titleReference?.trim() &&
      !parseSpreadsheetChartReference(content, ownerSheet, axis.titleReference)
    ) {
      return officeMessage(catalog, 'spreadsheet.chartDraft.axis.titleRefInvalid', {
        label,
      });
    }
    if (axis?.minimum !== undefined && !Number.isFinite(axis.minimum))
      return officeMessage(catalog, 'spreadsheet.chartDraft.axis.minInvalid', {
        label,
      });
    if (axis?.maximum !== undefined && !Number.isFinite(axis.maximum))
      return officeMessage(catalog, 'spreadsheet.chartDraft.axis.maxInvalid', {
        label,
      });
    if (
      axis?.minimum !== undefined &&
      axis.maximum !== undefined &&
      axis.minimum >= axis.maximum
    ) {
      return officeMessage(catalog, 'spreadsheet.chartDraft.axis.minGteMax', {
        label,
      });
    }
    if (
      axis?.majorUnit !== undefined &&
      (!Number.isFinite(axis.majorUnit) || axis.majorUnit <= 0)
    ) {
      return officeMessage(catalog, 'spreadsheet.chartDraft.axis.majorUnitInvalid', {
        label,
      });
    }
    if (
      axis?.labelInterval !== undefined &&
      (!Number.isInteger(axis.labelInterval) ||
        axis.labelInterval < 1 ||
        axis.labelInterval > 31_999)
    ) {
      return officeMessage(
        catalog,
        'spreadsheet.chartDraft.axis.labelIntervalInvalid',
        { label },
      );
    }
    if ((axis?.numberFormat?.length ?? 0) > 255)
      return officeMessage(
        catalog,
        'spreadsheet.chartDraft.axis.numberFormatTooLong',
        { label },
      );
  }
  return null;
}

function validateCustomErrorBarSource(
  content: WorkSpreadsheetContent,
  ownerSheet: WorkSpreadsheetContent['sheets'][number],
  reference: string | undefined,
  values: number[] | undefined,
  label: string,
): string | null {
  if (reference?.trim()) {
    return parseSpreadsheetChartReference(content, ownerSheet, reference)
      ? null
      : officeMessage(resolveOfficeMessages(), 'spreadsheet.chartDraft.refInvalid', {
          label,
        });
  }
  if (!values?.length)
    return officeMessage(
      resolveOfficeMessages(),
      'spreadsheet.chartDraft.refRequired',
      { label },
    );
  return values.every((value) => Number.isFinite(value) && value >= 0)
    ? null
    : officeMessage(resolveOfficeMessages(), 'spreadsheet.chartDraft.cacheInvalid', {
        label,
      });
}
