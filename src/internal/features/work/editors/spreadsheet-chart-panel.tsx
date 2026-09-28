import type { Selection } from '@fortune-sheet/core';
import { Plus, Trash2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  Button,
  CollectionState,
  InlineNotice,
  StateView,
} from '../../../design-system/primitives';
import { normalizeWorkSpreadsheetChartAxes } from '../work-spreadsheet-chart-axis';
import {
  normalizeWorkSpreadsheetChartLayout,
  workSpreadsheetChartSupportsSeriesAnalysis,
} from '../work-spreadsheet-chart-layout';
import { normalizeWorkSpreadsheetChartSeriesStyle } from '../work-spreadsheet-chart-series-style';
import { validateSpreadsheetChartSeriesTrendlines } from '../work-spreadsheet-chart-validation';
import {
  createSpreadsheetChartFromSelection,
  parseSpreadsheetChartReference,
} from '../work-spreadsheet-charts';
import {
  normalizeWorkSpreadsheetBubbleScale,
  normalizeWorkSpreadsheetBubbleSizeRepresents,
  normalizeWorkSpreadsheetChartAxisGroup,
  normalizeWorkSpreadsheetCombinationSeriesType,
  normalizeWorkSpreadsheetDataLabels,
  normalizeWorkSpreadsheetDoughnutHoleSize,
  normalizeWorkSpreadsheetErrorBars,
  normalizeWorkSpreadsheetRadarStyle,
  normalizeWorkSpreadsheetScatterStyle,
  normalizeWorkSpreadsheetTrendline,
  type WorkSpreadsheetBubbleSizeRepresents,
  type WorkSpreadsheetChart,
  type WorkSpreadsheetChartAxisGroup,
  type WorkSpreadsheetChartType,
  type WorkSpreadsheetCombinationSeriesType,
  type WorkSpreadsheetContent,
  type WorkSpreadsheetRadarStyle,
  type WorkSpreadsheetScatterStyle,
  workSpreadsheetChartSupportsAxes,
  workSpreadsheetChartSupportsErrorBars,
  workSpreadsheetChartSupportsTrendlines,
  workSpreadsheetChartTypeLabel,
  workSpreadsheetChartUsesNumericXAxis,
} from '../work-types';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { normalizeRequiredOfficeNumber } from './office-number-normalization';
import { SpreadsheetChartAxisEditor } from './spreadsheet-chart-axis-editor';
import {
  type ChartDraft,
  type ChartListItem,
  chartDraft,
  chartDraftWithType,
  chartKey,
  newChartSeries,
  replaceSeries,
  validateChartAxes,
  validateSeriesErrorBars,
} from './spreadsheet-chart-draft';
import { SpreadsheetChartLayoutEditor } from './spreadsheet-chart-layout-editor';
import { SpreadsheetChartSeriesStyleEditor } from './spreadsheet-chart-series-style-editor';
import { SpreadsheetDataLabelEditor } from './spreadsheet-data-label-editor';
import { SpreadsheetErrorBarEditor } from './spreadsheet-error-bar-editor';
import { SpreadsheetTrendlineEditor } from './spreadsheet-trendline-editor';
import { useOfficeDraft } from './use-office-draft';

interface SpreadsheetChartPanelProps {
  content: WorkSpreadsheetContent;
  activeSheetId: string;
  selection?: Selection;
  onChange: (content: WorkSpreadsheetContent) => void;
}

export function SpreadsheetChartPanel({
  content,
  activeSheetId,
  selection,
  onChange,
}: SpreadsheetChartPanelProps) {
  const messages = useOfficeMessages();
  const items = useMemo(
    () =>
      content.sheets.flatMap((sheet) =>
        (sheet.charts ?? []).flatMap((chart) =>
          sheet.id ? [{ sheetId: sheet.id, sheetName: sheet.name, chart }] : [],
        ),
      ),
    [content.sheets],
  );
  const [selectedKey, setSelectedKey] = useState<string | null>(() =>
    chartKey(items[0]),
  );
  const {
    cancelDraft: resetDraft,
    dirty,
    draft,
    draftRef,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft<ChartDraft | null>(() =>
    items[0] ? chartDraft(items[0]) : null,
  );
  const [error, setError] = useState('');
  const pendingCreatedKeyRef = useRef<string | null>(null);

  useEffect(() => {
    const current = items.find((item) => chartKey(item) === selectedKey);
    if (current) {
      const next = chartDraft(current);
      syncDraft(next, chartDraftKey(draftRef.current) !== selectedKey);
      if (pendingCreatedKeyRef.current === selectedKey) {
        pendingCreatedKeyRef.current = null;
      }
      return;
    }
    if (!selectedKey || pendingCreatedKeyRef.current === selectedKey) return;
    const first = items[0];
    const next = first ? chartDraft(first) : null;
    setSelectedKey(chartKey(first));
    replaceDraft(next);
  }, [draftRef, items, replaceDraft, syncDraft]);

  const selectChart = (item: ChartListItem) => {
    const nextKey = chartKey(item);
    if (nextKey === selectedKey) return;
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.unsaved'));
      return;
    }
    const next = chartDraft(item);
    pendingCreatedKeyRef.current = null;
    setSelectedKey(chartKey(item));
    replaceDraft(next);
    setError('');
  };
  const addChart = () => {
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.unsaved'));
      return;
    }
    const sheet =
      content.sheets.find((candidate) => candidate.id === activeSheetId) ??
      content.sheets.find((candidate) => !candidate.hide) ??
      content.sheets[0];
    if (!sheet?.id) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.noSheet'));
      return;
    }
    const fallbackSelection: Selection = {
      row: [0, Math.max(0, Math.min(4, (sheet.row ?? 5) - 1))],
      column: [0, Math.max(0, Math.min(1, (sheet.column ?? 2) - 1))],
    };
    const chart = createSpreadsheetChartFromSelection(
      content,
      sheet.id,
      selection ?? fallbackSelection,
    );
    if (!chart) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.needSelection'));
      return;
    }
    const next = content.sheets.map((candidate) =>
      candidate.id === sheet.id
        ? { ...candidate, charts: [...(candidate.charts ?? []), chart] }
        : candidate,
    );
    onChange({ ...content, sheets: next });
    const nextKey = `${sheet.id}:${chart.id}`;
    const nextDraft: ChartDraft = {
      ...chart,
      sheetId: sheet.id,
      series: chart.series.map((item) => ({ ...item })),
    };
    pendingCreatedKeyRef.current = nextKey;
    setSelectedKey(nextKey);
    replaceDraft(nextDraft);
    setError('');
  };
  const saveChart = () => {
    if (!draft) return;
    const ownerSheet = content.sheets.find(
      (sheet) => sheet.id === draft.sheetId,
    );
    if (!ownerSheet) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.sheetMissing'));
      return;
    }
    if (!draft.name.trim()) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.needName'));
      return;
    }
    if (
      !workSpreadsheetChartUsesNumericXAxis(draft.type) &&
      draft.categoryReference?.trim() &&
      !parseSpreadsheetChartReference(
        content,
        ownerSheet,
        draft.categoryReference,
      )
    ) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.categoryRange'));
      return;
    }
    if (!draft.series.length) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.needSeries'));
      return;
    }
    if (draft.type === 'pie' && draft.series.length > 1) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.pieOneSeries'));
      return;
    }
    if (draft.type === 'combination' && draft.series.length < 2) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.combinationMinSeries'));
      return;
    }
    if (
      draft.type === 'doughnut' &&
      (typeof draft.doughnutHoleSize !== 'number' ||
        !Number.isFinite(draft.doughnutHoleSize) ||
        (draft.doughnutHoleSize ?? 0) < 10 ||
        (draft.doughnutHoleSize ?? 0) > 90)
    ) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.doughnutHole'));
      return;
    }
    if (
      draft.type === 'bubble' &&
      (typeof draft.bubbleScale !== 'number' ||
        !Number.isFinite(draft.bubbleScale) ||
        draft.bubbleScale < 0 ||
        draft.bubbleScale > 300)
    ) {
      setError(officeMessage(messages, 'spreadsheet.chart.error.bubbleScale'));
      return;
    }
    const supportsSeriesAnalysis =
      workSpreadsheetChartSupportsSeriesAnalysis(draft);
    const hasSecondaryAxes =
      draft.type === 'combination' &&
      draft.series.some(
        (series) =>
          normalizeWorkSpreadsheetChartAxisGroup(series.axisGroup) ===
          'secondary',
      );
    if (workSpreadsheetChartSupportsAxes(draft.type)) {
      const axisError = validateChartAxes(
        content,
        ownerSheet,
        draft.axes,
        hasSecondaryAxes,
      );
      if (axisError) {
        setError(axisError);
        return;
      }
    }
    for (const [index, series] of draft.series.entries()) {
      if (
        series.nameReference?.trim() &&
        !parseSpreadsheetChartReference(
          content,
          ownerSheet,
          series.nameReference,
        )
      ) {
        setError(officeMessage(messages, 'spreadsheet.chart.error.seriesNameRef', { n: String(index + 1) }));
        return;
      }
      if (
        series.valuesReference?.trim() &&
        !parseSpreadsheetChartReference(
          content,
          ownerSheet,
          series.valuesReference,
        )
      ) {
        setError(officeMessage(messages, 'spreadsheet.chart.error.seriesValuesRef', { n: String(index + 1) }));
        return;
      }
      if (!series.valuesReference?.trim() && !series.values.length) {
        setError(officeMessage(messages, 'spreadsheet.chart.error.seriesValuesRequired', { n: String(index + 1) }));
        return;
      }
      if (workSpreadsheetChartUsesNumericXAxis(draft.type)) {
        if (
          series.xValuesReference?.trim() &&
          !parseSpreadsheetChartReference(
            content,
            ownerSheet,
            series.xValuesReference,
          )
        ) {
          setError(officeMessage(messages, 'spreadsheet.chart.error.seriesXRef', { n: String(index + 1) }));
          return;
        }
        if (!series.xValuesReference?.trim() && !series.xValues?.length) {
          setError(officeMessage(messages, 'spreadsheet.chart.error.seriesXRequired', { n: String(index + 1) }));
          return;
        }
      }
      if (draft.type === 'bubble') {
        if (
          series.bubbleSizesReference?.trim() &&
          !parseSpreadsheetChartReference(
            content,
            ownerSheet,
            series.bubbleSizesReference,
          )
        ) {
          setError(officeMessage(messages, 'spreadsheet.chart.error.seriesBubbleRef', { n: String(index + 1) }));
          return;
        }
        if (
          !series.bubbleSizesReference?.trim() &&
          !series.bubbleSizes?.length
        ) {
          setError(officeMessage(messages, 'spreadsheet.chart.error.seriesBubbleRequired', { n: String(index + 1) }));
          return;
        }
      }
      if (
        supportsSeriesAnalysis &&
        workSpreadsheetChartSupportsErrorBars(draft.type)
      ) {
        const errorBarError = validateSeriesErrorBars(
          content,
          ownerSheet,
          series,
          index,
          draft.type,
        );
        if (errorBarError) {
          setError(errorBarError);
          return;
        }
      }
      if (
        supportsSeriesAnalysis &&
        workSpreadsheetChartSupportsTrendlines(draft.type)
      ) {
        const trendlineError = validateSpreadsheetChartSeriesTrendlines(
          series,
          index,
        );
        if (trendlineError) {
          setError(trendlineError);
          return;
        }
      }
    }
    const numericXAxis = workSpreadsheetChartUsesNumericXAxis(draft.type);
    const supportsErrorBars =
      supportsSeriesAnalysis &&
      workSpreadsheetChartSupportsErrorBars(draft.type);
    const supportsTrendlines =
      supportsSeriesAnalysis &&
      workSpreadsheetChartSupportsTrendlines(draft.type);
    const saved: WorkSpreadsheetChart = {
      ...draft,
      name: draft.name.trim(),
      altText: draft.altText?.trim() || undefined,
      title: draft.title?.trim() || undefined,
      titleReference:
        draft.titleReference?.trim().replace(/^=/, '') || undefined,
      axes: normalizeWorkSpreadsheetChartAxes(
        draft.axes,
        draft.type,
        hasSecondaryAxes,
      ),
      ...normalizeWorkSpreadsheetChartLayout(draft),
      categoryReference: numericXAxis
        ? undefined
        : draft.categoryReference?.trim().replace(/^=/, '') || undefined,
      doughnutHoleSize:
        draft.type === 'doughnut'
          ? normalizeWorkSpreadsheetDoughnutHoleSize(draft.doughnutHoleSize)
          : undefined,
      radarStyle:
        draft.type === 'radar'
          ? normalizeWorkSpreadsheetRadarStyle(draft.radarStyle)
          : undefined,
      scatterStyle:
        draft.type === 'scatter'
          ? normalizeWorkSpreadsheetScatterStyle(draft.scatterStyle)
          : undefined,
      bubbleScale:
        draft.type === 'bubble'
          ? normalizeWorkSpreadsheetBubbleScale(draft.bubbleScale)
          : undefined,
      showNegativeBubbles:
        draft.type === 'bubble'
          ? draft.showNegativeBubbles === true
          : undefined,
      bubbleSizeRepresents:
        draft.type === 'bubble'
          ? normalizeWorkSpreadsheetBubbleSizeRepresents(
              draft.bubbleSizeRepresents,
            )
          : undefined,
      series: draft.series.map((series, index) => {
        const {
          xValues,
          xValuesReference,
          bubbleSizes,
          bubbleSizesReference,
          chartType,
          axisGroup,
          dataLabels,
          errorBars,
          trendlines,
          style,
          ...categorySeries
        } = series;
        return {
          ...categorySeries,
          name: series.name.trim() || officeMessage(messages, 'spreadsheet.chart.seriesDefaultName', { n: String(index + 1) }),
          nameReference:
            series.nameReference?.trim().replace(/^=/, '') || undefined,
          valuesReference:
            series.valuesReference?.trim().replace(/^=/, '') || undefined,
          ...(numericXAxis
            ? {
                xValues,
                xValuesReference:
                  xValuesReference?.trim().replace(/^=/, '') || undefined,
              }
            : {}),
          ...(draft.type === 'bubble'
            ? {
                bubbleSizes,
                bubbleSizesReference:
                  bubbleSizesReference?.trim().replace(/^=/, '') || undefined,
              }
            : {}),
          ...(draft.type === 'combination'
            ? {
                chartType:
                  normalizeWorkSpreadsheetCombinationSeriesType(chartType),
                axisGroup: normalizeWorkSpreadsheetChartAxisGroup(axisGroup),
              }
            : {}),
          ...(dataLabels
            ? {
                dataLabels: normalizeWorkSpreadsheetDataLabels(
                  dataLabels,
                  draft.type,
                ),
              }
            : {}),
          ...(style
            ? { style: normalizeWorkSpreadsheetChartSeriesStyle(style) }
            : {}),
          ...(supportsErrorBars && errorBars?.length
            ? {
                errorBars: errorBars.map((item) =>
                  normalizeWorkSpreadsheetErrorBars(item, draft.type),
                ),
              }
            : {}),
          ...(supportsTrendlines && trendlines?.length
            ? { trendlines: trendlines.map(normalizeWorkSpreadsheetTrendline) }
            : {}),
        };
      }),
    };
    const sheets = content.sheets.map((sheet) =>
      sheet.id === draft.sheetId
        ? {
            ...sheet,
            charts: (sheet.charts ?? []).map((chart) =>
              chart.id === saved.id ? saved : chart,
            ),
          }
        : sheet,
    );
    onChange({ ...content, sheets });
    const savedDraft = chartDraft({
      sheetId: draft.sheetId,
      sheetName: ownerSheet.name,
      chart: saved,
    });
    replaceDraft(savedDraft);
    setError('');
  };
  const cancelDraft = () => {
    resetDraft();
    setError('');
  };
  const deleteChart = () => {
    if (!draft) return;
    const sheets = content.sheets.map((sheet) => {
      if (sheet.id !== draft.sheetId) return sheet;
      const charts = (sheet.charts ?? []).filter(
        (chart) => chart.id !== draft.id,
      );
      return { ...sheet, charts: charts.length ? charts : undefined };
    });
    onChange({ ...content, sheets });
    const next = items.find(
      (item) => chartKey(item) !== `${draft.sheetId}:${draft.id}`,
    );
    const nextDraft = next ? chartDraft(next) : null;
    pendingCreatedKeyRef.current = null;
    setSelectedKey(chartKey(next));
    replaceDraft(nextDraft);
    setError('');
  };

  return (
    <fieldset
      className="work-spreadsheet-chart-manager"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancelDraft();
      }}
    >
      <legend className="sr-only">{officeMessage(messages, 'spreadsheet.chart.legend')}</legend>
      <aside aria-label={officeMessage(messages, 'spreadsheet.chart.listAria')}>
        <Button className="create" tone="secondary" onClick={addChart}>
          <Plus size={13} />
          {officeMessage(messages, 'spreadsheet.chart.createFromSelection')}
        </Button>
        <div className="work-spreadsheet-chart-list">
          {items.map((item) => (
            <button
              type="button"
              className={chartKey(item) === selectedKey ? 'active' : ''}
              key={chartKey(item)}
              onClick={() => selectChart(item)}
            >
              <strong>{item.chart.title || item.chart.name}</strong>
              <span>
                {item.sheetName} ·{' '}
                {workSpreadsheetChartTypeLabel(item.chart.type, messages)}
              </span>
            </button>
          ))}
          {!items.length && (
            <CollectionState
              className="work-office-collection-empty"
              role="status"
            >
              {officeMessage(messages, 'spreadsheet.chart.emptyList')}
            </CollectionState>
          )}
        </div>
      </aside>
      {draft ? (
        <form
          onSubmit={(event) => {
            event.preventDefault();
            saveChart();
          }}
        >
          <div className="work-spreadsheet-chart-fields">
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.chart.objectName')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.chart.objectNameAria')}
                value={draft.name}
                maxLength={255}
                onChange={(event) =>
                  setDraft({ ...draft, name: event.target.value })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.chart.type')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'spreadsheet.chart.typeAria')}
                value={draft.type}
                options={[
                  { value: 'column', label: officeMessage(messages, 'spreadsheet.chart.type.column') },
                  { value: 'bar', label: officeMessage(messages, 'spreadsheet.chart.type.bar') },
                  { value: 'line', label: officeMessage(messages, 'spreadsheet.chart.type.line') },
                  { value: 'pie', label: officeMessage(messages, 'spreadsheet.chart.type.pie') },
                  { value: 'doughnut', label: officeMessage(messages, 'spreadsheet.chart.type.doughnut') },
                  { value: 'area', label: officeMessage(messages, 'spreadsheet.chart.type.area') },
                  { value: 'radar', label: officeMessage(messages, 'spreadsheet.chart.type.radar') },
                  { value: 'scatter', label: officeMessage(messages, 'spreadsheet.chart.type.scatter') },
                  { value: 'bubble', label: officeMessage(messages, 'spreadsheet.chart.type.bubble') },
                  { value: 'combination', label: officeMessage(messages, 'spreadsheet.chart.type.combination') },
                ]}
                onValueChange={(nextType) => {
                  const type = nextType as WorkSpreadsheetChartType;
                  setDraft(chartDraftWithType(draft, type));
                }}
              />
            </div>
            {draft.type === 'doughnut' && (
              <div className="work-office-field">
                <span>{officeMessage(messages, 'spreadsheet.chart.doughnutHole')}</span>
                <CommittedOfficeNumberField
                  ariaLabel={officeMessage(messages, 'spreadsheet.chart.doughnutHoleAria')}
                  min={10}
                  max={90}
                  step={1}
                  value={draft.doughnutHoleSize ?? 50}
                  normalizeValue={(value) =>
                    normalizeRequiredOfficeNumber(value, {
                      integer: true,
                      minimum: 10,
                      maximum: 90,
                    })
                  }
                  onValueCommit={(doughnutHoleSize) =>
                    setDraft({
                      ...draft,
                      doughnutHoleSize,
                    })
                  }
                />
              </div>
            )}
            {draft.type === 'radar' && (
              <div className="work-office-field">
                <span>{officeMessage(messages, 'spreadsheet.chart.radarStyle')}</span>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'spreadsheet.chart.radarStyleAria')}
                  value={normalizeWorkSpreadsheetRadarStyle(draft.radarStyle)}
                  options={[
                    { value: 'standard', label: officeMessage(messages, 'spreadsheet.chart.radar.standard') },
                    { value: 'marker', label: officeMessage(messages, 'spreadsheet.chart.radar.marker') },
                    { value: 'filled', label: officeMessage(messages, 'spreadsheet.chart.radar.filled') },
                  ]}
                  onValueChange={(radarStyle) =>
                    setDraft({
                      ...draft,
                      radarStyle: radarStyle as WorkSpreadsheetRadarStyle,
                    })
                  }
                />
              </div>
            )}
            {draft.type === 'scatter' && (
              <div className="work-office-field">
                <span>{officeMessage(messages, 'spreadsheet.chart.scatterStyle')}</span>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'spreadsheet.chart.scatterStyleAria')}
                  value={normalizeWorkSpreadsheetScatterStyle(
                    draft.scatterStyle,
                  )}
                  options={[
                    { value: 'marker', label: officeMessage(messages, 'spreadsheet.chart.scatter.marker') },
                    { value: 'line', label: officeMessage(messages, 'spreadsheet.chart.scatter.line') },
                    { value: 'lineMarker', label: officeMessage(messages, 'spreadsheet.chart.scatter.lineMarker') },
                    { value: 'smooth', label: officeMessage(messages, 'spreadsheet.chart.scatter.smooth') },
                    { value: 'smoothMarker', label: officeMessage(messages, 'spreadsheet.chart.scatter.smoothMarker') },
                  ]}
                  onValueChange={(scatterStyle) =>
                    setDraft({
                      ...draft,
                      scatterStyle: scatterStyle as WorkSpreadsheetScatterStyle,
                    })
                  }
                />
              </div>
            )}
            {draft.type === 'bubble' && (
              <>
                <div className="work-office-field">
                  <span>{officeMessage(messages, 'spreadsheet.chart.bubbleScale')}</span>
                  <CommittedOfficeNumberField
                    ariaLabel={officeMessage(messages, 'spreadsheet.chart.bubbleScaleAria')}
                    min={0}
                    max={300}
                    step={1}
                    value={draft.bubbleScale ?? 100}
                    normalizeValue={(value) =>
                      normalizeRequiredOfficeNumber(value, {
                        integer: true,
                        minimum: 0,
                        maximum: 300,
                      })
                    }
                    onValueCommit={(bubbleScale) =>
                      setDraft({ ...draft, bubbleScale })
                    }
                  />
                </div>
                <div className="work-office-field">
                  <span>{officeMessage(messages, 'spreadsheet.chart.bubbleSizeRepresents')}</span>
                  <OfficeSelect
                    ariaLabel={officeMessage(messages, 'spreadsheet.chart.bubbleSizeRepresentsAria')}
                    value={normalizeWorkSpreadsheetBubbleSizeRepresents(
                      draft.bubbleSizeRepresents,
                    )}
                    options={[
                      { value: 'area', label: officeMessage(messages, 'spreadsheet.chart.bubbleSize.area') },
                      { value: 'width', label: officeMessage(messages, 'spreadsheet.chart.bubbleSize.width') },
                    ]}
                    onValueChange={(bubbleSizeRepresents) =>
                      setDraft({
                        ...draft,
                        bubbleSizeRepresents:
                          bubbleSizeRepresents as WorkSpreadsheetBubbleSizeRepresents,
                      })
                    }
                  />
                </div>
                <OfficeCheckbox
                  className="check"
                  ariaLabel={officeMessage(messages, 'spreadsheet.chart.showNegativeBubblesAria')}
                  checked={draft.showNegativeBubbles === true}
                  onCheckedChange={(showNegativeBubbles) =>
                    setDraft({ ...draft, showNegativeBubbles })
                  }
                >
                  {officeMessage(messages, 'spreadsheet.chart.showNegativeBubbles')}
                </OfficeCheckbox>
              </>
            )}
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.chart.title')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.chart.titleAria')}
                value={draft.title ?? ''}
                maxLength={255}
                onChange={(event) =>
                  setDraft({ ...draft, title: event.target.value })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.chart.titleReference')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.chart.titleReferenceAria')}
                value={draft.titleReference ?? ''}
                placeholder={officeMessage(messages, 'spreadsheet.chart.placeholder.titleRef')}
                onChange={(event) =>
                  setDraft({ ...draft, titleReference: event.target.value })
                }
              />
            </div>
            {workSpreadsheetChartSupportsAxes(draft.type) && (
              <SpreadsheetChartAxisEditor
                axes={draft.axes}
                chartType={draft.type}
                showSecondaryAxes={
                  draft.type === 'combination' &&
                  draft.series.some(
                    (series) =>
                      normalizeWorkSpreadsheetChartAxisGroup(
                        series.axisGroup,
                      ) === 'secondary',
                  )
                }
                onChange={(axes) => setDraft({ ...draft, axes })}
              />
            )}
            {!workSpreadsheetChartUsesNumericXAxis(draft.type) && (
              <div className="work-office-field reference">
                <span>{officeMessage(messages, 'spreadsheet.chart.categories')}</span>
                <OfficeTextField
                  aria-label={officeMessage(messages, 'spreadsheet.chart.categoriesAria')}
                  value={draft.categoryReference ?? ''}
                  placeholder={officeMessage(messages, 'spreadsheet.chart.placeholder.categories')}
                  onChange={(event) =>
                    setDraft({
                      ...draft,
                      categoryReference: event.target.value,
                    })
                  }
                />
              </div>
            )}
            <div className="work-office-field alternative-text">
              <span>{officeMessage(messages, 'spreadsheet.chart.altText')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.chart.altTextAria')}
                value={draft.altText ?? ''}
                maxLength={1_024}
                placeholder={officeMessage(messages, 'spreadsheet.chart.altTextPlaceholder')}
                onChange={(event) =>
                  setDraft({ ...draft, altText: event.target.value })
                }
              />
            </div>
            <SpreadsheetChartLayoutEditor
              chart={draft}
              onChange={(change) => setDraft({ ...draft, ...change })}
            />
          </div>
          <section
            className="work-spreadsheet-chart-series"
            aria-label={officeMessage(messages, 'spreadsheet.chart.seriesAria')}
          >
            <header>
              <strong>{officeMessage(messages, 'spreadsheet.chart.seriesHeading')}</strong>
              <button
                type="button"
                onClick={() =>
                  setDraft({
                    ...draft,
                    series: [
                      ...draft.series,
                      newChartSeries(draft.type, draft.series.length),
                    ],
                  })
                }
              >
                <Plus size={12} />
                {officeMessage(messages, 'spreadsheet.chart.addSeries')}
              </button>
            </header>
            {workSpreadsheetChartUsesNumericXAxis(draft.type) && (
              <p className="xy-note">
                {officeMessage(messages, 'spreadsheet.chart.xyNote')}
              </p>
            )}
            {draft.series.map((series, index) => (
              <div
                className={`work-spreadsheet-chart-series-row${
                  workSpreadsheetChartUsesNumericXAxis(draft.type) ? ' xy' : ''
                }${draft.type === 'bubble' ? ' bubble' : ''}${draft.type === 'combination' ? ' combination' : ''}`}
                key={`${draft.id}-series-${index}`}
              >
                <div className="work-office-field series-name">
                  <span>{officeMessage(messages, 'spreadsheet.chart.seriesName', { n: String(index + 1) })}</span>
                  <OfficeTextField
                    aria-label={officeMessage(messages, 'spreadsheet.chart.seriesNameAria', { n: String(index + 1) })}
                    value={series.name}
                    onChange={(event) =>
                      setDraft({
                        ...draft,
                        series: replaceSeries(draft.series, index, {
                          name: event.target.value,
                        }),
                      })
                    }
                  />
                </div>
                <div className="work-office-field name-reference">
                  <span>{officeMessage(messages, 'spreadsheet.chart.nameReference')}</span>
                  <OfficeTextField
                    aria-label={officeMessage(messages, 'spreadsheet.chart.seriesNameRefAria', { n: String(index + 1) })}
                    value={series.nameReference ?? ''}
                    placeholder={officeMessage(messages, 'spreadsheet.chart.placeholder.seriesNameRef')}
                    onChange={(event) =>
                      setDraft({
                        ...draft,
                        series: replaceSeries(draft.series, index, {
                          nameReference: event.target.value,
                        }),
                      })
                    }
                  />
                </div>
                {workSpreadsheetChartUsesNumericXAxis(draft.type) && (
                  <div className="work-office-field x-reference">
                    <span>{officeMessage(messages, 'spreadsheet.chart.xValues')}</span>
                    <OfficeTextField
                      aria-label={officeMessage(messages, 'spreadsheet.chart.seriesXAria', { n: String(index + 1) })}
                      value={series.xValuesReference ?? ''}
                      placeholder={officeMessage(messages, 'spreadsheet.chart.placeholder.xValues')}
                      onChange={(event) =>
                        setDraft({
                          ...draft,
                          series: replaceSeries(draft.series, index, {
                            xValuesReference: event.target.value,
                          }),
                        })
                      }
                    />
                  </div>
                )}
                <div
                  className={`work-office-field ${
                    workSpreadsheetChartUsesNumericXAxis(draft.type)
                      ? 'y-reference'
                      : 'reference'
                  }`}
                >
                  <span>
                    {workSpreadsheetChartUsesNumericXAxis(draft.type)
                      ? officeMessage(messages, 'spreadsheet.chart.yValues')
                      : officeMessage(messages, 'spreadsheet.chart.values')}
                  </span>
                  <OfficeTextField
                    aria-label={officeMessage(
                      messages,
                      workSpreadsheetChartUsesNumericXAxis(draft.type)
                        ? 'spreadsheet.chart.seriesYAria'
                        : 'spreadsheet.chart.seriesValuesAria',
                      { n: String(index + 1) },
                    )}
                    value={series.valuesReference ?? ''}
                    placeholder={officeMessage(messages, 'spreadsheet.chart.placeholder.values')}
                    onChange={(event) =>
                      setDraft({
                        ...draft,
                        series: replaceSeries(draft.series, index, {
                          valuesReference: event.target.value,
                        }),
                      })
                    }
                  />
                </div>
                {draft.type === 'combination' && (
                  <>
                    <div className="work-office-field combination-chart-type">
                      <span>{officeMessage(messages, 'spreadsheet.chart.seriesChartType')}</span>
                      <OfficeSelect
                        ariaLabel={officeMessage(messages, 'spreadsheet.chart.seriesChartTypeAria', { n: String(index + 1) })}
                        value={normalizeWorkSpreadsheetCombinationSeriesType(
                          series.chartType,
                        )}
                        options={[
                          { value: 'column', label: officeMessage(messages, 'spreadsheet.chart.typeShort.column') },
                          { value: 'line', label: officeMessage(messages, 'spreadsheet.chart.typeShort.line') },
                          { value: 'area', label: officeMessage(messages, 'spreadsheet.chart.typeShort.area') },
                        ]}
                        onValueChange={(chartType) =>
                          setDraft({
                            ...draft,
                            series: replaceSeries(draft.series, index, {
                              chartType:
                                chartType as WorkSpreadsheetCombinationSeriesType,
                            }),
                          })
                        }
                      />
                    </div>
                    <div className="work-office-field combination-axis-group">
                      <span>{officeMessage(messages, 'spreadsheet.chart.axis')}</span>
                      <OfficeSelect
                        ariaLabel={officeMessage(messages, 'spreadsheet.chart.seriesAxisAria', { n: String(index + 1) })}
                        value={normalizeWorkSpreadsheetChartAxisGroup(
                          series.axisGroup,
                        )}
                        options={[
                          { value: 'primary', label: officeMessage(messages, 'spreadsheet.chart.axis.primary') },
                          { value: 'secondary', label: officeMessage(messages, 'spreadsheet.chart.axis.secondary') },
                        ]}
                        onValueChange={(axisGroup) =>
                          setDraft({
                            ...draft,
                            series: replaceSeries(draft.series, index, {
                              axisGroup:
                                axisGroup as WorkSpreadsheetChartAxisGroup,
                            }),
                          })
                        }
                      />
                    </div>
                  </>
                )}
                {draft.type === 'bubble' && (
                  <div className="work-office-field bubble-reference">
                    <span>{officeMessage(messages, 'spreadsheet.chart.bubbleSizes')}</span>
                    <OfficeTextField
                      aria-label={officeMessage(messages, 'spreadsheet.chart.seriesBubbleAria', { n: String(index + 1) })}
                      value={series.bubbleSizesReference ?? ''}
                      placeholder={officeMessage(messages, 'spreadsheet.chart.placeholder.bubbleSizes')}
                      onChange={(event) =>
                        setDraft({
                          ...draft,
                          series: replaceSeries(draft.series, index, {
                            bubbleSizesReference: event.target.value,
                          }),
                        })
                      }
                    />
                  </div>
                )}
                <button
                  type="button"
                  className="remove-series"
                  aria-label={officeMessage(messages, 'spreadsheet.chart.deleteSeriesAria', { n: String(index + 1) })}
                  disabled={draft.series.length <= 1}
                  onClick={() =>
                    setDraft({
                      ...draft,
                      series: draft.series.filter(
                        (_, candidate) => candidate !== index,
                      ),
                    })
                  }
                >
                  <Trash2 size={13} />
                </button>
                <SpreadsheetDataLabelEditor
                  chartType={draft.type}
                  seriesNumber={index + 1}
                  value={series.dataLabels}
                  onChange={(dataLabels) =>
                    setDraft({
                      ...draft,
                      series: replaceSeries(draft.series, index, {
                        dataLabels,
                      }),
                    })
                  }
                />
                <SpreadsheetChartSeriesStyleEditor
                  seriesNumber={index + 1}
                  supportsMarkers={
                    draft.type === 'line' ||
                    draft.type === 'radar' ||
                    draft.type === 'scatter' ||
                    (draft.type === 'combination' &&
                      normalizeWorkSpreadsheetCombinationSeriesType(
                        series.chartType,
                      ) === 'line')
                  }
                  value={series.style}
                  onChange={(style) =>
                    setDraft({
                      ...draft,
                      series: replaceSeries(draft.series, index, { style }),
                    })
                  }
                />
                {workSpreadsheetChartSupportsSeriesAnalysis(draft) &&
                  workSpreadsheetChartSupportsErrorBars(draft.type) && (
                    <SpreadsheetErrorBarEditor
                      chartType={draft.type}
                      seriesNumber={index + 1}
                      errorBars={series.errorBars ?? []}
                      onChange={(errorBars) =>
                        setDraft({
                          ...draft,
                          series: replaceSeries(draft.series, index, {
                            errorBars,
                          }),
                        })
                      }
                    />
                  )}
                {workSpreadsheetChartSupportsSeriesAnalysis(draft) &&
                  workSpreadsheetChartSupportsTrendlines(draft.type) && (
                    <SpreadsheetTrendlineEditor
                      seriesNumber={index + 1}
                      trendlines={series.trendlines ?? []}
                      onChange={(trendlines) =>
                        setDraft({
                          ...draft,
                          series: replaceSeries(draft.series, index, {
                            trendlines,
                          }),
                        })
                      }
                    />
                  )}
              </div>
            ))}
          </section>
          <div className="actions">
            {error && (
              <InlineNotice
                className="work-office-form-error"
                tone="danger"
                role="alert"
              >
                {error}
              </InlineNotice>
            )}
            <Button tone="danger" onClick={deleteChart}>
              <Trash2 size={13} />
              {officeMessage(messages, 'spreadsheet.chart.delete')}
            </Button>
            <Button tone="secondary" disabled={!dirty} onClick={cancelDraft}>
              {officeMessage(messages, 'spreadsheet.chart.cancel')}
            </Button>
            <Button type="submit" tone="primary" disabled={!dirty}>
              {officeMessage(messages, 'spreadsheet.chart.save')}
            </Button>
          </div>
        </form>
      ) : (
        <StateView
          className="work-spreadsheet-chart-empty"
          size="compact"
          title={officeMessage(messages, 'spreadsheet.chart.emptyTitle')}
          description={officeMessage(messages, 'spreadsheet.chart.emptyDescription')}
        >
          {error && (
            <InlineNotice
              className="work-office-form-error"
              tone="danger"
              role="alert"
            >
              {error}
            </InlineNotice>
          )}
        </StateView>
      )}
    </fieldset>
  );
}

function chartDraftKey(draft: ChartDraft | null): string | null {
  return draft ? `${draft.sheetId}:${draft.id}` : null;
}
