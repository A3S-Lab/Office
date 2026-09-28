import { Plus, Trash2, X } from 'lucide-react';
import { useRef } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, IconButton } from '../../../design-system/primitives';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import {
  createPresentationChartSeries,
  normalizeDoughnutHoleSize,
  normalizePresentationBubbleScale,
  normalizePresentationBubbleSizeRepresents,
  normalizePresentationScatterStyle,
  parsePresentationChartCategories,
  parsePresentationChartValueDraft,
  presentationChartSupportsAxisTitles,
  presentationChartSupportsSeriesMarkers,
  presentationChartTypeLabel,
  presentationChartUsesNumericXAxis,
  withPresentationChartDataLabels,
  withPresentationChartSeriesStyle,
  withPresentationChartType,
} from '../work-presentation-charts';
import type {
  WorkSlideBubbleSizeRepresents,
  WorkSlideChart,
  WorkSlideChartType,
  WorkSlideRadarStyle,
  WorkSlideScatterStyle,
} from '../work-types';
import {
  CommittedOfficeNumberField,
  CommittedOfficeTextArea,
  CommittedOfficeTextField,
  OfficeCheckbox,
  OfficeSelect,
} from './office-controls';
import { PresentationChartAxisEditor } from './presentation-chart-axis-editor';
import { PresentationChartDataLabelEditor } from './presentation-chart-data-label-editor';
import { PresentationChartLayoutEditor } from './presentation-chart-layout-editor';
import { PresentationChartSeriesAnalysisEditor } from './presentation-chart-series-analysis-editor';
import { SpreadsheetChartSeriesStyleEditor } from './spreadsheet-chart-series-style-editor';
import { useOfficeMessages } from './office-messages-context';
import {
  handleOfficeTaskPaneKeyDown,
  useOfficeTaskPaneModal,
} from './office-task-pane';

const PRESENTATION_CHART_PANE_MODAL_QUERY = '(max-width: 1100px)';

const CHART_TYPES: WorkSlideChartType[] = [
  'column',
  'bar',
  'line',
  'area',
  'pie',
  'doughnut',
  'radar',
  'scatter',
  'bubble',
];

export function PresentationChartPanel({
  chart,
  onChange,
  onDelete,
  onClose,
  restoreFocusTarget,
}: {
  chart: WorkSlideChart;
  onChange: (chart: WorkSlideChart) => void;
  onDelete: () => void;
  onClose: () => void;
  restoreFocusTarget?: () => HTMLElement | null;
}) {
  const messages = useOfficeMessages();
  const panelRef = useRef<HTMLElement>(null);
  const closeRef = useRef<HTMLButtonElement>(null);
  const modal = useOfficeTaskPaneModal(PRESENTATION_CHART_PANE_MODAL_QUERY);
  const modalAttributes = modal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};
  const focusScope = useDialogFocusScope<HTMLElement>({
    active: modal,
    initialFocus: () => closeRef.current,
    getActiveScope: () => panelRef.current,
    restoreFocusTarget,
  });
  const numericXAxis = presentationChartUsesNumericXAxis(chart.type);
  const updateSeries = (
    index: number,
    patch: Partial<WorkSlideChart['series'][number]>,
  ) => {
    onChange({
      ...chart,
      series: chart.series.map((series, current) =>
        current === index ? { ...series, ...patch } : series,
      ),
    });
  };
  return (
    <section
      {...modalAttributes}
      ref={panelRef}
      className="work-presentation-chart-panel"
      aria-label={officeMessage(messages, 'presentation.chart.panelAria')}
      onKeyDown={(event) => {
        focusScope.handleKeyDown(event);
        if (!event.defaultPrevented)
          handleOfficeTaskPaneKeyDown(event, onClose);
      }}
    >
      <header>
        <div>
          <strong>{officeMessage(messages, 'presentation.chart.panelTitle')}</strong>
          <span>{officeMessage(messages, 'presentation.chart.panelHint')}</span>
        </div>
        <div>
          <Button tone="danger" aria-label={officeMessage(messages, 'presentation.chart.deleteAria')} onClick={onDelete}>
            <Trash2 size={13} />
            {officeMessage(messages, 'presentation.chart.delete')}
          </Button>
          <IconButton
            ref={closeRef}
            className="close"
            label={officeMessage(messages, 'presentation.chart.close')}
            onClick={(event) => {
              event.stopPropagation();
              onClose();
            }}
          >
            <X size={14} />
          </IconButton>
        </div>
      </header>
      <div className="work-presentation-chart-controls">
        <div className="work-office-field">
          <span>{officeMessage(messages, 'presentation.chart.type')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'presentation.chart.typeAria')}
            value={chart.type}
            options={CHART_TYPES.map((type) => ({
              value: type,
              label: presentationChartTypeLabel(type, messages),
            }))}
            onValueChange={(type) =>
              onChange(
                withPresentationChartType(chart, type as WorkSlideChartType),
              )
            }
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'presentation.chart.title')}</span>
          <CommittedOfficeTextField
            aria-label={officeMessage(messages, 'presentation.chart.titleAria')}
            value={chart.title}
            formatValue={(title) => title ?? ''}
            parseValue={(draft) => draft.trim().slice(0, 255) || undefined}
            onValueCommit={(title) => onChange({ ...chart, title })}
          />
        </div>
        <PresentationChartLayoutEditor chart={chart} onChange={onChange} />
        {presentationChartSupportsAxisTitles(chart) && (
          <PresentationChartAxisEditor chart={chart} onChange={onChange} />
        )}
        <PresentationChartDataLabelEditor
          chartType={chart.type}
          value={chart.dataLabels}
          onChange={(dataLabels) =>
            onChange(withPresentationChartDataLabels(chart, dataLabels))
          }
        />
        {chart.type === 'doughnut' && (
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.chart.holeSize')}</span>
            <CommittedOfficeNumberField
              ariaLabel={officeMessage(messages, 'presentation.chart.holeSizeAria')}
              min={10}
              max={90}
              value={normalizeDoughnutHoleSize(chart.doughnutHoleSize)}
              normalizeValue={(value) =>
                normalizePresentationChartPercent(value, 10, 90)
              }
              onValueCommit={(doughnutHoleSize) =>
                onChange({
                  ...chart,
                  doughnutHoleSize,
                })
              }
            />
          </div>
        )}
        {chart.type === 'radar' && (
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.chart.style')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.chart.radarStyleAria')}
              value={chart.radarStyle ?? 'standard'}
              options={[
                { value: 'standard', label: officeMessage(messages, 'presentation.chart.radar.standard') },
                { value: 'marker', label: officeMessage(messages, 'presentation.chart.radar.marker') },
                { value: 'filled', label: officeMessage(messages, 'presentation.chart.radar.filled') },
              ]}
              onValueChange={(radarStyle) =>
                onChange({
                  ...chart,
                  radarStyle: radarStyle as WorkSlideRadarStyle,
                })
              }
            />
          </div>
        )}
        {chart.type === 'scatter' && (
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.chart.scatterStyle')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.chart.scatterStyleAria')}
              value={normalizePresentationScatterStyle(chart.scatterStyle)}
              options={[
                { value: 'marker', label: officeMessage(messages, 'presentation.chart.scatter.marker') },
                { value: 'line', label: officeMessage(messages, 'presentation.chart.scatter.line') },
                { value: 'lineMarker', label: officeMessage(messages, 'presentation.chart.scatter.lineMarker') },
                { value: 'smooth', label: officeMessage(messages, 'presentation.chart.scatter.smooth') },
                { value: 'smoothMarker', label: officeMessage(messages, 'presentation.chart.scatter.smoothMarker') },
              ]}
              onValueChange={(scatterStyle) =>
                onChange({
                  ...chart,
                  scatterStyle: scatterStyle as WorkSlideScatterStyle,
                })
              }
            />
          </div>
        )}
        {chart.type === 'bubble' && (
          <>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'presentation.chart.bubbleScale')}</span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(messages, 'presentation.chart.bubbleScaleAria')}
                min={5}
                max={300}
                value={normalizePresentationBubbleScale(chart.bubbleScale)}
                normalizeValue={(value) =>
                  normalizePresentationChartPercent(value, 5, 300)
                }
                onValueCommit={(bubbleScale) =>
                  onChange({
                    ...chart,
                    bubbleScale,
                  })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'presentation.chart.sizeRepresents')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'presentation.chart.sizeRepresentsAria')}
                value={normalizePresentationBubbleSizeRepresents(
                  chart.bubbleSizeRepresents,
                )}
                options={[
                  { value: 'area', label: officeMessage(messages, 'presentation.chart.size.area') },
                  { value: 'width', label: officeMessage(messages, 'presentation.chart.size.width') },
                ]}
                onValueChange={(bubbleSizeRepresents) =>
                  onChange({
                    ...chart,
                    bubbleSizeRepresents:
                      bubbleSizeRepresents as WorkSlideBubbleSizeRepresents,
                  })
                }
              />
            </div>
            <div className="check">
              <span>{officeMessage(messages, 'presentation.chart.negativeBubbles')}</span>
              <OfficeCheckbox
                className="work-presentation-chart-check-control"
                ariaLabel={officeMessage(messages, 'presentation.chart.negativeBubblesAria')}
                checked={chart.showNegativeBubbles === true}
                onCheckedChange={(showNegativeBubbles) =>
                  onChange({ ...chart, showNegativeBubbles })
                }
              >
                {officeMessage(messages, 'presentation.chart.show')}
              </OfficeCheckbox>
            </div>
          </>
        )}
        <div className="work-office-field categories">
          <span>{(numericXAxis ? officeMessage(messages, 'presentation.chart.xValues') : officeMessage(messages, 'presentation.chart.categories')) + officeMessage(messages, 'presentation.chart.perLine')}</span>
          <CommittedOfficeTextArea
            aria-label={numericXAxis ? officeMessage(messages, 'presentation.chart.xValuesAria') : officeMessage(messages, 'presentation.chart.categoriesAria')}
            value={chart.categories}
            formatValue={(categories) => categories.join('\n')}
            parseValue={(draft) => {
              if (!numericXAxis) return parsePresentationChartCategories(draft);
              const values = parsePresentationChartValueDraft(draft);
              return values?.map(String) ?? null;
            }}
            onValueCommit={(categories) =>
              onChange({
                ...chart,
                categories,
              })
            }
          />
        </div>
        <div className="work-presentation-chart-series-list">
          {chart.series.map((series, index) => (
            <div className="work-presentation-chart-series-card" key={index}>
              <fieldset>
                <legend>{officeMessage(messages, 'presentation.chart.seriesLegend', { index: String(index + 1) })}</legend>
                <CommittedOfficeTextField
                  aria-label={officeMessage(messages, 'presentation.chart.seriesNameAria', { index: String(index + 1) })}
                  value={series.name}
                  formatValue={(name) => name}
                  parseValue={(draft) => draft.trim().slice(0, 255)}
                  onValueCommit={(name) => updateSeries(index, { name })}
                />
                <CommittedOfficeTextArea
                  aria-label={officeMessage(messages, numericXAxis ? 'presentation.chart.seriesYAria' : 'presentation.chart.seriesDataAria', { index: String(index + 1) })}
                  value={series.values}
                  formatValue={(values) => values.join(', ')}
                  parseValue={parsePresentationChartValueDraft}
                  onValueCommit={(values) =>
                    updateSeries(index, {
                      values,
                    })
                  }
                />
                {chart.type === 'bubble' && (
                  <CommittedOfficeTextArea
                    aria-label={officeMessage(messages, 'presentation.chart.seriesBubbleAria', { index: String(index + 1) })}
                    value={series.bubbleSizes ?? []}
                    formatValue={(values) => values.join(', ')}
                    parseValue={parsePresentationChartValueDraft}
                    onValueCommit={(bubbleSizes) =>
                      updateSeries(index, {
                        bubbleSizes,
                      })
                    }
                  />
                )}
                <IconButton
                  label={officeMessage(messages, 'presentation.chart.deleteSeriesAria', { index: String(index + 1) })}
                  disabled={chart.series.length === 1}
                  onClick={() =>
                    onChange({
                      ...chart,
                      series: chart.series.filter(
                        (_, current) => current !== index,
                      ),
                    })
                  }
                >
                  <Trash2 size={12} />
                </IconButton>
              </fieldset>
              <SpreadsheetChartSeriesStyleEditor
                seriesNumber={index + 1}
                supportsMarkers={presentationChartSupportsSeriesMarkers(
                  chart.type,
                )}
                value={series.style}
                onChange={(style) =>
                  onChange(
                    withPresentationChartSeriesStyle(chart, index, style),
                  )
                }
              />
              <PresentationChartSeriesAnalysisEditor
                chart={chart}
                seriesIndex={index}
                onChange={onChange}
              />
            </div>
          ))}
          <Button
            tone="secondary"
            className="add-series"
            aria-label={officeMessage(messages, 'presentation.chart.addSeriesAria')}
            onClick={() =>
              onChange({
                ...chart,
                series: [...chart.series, createPresentationChartSeries(chart, messages)],
              })
            }
          >
            <Plus size={13} />
            {officeMessage(messages, 'presentation.chart.addSeries')}
          </Button>
        </div>
      </div>
    </section>
  );
}

function normalizePresentationChartPercent(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number)))
    : null;
}
