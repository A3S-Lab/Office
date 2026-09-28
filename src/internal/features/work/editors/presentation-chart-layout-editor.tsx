import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizePresentationChartLegendPosition,
  presentationChartShowsLegend,
  withPresentationChartLayout,
} from '../work-presentation-charts';
import {
  normalizeWorkSpreadsheetChartGapWidth,
  normalizeWorkSpreadsheetChartGrouping,
  normalizeWorkSpreadsheetChartLegendOverlay,
  normalizeWorkSpreadsheetChartOverlap,
  normalizeWorkSpreadsheetChartSmoothLines,
  type WorkSpreadsheetChartGrouping,
  workSpreadsheetChartGroupingIsStacked,
  workSpreadsheetChartSupportsBarSpacing,
  workSpreadsheetChartSupportsGrouping,
  workSpreadsheetChartSupportsSmoothLines,
} from '../work-spreadsheet-chart-layout';
import type {
  WorkSlideChart,
  WorkSlideChartLegendPosition,
} from '../work-types';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export function PresentationChartLayoutEditor({
  chart,
  onChange,
}: {
  chart: WorkSlideChart;
  onChange: (chart: WorkSlideChart) => void;
}) {
  const messages = useOfficeMessages();
  const showLegend = presentationChartShowsLegend(chart);
  const grouping = normalizeWorkSpreadsheetChartGrouping(
    chart.grouping,
    chart.type,
  );
  const change = (patch: Parameters<typeof withPresentationChartLayout>[1]) =>
    onChange(withPresentationChartLayout(chart, patch));
  return (
    <section
      className="work-presentation-chart-layout"
      aria-label={officeMessage(messages, 'presentation.chart.layout.settingsAria')}
    >
      <header>
        <strong>{officeMessage(messages, 'presentation.chart.layout.title')}</strong>
        <span>{officeMessage(messages, 'presentation.chart.layout.subtitle')}</span>
      </header>
      <div>
        <div className="check">
          <span>{officeMessage(messages, 'presentation.chart.layout.legend')}</span>
          <OfficeCheckbox
            className="work-presentation-chart-check-control"
            ariaLabel={officeMessage(messages, 'presentation.chart.layout.showLegendAria')}
            checked={showLegend}
            onCheckedChange={(showLegend) => change({ showLegend })}
          >
            {officeMessage(messages, 'presentation.chart.layout.show')}
          </OfficeCheckbox>
        </div>
        {showLegend && (
          <>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'presentation.chart.layout.legendPos')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'presentation.chart.layout.legendPosAria')}
                value={normalizePresentationChartLegendPosition(
                  chart.legendPosition,
                )}
                options={[
                  { value: 'right', label: officeMessage(messages, 'presentation.chart.legend.right') },
                  { value: 'left', label: officeMessage(messages, 'presentation.chart.legend.left') },
                  { value: 'top', label: officeMessage(messages, 'presentation.chart.legend.top') },
                  { value: 'bottom', label: officeMessage(messages, 'presentation.chart.legend.bottom') },
                  { value: 'topRight', label: officeMessage(messages, 'presentation.chart.legend.topRight') },
                ]}
                onValueChange={(legendPosition) =>
                  change({
                    legendPosition:
                      legendPosition as WorkSlideChartLegendPosition,
                  })
                }
              />
            </div>
            <div className="check">
              <span>{officeMessage(messages, 'presentation.chart.layout.legendLayout')}</span>
              <OfficeCheckbox
                className="work-presentation-chart-check-control"
                ariaLabel={officeMessage(messages, 'presentation.chart.layout.overlayAria')}
                checked={normalizeWorkSpreadsheetChartLegendOverlay(
                  chart.legendOverlay,
                )}
                onCheckedChange={(legendOverlay) => change({ legendOverlay })}
              >
                {officeMessage(messages, 'presentation.chart.layout.overlay')}
              </OfficeCheckbox>
            </div>
          </>
        )}
        {workSpreadsheetChartSupportsGrouping(chart.type) && (
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.chart.layout.grouping')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.chart.layout.groupingAria')}
              value={grouping}
              options={[
                ...(workSpreadsheetChartSupportsBarSpacing(chart.type)
                  ? [{ value: 'clustered', label: officeMessage(messages, 'presentation.chart.layout.grouping.clustered') } as const]
                  : []),
                { value: 'standard', label: officeMessage(messages, 'presentation.chart.layout.grouping.standard') },
                { value: 'stacked', label: officeMessage(messages, 'presentation.chart.layout.grouping.stacked') },
                { value: 'percentStacked', label: officeMessage(messages, 'presentation.chart.layout.grouping.percentStacked') },
              ]}
              onValueChange={(value) => {
                const nextGrouping = value as WorkSpreadsheetChartGrouping;
                if (!workSpreadsheetChartSupportsBarSpacing(chart.type)) {
                  change({ grouping: nextGrouping });
                  return;
                }
                const currentDefault = workSpreadsheetChartGroupingIsStacked(
                  grouping,
                )
                  ? 100
                  : 0;
                const nextDefault = workSpreadsheetChartGroupingIsStacked(
                  nextGrouping,
                )
                  ? 100
                  : 0;
                const currentOverlap = normalizeWorkSpreadsheetChartOverlap(
                  chart.overlap,
                  grouping,
                );
                change({
                  grouping: nextGrouping,
                  overlap:
                    currentOverlap === currentDefault
                      ? nextDefault
                      : currentOverlap,
                });
              }}
            />
          </div>
        )}
        {workSpreadsheetChartSupportsBarSpacing(chart.type) && (
          <>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'presentation.chart.layout.gapWidth')}</span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(messages, 'presentation.chart.layout.gapWidthAria')}
                min={0}
                max={500}
                step={1}
                value={normalizeWorkSpreadsheetChartGapWidth(chart.gapWidth)}
                normalizeValue={(value) =>
                  normalizePresentationChartInteger(value, 0, 500)
                }
                onValueCommit={(gapWidth) => change({ gapWidth })}
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'presentation.chart.layout.overlap')}</span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(messages, 'presentation.chart.layout.overlapAria')}
                min={-100}
                max={100}
                step={1}
                value={normalizeWorkSpreadsheetChartOverlap(
                  chart.overlap,
                  grouping,
                )}
                normalizeValue={(value) =>
                  normalizePresentationChartInteger(value, -100, 100)
                }
                onValueCommit={(overlap) => change({ overlap })}
              />
            </div>
          </>
        )}
        {workSpreadsheetChartSupportsSmoothLines(chart.type) && (
          <div className="check">
            <span>{officeMessage(messages, 'presentation.chart.layout.line')}</span>
            <OfficeCheckbox
              className="work-presentation-chart-check-control"
              ariaLabel={officeMessage(messages, 'presentation.chart.layout.smoothAria')}
              checked={normalizeWorkSpreadsheetChartSmoothLines(
                chart.smoothLines,
              )}
              onCheckedChange={(smoothLines) => change({ smoothLines })}
            >
              {officeMessage(messages, 'presentation.chart.layout.smooth')}
            </OfficeCheckbox>
          </div>
        )}
      </div>
      {workSpreadsheetChartGroupingIsStacked(grouping) && (
        <p>
          {officeMessage(messages, 'presentation.chart.layout.stackNote')}
        </p>
      )}
    </section>
  );
}

function normalizePresentationChartInteger(
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
