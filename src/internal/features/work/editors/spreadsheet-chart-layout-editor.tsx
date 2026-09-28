import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizeWorkSpreadsheetChartGapWidth,
  normalizeWorkSpreadsheetChartGrouping,
  normalizeWorkSpreadsheetChartLegendOverlay,
  normalizeWorkSpreadsheetChartLegendPosition,
  normalizeWorkSpreadsheetChartOverlap,
  normalizeWorkSpreadsheetChartSmoothLines,
  type WorkSpreadsheetChartGrouping,
  type WorkSpreadsheetChartLayout,
  type WorkSpreadsheetChartLegendPosition,
  workSpreadsheetChartGroupingIsStacked,
  workSpreadsheetChartSupportsBarSpacing,
  workSpreadsheetChartSupportsGrouping,
  workSpreadsheetChartSupportsSmoothLines,
} from '../work-spreadsheet-chart-layout';
import type { WorkSpreadsheetChartType } from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
} from './office-controls';
import { normalizeRequiredOfficeNumber } from './office-number-normalization';

interface SpreadsheetChartLayoutEditorProps {
  chart: WorkSpreadsheetChartLayout & {
    type: WorkSpreadsheetChartType;
    showLegend: boolean;
  };
  onChange: (
    change: Partial<WorkSpreadsheetChartLayout & { showLegend: boolean }>,
  ) => void;
}

export function SpreadsheetChartLayoutEditor({
  chart,
  onChange,
}: SpreadsheetChartLayoutEditorProps) {
  const messages = useOfficeMessages();
  const grouping = normalizeWorkSpreadsheetChartGrouping(
    chart.grouping,
    chart.type,
  );
  return (
    <section
      className="work-spreadsheet-chart-layout"
      aria-label={officeMessage(messages, 'spreadsheet.chart.layout.settingsAria')}
    >
      <header>
        <strong>
          {officeMessage(messages, 'spreadsheet.chart.layout.title')}
        </strong>
        <span>
          {officeMessage(messages, 'spreadsheet.chart.layout.subtitle')}
        </span>
      </header>
      <div>
        <OfficeCheckbox
          className="check"
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.chart.layout.showLegendAria',
          )}
          checked={chart.showLegend}
          onCheckedChange={(showLegend) => onChange({ showLegend })}
        >
          {officeMessage(messages, 'spreadsheet.chart.layout.showLegend')}
        </OfficeCheckbox>
        {chart.showLegend ? (
          <>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.layout.legendPos')}
              </span>
              <OfficeSelect
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.layout.legendPosAria',
                )}
                value={normalizeWorkSpreadsheetChartLegendPosition(
                  chart.legendPosition,
                )}
                options={[
                  {
                    value: 'right',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.layout.legend.right',
                    ),
                  },
                  {
                    value: 'left',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.layout.legend.left',
                    ),
                  },
                  {
                    value: 'top',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.layout.legend.top',
                    ),
                  },
                  {
                    value: 'bottom',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.layout.legend.bottom',
                    ),
                  },
                  {
                    value: 'topRight',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.layout.legend.topRight',
                    ),
                  },
                ]}
                onValueChange={(legendPosition) =>
                  onChange({
                    legendPosition:
                      legendPosition as WorkSpreadsheetChartLegendPosition,
                  })
                }
              />
            </div>
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.layout.overlayAria',
              )}
              checked={normalizeWorkSpreadsheetChartLegendOverlay(
                chart.legendOverlay,
              )}
              onCheckedChange={(legendOverlay) => onChange({ legendOverlay })}
            >
              {officeMessage(messages, 'spreadsheet.chart.layout.overlay')}
            </OfficeCheckbox>
          </>
        ) : null}
        {workSpreadsheetChartSupportsGrouping(chart.type) ? (
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'spreadsheet.chart.layout.grouping')}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.layout.groupingAria',
              )}
              value={grouping}
              options={[
                ...(workSpreadsheetChartSupportsBarSpacing(chart.type)
                  ? [
                      {
                        value: 'clustered',
                        label: officeMessage(
                          messages,
                          'spreadsheet.chart.layout.grouping.clustered',
                        ),
                      } as const,
                    ]
                  : []),
                {
                  value: 'standard',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.layout.grouping.standard',
                  ),
                },
                {
                  value: 'stacked',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.layout.grouping.stacked',
                  ),
                },
                {
                  value: 'percentStacked',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.layout.grouping.percentStacked',
                  ),
                },
              ]}
              onValueChange={(value) => {
                const nextGrouping = value as WorkSpreadsheetChartGrouping;
                if (!workSpreadsheetChartSupportsBarSpacing(chart.type)) {
                  onChange({ grouping: nextGrouping });
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
                onChange({
                  grouping: nextGrouping,
                  overlap:
                    currentOverlap === currentDefault
                      ? nextDefault
                      : currentOverlap,
                });
              }}
            />
          </div>
        ) : null}
        {workSpreadsheetChartSupportsBarSpacing(chart.type) ? (
          <>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.layout.gapWidth')}
              </span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.layout.gapWidth',
                )}
                min={0}
                max={500}
                step={1}
                value={normalizeWorkSpreadsheetChartGapWidth(chart.gapWidth)}
                normalizeValue={(value) =>
                  normalizeRequiredOfficeNumber(value, {
                    integer: true,
                    minimum: 0,
                    maximum: 500,
                  })
                }
                onValueCommit={(gapWidth) => onChange({ gapWidth })}
              />
            </div>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.layout.overlap')}
              </span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.layout.overlap',
                )}
                min={-100}
                max={100}
                step={1}
                value={normalizeWorkSpreadsheetChartOverlap(
                  chart.overlap,
                  grouping,
                )}
                normalizeValue={(value) =>
                  normalizeRequiredOfficeNumber(value, {
                    integer: true,
                    minimum: -100,
                    maximum: 100,
                  })
                }
                onValueCommit={(overlap) => onChange({ overlap })}
              />
            </div>
          </>
        ) : null}
        {workSpreadsheetChartSupportsSmoothLines(chart.type) ? (
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.chart.layout.smoothAria',
            )}
            checked={normalizeWorkSpreadsheetChartSmoothLines(
              chart.smoothLines,
            )}
            onCheckedChange={(smoothLines) => onChange({ smoothLines })}
          >
            {officeMessage(messages, 'spreadsheet.chart.layout.smooth')}
          </OfficeCheckbox>
        ) : null}
      </div>
      {workSpreadsheetChartGroupingIsStacked(grouping) ? (
        <p>{officeMessage(messages, 'spreadsheet.chart.layout.stackNote')}</p>
      ) : null}
    </section>
  );
}
