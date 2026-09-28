import { officeMessage } from '../../../i18n/office-locale';
import {
  presentationChartAxes,
  withPresentationChartAxes,
} from '../work-presentation-chart-axes';
import {
  workSpreadsheetChartAxisDefaultLabelPosition,
  workSpreadsheetChartAxisIsCategoryAxis,
  workSpreadsheetChartAxisIsValueAxis,
  workSpreadsheetChartAxisShowsMajorGridlinesByDefault,
} from '../work-spreadsheet-chart-axis';
import type {
  WorkSlideChart,
  WorkSlideChartAxis,
  WorkSpreadsheetChartAxisPosition,
} from '../work-types';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { normalizeOptionalOfficeNumber } from './office-number-normalization';

const PRIMARY_AXIS_POSITIONS = ['bottom', 'left'] as const;

const AXIS_LABEL_KEYS = {
  bottom: 'presentation.chart.axis.bottom',
  left: 'presentation.chart.axis.left',
} as const;

export function PresentationChartAxisEditor({
  chart,
  onChange,
}: {
  chart: WorkSlideChart;
  onChange: (chart: WorkSlideChart) => void;
}) {
  const messages = useOfficeMessages();
  const axes = presentationChartAxes(chart);
  const updateAxis = (
    position: WorkSpreadsheetChartAxisPosition,
    changes: Partial<WorkSlideChartAxis>,
  ) => {
    onChange(
      withPresentationChartAxes(chart, {
        ...axes,
        [position]: { ...axes?.[position], ...changes },
      }),
    );
  };

  return (
    <section
      className="work-spreadsheet-chart-axes work-presentation-chart-axes"
      aria-label={officeMessage(messages, 'presentation.chart.axis.settingsAria')}
    >
      <header>
        <strong>{officeMessage(messages, 'presentation.chart.axis.title')}</strong>
        <span>{officeMessage(messages, 'presentation.chart.axis.subtitle')}</span>
      </header>
      <div>
        {PRIMARY_AXIS_POSITIONS.map((position) => {
          const label = officeMessage(messages, AXIS_LABEL_KEYS[position]);
          const axis = axes?.[position];
          const valueAxis = workSpreadsheetChartAxisIsValueAxis(
            chart.type,
            position,
          );
          const categoryAxis = workSpreadsheetChartAxisIsCategoryAxis(
            chart.type,
            position,
          );
          const labelPosition =
            axis?.labelPosition ??
            workSpreadsheetChartAxisDefaultLabelPosition(chart.type, position);
          const showMajorGridlines =
            axis?.showMajorGridlines ??
            workSpreadsheetChartAxisShowsMajorGridlinesByDefault(
              chart.type,
              position,
            );
          const sourceLinked =
            axis?.numberFormatSourceLinked ?? !axis?.numberFormat;
          const ariaName = officeMessage(messages, 'presentation.chart.axis.ariaPrefix', { axis: label });
          return (
            <fieldset key={position}>
              <legend>{label}</legend>
              <div className="work-office-field">
                <span>{officeMessage(messages, 'presentation.chart.axis.caption')}</span>
                <OfficeTextField
                  aria-label={officeMessage(messages, 'presentation.chart.axis.captionAria', { axis: ariaName })}
                  value={axis?.title ?? ''}
                  maxLength={255}
                  onChange={(event) =>
                    updateAxis(position, { title: event.target.value })
                  }
                />
              </div>
              <OfficeCheckbox
                className="axis-check"
                ariaLabel={officeMessage(messages, 'presentation.chart.axis.reverseAria', { axis: ariaName })}
                checked={axis?.reverseOrder === true}
                onCheckedChange={(reverseOrder) =>
                  updateAxis(position, { reverseOrder })
                }
              >
                {officeMessage(messages, 'presentation.chart.axis.reverse')}
              </OfficeCheckbox>
              <div className="work-office-field">
                <span>{officeMessage(messages, 'presentation.chart.axis.labelPos')}</span>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'presentation.chart.axis.labelPosAria', { axis: ariaName })}
                  value={labelPosition}
                  options={[
                    { value: 'nextTo', label: officeMessage(messages, 'presentation.chart.axis.label.nextTo') },
                    { value: 'high', label: officeMessage(messages, 'presentation.chart.axis.label.high') },
                    { value: 'low', label: officeMessage(messages, 'presentation.chart.axis.label.low') },
                    { value: 'none', label: officeMessage(messages, 'presentation.chart.axis.label.none') },
                  ]}
                  onValueChange={(value) =>
                    updateAxis(position, {
                      labelPosition: value as NonNullable<
                        WorkSlideChartAxis['labelPosition']
                      >,
                    })
                  }
                />
              </div>
              <div className="work-office-field">
                <span>{officeMessage(messages, 'presentation.chart.axis.majorTick')}</span>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'presentation.chart.axis.majorTickAria', { axis: ariaName })}
                  value={axis?.majorTickMark ?? 'none'}
                  options={[
                    { value: 'none', label: officeMessage(messages, 'presentation.chart.axis.tick.none') },
                    { value: 'inside', label: officeMessage(messages, 'presentation.chart.axis.tick.inside') },
                    { value: 'outside', label: officeMessage(messages, 'presentation.chart.axis.tick.outside') },
                    { value: 'cross', label: officeMessage(messages, 'presentation.chart.axis.tick.cross') },
                  ]}
                  onValueChange={(value) =>
                    updateAxis(position, {
                      majorTickMark: value as NonNullable<
                        WorkSlideChartAxis['majorTickMark']
                      >,
                    })
                  }
                />
              </div>
              {categoryAxis && (
                <div className="work-office-field">
                  <span>{officeMessage(messages, 'presentation.chart.axis.labelInterval')}</span>
                  <CommittedOfficeNumberField
                    min={1}
                    max={31_999}
                    step={1}
                    ariaLabel={officeMessage(messages, 'presentation.chart.axis.labelIntervalAria', { axis: ariaName })}
                    value={axis?.labelInterval}
                    placeholder={officeMessage(messages, 'presentation.chart.axis.auto')}
                    normalizeValue={(value) =>
                      normalizeOptionalOfficeNumber(value, {
                        integer: true,
                        minimum: 1,
                        maximum: 31_999,
                      })
                    }
                    onValueCommit={(labelInterval) =>
                      updateAxis(position, {
                        labelInterval,
                      })
                    }
                  />
                </div>
              )}
              {valueAxis && (
                <>
                  <div className="work-office-field">
                    <span>{officeMessage(messages, 'presentation.chart.axis.min')}</span>
                    <CommittedOfficeNumberField
                      step={0.1}
                      ariaLabel={officeMessage(messages, 'presentation.chart.axis.minAria', { axis: ariaName })}
                      value={axis?.minimum}
                      placeholder={officeMessage(messages, 'presentation.chart.axis.auto')}
                      normalizeValue={(value) =>
                        normalizeOptionalOfficeNumber(value, {
                          isValid: (minimum) =>
                            axis?.maximum === undefined ||
                            minimum < axis.maximum,
                        })
                      }
                      onValueCommit={(minimum) =>
                        updateAxis(position, { minimum })
                      }
                    />
                  </div>
                  <div className="work-office-field">
                    <span>{officeMessage(messages, 'presentation.chart.axis.max')}</span>
                    <CommittedOfficeNumberField
                      step={0.1}
                      ariaLabel={officeMessage(messages, 'presentation.chart.axis.maxAria', { axis: ariaName })}
                      value={axis?.maximum}
                      placeholder={officeMessage(messages, 'presentation.chart.axis.auto')}
                      normalizeValue={(value) =>
                        normalizeOptionalOfficeNumber(value, {
                          isValid: (maximum) =>
                            axis?.minimum === undefined ||
                            maximum > axis.minimum,
                        })
                      }
                      onValueCommit={(maximum) =>
                        updateAxis(position, { maximum })
                      }
                    />
                  </div>
                  <div className="work-office-field">
                    <span>{officeMessage(messages, 'presentation.chart.axis.majorUnit')}</span>
                    <CommittedOfficeNumberField
                      min={0}
                      step={0.1}
                      ariaLabel={officeMessage(messages, 'presentation.chart.axis.majorUnitAria', { axis: ariaName })}
                      value={axis?.majorUnit}
                      placeholder={officeMessage(messages, 'presentation.chart.axis.auto')}
                      normalizeValue={(value) =>
                        normalizeOptionalOfficeNumber(value, {
                          isValid: (majorUnit) => majorUnit > 0,
                        })
                      }
                      onValueCommit={(majorUnit) =>
                        updateAxis(position, { majorUnit })
                      }
                    />
                  </div>
                  <div className="work-office-field">
                    <span>{officeMessage(messages, 'presentation.chart.axis.numberFormat')}</span>
                    <OfficeTextField
                      aria-label={officeMessage(messages, 'presentation.chart.axis.numberFormatAria', { axis: ariaName })}
                      value={axis?.numberFormat ?? ''}
                      maxLength={255}
                      placeholder={officeMessage(messages, 'presentation.chart.axis.numberFormatPlaceholder')}
                      onChange={(event) => {
                        const numberFormat = event.target.value;
                        updateAxis(position, {
                          numberFormat,
                          numberFormatSourceLinked: numberFormat.trim()
                            ? false
                            : undefined,
                        });
                      }}
                    />
                  </div>
                  <OfficeCheckbox
                    className="axis-check"
                    ariaLabel={officeMessage(messages, 'presentation.chart.axis.majorGridAria', { axis: ariaName })}
                    checked={showMajorGridlines}
                    onCheckedChange={(showMajorGridlines) =>
                      updateAxis(position, { showMajorGridlines })
                    }
                  >
                    {officeMessage(messages, 'presentation.chart.axis.majorGrid')}
                  </OfficeCheckbox>
                  <OfficeCheckbox
                    className="axis-check"
                    ariaLabel={officeMessage(messages, 'presentation.chart.axis.linkNumberAria', { axis: ariaName })}
                    checked={sourceLinked}
                    onCheckedChange={(numberFormatSourceLinked) =>
                      updateAxis(position, { numberFormatSourceLinked })
                    }
                  >
                    {officeMessage(messages, 'presentation.chart.axis.linkNumber')}
                  </OfficeCheckbox>
                </>
              )}
            </fieldset>
          );
        })}
      </div>
    </section>
  );
}
