import { officeMessage } from '../../../i18n/office-locale';
import {
  workSpreadsheetChartAxisDefaultLabelPosition,
  workSpreadsheetChartAxisIsCategoryAxis,
  workSpreadsheetChartAxisIsValueAxis,
  workSpreadsheetChartAxisShowsMajorGridlinesByDefault,
} from '../work-spreadsheet-chart-axis';
import type {
  WorkSpreadsheetChartAxes,
  WorkSpreadsheetChartAxis,
  WorkSpreadsheetChartAxisPosition,
  WorkSpreadsheetChartType,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { normalizeOptionalOfficeNumber } from './office-number-normalization';

interface SpreadsheetChartAxisEditorProps {
  axes: WorkSpreadsheetChartAxes | undefined;
  chartType: WorkSpreadsheetChartType;
  showSecondaryAxes: boolean;
  onChange: (axes: WorkSpreadsheetChartAxes) => void;
}

const PRIMARY_AXIS_POSITIONS = ['bottom', 'left'] as const;
const SECONDARY_AXIS_POSITIONS = ['top', 'right'] as const;

const AXIS_LABEL_KEYS = {
  bottom: 'spreadsheet.chart.axis.bottom',
  left: 'spreadsheet.chart.axis.left',
  top: 'spreadsheet.chart.axis.top',
  right: 'spreadsheet.chart.axis.right',
} as const;

export function SpreadsheetChartAxisEditor({
  axes,
  chartType,
  showSecondaryAxes,
  onChange,
}: SpreadsheetChartAxisEditorProps) {
  const messages = useOfficeMessages();
  const positions = showSecondaryAxes
    ? [...PRIMARY_AXIS_POSITIONS, ...SECONDARY_AXIS_POSITIONS]
    : [...PRIMARY_AXIS_POSITIONS];
  const updateAxis = (
    position: WorkSpreadsheetChartAxisPosition,
    changes: Partial<WorkSpreadsheetChartAxis>,
  ) => {
    onChange({
      ...axes,
      [position]: {
        ...axes?.[position],
        ...changes,
      },
    });
  };

  return (
    <section
      className="work-spreadsheet-chart-axes"
      aria-label={officeMessage(messages, 'spreadsheet.chart.axis.settingsAria')}
    >
      <header>
        <strong>{officeMessage(messages, 'spreadsheet.chart.axis.title')}</strong>
        <span>{officeMessage(messages, 'spreadsheet.chart.axis.subtitle')}</span>
      </header>
      <div>
        {positions.map((position) => {
          const label = officeMessage(messages, AXIS_LABEL_KEYS[position]);
          const axis = axes?.[position];
          const valueAxis = workSpreadsheetChartAxisIsValueAxis(
            chartType,
            position,
          );
          const categoryAxis = workSpreadsheetChartAxisIsCategoryAxis(
            chartType,
            position,
          );
          const labelPosition =
            axis?.labelPosition ??
            workSpreadsheetChartAxisDefaultLabelPosition(chartType, position);
          const showMajorGridlines =
            axis?.showMajorGridlines ??
            workSpreadsheetChartAxisShowsMajorGridlinesByDefault(
              chartType,
              position,
            );
          const sourceLinked =
            axis?.numberFormatSourceLinked ?? !axis?.numberFormat;
          return (
            <fieldset key={position}>
              <legend>{label}</legend>
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.axis.caption')}
                </span>
                <OfficeTextField
                  aria-label={officeMessage(
                    messages,
                    'spreadsheet.chart.axis.captionAria',
                    { axis: label },
                  )}
                  value={axis?.title ?? ''}
                  maxLength={255}
                  onChange={(event) =>
                    updateAxis(position, { title: event.target.value })
                  }
                />
              </div>
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.axis.captionRef')}
                </span>
                <OfficeTextField
                  aria-label={officeMessage(
                    messages,
                    'spreadsheet.chart.axis.captionRefAria',
                    { axis: label },
                  )}
                  value={axis?.titleReference ?? ''}
                  placeholder={officeMessage(
                    messages,
                    'spreadsheet.chart.axis.captionRefPlaceholder',
                  )}
                  onChange={(event) =>
                    updateAxis(position, { titleReference: event.target.value })
                  }
                />
              </div>
              <OfficeCheckbox
                className="axis-check"
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.axis.reverseAria',
                  { axis: label },
                )}
                checked={axis?.reverseOrder === true}
                onCheckedChange={(reverseOrder) =>
                  updateAxis(position, { reverseOrder })
                }
              >
                {officeMessage(messages, 'spreadsheet.chart.axis.reverse')}
              </OfficeCheckbox>
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.axis.labelPos')}
                </span>
                <OfficeSelect
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.axis.labelPosAria',
                    { axis: label },
                  )}
                  value={labelPosition}
                  options={[
                    {
                      value: 'nextTo',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.label.nextTo',
                      ),
                    },
                    {
                      value: 'high',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.label.high',
                      ),
                    },
                    {
                      value: 'low',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.label.low',
                      ),
                    },
                    {
                      value: 'none',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.label.none',
                      ),
                    },
                  ]}
                  onValueChange={(value) =>
                    updateAxis(position, {
                      labelPosition: value as NonNullable<
                        WorkSpreadsheetChartAxis['labelPosition']
                      >,
                    })
                  }
                />
              </div>
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.axis.majorTick')}
                </span>
                <OfficeSelect
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.axis.majorTickAria',
                    { axis: label },
                  )}
                  value={axis?.majorTickMark ?? 'none'}
                  options={[
                    {
                      value: 'none',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.tick.none',
                      ),
                    },
                    {
                      value: 'inside',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.tick.inside',
                      ),
                    },
                    {
                      value: 'outside',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.tick.outside',
                      ),
                    },
                    {
                      value: 'cross',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.axis.tick.cross',
                      ),
                    },
                  ]}
                  onValueChange={(value) =>
                    updateAxis(position, {
                      majorTickMark: value as NonNullable<
                        WorkSpreadsheetChartAxis['majorTickMark']
                      >,
                    })
                  }
                />
              </div>
              {categoryAxis ? (
                <div className="work-office-field">
                  <span>
                    {officeMessage(
                      messages,
                      'spreadsheet.chart.axis.labelInterval',
                    )}
                  </span>
                  <CommittedOfficeNumberField
                    min={1}
                    max={31_999}
                    step={1}
                    ariaLabel={officeMessage(
                      messages,
                      'spreadsheet.chart.axis.labelIntervalAria',
                      { axis: label },
                    )}
                    value={axis?.labelInterval}
                    placeholder={officeMessage(
                      messages,
                      'spreadsheet.chart.axis.auto',
                    )}
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
              ) : null}
              {valueAxis ? (
                <>
                  <div className="work-office-field">
                    <span>
                      {officeMessage(messages, 'spreadsheet.chart.axis.min')}
                    </span>
                    <CommittedOfficeNumberField
                      step={0.1}
                      ariaLabel={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.minAria',
                        { axis: label },
                      )}
                      value={axis?.minimum}
                      placeholder={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.auto',
                      )}
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
                    <span>
                      {officeMessage(messages, 'spreadsheet.chart.axis.max')}
                    </span>
                    <CommittedOfficeNumberField
                      step={0.1}
                      ariaLabel={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.maxAria',
                        { axis: label },
                      )}
                      value={axis?.maximum}
                      placeholder={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.auto',
                      )}
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
                    <span>
                      {officeMessage(
                        messages,
                        'spreadsheet.chart.axis.majorUnit',
                      )}
                    </span>
                    <CommittedOfficeNumberField
                      min={0}
                      step={0.1}
                      ariaLabel={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.majorUnitAria',
                        { axis: label },
                      )}
                      value={axis?.majorUnit}
                      placeholder={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.auto',
                      )}
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
                    <span>
                      {officeMessage(
                        messages,
                        'spreadsheet.chart.axis.numberFormat',
                      )}
                    </span>
                    <OfficeTextField
                      aria-label={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.numberFormatAria',
                        { axis: label },
                      )}
                      value={axis?.numberFormat ?? ''}
                      maxLength={255}
                      placeholder={officeMessage(
                        messages,
                        'spreadsheet.chart.axis.numberFormatPlaceholder',
                      )}
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
                    ariaLabel={officeMessage(
                      messages,
                      'spreadsheet.chart.axis.majorGridAria',
                      { axis: label },
                    )}
                    checked={showMajorGridlines}
                    onCheckedChange={(showMajorGridlines) =>
                      updateAxis(position, { showMajorGridlines })
                    }
                  >
                    {officeMessage(messages, 'spreadsheet.chart.axis.majorGrid')}
                  </OfficeCheckbox>
                  <OfficeCheckbox
                    className="axis-check"
                    ariaLabel={officeMessage(
                      messages,
                      'spreadsheet.chart.axis.linkNumberAria',
                      { axis: label },
                    )}
                    checked={sourceLinked}
                    onCheckedChange={(numberFormatSourceLinked) =>
                      updateAxis(position, { numberFormatSourceLinked })
                    }
                  >
                    {officeMessage(
                      messages,
                      'spreadsheet.chart.axis.linkNumber',
                    )}
                  </OfficeCheckbox>
                </>
              ) : null}
            </fieldset>
          );
        })}
      </div>
    </section>
  );
}
