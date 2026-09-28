import { officeMessage } from '../../../i18n/office-locale';
import {
  defaultWorkSpreadsheetChartSeriesStyle,
  normalizeWorkSpreadsheetChartSeriesStyle,
} from '../work-spreadsheet-chart-series-style';
import type {
  WorkSpreadsheetChartLineDash,
  WorkSpreadsheetChartMarkerStyle,
  WorkSpreadsheetChartMarkerSymbol,
  WorkSpreadsheetChartSeriesStyle,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeSelect,
} from './office-controls';
import { normalizeRequiredOfficeNumber } from './office-number-normalization';

interface SpreadsheetChartSeriesStyleEditorProps {
  seriesNumber: number;
  supportsMarkers: boolean;
  value?: WorkSpreadsheetChartSeriesStyle;
  onChange: (style: WorkSpreadsheetChartSeriesStyle | undefined) => void;
}

export function SpreadsheetChartSeriesStyleEditor({
  seriesNumber,
  supportsMarkers,
  value,
  onChange,
}: SpreadsheetChartSeriesStyleEditorProps) {
  const messages = useOfficeMessages();
  const n = String(seriesNumber);
  const defaults = seriesStyleDefaults(seriesNumber - 1, supportsMarkers);
  const style = normalizeWorkSpreadsheetChartSeriesStyle(value) ?? defaults;
  const marker =
    style.marker ??
    defaultWorkSpreadsheetChartSeriesStyle(seriesNumber - 1).marker!;
  const update = (change: Partial<WorkSpreadsheetChartSeriesStyle>) =>
    onChange(normalizeWorkSpreadsheetChartSeriesStyle({ ...style, ...change }));
  const updateMarker = (change: Partial<WorkSpreadsheetChartMarkerStyle>) =>
    update({ marker: { ...marker, ...change } });
  return (
    <section
      className="work-spreadsheet-chart-series-style"
      aria-label={officeMessage(
        messages,
        'spreadsheet.chart.seriesStyle.settingsAria',
        { n },
      )}
    >
      <OfficeCheckbox
        className="check enable-series-style"
        ariaLabel={officeMessage(
          messages,
          'spreadsheet.chart.seriesStyle.enableAria',
          { n },
        )}
        checked={Boolean(value)}
        onCheckedChange={(checked) => onChange(checked ? defaults : undefined)}
      >
        {officeMessage(messages, 'spreadsheet.chart.seriesStyle.enable')}
      </OfficeCheckbox>
      {value ? (
        <div>
          <div className="work-office-field">
            <span>
              {officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.fillColor',
              )}
            </span>
            <OfficeColorPicker
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.fillColorAria',
                { n },
              )}
              value={style.fillColor ?? '#4F6BED'}
              onValueChange={(fillColor) => update({ fillColor })}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.fillTransparency',
              )}
            </span>
            <CommittedOfficeNumberField
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.fillTransparencyAria',
                { n },
              )}
              min={0}
              max={100}
              step={1}
              value={style.fillTransparency ?? 0}
              normalizeValue={(value) =>
                normalizeRequiredOfficeNumber(value, {
                  integer: true,
                  minimum: 0,
                  maximum: 100,
                })
              }
              onValueCommit={(fillTransparency) => update({ fillTransparency })}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.lineColor',
              )}
            </span>
            <OfficeColorPicker
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.lineColorAria',
                { n },
              )}
              value={style.lineColor ?? '#4F6BED'}
              onValueChange={(lineColor) => update({ lineColor })}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.lineWidth',
              )}
            </span>
            <CommittedOfficeNumberField
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.lineWidthAria',
                { n },
              )}
              min={0.25}
              max={20}
              step={0.25}
              value={style.lineWidth ?? 2.25}
              normalizeValue={(value) =>
                normalizeRequiredOfficeNumber(value, {
                  decimalPlaces: 2,
                  minimum: 0.25,
                  maximum: 20,
                })
              }
              onValueCommit={(lineWidth) => update({ lineWidth })}
            />
          </div>
          <div className="work-office-field">
            <span>
              {officeMessage(messages, 'spreadsheet.chart.seriesStyle.lineDash')}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.seriesStyle.lineDashAria',
                { n },
              )}
              value={style.lineDash ?? 'solid'}
              options={[
                {
                  value: 'solid',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.lineDash.solid',
                  ),
                },
                {
                  value: 'dash',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.lineDash.dash',
                  ),
                },
                {
                  value: 'dot',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.lineDash.dot',
                  ),
                },
                {
                  value: 'dashDot',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.lineDash.dashDot',
                  ),
                },
              ]}
              onValueChange={(lineDash) =>
                update({ lineDash: lineDash as WorkSpreadsheetChartLineDash })
              }
            />
          </div>
          {supportsMarkers ? (
            <>
              <div className="work-office-field">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerSymbol',
                  )}
                </span>
                <OfficeSelect
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerSymbolAria',
                    { n },
                  )}
                  value={marker.symbol ?? 'circle'}
                  options={[
                    {
                      value: 'none',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.none',
                      ),
                    },
                    {
                      value: 'circle',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.circle',
                      ),
                    },
                    {
                      value: 'square',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.square',
                      ),
                    },
                    {
                      value: 'diamond',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.diamond',
                      ),
                    },
                    {
                      value: 'triangle',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.triangle',
                      ),
                    },
                    {
                      value: 'plus',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.plus',
                      ),
                    },
                    {
                      value: 'x',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.x',
                      ),
                    },
                    {
                      value: 'star',
                      label: officeMessage(
                        messages,
                        'spreadsheet.chart.seriesStyle.marker.star',
                      ),
                    },
                  ]}
                  onValueChange={(symbol) =>
                    updateMarker({
                      symbol: symbol as WorkSpreadsheetChartMarkerSymbol,
                    })
                  }
                />
              </div>
              <div className="work-office-field">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerSize',
                  )}
                </span>
                <CommittedOfficeNumberField
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerSizeAria',
                    { n },
                  )}
                  min={2}
                  max={72}
                  step={1}
                  value={marker.size ?? 5}
                  normalizeValue={(value) =>
                    normalizeRequiredOfficeNumber(value, {
                      integer: true,
                      minimum: 2,
                      maximum: 72,
                    })
                  }
                  onValueCommit={(size) => updateMarker({ size })}
                />
              </div>
              <div className="work-office-field">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerFill',
                  )}
                </span>
                <OfficeColorPicker
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerFillAria',
                    { n },
                  )}
                  value={marker.fillColor ?? '#FFFFFF'}
                  onValueChange={(fillColor) => updateMarker({ fillColor })}
                />
              </div>
              <div className="work-office-field">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerLine',
                  )}
                </span>
                <OfficeColorPicker
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.seriesStyle.markerLineAria',
                    { n },
                  )}
                  value={marker.lineColor ?? style.lineColor ?? '#4F6BED'}
                  onValueChange={(lineColor) => updateMarker({ lineColor })}
                />
              </div>
            </>
          ) : null}
        </div>
      ) : null}
    </section>
  );
}

function seriesStyleDefaults(
  seriesIndex: number,
  supportsMarkers: boolean,
): WorkSpreadsheetChartSeriesStyle {
  const style = defaultWorkSpreadsheetChartSeriesStyle(seriesIndex);
  return supportsMarkers ? style : { ...style, marker: undefined };
}
