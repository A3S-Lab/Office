import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizeWorkSpreadsheetDataLabelPosition,
  type WorkSpreadsheetChartType,
  type WorkSpreadsheetDataLabelPosition,
  type WorkSpreadsheetDataLabels,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';

interface SpreadsheetDataLabelEditorProps {
  chartType: WorkSpreadsheetChartType;
  seriesNumber: number;
  value?: WorkSpreadsheetDataLabels;
  onChange: (value: WorkSpreadsheetDataLabels | undefined) => void;
}

export function SpreadsheetDataLabelEditor({
  chartType,
  seriesNumber,
  value,
  onChange,
}: SpreadsheetDataLabelEditorProps) {
  const messages = useOfficeMessages();
  const labelPrefix = officeMessage(
    messages,
    'spreadsheet.chart.dataLabel.prefix',
    { n: String(seriesNumber) },
  );
  const change = (update: Partial<WorkSpreadsheetDataLabels>) =>
    onChange({ ...value, ...update });

  return (
    <section className="work-spreadsheet-data-labels" aria-label={labelPrefix}>
      <OfficeCheckbox
        className="check enable-data-labels"
        ariaLabel={officeMessage(messages, 'spreadsheet.chart.dataLabel.showAria', {
          n: String(seriesNumber),
        })}
        checked={Boolean(value)}
        onCheckedChange={(checked) =>
          onChange(checked ? { showValue: true } : undefined)
        }
      >
        {officeMessage(messages, 'spreadsheet.chart.dataLabel.show')}
      </OfficeCheckbox>
      {value ? (
        <div>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.chart.dataLabel.showValueAria',
              { prefix: labelPrefix },
            )}
            checked={value.showValue === true}
            onCheckedChange={(showValue) => change({ showValue })}
          >
            {officeMessage(messages, 'spreadsheet.chart.dataLabel.showValue')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.chart.dataLabel.showCategoryAria',
              { prefix: labelPrefix },
            )}
            checked={value.showCategoryName === true}
            onCheckedChange={(showCategoryName) => change({ showCategoryName })}
          >
            {officeMessage(messages, 'spreadsheet.chart.dataLabel.showCategory')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.chart.dataLabel.showSeriesAria',
              { prefix: labelPrefix },
            )}
            checked={value.showSeriesName === true}
            onCheckedChange={(showSeriesName) => change({ showSeriesName })}
          >
            {officeMessage(messages, 'spreadsheet.chart.dataLabel.showSeries')}
          </OfficeCheckbox>
          {chartType === 'pie' || chartType === 'doughnut' ? (
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.dataLabel.showPercentAria',
                { prefix: labelPrefix },
              )}
              checked={value.showPercentage === true}
              onCheckedChange={(showPercentage) => change({ showPercentage })}
            >
              {officeMessage(
                messages,
                'spreadsheet.chart.dataLabel.showPercent',
              )}
            </OfficeCheckbox>
          ) : null}
          {chartType === 'bubble' ? (
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.dataLabel.showBubbleAria',
                { prefix: labelPrefix },
              )}
              checked={value.showBubbleSize === true}
              onCheckedChange={(showBubbleSize) => change({ showBubbleSize })}
            >
              {officeMessage(
                messages,
                'spreadsheet.chart.dataLabel.showBubble',
              )}
            </OfficeCheckbox>
          ) : null}
          <div className="work-office-field data-label-position">
            <span>
              {officeMessage(messages, 'spreadsheet.chart.dataLabel.position')}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.dataLabel.positionAria',
                { prefix: labelPrefix },
              )}
              value={normalizeWorkSpreadsheetDataLabelPosition(value.position)}
              options={[
                {
                  value: 'bestFit',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.bestFit',
                  ),
                },
                {
                  value: 'center',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.center',
                  ),
                },
                {
                  value: 'insideBase',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.insideBase',
                  ),
                },
                {
                  value: 'insideEnd',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.insideEnd',
                  ),
                },
                {
                  value: 'outsideEnd',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.outsideEnd',
                  ),
                },
                {
                  value: 'left',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.left',
                  ),
                },
                {
                  value: 'right',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.right',
                  ),
                },
                {
                  value: 'above',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.above',
                  ),
                },
                {
                  value: 'below',
                  label: officeMessage(
                    messages,
                    'spreadsheet.chart.dataLabel.position.below',
                  ),
                },
              ]}
              onValueChange={(position) =>
                change({
                  position: position as WorkSpreadsheetDataLabelPosition,
                })
              }
            />
          </div>
          <div className="work-office-field data-label-separator">
            <span>
              {officeMessage(messages, 'spreadsheet.chart.dataLabel.separator')}
            </span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'spreadsheet.chart.dataLabel.separatorAria',
                { prefix: labelPrefix },
              )}
              value={value.separator ?? ', '}
              maxLength={64}
              onChange={(event) => change({ separator: event.target.value })}
            />
          </div>
        </div>
      ) : null}
    </section>
  );
}
